from __future__ import annotations

import csv
import io
import os
import re
import unicodedata
import zipfile
from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Dict, Iterable, Iterator, List, Optional, Tuple

import pandas as pd
from pypdf import PdfReader

MONTHS = {
    'janeiro': '01',
    'fevereiro': '02',
    'marco': '03',
    'abril': '04',
    'maio': '05',
    'junho': '06',
    'julho': '07',
    'agosto': '08',
    'setembro': '09',
    'outubro': '10',
    'novembro': '11',
    'dezembro': '12',
}

COMPANY_PATTERNS = [
    ('ATIVA TERCEIRIZACAO', 'ATIVA'),
    ('ATIVA TERCEIRIZAÇÃO', 'ATIVA'),
    ('LIDER LIMPE LIMPEZA COMERCIAL', 'LIDER COMERCIAL'),
    ('LIDER LIMPEZA COMERCIAL', 'LIDER COMERCIAL'),
    ('LIDER MULTISSERV', 'LIDER MULTISSERVICOS'),
    ('LIDER MULTISERV', 'LIDER MULTISSERVICOS'),
    ('VSP VIGILANCIA', 'VSP'),
    ('VSP VIGILANCIA E SEGURANCA', 'VSP'),
]

SUBGROUP_PATTERNS = [
    ('PMV', 'PMV'),
    ('PREFEITURA MUNICIPAL DE VITORIA', 'PMV'),
    ('SEDU', 'SEDU'),
]


@dataclass
class Record:
    original_file: str
    new_name: str
    cpf: str
    code: str
    company: str
    subgroup: str
    period: str
    status: str
    included: bool
    reason: str
    source_company: str
    source_period: str
    source_cpf: str


def normalize_text(value: object) -> str:
    if value is None:
        return ''
    text = str(value)
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(ch for ch in text if not unicodedata.combining(ch))
    return text.strip()


NORMALIZED_COMPANY_PATTERNS = [(normalize_text(pattern).upper(), company) for pattern, company in COMPANY_PATTERNS]
NORMALIZED_SUBGROUP_PATTERNS = [(normalize_text(pattern).upper(), subgroup) for pattern, subgroup in SUBGROUP_PATTERNS]


def clean_digits(value: object) -> str:
    return re.sub(r'\D+', '', '' if value is None else str(value))


def format_cpf(cpf: object) -> str:
    digits = clean_digits(cpf).zfill(11)[-11:]
    return f'{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}'


def normalize_column(name: object) -> str:
    text = normalize_text(name).lower()
    text = re.sub(r'[^a-z0-9]+', '_', text).strip('_')
    return text


def detect_company(text: str) -> str:
    upper = normalize_text(text).upper()
    for pattern, company in NORMALIZED_COMPANY_PATTERNS:
        if pattern in upper:
            return company
    return 'NAO_IDENTIFICADA'


def detect_ativa_subgroup(text: str) -> str:
    upper = normalize_text(text).upper()
    for pattern, subgroup in NORMALIZED_SUBGROUP_PATTERNS:
        if pattern in upper:
            return subgroup
    return 'GERAL'


def detect_period_from_text(text: str) -> Optional[str]:
    text_norm = normalize_text(text).lower()
    month_pattern = '|'.join(MONTHS.keys())
    match = re.search(rf'({month_pattern})\s+de\s+(20\d{{2}})', text_norm)
    if match:
        month_name, year = match.groups()
        return f'{MONTHS[month_name]}-{year}'
    return None


def detect_period_from_filename(filename: str) -> Optional[str]:
    name = Path(filename).name
    match = re.search(r'(?<!\d)(0?[1-9]|1[0-2])[-_](20\d{2})(?!\d)', name)
    if match:
        return f"{match.group(1).zfill(2)}-{match.group(2)}"
    old_match = re.search(r'\b\d+-([0-1]?\d)-(20\d{2})-', name)
    if old_match:
        return f"{old_match.group(1).zfill(2)}-{old_match.group(2)}"
    return None


def detect_code_from_filename(filename: str) -> str:
    name = Path(filename).name
    patterns = [
        r'-M-(\d+)-',
        r'-m-(\d+)-',
        r'-(\d+)-Recibo',
        r'\b(\d{1,10})\b(?=.*Recibo)',
    ]
    for pattern in patterns:
        match = re.search(pattern, name, flags=re.IGNORECASE)
        if match:
            return clean_digits(match.group(1))
    return ''


def detect_code_from_text(text: str) -> str:
    cleaned = normalize_text(text)
    patterns = [
        r'(\d{1,10})\s*Codigo\s*Nome do Funcionario',
        r'(\d{1,10})\s*Codigo\b',
        r'Codigo\s*(\d{1,10})\s*Nome do Funcionario',
        r'Codigo\s*:?\s*(\d{1,10})\b',
    ]
    for pattern in patterns:
        match = re.search(pattern, cleaned, flags=re.IGNORECASE)
        if match:
            return clean_digits(match.group(1))
    return ''


def detect_cpf_from_text(text: str) -> str:
    matches = re.findall(r'(\d{3}\.\d{3}\.\d{3}-\d{2})', text)
    if matches:
        return format_cpf(matches[0])
    normalized = normalize_text(text)
    match = re.search(r'CPF\s*:?\s*(\d{11})', normalized, flags=re.IGNORECASE)
    if match:
        return format_cpf(match.group(1))
    generic = re.search(r'\b(\d{11})\b', normalized)
    if generic:
        return format_cpf(generic.group(1))
    return ''


def read_excel_flexible(file_obj_or_path) -> pd.DataFrame:
    name = str(getattr(file_obj_or_path, 'name', file_obj_or_path)).lower()
    if name.endswith('.xls'):
        return pd.read_excel(file_obj_or_path, engine='xlrd')
    return pd.read_excel(file_obj_or_path, engine='openpyxl')


def load_employee_map(file_obj_or_path) -> Dict[str, dict]:
    df = read_excel_flexible(file_obj_or_path)
    col_map = {normalize_column(c): c for c in df.columns}

    aliases = {
        'code': ['codigo', 'codigo_empregado', 'i_empregados', 'cod', 'codigo_colaborador'],
        'cpf': ['cpf'],
        'situacao': ['situacao', 'status', 'sit'],
        'company': ['cp_nome_emp', 'empresa', 'nome_empresa', 'razao_social'],
        'worksite': ['nome_quebra', 'posto', 'centro_custo', 'cc', 'lotacao'],
    }

    resolved = {}
    for target, options in aliases.items():
        for opt in options:
            if opt in col_map:
                resolved[target] = col_map[opt]
                break

    if 'code' not in resolved or 'cpf' not in resolved:
        raise ValueError('A planilha precisa ter pelo menos código e CPF (ex.: i_empregados/código e cpf).')

    employee_map: Dict[str, dict] = {}
    for _, row in df.iterrows():
        code = clean_digits(row.get(resolved['code']))
        cpf_digits = clean_digits(row.get(resolved['cpf']))
        if not code or not cpf_digits:
            continue

        situ_raw = row.get(resolved.get('situacao', ''), '') if resolved.get('situacao') else ''
        situ_digits = clean_digits(situ_raw)
        if situ_digits == '8':
            status = 'Demitido'
        elif situ_digits == '1':
            status = 'Trabalhando'
        elif situ_digits:
            status = 'Outros'
        else:
            status = 'Nao informado'

        company_raw = str(row.get(resolved.get('company', ''), '') if resolved.get('company') else '')
        worksite_raw = str(row.get(resolved.get('worksite', ''), '') if resolved.get('worksite') else '')
        company = detect_company(company_raw) if company_raw else 'NAO_IDENTIFICADA'
        subgroup = detect_ativa_subgroup(worksite_raw) if company == 'ATIVA' else ''

        employee_map[code] = {
            'cpf': format_cpf(cpf_digits),
            'status': status,
            'company': company,
            'company_raw': company_raw,
            'worksite_raw': worksite_raw,
            'subgroup': subgroup,
        }
    return employee_map


def extract_text_from_pdf_bytes(data: bytes, max_pages: int = 2) -> str:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    total = min(max_pages, len(reader.pages))
    for i in range(total):
        pages.append(reader.pages[i].extract_text() or '')
    return '\n'.join(pages)


def iter_pdf_uploads(uploaded_files: Iterable) -> Iterator[Tuple[str, bytes]]:
    for uploaded in uploaded_files or []:
        name = Path(uploaded.name).name
        raw = uploaded.read()
        if name.lower().endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(raw)) as zf:
                for member in zf.infolist():
                    if member.is_dir() or not member.filename.lower().endswith('.pdf'):
                        continue
                    yield Path(member.filename).name, zf.read(member)
        elif name.lower().endswith('.pdf'):
            yield name, raw


def build_record(filename: str, data: bytes, employee_map: Optional[Dict[str, dict]] = None) -> Record:
    employee_map = employee_map or {}
    text = extract_text_from_pdf_bytes(data)

    company_from_pdf = detect_company(text)
    subgroup_from_pdf = detect_ativa_subgroup(text) if company_from_pdf == 'ATIVA' else ''

    period_pdf = detect_period_from_text(text)
    period_file = detect_period_from_filename(filename)
    period = period_file or period_pdf or 'SEM_PERIODO'

    code = detect_code_from_text(text) or detect_code_from_filename(filename)
    cpf_pdf = detect_cpf_from_text(text)

    employee = employee_map.get(code, {}) if code else {}
    status = employee.get('status', 'Nao informado')
    company = employee.get('company') or company_from_pdf
    subgroup = employee.get('subgroup') or subgroup_from_pdf
    cpf = employee.get('cpf') or cpf_pdf

    included = True
    reason = 'OK'

    if status == 'Demitido':
        included = False
        reason = 'Ignorado por situação 8 (Demitido)'
    elif not cpf:
        included = False
        reason = 'CPF não encontrado no PDF nem na planilha'
    elif period == 'SEM_PERIODO':
        included = False
        reason = 'Período não identificado'
    elif company == 'NAO_IDENTIFICADA':
        reason = 'Empresa não identificada com segurança'

    if company == 'ATIVA' and not subgroup:
        subgroup = 'GERAL'

    new_name = f'{period}-{cpf}.pdf' if cpf and period != 'SEM_PERIODO' else filename

    source_period = 'arquivo' if period_file else ('pdf' if period_pdf else 'nao_encontrado')
    source_cpf = 'planilha' if employee.get('cpf') else ('pdf' if cpf_pdf else 'nao_encontrado')

    return Record(
        original_file=filename,
        new_name=new_name,
        cpf=cpf,
        code=code,
        company=company,
        subgroup=subgroup,
        period=period,
        status=status,
        included=included,
        reason=reason,
        source_company=company_from_pdf,
        source_period=source_period,
        source_cpf=source_cpf,
    )


def unique_period_label(records: List[Record]) -> str:
    periods = sorted({r.period for r in records if r.period != 'SEM_PERIODO'})
    if not periods:
        return 'SEM-PERIODO'
    if len(periods) == 1:
        return periods[0]
    return 'VARIOS-PERIODOS'


def _write_report_files(zf: zipfile.ZipFile, records: List[Record], errors: List[str]) -> None:
    fieldnames = list(asdict(records[0]).keys()) if records else [
        'original_file', 'new_name', 'cpf', 'code', 'company', 'subgroup', 'period', 'status', 'included', 'reason', 'source_company', 'source_period', 'source_cpf'
    ]
    csv_buffer = io.StringIO()
    writer = csv.DictWriter(csv_buffer, fieldnames=fieldnames)
    writer.writeheader()
    for record in records:
        writer.writerow(asdict(record))
    zf.writestr('relatorio_processamento.csv', csv_buffer.getvalue().encode('utf-8-sig'))

    summary = Counter()
    for record in records:
        if record.included:
            summary['incluidos'] += 1
            summary[f'empresa_{record.company}'] += 1
            if record.company == 'ATIVA':
                summary[f'ativa_{record.subgroup or "GERAL"}'] += 1
        else:
            summary['ignorados'] += 1
    lines = ['Resumo do processamento']
    for key, value in summary.items():
        lines.append(f'- {key}: {value}')
    if errors:
        lines.append('')
        lines.append('Ocorrências:')
        lines.extend(f'- {e}' for e in errors)
    zf.writestr('log_processamento.txt', '\n'.join(lines).encode('utf-8'))


def create_output_zips(
    pdf_items: Iterable[Tuple[str, bytes]],
    employee_map: Optional[Dict[str, dict]] = None,
    make_separate_zip: bool = True,
    make_general_zip: bool = True,
) -> Tuple[Optional[bytes], Optional[bytes], List[Record]]:
    employee_map = employee_map or {}
    records: List[Record] = []
    errors: List[str] = []

    separate_buffer = io.BytesIO() if make_separate_zip else None
    combined_buffer = io.BytesIO() if make_general_zip else None
    used_separate = set()
    used_combined = set()

    zsep = zipfile.ZipFile(separate_buffer, 'w', compression=zipfile.ZIP_DEFLATED) if separate_buffer else None
    zall = zipfile.ZipFile(combined_buffer, 'w', compression=zipfile.ZIP_DEFLATED) if combined_buffer else None

    try:
        for filename, data in pdf_items:
            try:
                record = build_record(filename, data, employee_map)
                records.append(record)

                if not record.included:
                    continue

                if zall is not None:
                    combined_name = record.new_name
                    base, ext = os.path.splitext(combined_name)
                    counter = 2
                    while combined_name in used_combined:
                        combined_name = f'{base}__{counter}{ext}'
                        counter += 1
                    used_combined.add(combined_name)
                    zall.writestr(combined_name, data)

                if zsep is not None:
                    folder_parts = [record.company.replace('/', '-').strip() or 'NAO_IDENTIFICADA']
                    if record.company == 'ATIVA':
                        folder_parts.append(record.subgroup or 'GERAL')
                    sep_path = '/'.join(folder_parts + [record.new_name])
                    base_path, ext = os.path.splitext(sep_path)
                    counter = 2
                    while sep_path in used_separate:
                        sep_path = f'{base_path}__{counter}{ext}'
                        counter += 1
                    used_separate.add(sep_path)
                    zsep.writestr(sep_path, data)
            except Exception as exc:
                record = Record(
                    original_file=filename,
                    new_name=filename,
                    cpf='',
                    code='',
                    company='NAO_IDENTIFICADA',
                    subgroup='',
                    period='SEM_PERIODO',
                    status='Erro',
                    included=False,
                    reason=f'Falha ao ler PDF: {exc}',
                    source_company='NAO_IDENTIFICADA',
                    source_period='erro',
                    source_cpf='erro',
                )
                records.append(record)
                errors.append(f'{filename}: {exc}')

        if zsep is not None:
            _write_report_files(zsep, records, errors)
        if zall is not None:
            _write_report_files(zall, records, errors)
    finally:
        if zsep is not None:
            zsep.close()
        if zall is not None:
            zall.close()

    return (
        separate_buffer.getvalue() if separate_buffer else None,
        combined_buffer.getvalue() if combined_buffer else None,
        records,
    )


def summarize_records(records: List[Record]) -> dict:
    included = [r for r in records if r.included]
    ignored = [r for r in records if not r.included]
    company_counter = Counter(r.company for r in included)
    subgroup_counter = Counter((r.subgroup or 'GERAL') for r in included if r.company == 'ATIVA')
    source_cpf_counter = Counter(r.source_cpf for r in records)
    return {
        'total': len(records),
        'incluidos': len(included),
        'ignorados': len(ignored),
        'empresas': dict(company_counter),
        'ativa_subgrupos': dict(subgroup_counter),
        'periodo': unique_period_label(records),
        'cpf_por_origem': dict(source_cpf_counter),
    }
