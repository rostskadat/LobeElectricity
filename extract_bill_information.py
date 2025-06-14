#!/usr/bin/env python
"""
Extract bill information from PDF files and generate a report.
"""

try:
    import coloredlogs
    import locale
    import logging
    import os
    import pdfplumber
    import re
    import sys
    import yaml
    from argparse import ArgumentParser, RawTextHelpFormatter
    from datetime import datetime
    from dotenv import load_dotenv
    from openpyxl import Workbook
    from openpyxl.workbook.defined_name import DefinedName
    from os.path import basename
    from pprint import pprint
except ModuleNotFoundError as e:
    print(f"{e}. Did you load your environment?")
    sys.exit(1)

load_dotenv()
logger = logging.getLogger()

# i.e.: +/-0.000,00 €
AMOUNT_PATTERN = re.compile(r"[+-]?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2} ?€")
# i.e.: P4 1.18.4 1.278,00 2.578,00 1,00 0,00 1.300,00
ENDESA_PX_PATTERN = re.compile(r"^(P[123456]) 1\.18\.[123456]( (?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}){5}.*")
ENDESA_PV_PATTERN = re.compile(r".*(Punta|Llano|Valle)( (?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}){5}.*")
# i.e.: Px 7.994 7.994 0 66 66 0 0,00 0,00 0,00
TOTAL_PV_PATTERN = re.compile(r"^(P[123456])( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){3}( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){6}")
# i.e.: Px 593 665 72 0 0 3,81
NUFRI_PX_PATTERN = re.compile(r"^(P[123456])( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){3}( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){3}")

def main():
    with open(os.path.join(os.path.dirname(__file__), 'defaults.yaml'), 'r', encoding='utf-8') as f:
        defaults = yaml.safe_load(f)
    args = parse_command_line(defaults)
    args.defaults = defaults
    try:
        # locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # On Windows
        locale.setlocale(locale.LC_NUMERIC, args.input_locale)
        locale.setlocale(locale.LC_TIME, args.input_locale)
        for verbose_package in ['pdfminer.pdfpage']:
            verbose_logger = logging.getLogger(verbose_package)
            verbose_logger.setLevel(level=logging.ERROR)
        if args.debug:
            coloredlogs.install(level=logging.DEBUG, logger=logger)
        elif args.quiet:
            coloredlogs.install(level=logging.ERROR, logger=logger)
        else:
            coloredlogs.install(level=logging.INFO, logger=logger)
        return args.func(args)
    except Exception as e:
        logging.error(e, exc_info=True)
        return 1


def parse_command_line(defaults):
    parser = ArgumentParser(prog='extract_bill_information',
                            description=__doc__, formatter_class=RawTextHelpFormatter)
    parser.add_argument(
        '--debug', action="store_true", help='Run the program in debug', required=False, default=False)
    parser.add_argument(
        '--quiet', action="store_true", help='Run the program in quiet mode', required=False, default=False)
    parser.add_argument(
        '--input-dir', help=f"The path where the bills are located. By default '{defaults['input-dir']}'.", required=False, default=defaults['input-dir'])
    parser.add_argument(
        '--input-locale', help=f"The locale to use when reading the bills", required=False, default=defaults['input-locale'])
    parser.add_argument(
        '--output', help=f"The name of the report.", required=False, default=defaults['output'])
    parser.add_argument(
        '--limit', type=int, help=f"Limit the number of ticket to download. By default '{defaults['limit']}' (-1 is no limit).", required=False, default=defaults['limit'])
    parser.set_defaults(func=extract_bill_information)
    return parser.parse_args()


def extract_bill_information(args):
    """
    Extracts and processes bill information from PDF files in a specified input directory.

    This function scans the given input directory for PDF files, extracts bill information using
    a dispatcher, sanitizes the extracted data, and organizes it by CUPS and bill ID. It also
    supports limiting the number of files processed and generates a final bill report.

    Args:
        args: An object containing the following attributes:
            - input_dir (str): Path to the directory containing PDF files.
            - defaults (dict): Default configuration, including 'dispatchers' for extraction.
            - limit (int): Maximum number of files to process. If 0 or less, all files are processed.

    Returns:
        int: 1 if an error occurs (e.g., input directory does not exist or no files found), otherwise None.
    """
    if not os.path.isdir(args.input_dir):
        logger.error(f"Input directory '{args.input_dir}' does not exist.")
        return 1
    logger.info(f"Reading files from '{args.input_dir}' ...")
    files = [os.path.join(dp, f) for dp, dn, filenames in os.walk(args.input_dir) for f in filenames if f.lower().endswith('.pdf')]
    if not files:
        logger.error(f"No files found in '{args.input_dir}'.")
        return 1
    logger.debug(f"Found {len(files)} files in '{args.input_dir}'.")
    if args.limit > 0:
        files = files[:args.limit]
        logger.warning(f"Limiting to {len(files)} files as per the --limit argument.")

    bills = {}
    for file in files:
        bill_info = extract_dispatcher(args.defaults['dispatchers'], file)
        if not bill_info:
            continue
        bill_info = sanitize_bill(bill_info)
        if not bill_info:
            continue
        bill_info['file'] = basename(file)
        cups = bill_info['cups']
        bill_id = bill_info['bill_id']
        if cups not in bills:
            bills[cups] = {}
        bills[cups][bill_id] = bill_info

    generate_bill_report(args, bills)


def extract_dispatcher(dispatchers:dict, file:str):
    """
    Extracts bill information by selecting and invoking the
    appropriate extractor function based on the content of
    the first page of a PDF bill.

    Args:
        dispatchers (dict): A dictionary mapping dispatcher names (str) to extractor function names (str).
        file (str): The path to the PDF file to be processed.

    Returns:
        Any: The result of the extractor function if a matching dispatcher is found; otherwise, None.

    Raises:
        ValueError: If the specified extractor function is not found or is not callable.

    """
    logger.info(f"Extracting information from bill '{file}' ...")
    with pdfplumber.open(file) as pdf:
        first_page = pdf.pages[0].extract_text()
        found_extractor = False
        for dispatcher, extractor in dispatchers.items():
            if extractor not in globals() or not callable(globals()[extractor]):
                raise ValueError(f"Extractor '{extractor}' not found or not callable.")
            if dispatcher in first_page:
                logger.debug(f"Detected '{dispatcher}' bill. Using {extractor} ...")
                return globals()[extractor](file, pdf)
        if not found_extractor:
            logger.error(f"Could not detect the type of bill for '{file}'. Skipping")
            return None


def extract_nufri_bill(file, pdf):
    bill_info = _get_default_bill_info()
    first_page = pdf.pages[0]
    for line in first_page.extract_text().split('\n'):
        if 'CAMPANILLA' in line:
            bill_info['is_ours'] = True
            continue
        if 'Nº de Factura:' in line:
            bill_info['bill_id'] = ' '.join(line.split(':')[-1].strip().split(' ')[0:2])
            continue
        if 'Fecha de Factura:' in line:
            date_obj = datetime.strptime(line.split(':')[-1].strip(), "%d/%m/%Y")
            bill_info['billing_date'] = date_obj.strftime("%d/%m/%Y")
            continue
        if 'Periodo de Consumo:' in line:
            bill_info['billing_period'] = line.split(':')[-1].strip()
        if 'Por Potencia Contratada' in line:
            bill_info['billed_power_capacity'] = re.match(r".*Por Potencia Contratada ([\,0-9]+ €).*", line).group(1)
        if 'Por Energía Consumida' in line:
            bill_info['billed_energy_consumed'] = re.match(r".*Por Energía Consumida ([\,0-9]+ €).*", line).group(1)
        if 'Total Factura' in line:
            amount = re.match(r".*Total Factura ([\,0-9]+ €).*", line).group(1)
            bill_info['billed_amount_0'] = amount
            bill_info['billed_amount_1'] = amount
            continue
        if line.startswith('CUPS:'):
            bill_info['cups'] = line.split(':')[-1].strip()
            continue
        if line.startswith('P'):
            power_type, power_amount = _extract_nufri_power(line)
            if power_type is not None:
                bill_info[power_type] = power_amount
                continue

    if not bill_info['is_ours']:
        logger.warning(f"File '{file}' does not belong to us. Skipping.")
        return None
    if 'P1' not in bill_info:
        logger.warning(f"Power reading not on 2nd nor on the 3rd page for '{file}'. Skipping.")
        return None
    return bill_info


def _extract_nufri_power(line:str) -> str:
    power_type = None
    power_amount = None
    match = NUFRI_PX_PATTERN.match(line)
    if match:
        power_type = match.group(1)
        power_amount = match.group(2)
    return power_type, power_amount


def extract_total_bill(file, pdf):
    bill_info = _get_default_bill_info()
    first_page = pdf.pages[0]
    for line in first_page.extract_text().split('\n'):
        if 'CAMPANILLA' in line:
            bill_info['is_ours'] = True
            continue
        if 'Nº Factura' in line:
            bill_info['bill_id'] = line.split(' ')[-1].strip()
            continue
        if 'Fecha emisión factura:' in line:
            date_obj = datetime.strptime(line.split(':')[-1].strip(), "%d de %B de %Y")
            bill_info['billing_date'] = date_obj.strftime("%d/%m/%Y")
            continue
        if 'Periodo de facturación:' in line:
            bill_info['billing_period'] = line.split(':')[-1].strip()
            continue
        if line.startswith('Potencia '):
            bill_info['billed_power_capacity'] = line
            continue
        if line.startswith('Energía '):
            bill_info['billed_energy_consumed'] = line
            continue
        if line.startswith('TOTAL '):
            bill_info['billed_amount_0'] = line
            bill_info['billed_amount_1'] = line
            continue
        if '(CUPS):' in line :
            bill_info['cups'] = line.split(':')[-1].strip()
            continue

    second_page = pdf.pages[1]
    for line in second_page.extract_text().split('\n'):
        power_type, power_amount = _extract_total_power(line)
        if power_type is not None:
            bill_info[power_type] = power_amount
            continue

    if not bill_info['is_ours']:
        logger.warning(f"File '{file}' does not belong to us. Skipping.")
        return None
    if 'P1' not in bill_info:
        logger.warning(f"Power reading not on 2nd nor on the 3rd page for '{file}'. Skipping.")
        return None
    return bill_info


def _extract_total_power(line:str) -> str:
    power_type = None
    power_amount = None
    match = TOTAL_PV_PATTERN.match(line)
    if match:
        power_type = match.group(1)
        power_amount = match.group(2)
    return power_type, power_amount


def extract_endesa_bill(file, pdf):
    bill_info = _get_default_bill_info()
    first_page = pdf.pages[0]
    for line in first_page.extract_text().split('\n'):
        if 'CAMPANILLA' in line:
            bill_info['is_ours'] = True
            continue
        if 'Nº factura:' in line or 'Nº de factura:' in line or 'Nºfactura:' in line:
            bill_info['bill_id'] = line.split(':')[-1].strip()
            continue
        if 'Fecha emisión factura:' in line or 'Fechaemisiónfactura:' in line:
            bill_info['billing_date'] = line.split(':')[-1].strip()
            continue
        if 'Periodo de facturación:' in line or 'Periododefacturación' in line:
            bill_info['billing_period'] = line.split(':')[-1].strip()
            continue
        if line.startswith('Potencia '):
            bill_info['billed_power_capacity'] = line
            continue
        if line.startswith('Energía '):
            bill_info['billed_energy_consumed'] = line
            continue
        if line.startswith('Total '):
            bill_info['billed_amount_0'] = line
            continue

    second_page = pdf.pages[1]
    for line in second_page.extract_text().split('\n'):
        if 'CUPS' in line:
            bill_info['cups'] = line.split(':')[-1].strip()
            bill_info['cups'] = bill_info['cups'].split('(')[0].strip()  # Remove any additional info in parentheses
            continue
        if 'TOTAL ' in line:
            bill_info['billed_amount_1'] = line
            # continue # Sometimes the TOTAL and the power readings are on the same line
        power_type, power_amount = _extract_endesa_power(line)
        if power_type is not None:
            bill_info[power_type] = power_amount
            continue

    if 'P1' not in bill_info:
        if len(pdf.pages) > 2:
            third_page = pdf.pages[2]
            for line in third_page.extract_text().split('\n'):
                power_type, power_amount = _extract_endesa_power(line)
                if power_type is not None:
                    bill_info[power_type] = power_amount

    if not bill_info['is_ours']:
        logger.warning(f"File '{file}' does not belong to us. Skipping.")
        return None
    if 'P1' not in bill_info:
        logger.warning(f"Power reading not on 2nd nor on the 3rd page for '{file}'. Skipping.")
        return None
    return bill_info


def _extract_endesa_power(line:str) -> str:
    power_type = None
    power_amount = None
    match = ENDESA_PX_PATTERN.match(line)
    if match:
        power_type = match.group(1)
        power_amount = match.group(2)
    else:
        match = ENDESA_PV_PATTERN.match(line)
        if match:
            power_type = match.group(1)
            power_amount = match.group(2)
    if power_type == 'Punta':
        power_type = 'P1'
    elif power_type == 'Llano':
        power_type = 'P2'
    elif power_type == 'Valle':
        power_type = 'P3'
    return power_type, power_amount


def _extract_billed_amount(data:dict, key:str) -> str:
    billed_amount = AMOUNT_PATTERN.findall( data[key])
    if len(billed_amount) != 1:
        return None
    return locale.atof(billed_amount[0].replace('€', '').strip())


def _get_default_bill_info():
    return {
        'is_ours': False,
        'bill_id': None,
        'billing_date': None,
        'billing_period': None,
        'billed_power_capacity': None,
        'billed_energy_consumed': None,
        'billed_amount_0': None,
        'billed_amount_1': None,
        'is_rectification': False,
        # 'P1': None,
        # 'P2': None,
        # 'P3': None,
        # 'P4': None,
        # 'P5': None,
        # 'P6': None,
        'cups': None,
    }

def sanitize_bill(data:dict):
    """
    Sanitize the bill data by converting strings to appropriate types and formatting.
    """
    is_sane, missing_keys = is_bill_sane(data)
    if not is_sane:
        logger.error(f"Missing keys for CUPS {data['cups']} and bill {data['bill_id']}: {missing_keys}")
        return None

    try:
        data['billing_date'] = datetime.strptime(data['billing_date'], "%d/%m/%Y").date()
    except ValueError as e:
        logger.error(f"Invalid billing date '{data['billing_date']}' for CUPS {data['cups']} and bill {data['bill_id']}: {e}")
        return None

    billing_period = re.findall(r"\d{2}/\d{2}/\d{4}", data['billing_period'])
    if len(billing_period) != 2:
        logger.error(f"Billing period '{data['billing_period']}' is not in the expected format 'dd/mm/yyyy - dd/mm/yyyy' for CUPS {data['cups']} and bill {data['bill_id']}.")
        return None
    try:
        data['billing_period_start'] = datetime.strptime(billing_period[0], "%d/%m/%Y").date()
        data['billing_period_end'] = datetime.strptime(billing_period[1], "%d/%m/%Y").date()
    except ValueError as e:
        logger.error(f"Invalid billing period '{data['billing_period']}' for CUPS {data['cups']} and bill {data['bill_id']}: {e}")
        return None

    billed_power_capacity = _extract_billed_amount(data, 'billed_power_capacity')
    if billed_power_capacity is None:
        logger.error(f"Billed power capacity '{data['billed_power_capacity']}' is not in the expected format '+/-0.000,00 €' for CUPS {data['cups']} and bill {data['bill_id']}.")
        return None
    data['billed_power_capacity'] = billed_power_capacity

    billed_energy_consumed = _extract_billed_amount(data, 'billed_energy_consumed')
    if billed_energy_consumed is None:
        logger.error(f"Billed enerygy consumed '{data['billed_energy_consumed']}' is not in the expected format '+/-0.000,00 €' for CUPS {data['cups']} and bill {data['bill_id']}.")
        return None
    data['billed_energy_consumed'] = billed_energy_consumed

    billed_amount_0 = _extract_billed_amount(data, 'billed_amount_0')
    if billed_amount_0 is None:
        logger.error(f"Billed amount 0 '{data['billed_amount_0']}' is not in the expected format '+/-0.000,00 €' for CUPS {data['cups']} and bill {data['bill_id']}")
        return None
    data['billed_amount_0'] = billed_amount_0

    billed_amount_1 = _extract_billed_amount(data, 'billed_amount_1')
    if billed_amount_1 is None:
        logger.error(f"Billed amount 1 '{data['billed_amount_1']}' is not in the expected format '+/-0.000,00 €' for CUPS {data['cups']} and bill {data['bill_id']}.")
        return None
    data['billed_amount_1'] = billed_amount_1

    if data['billed_amount_1'] != data['billed_amount_0']:
        logger.debug(f"Billed amount 1 '{data['billed_amount_1']}' is different from billed amount 0 '{data['billed_amount_0']}' for CUPS {data['cups']} and bill {data['bill_id']}.")
        data['is_rectification'] = True

    for power_type in ['P1', 'P2', 'P3', 'P4', 'P5', 'P6']:
        if power_type not in data and power_type in ['P1', 'P2', 'P3']:
            logger.error(f"Mandatory '{power_type}' comsumption has not been extracted from CUPS {data['cups']} and bill {data['bill_id']}.")
            return None
        if power_type in data and data[power_type] is not None: # P4, P5, P6
            data[power_type] = locale.atof(data[power_type].strip())

    return data


def is_bill_sane(data:dict):
    # check that all the keys have a value
    is_sane = True
    missing_keys = []
    for key, value in data.items():
        if value is None or (isinstance(value, str) and not value.strip()):
            missing_keys.append(key)
            is_sane = False
    return is_sane, missing_keys


def generate_bill_report(args, bills:dict):
    """
    Generates an Excel workbook report from provided bill information.

    For each CUPS (supply point) in the `bills` dictionary, a worksheet is created and populated
    with bill data. The worksheet names and column headers are determined by the `args.defaults`
    configuration. The resulting workbook is saved to the file path specified by `args.output`.

    Args:
        args: An object containing configuration options, including:
            - defaults['column_labels']: A dictionary mapping column keys to their display labels.
            - defaults['sheet_names']: A dictionary mapping CUPS to worksheet names.
            - output: The file path where the Excel workbook will be saved.
        bills (dict): A dictionary where each key is a CUPS identifier and each value is another
            dictionary containing bill information for that CUPS.

    Returns:
        None. The function saves the generated Excel workbook to the specified output file.
    """
    wb = Workbook()

    column_keys = list(args.defaults['column_labels'].keys())
    column_headers = list(args.defaults['column_labels'].values())
    sheet_names = args.defaults['sheet_names']

    ws = wb.active
    ws.title = "Simulación"
    ws.append(["-", "P1", "P2", "P3", "P4", "P5", "P6"])
    for k,v in args.defaults['tariffs'].items():
        ws.append([k, *v])

    sheets = {}
    # Add a new worksheet
    for cups, bill_infos in bills.items():
        ws = wb.create_sheet(title=sheet_names.get(cups, cups))
        sheets[ws.title] = ws

        # Add data to each worksheet
        is_first_row = True
        for bill_info in bill_infos.values():
            if is_first_row:
                ws.append(column_headers)
                is_first_row = False
            ws.append([ bill_info.get(h, None) for h in column_keys ])

    ordered_sheets = []
    for title in args.defaults['sheet_names'].values():
        if title in sheets:
            ordered_sheets.append(sheets[title])
    if len(ordered_sheets) == 0:
        ordered_sheets.extend(list(sheets.values()))
    wb._sheets = [ wb.worksheets[0] ] + ordered_sheets

    # Save the workbook
    wb.save(args.output)

if __name__ == '__main__':
    sys.exit(main())


