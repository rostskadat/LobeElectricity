#!/usr/bin/env python
"""
Extract bill information from PDF files and generate a report.
"""

try:
    import coloredlogs
    import csv
    import itertools
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
    from pprint import pprint
except ModuleNotFoundError as e:
    print(f"{e}. Did you load your environment?")
    sys.exit(1)

load_dotenv()
logger = logging.getLogger()

AMOUNT_PATTERN = re.compile(r"[+-]?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2} ?€")
# i.e.: P4 1.18.4 1.278,00 2.578,00 1,00 0,00 1.300,00
PX_PATTERN = re.compile(r"^(P[123456]) 1\.18\.[123456]( (?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}){5}.*")
PV_PATTERN = re.compile(r".*(Punta|Llano|Valle)( (?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}){5}.*")

# def generate_csv_report(args, tickets_by_room:map):
#     todays_date = datetime.today().strftime(DEFAULT_FMT)
#     document_name = f"tickets-{todays_date}.csv"
#     if not os.path.isdir(args.output_dir):
#         os.makedirs(args.output_dir)
#     report_filename = os.path.join(args.output_dir, document_name)
#     logger.info(f"Saving report {report_filename} ...")
#     with open(report_filename, 'w', newline='') as csvfile:
#         fieldnames = ['room', 'ticket_number', 'state', 'accepted', 'creation_time', 'description', 'answer']
#         writer = csv.DictWriter(csvfile, fieldnames=fieldnames, extrasaction='ignore')
#         writer.writeheader()
#         rows = list(itertools.chain(*tickets_by_room.values()))
#         writer.writerows(rows)

def is_bill_sane(data:dict):
    # check that all the keys have a value
    is_sane = True
    missing_keys = []
    for key, value in data.items():
        if value is None or (isinstance(value, str) and not value.strip()):
            missing_keys.append(key)
            is_sane = False
    return is_sane, missing_keys

def _extract_billed_amount(data:dict, key:str) -> str:
    billed_amount = AMOUNT_PATTERN.findall( data[key])
    if len(billed_amount) != 1:
        return None
    return billed_amount[0]
    # return locale.atof(billed_amount[0].replace(' ','').replace('€',''))

def _extract_power(line) -> str:
    power_type = None
    power_amount = None
    match = PX_PATTERN.match(line)
    if match:
        power_type = match.group(1)
        power_amount = match.group(2)
    else:
        match = PV_PATTERN.match(line)
        if match:
            power_type = match.group(1)
            power_amount = match.group(2)
    return power_type, power_amount

def extract_endesa_bill(file):
    logger.debug(f"Extracting bill info from ENDESA file '{file}' ...")
    bill_info = {
        'is_ours': False,
        'bill_id': None,
        'billing_date': None,
        'billing_period': None,
        'billed_power_capacity': None,
        'billed_energy_consumed': None,
        'billed_amount_0': None,
        'billed_amount_1': None,
        'is_rectification': False,
        'cups': None,
    }
    with pdfplumber.open(file) as pdf:
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
            power_type, power_amount = _extract_power(line)
            if power_type is not None:
                bill_info[power_type] = power_amount
                continue

        if 'P1' not in bill_info and 'Punta' not in bill_info:
            if len(pdf.pages) > 2:
                third_page = pdf.pages[2]
                for line in third_page.extract_text().split('\n'):
                    power_type, power_amount = _extract_power(line)
                    if power_type is not None:
                        bill_info[power_type] = power_amount

    if not bill_info['is_ours']:
        logger.warning(f"File '{file}' does not belong to us. Skipping.")
        return None
    if 'P1' not in bill_info and 'Punta' not in bill_info:
        logger.warning(f"Power reading not on 2nd nor on the 3rd page for '{file}'. Skipping.")
        return None
    return bill_info

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

    if 'P1' in data:
        for power_type in ['P1', 'P2', 'P3', 'P4', 'P5', 'P6']:
            if power_type not in data:
                logger.error(f"'{power_type}' comsumption has not been extracted from CUPS {data['cups']} and bill {data['bill_id']}.")
                return None
            data[power_type] = data[power_type].strip()
    elif 'Punta' in data:
        for power_type in ['Punta', 'Llano', 'Valle']:
            if power_type not in data:
                logger.error(f"'{power_type}' comsumption has not been extracted from CUPS {data['cups']} and bill {data['bill_id']}.")
                return None
            data[power_type] = data[power_type].strip()

    return data


def generate_bill_report(args, bills:dict):
    wb = Workbook()

    column_keys = list(args.defaults['columns'].keys())
    column_headers = list(args.defaults['columns'].values())

    # Add a new worksheet
    for cups, bill_infos in bills.items():
        ws = wb.create_sheet(title=cups)

        # Add data to each worksheet
        is_first_row = True
        for bill_info in bill_infos.values():
            if is_first_row:
                ws.append(column_headers)
                is_first_row = False
            ws.append([ bill_info.get(h, None) for h in column_keys ])

    # Save the workbook
    wb.save(args.output)


def extract_bill_information(args):
    # Recursively list all files found in the args.input_dir
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
        bill_info = extract_endesa_bill(file)
        if not bill_info:
            continue
        bill_info = sanitize_bill(bill_info)
        if not bill_info:
            continue
        cups = bill_info['cups']
        bill_id = bill_info['bill_id']
        if cups not in bills:
            bills[cups] = {}
        bills[cups][bill_id] = bill_info

    generate_bill_report(args, bills)

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


def main():
    with open(os.path.join(os.path.dirname(__file__), 'defaults.yaml'), 'r', encoding='utf-8') as f:
        defaults = yaml.safe_load(f)
    args = parse_command_line(defaults)
    args.defaults = defaults
    try:
        locale.setlocale(locale.LC_NUMERIC, args.input_locale)
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


if __name__ == '__main__':
    sys.exit(main())


