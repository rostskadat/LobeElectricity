#!/usr/bin/env python
"""
Extract bill information from PDF files and generate a report.
"""

from datetime import date
from email.mime import text
import sys

try:
    import coloredlogs
    import csv
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
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from os.path import basename, join, dirname
    from pprint import pprint
    from pathlib import Path
    import textfsm
    import gspread
    import pandas as pd
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build
except ModuleNotFoundError as e:
    print(f"{e}. Did you load your environment?")
    sys.exit(1)

load_dotenv()
logger = logging.getLogger()

# i.e.: +/-0.000,00 €
AMOUNT_PATTERN = re.compile(r"[+-]?(?:\d{1,3}(?:\.\d{3})*|\d+),\d{2} ?€")
# i.e.: P4 1.18.4 1.278,00 2.578,00 1,00 0,00 1.300,00
ENDESA_PX_PATTERN = re.compile(
    r"^(P[123456]) 1\.18\.[123456]( (?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}){5}.*"
)
ENDESA_PV_PATTERN = re.compile(
    r".*(Punta|Llano|Valle)( (?:\d{1,3}(?:\.\d{3})*|\d+),\d{2}){5}.*"
)
# i.e.: Px 7.994 7.994 0 66 66 0 0,00 0,00 0,00
TOTAL_PV_PATTERN = re.compile(
    r"^(P[123456])( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){3}( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){6}"
)
# i.e.: Px 593 665 72 0 0 3,81
NUFRI_PX_PATTERN = re.compile(
    r"^(P[123456])( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){3}( (?:\d{1,3}(?:\.\d{3})*|\d+)(,\d{2})?){3}"
)

STD_DATE_FMT = r"\d{2}/\d{2}/\d{4}"


def main():
    with open(join(dirname(__file__), "defaults.yaml"), "r", encoding="utf-8") as f:
        defaults = yaml.safe_load(f)
    args = parse_command_line(defaults)
    args.defaults = defaults
    try:
        # locale.setlocale(locale.LC_TIME, 'Spanish_Spain.1252')  # On Windows
        locale.setlocale(locale.LC_NUMERIC, args.locale)
        locale.setlocale(locale.LC_TIME, args.locale)
        for verbose_package in [
            "pdfminer.pdfpage",
            "pdfminer.pdfdocument",
            "pdfminer.psparser",
            "pdfminer.pdfinterp",
            "pdfminer.cmapdb",
            "pdfminer.pdfparser",
            "googleapiclient.discovery_cache",
        ]:
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
    parser = ArgumentParser(
        prog="extract_bill_information",
        description=__doc__,
        formatter_class=RawTextHelpFormatter,
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="Run the program in debug. Default to 'False'.",
        required=False,
        default=False,
    )
    parser.add_argument(
        "--quiet",
        action="store_true",
        help="Run the program in quiet mode. Default to 'False'.",
        required=False,
        default=False,
    )
    parser.add_argument(
        "--locale",
        help=f"The locale to use when reading the bills. By default '{defaults['locale']}'.",
        required=False,
        default=defaults["locale"],
    )
    parser.add_argument(
        "--no-refresh",
        action="store_true",
        help=f"Use the Excel output file to upload to Google Sheets. Default to False.",
        required=False,
        default=False,
    )
    parser.add_argument(
        "--bill-input",
        help=f"The path where the bills are located. Either a directory where all bills will be processed or a file for just one bill. By default '{defaults['bill-input']}'.",
        required=False,
        default=defaults["bill-input"],
    )
    parser.add_argument(
        "--bill-input-excludes",
        type=lambda s: s.split(","),
        nargs="+",
        help=f"A comma separated list of paths to exclude. By default '{','.join(defaults['bill-input-excludes'])}'.",
        required=False,
        default=",".join(defaults["bill-input-excludes"]),
    )
    parser.add_argument(
        "--include-loads",
        action="store_true",
        help=f"Wether to add the load in the output. By default '{defaults['include-loads']}'.",
        required=False,
        default=defaults["include-loads"],
    )
    parser.add_argument(
        "--load-input-dir",
        help=f"The path where the laod CSV are located. By default '{defaults['load-input-dir']}'.",
        required=False,
        default=defaults["load-input-dir"],
    )
    parser.add_argument(
        "--output",
        help=f"The name of the report. By default '{defaults['output']}'.",
        required=False,
        default=defaults["output"],
    )
    parser.add_argument(
        "--limit",
        type=int,
        help=f"Limit the number of files processed. By default '{defaults['limit']}' (-1 is no limit).",
        required=False,
        default=defaults["limit"],
    )
    parser.add_argument(
        "--no-trim",
        action="store_true",
        help=f"Trim unwanted columns. Useful for debugging. Default to 'False'.",
        required=False,
        default=False,
    )
    parser.add_argument(
        "--dump",
        action="store_true",
        help="Just dump the text of the bills to stdout. Useful for debugging. Default to 'False'.",
        required=False,
        default=False,
    )
    parser.add_argument(
        "--dump-output-prefix",
        help=f"The prefix for the dump output files. By default '{defaults['dump-output-prefix']}'.",
        required=False,
        default=defaults["dump-output-prefix"],
    )
    parser.add_argument(
        "--no-upload",
        action="store_true",
        help=f"Upload the extracted bill information to Google Sheets. Default to False.",
        required=False,
        default=False,
    )
    parser.add_argument(
        "--gsheet-id",
        help=f"The ID of the Google Sheet to upload to. By default '{defaults['gsheet-id']}'.",
        required=False,
        default=defaults["gsheet-id"],
    )
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
    if not os.path.isdir(args.bill_input) and not os.path.isfile(args.bill_input):
        logger.error(f"Input directory '{args.bill_input}' does not exist.")
        return 1
    if not args.no_refresh:
        logger.info(f"Reading files from '{args.bill_input}' ...")
    bill_input_excludes = [
        re.compile(bill_input_exclude)
        for bill_input_exclude in args.bill_input_excludes
    ]
    if os.path.isfile(args.bill_input):
        files = [args.bill_input]
    else:
        files = [
            full_path
            for dp, dn, filenames in os.walk(args.bill_input)
            for pdf_filename in filenames
            for full_path in [join(dp, pdf_filename)]
            if pdf_filename.lower().endswith(".pdf")
            and not any(regex.search(full_path) for regex in bill_input_excludes)
        ]
    if not files:
        logger.error(f"No files found in '{args.bill_input}'.")
        return 1
    logger.debug(f"Found {len(files)} files in '{args.bill_input}'.")
    if args.limit > 0:
        files = files[: args.limit]
        logger.warning(f"Limiting to {len(files)} files as per the --limit argument.")

    if args.dump:
        for file in files:
            logger.info(file)
            with open(
                f"{args.dump_output_prefix}{basename(file)}.txt", "w", encoding="utf-8"
            ) as f:
                with pdfplumber.open(file) as pdf:
                    for page in pdf.pages:
                        f.write(page.extract_text())
        logger.info(
            f"Dumped {len(files)} files to {basename(args.dump_output_prefix)}. Exiting."
        )
        return

    if args.no_refresh:
        bills = {}
        loads = {}
    else:
        bills = extract_bills(args, files)
        loads = {}
        if args.include_loads:
            files = [
                join(dp, f)
                for dp, dn, filenames in os.walk(args.load_input_dir)
                for f in filenames
                if f.lower().endswith(".csv")
            ]
            if not files:
                logger.warning(f"No files found in '{args.load_input_dir}'.")
            else:
                logger.debug(f"Found {len(files)} files in '{args.bill_input}'.")
                if args.limit > 0:
                    files = files[: args.limit]
                    logger.warning(
                        f"Limiting to {len(files)} files as per the --limit argument."
                    )
                loads = extract_loads(args, files)

    wb = generate_report(args, bills, loads)

    if not args.no_upload:
        if not args.gsheet_id:
            logger.error(
                f"When uploading to GoogleSheet you must set the 'gsheet-id' parameter."
            )
            return 1
        upload_report(args.gsheet_id, wb)


def extract_bills(args, files: list):
    """Process all the input bills and construct a dict by CUPS/BILL_ID with all the relevant information

    Args:
        args (namespace): the input arguments
        files (list): the list of files to process

    Returns:
        dict: a dict of CUPS/BILL_ID with all the extracted information
    """
    bills = {}
    for file in files:
        bill_info = extract_dispatcher(args.defaults["dispatchers"], file, not args.no_trim)
        if not bill_info:
            continue
        bill_info = sanitize_bill(bill_info)
        if not bill_info:
            logger.debug(f"Skipping {file} ...")
            continue
        bill_info["file"] = basename(file)
        cups = bill_info["cups"]
        if cups not in bills:
            bills[cups] = []
        bills[cups].append(bill_info)
    return bills


def extract_dispatcher(dispatchers: dict, file: str, trim: bool):
    """
    Extracts bill information by selecting and invoking the
    appropriate extractor function based on the content of
    the first page of a PDF bill.

    Args:
        dispatchers (dict): A dictionary mapping dispatcher names (str) to extractor function names (str).
        file (str): The path to the PDF file to be processed.
        trim (bool): whether to trim unwanted columns.

    Returns:
        Any: The result of the extractor function if a matching dispatcher is found; otherwise, None.

    Raises:
        ValueError: If the specified extractor function is not found or is not callable.

    """
    logger.info("Extracting information from bill '%s' ...", _readable_path(file))
    with pdfplumber.open(file) as pdf:
        first_page = pdf.pages[0].extract_text()
        found_extractor = False
        for dispatcher, extractor in dispatchers.items():
            if extractor not in globals() or not callable(globals()[extractor]):
                raise ValueError(f"Extractor '{extractor}' not found or not callable.")
            if dispatcher in first_page:
                logger.debug(
                    f"Detected '{dispatcher}' bill. Using {extractor} to extract information ..."
                )
                bill_info = globals()[extractor](pdf, trim)
                return bill_info
        if not found_extractor:
            logger.error(f"Could not detect the type of bill for '{file}'. Skipping")
            return None, None, None


def extract_plenitude_bill(pdf, trim: bool):
    """Extractor for plenitude bills

    Args:
        pdf (PDFfile): The corresponding pdf file
        trim (bool): whether to trim unwanted columns.

    Returns:
        dict: A dictionary containing the extracted bill information.
    """
    if not hasattr(extract_plenitude_bill, "re_table"):
        extract_plenitude_bill.re_table = None
    numeric_cols = [
        "billed_amount",
        "billed_energy",
        "billed_power",
        "P1",
        "P2",
        "P3",
    ]
    with open(
        join(dirname(__file__), "assets/templates/es.plenitude.textfsm"),
        "r",
        encoding="utf-8",
    ) as template:
        extract_plenitude_bill.re_table = textfsm.TextFSM(template)
    df = _extract_bill(pdf, extract_plenitude_bill.re_table, numeric_cols)
    return df.iloc[0].to_dict()


def extract_nufri_bill(pdf, trim: bool):
    """Extractor for Nufri bills

    Args:
        pdf (PDFfile): The corresponding pdf file
        trim (bool): whether to trim unwanted columns.

    Returns:
        dict: A dictionary containing the extracted bill information.
    """
    if not hasattr(extract_nufri_bill, "re_table"):
        extract_nufri_bill.re_table = None

    # TODO: extract CPx information
    numeric_cols = [
        "billed_amount",
        "billed_energy",
        "billed_power",
        "P1",
        "P2",
        "P3",
    ]
    with open(
        join(dirname(__file__), "assets/templates/es.nufri.textfsm"),
        "r",
        encoding="utf-8",
    ) as template:
        extract_nufri_bill.re_table = textfsm.TextFSM(template)
    df = _extract_bill(pdf, extract_nufri_bill.re_table, numeric_cols)
    return df.iloc[0].to_dict()


def extract_te_bill(pdf, trim: bool):
    """Extractor for Total Energie bills

    Args:
        pdf (PDFfile): The corresponding pdf file
        trim (bool): whether to trim unwanted columns.

    Returns:
        dict: A dictionary containing the extracted bill information.
    """
    if not hasattr(extract_te_bill, "re_table"):
        extract_te_bill.re_table = None

    numeric_cols = [
        "billed_amount",
        "te_power_access_P1",
        "te_power_access_P2",
        "te_power_access_P3",
        "te_power_access_P4",
        "te_power_access_P5",
        "te_power_access_P6",
        "te_power_P1",
        "te_power_P2",
        "te_power_P3",
        "te_power_P4",
        "te_power_P5",
        "te_power_P6",
        "te_power_charge_P1",
        "te_power_charge_P2",
        "te_power_charge_P3",
        "te_power_charge_P4",
        "te_power_charge_P5",
        "te_power_charge_P6",
        "te_energy_access_P1",
        "te_energy_access_P2",
        "te_energy_access_P3",
        "te_energy_access_P4",
        "te_energy_access_P5",
        "te_energy_access_P6",
        "te_energy_P1",
        "te_energy_P2",
        "te_energy_P3",
        "te_energy_P4",
        "te_energy_P5",
        "te_energy_P6",
        "te_energy_charge_P1",
        "te_energy_charge_P2",
        "te_energy_charge_P3",
        "te_energy_charge_P4",
        "te_energy_charge_P5",
        "te_energy_charge_P6",
        "P1",
        "P2",
        "P3",
        "P4",
        "P5",
        "P6",
    ]

    with open(
        join(dirname(__file__), "assets/templates/es.totalenergies.textfsm"),
        "r",
        encoding="utf-8",
    ) as template:
        extract_te_bill.re_table = textfsm.TextFSM(template)
    df = _extract_bill(pdf, extract_te_bill.re_table, numeric_cols)
    # Do some summing
    df["billed_energy"] = df[
        [
            f"{prefix}{i}"
            for i in range(1, 7)
            for prefix in ("te_energy_access_P", "te_energy_P", "te_energy_charge_P")
        ]
    ].sum(axis=1)
    df["billed_power"] = df[
        [
            f"{prefix}{i}"
            for i in range(1, 7)
            for prefix in ("te_power_access_P", "te_power_P", "te_power_charge_P")
        ]
    ].sum(axis=1)
    # unfortunately te_contracted_power is not easily extractible
    matches = re.findall(r"(P[1-6])\s+([\d,]+)", df["te_contracted_power"][0])
    if matches:
        for i in range(1, 7):
            df[f"CP{i}"] = 0.0
        for k, v in dict(matches).items():
            df[f"C{k}"] = locale.atof(v)
    # Dropping all specific columns
    if trim:
        logger.debug("Removing all specific te_* columns ...")
        df = df.loc[:, ~df.columns.str.startswith("te_")]
    return df.iloc[0].to_dict()


def extract_endesa_bill(pdf, trim: bool):
    """Extractor for Endesa bills

    Args:
        pdf (PDFfile): The corresponding pdf file
        trim (bool): whether to trim unwanted columns.

    Returns:
        dict: A dictionary containing the extracted bill information.
    """
    if not hasattr(extract_endesa_bill, "re_table"):
        extract_endesa_bill.re_table = None
    numeric_cols = [
        "billed_amount",
        "billed_power",
        "billed_energy",
        "P1",
        "P2",
        "P3",
        "P4",
        "P5",
        "P6",
        "CP1",
        "CP2",
        "CP3",
        "CP4",
        "CP5",
        "CP6",
    ]

    with open(
        join(dirname(__file__), "assets/templates/es.endesa.textfsm"),
        "r",
        encoding="utf-8",
    ) as template:
        extract_endesa_bill.re_table = textfsm.TextFSM(template)
    df = _extract_bill(pdf, extract_endesa_bill.re_table, numeric_cols)
    # I need to adjust the billing period end by -1 otherwise I
    # overshoot the contracted power calculation. Except when it's
    # just one day
    # if df["billing_period_start"][0] != df["billing_period_end"][0]:
        # XXX: check that is correct
        # df["billing_period_end"][0] = df["billing_period_end"][0] - 1
        # pass
    matches = re.findall(
        r"\s*(punta|punta-llano|valle)\s*([\d,]+)\s*kW;?",
        df["endesa_contracted_power"][0],
    )
    if matches:
        for i in range(1, 7):
            df[f"CP{i}"] = 0.0
        for k, v in dict(matches).items():
            if k == "punta" or k == "punta-llano":
                df[f"CP1"] = locale.atof(v)
            elif k == "valle":
                df[f"CP3"] = locale.atof(v)
            else:
                logger.warning(f"Unknown contracted power {k}")
    # Dropping all specific columns
    if trim:
        df = df.loc[:, ~df.columns.str.startswith("endesa_")]
    return df.iloc[0].to_dict()


def extract_qener_bill(pdf, trim: bool):
    """Extractor for Qener bills

    Args:
        pdf (PDFfile): The corresponding pdf file
        trim (bool): whether to trim unwanted columns.

    Returns:
        dict: A dictionary containing the extracted bill information.
    """
    if not hasattr(extract_qener_bill, "re_table"):
        extract_qener_bill.re_table = None
    numeric_cols = [
        "billed_amount",
        "billed_energy",
        "billed_power",
        "P1",
        "P2",
        "P3",
        "P4",
        "P5",
        "P6",
        "CP1",
        "CP2",
        "CP3",
        "CP4",
        "CP5",
        "CP6",
    ]
    with open(
        join(dirname(__file__), "assets/templates/es.qener.textfsm"),
        "r",
        encoding="utf-8",
    ) as template:
        extract_qener_bill.re_table = textfsm.TextFSM(template)
    df = _extract_bill(pdf, extract_qener_bill.re_table, numeric_cols)
    return df.iloc[0].to_dict()


def _extract_bill(pdf, re_table, numeric_columns):
    pages = [page.extract_text() for page in pdf.pages]
    headers = re_table.header
    data = re_table.ParseText("\n".join(pages))
    df = pd.DataFrame(data, columns=headers)
    # Convert to float and fill with 0
    df[numeric_columns] = (
        df[numeric_columns]
        .apply(
            lambda serie: pd.to_numeric(
                serie.astype(str)
                .str.replace(".", "", regex=False)
                .str.replace(",", ".", regex=False),
                errors="coerce",
            )
        )
        .fillna(0)
    )
    return df


def sanitize_bill(data: dict):
    """
    Sanitize the bill data by converting strings to appropriate types and formatting.
    """
    is_sane, missing_keys = _is_bill_sane(data)
    if not is_sane:
        logger.error(
            f"Missing keys for CUPS {data['cups']} and bill {data['bill_id']}: {missing_keys}"
        )
        return None

    if not data["cups"].endswith("0F"):
        # For some reason all CUPS ends with 0F except for Qener bills...
        cups = f"{data['cups']}0F"
        data["cups"] = cups

    if not data["billing_date"]:
        logger.error(
            f"Invalid billing date for CUPS {data['cups']} and bill {data['bill_id']}"
        )
        return None

    if not data["billing_period_start"] or not data["billing_period_end"]:
        logger.error(
            f"Invalid billing period for CUPS {data['cups']} and bill {data['bill_id']}"
        )
        return None

    if "billed_power" not in data:
        logger.error(
            f"Invalid billed power capacity (not in the expected format '+/-0.000,00 €') for CUPS {data['cups']} and bill {data['bill_id']}."
        )
        return None

    if "billed_energy" not in data:
        logger.error(
            f"Invalid billed energy consumed (not in the expected format '+/-0.000,00 €') for CUPS {data['cups']} and bill {data['bill_id']}."
        )
        return None

    if "billed_amount" not in data:
        logger.error(
            f"Invalid billed amount (not in the expected format '+/-0.000,00 €') for CUPS {data['cups']} and bill {data['bill_id']}."
        )
        return None

    for power_type in ["P1", "P2", "P3", "P4", "P5", "P6"]:
        if power_type not in data and power_type in ["P1", "P2", "P3"]:
            logger.error(
                f"Mandatory '{power_type}' comsumption has not been extracted from CUPS {data['cups']} and bill {data['bill_id']}."
            )
            return None
        if (
            power_type in data
            and data[power_type] is not None
            and type(data[power_type]) is str
        ):  # P4, P5, P6
            data[power_type] = locale.atof(data[power_type].strip())

    return data


def _is_bill_sane(data: dict):
    # check that all the keys have a value
    is_sane = True
    missing_keys = []
    for key, value in data.items():
        if value is None or (isinstance(value, str) and not value.strip()):
            missing_keys.append(key)
            is_sane = False
    return is_sane, missing_keys


def _readable_path(path: str, max_len: int = 72) -> str:
    path = Path(path)
    parts = path.parts

    # Return unchanged if already short enough
    if len(str(path)) <= max_len:
        return str(path)

    # Keep first element and as many trailing elements as possible
    for n in range(len(parts) - 1, 1, -1):
        shortened = str(Path(parts[0], "...", *parts[-n:]))
        if len(shortened) <= max_len:
            return shortened

    # Fallback
    return f"{parts[0]}/.../{parts[-1]}"


def extract_loads(args, files: list):
    loads = {}
    for file in files:
        with open(file, encoding="utf-8") as f:
            reader = csv.DictReader(f, delimiter=";")
            for row in reader:
                cups = row["CUPS"].strip()
                fecha = row["Fecha"].strip()
                hora = row["Hora"].strip()
                ae_kwh = row["AE_kWh"].strip()

                # Hora starts at 1, while datetime starts at 0
                dt_str = f"{fecha} {int(hora)-1}"
                try:
                    dt = datetime.strptime(dt_str, "%d/%m/%Y %H")
                except ValueError as e:
                    logger.error(
                        f"Could not parse datetime '{dt_str}' in file '{file}': {str(e)}"
                    )
                    continue

                # Convert AE_kWh to float
                try:
                    if ae_kwh != "":
                        ae_kwh_val = locale.atof(ae_kwh)
                    else:
                        ae_kwh_val = 0.0
                except ValueError as e:
                    logger.error(f"Could not parse AE_kWh '{ae_kwh}' in file '{file}'")
                    continue

                # Insert into loads dict
                if cups not in loads:
                    loads[cups] = {}
                loads[cups][dt] = ae_kwh_val
    return loads


def generate_report(args, bills: dict, loads: dict) -> Workbook:
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
        loads (dict): A dictionary where each key is a CUPS identifier and each value is another
            dictionary containing load information for that CUPS.

    Returns:
        Workbook. The function generated workbook
    """
    if args.no_refresh:
        return load_workbook(args.output, data_only=True)

    wb = Workbook()

    column_keys = list(args.defaults["column_labels"].keys())
    column_headers = list(args.defaults["column_labels"].values())
    sheet_names = args.defaults["sheet_names"]

    ws = wb.active
    ws.title = "Loads"
    ws.append(["CUPS", "Fecha", "AE_kWh"])
    for cups, cups_loads in loads.items():
        for dt, load in cups_loads.items():
            ws.append([cups, dt, load])

    sheets = {}
    # Add a new worksheet for each CUPS
    for cups, bill_infos in bills.items():
        logger.info(
            f"Adding worksheet for CUPS '{cups}' with {len(bill_infos)} bills ..."
        )
        ws = wb.create_sheet(title=sheet_names.get(cups, cups))
        sheets[ws.title] = ws

        # Enrich the report
        if args.no_trim:
            df = pd.DataFrame(bill_infos)
        else:
            df = pd.DataFrame(bill_infos, columns=column_keys)
        df["gross_amount"] = df["billed_power"] + df["billed_energy"]
        df = df.rename(columns=args.defaults["column_labels"])

        # And write it in the worksheet
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

    ordered_sheets = []
    for title in args.defaults["sheet_names"].values():
        if title in sheets:
            ordered_sheets.append(sheets[title])
    if len(ordered_sheets) == 0:
        ordered_sheets.extend(list(sheets.values()))
    wb._sheets = [wb.worksheets[0]] + ordered_sheets

    # Save the workbook
    wb.save(args.output)
    return wb


def upload_report(gsheet_id: str, wb: Workbook):
    logger.info(f"Uploading report to Google Sheets with ID '{gsheet_id}' ...")

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_file(
        filename="service-account.json", scopes=scopes
    )
    gc = gspread.authorize(creds)

    # gc = gspread.service_account(filename="service-account.json", scopes = scopes)
    spreadsheet = gc.open_by_key(gsheet_id)
    for ws in wb.worksheets:
        try:
            worksheet = spreadsheet.worksheet(ws.title)
            spreadsheet.del_worksheet(worksheet)
        except gspread.exceptions.WorksheetNotFound:
            pass
        worksheet = spreadsheet.add_worksheet(
            title=ws.title, rows=ws.max_row, cols=ws.max_column
        )
        values = []
        for row in ws.iter_rows(values_only=True):
            values.append([convert_value(value) for value in row])
        # Skip completely empty worksheets
        if not values:
            values = [[""]]
        worksheet.update(values)

        # set_with_dataframe(worksheet, data, include_index=False, include_column_header=False)

    drive = build("drive", "v3", credentials=creds)
    sheet = (
        drive.files()
        .get(
            fileId=gsheet_id,
            fields="id, kind, mimeType, name, size, modifiedTime, webViewLink",
        )
        .execute()
    )
    logger.info(f"Report '{sheet['name']}' saved to Google Drive ...")


def get_drive_path(drive, file_id):
    parts = []

    while file_id:
        item = drive.files().get(fileId=file_id, fields="id,name,parents").execute()

        parts.insert(0, item["name"])

        parents = item.get("parents")
        file_id = parents[0] if parents else None

    return "/".join(parts)


def convert_value(value):
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    elif isinstance(value, date):
        return value.isoformat()
    elif value is None:
        return ""
    return value


if __name__ == "__main__":
    sys.exit(main())
