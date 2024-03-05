import datetime
import logging
import sys
import argparse
import time
import openpyxl
from farmer_db import easyfarmer

SNR_TYPE_ESP  = 1
SNR_TYPE_LORA = 2
SNR_TYPE_WM   = 3

DEVEL = True
gid_list = ['2037314b-50525007-440020']

def validate_date(date_str):
    try:
        return datetime.datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        msg = f"not a valid date: {date_str}"
        raise argparse.ArgumentTypeError(msg)


def logger_init():
    logger = logging.getLogger()
    logger.handlers = []
    logger.setLevel(logging.DEBUG)

    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

    stdout_handler = logging.StreamHandler(sys.stdout)
    stdout_handler.setLevel(logging.DEBUG)
    stdout_handler.setFormatter(formatter)
    logger.addHandler(stdout_handler)

    file_handler = logging.FileHandler(args.output if args.output else "gscand.log" if DEVEL else "/var/log/gscand/gscand.log")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)


def check_sensor(g, snr_type, sheet):
    i = 2
    dns = ez.get_dns(g, snr_type)
    for dn in dns:
        ds = ez.get_data(g.gateway_id, dn.dnid, int(start_dt.timestamp()), int(end_dt.timestamp()))

        sheet.cell(row=i, column=1).value = dn.name
        if ds is not None:
            ds_dezip = list(zip(*ds))
            ds_values = list(ds_dezip[0])

            values_count = len(ds_values)
            failed_count = ds_values.count(-9999)

            if values_count - failed_count == 0:
                sheet.cell(row=i, column=2).value = '異常'
            elif failed_count > 0:
                sheet.cell(row=i, column=2).value = f'{failed_count}筆遺漏'
        i = i + 1


if __name__ == '__main__':
    ez = easyfarmer.EasyFarmer()

    parser = argparse.ArgumentParser(description="gscan daemon")
    parser.add_argument("-o", "--output", help="Output log file")
    parser.add_argument("-d", "--date", help="Date", type=validate_date)
    args = parser.parse_args()

    today = datetime.date.today()
    yesterday = today - datetime.timedelta(days=1)
    target_date = args.date if args.date else yesterday
    start_dt = datetime.datetime.combine(target_date, datetime.datetime.min.time())
    end_dt = datetime.datetime.combine(target_date, datetime.datetime.max.time())

    logger_init()
    logging.info("-" * 50)
    logging.info("<<< gscan daemon >>>")
    logging.info(f"Target Date: {target_date}")
    logging.info(f"{start_dt} - {end_dt}")
    logging.info("-" * 50)

    logging.info("gscand starting...")

    now_t = int(time.time())

    workbook = openpyxl.Workbook()

    sheet0 = workbook.create_sheet('氣象站', 0)
    sheet0.cell(row=1, column=1).value = '感測器'
    sheet0.cell(row=1, column=2).value = f'{target_date.month}-{target_date.day}'

    sheet1 = workbook.create_sheet('水位計', 1)
    sheet1.cell(row=1, column=1).value = '感測器'
    sheet1.cell(row=1, column=2).value = f'{target_date.month}-{target_date.day}'

    sheet2 = workbook.create_sheet('水錶', 2)
    sheet2.cell(row=1, column=1).value = '感測器'
    sheet2.cell(row=1, column=2).value = f'{target_date.month}-{target_date.day}'

    for gid in gid_list:
        g = ez.get_gw(gid)
        check_sensor(g, ez.SNR_TYPE_ESP, sheet0)
        check_sensor(g, ez.SNR_TYPE_LORA, sheet1)
        check_sensor(g, ez.SNR_TYPE_WM, sheet2)

    workbook.save("test.xlsx")

    logging.info("gscand terminated.")
