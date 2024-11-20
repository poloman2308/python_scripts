import pandas as pd
from openpyxl.reader.excel import load_workbook

from devices_details import \
    set_recommended_action
from executive_summary import \
    ExecutiveSummary
from hw_strategy import \
    HWStrategy
from sw_strategy import \
    SWStrategy


def enrich(filename):
    df = pd.read_excel(
        filename,
        sheet_name='Report',
        header=0,
    )
    set_recommended_action(df)
    wb = load_workbook(filename)
    ExecutiveSummary(wb, df).create()
    HWStrategy(wb, df).create()
    SWStrategy(wb, df).create()
    wb.save(filename)