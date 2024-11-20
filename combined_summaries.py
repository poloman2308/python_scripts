from openpyxl.chart import Reference
from openpyxl.styles import Alignment, Border, Side
from openpyxl.workbook import Workbook

from .utils import (
    calculate_values, column_sizing, create_bar_chart, ensure_tabs_exist,
    format_cells, hide_columns, input_dates,
    input_values, standart_bold_font,
    standart_pattern_fill, thin_border,
)

cell_values = {
            "B1": "Install Base (IB)",
            "C1": "HW LDoS",
            "D1": "SW EoSWM non-compliance",
            "E1": "SW EoVSS non-compliance",
            "F1": "SW LDoS",
            "A2": "All NPL reported Chassis",
            "A3": "NPL Chassis excluding APs",
            "A4": "Not Available (No Owner Provided)",
            "A5": "Business Partner",
            "A6": "Core WAN & Telecom",
            "A7": "Corporate",
            "A8": "Electronic Trading Services",
            "A9": "EUS Voice",
            "A10": "Internet",
            "A11": "Application Traffic Routing",
            "A12": "Retail Branch",
            "A13": "Switch and Routing",
            "A14": "Shared Network Products",
            "A15": "Swiss Owned Devices",
            "A16": "Unknown",
            "A17": "Wireless",
            "A18": "Wireless EXcluding APs"
        }
cell_formulas = {
            "B2:B2": "='NPL - HW EOL Timeline'!F2",
            "B3:B3": "='NPL - HW EOL Timeline'!F3",
            "B4:B4": "='NPL - HW EOL Timeline'!F4",
            "B5:B5": "='NPL - HW EOL Timeline'!F5",
            "B6:B6": "='NPL - HW EOL Timeline'!F6",
            "B7:B7": "='NPL - HW EOL Timeline'!F7",
            "B8:B8": "='NPL - HW EOL Timeline'!F8",
            "B9:B9": "='NPL - HW EOL Timeline'!F9",
            "B10:B10": "='NPL - HW EOL Timeline'!F10",
            "B11:B11": "='NPL - HW EOL Timeline'!F11",
            "B12:B12": "='NPL - HW EOL Timeline'!F12",
            "B13:B13": "='NPL - HW EOL Timeline'!F13",
            "B14:B14": "='NPL - HW EOL Timeline'!F14",
            "B15:B15": "='NPL - HW EOL Timeline'!F15",
            "B16:B16": "='NPL - HW EOL Timeline'!F16",
            "B17:B17": "='NPL - HW EOL Timeline'!F17",
            "B18:B18": "='NPL - HW EOL Timeline'!F18",
            "C2:C2": "='NPL - HW EOL Timeline'!G2",
            "C3:C3": "='NPL - HW EOL Timeline'!G3",
            "C4:C4": "='NPL - HW EOL Timeline'!G4",
            "C5:C5": "='NPL - HW EOL Timeline'!G5",
            "C6:C6": "='NPL - HW EOL Timeline'!G6",
            "C7:C7": "='NPL - HW EOL Timeline'!G7",
            "C8:C8": "='NPL - HW EOL Timeline'!G8",
            "C9:C9": "='NPL - HW EOL Timeline'!G9",
            "C10:C10": "='NPL - HW EOL Timeline'!G10",
            "C11:C11": "='NPL - HW EOL Timeline'!G11",
            "C12:C12": "='NPL - HW EOL Timeline'!G12",
            "C13:C13": "='NPL - HW EOL Timeline'!G13",
            "C14:C14": "='NPL - HW EOL Timeline'!G14",
            "C15:C15": "='NPL - HW EOL Timeline'!G15",
            "C16:C16": "='NPL - HW EOL Timeline'!G16",
            "C17:C17": "='NPL - HW EOL Timeline'!G17",
            "C18:C18": "='NPL - HW EOL Timeline'!G18",
            "D2:D2": "='NPL - SW EOL Timeline'!G2",
            "D3:D3": "='NPL - SW EOL Timeline'!G3",
            "D4:D4": "='NPL - SW EOL Timeline'!G4",
            "D5:D5": "='NPL - SW EOL Timeline'!G5",
            "D6:D6": "='NPL - SW EOL Timeline'!G6",
            "D7:D7": "='NPL - SW EOL Timeline'!G7",
            "D8:D8": "='NPL - SW EOL Timeline'!G8",
            "D9:D9": "='NPL - SW EOL Timeline'!G9",
            "D10:D10": "='NPL - SW EOL Timeline'!G10",
            "D11:D11": "='NPL - SW EOL Timeline'!G11",
            "D12:D12": "='NPL - SW EOL Timeline'!G12",
            "D13:D13": "='NPL - SW EOL Timeline'!G13",
            "D14:D14": "='NPL - SW EOL Timeline'!G14",
            "D15:D15": "='NPL - SW EOL Timeline'!G15",
            "D16:D16": "='NPL - SW EOL Timeline'!G16",
            "D17:D17": "='NPL - SW EOL Timeline'!G17",
            "D18:D18": "='NPL - SW EOL Timeline'!G18",
            "E2:E2": "='NPL - SW EOL Timeline'!H2",
            "E3:E3": "='NPL - SW EOL Timeline'!H3",
            "E4:E4": "='NPL - SW EOL Timeline'!H4",
            "E5:E5": "='NPL - SW EOL Timeline'!H5",
            "E6:E6": "='NPL - SW EOL Timeline'!H6",
            "E7:E7": "='NPL - SW EOL Timeline'!H7",
            "E8:E8": "='NPL - SW EOL Timeline'!H8",
            "E9:E9": "='NPL - SW EOL Timeline'!H9",
            "E10:E10": "='NPL - SW EOL Timeline'!H10",
            "E11:E11": "='NPL - SW EOL Timeline'!H11",
            "E12:E12": "='NPL - SW EOL Timeline'!H12",
            "E13:E13": "='NPL - SW EOL Timeline'!H13",
            "E14:E14": "='NPL - SW EOL Timeline'!H14",
            "E15:E15": "='NPL - SW EOL Timeline'!H15",
            "E16:E16": "='NPL - SW EOL Timeline'!H16",
            "E17:E17": "='NPL - SW EOL Timeline'!H17",
            "E18:E18": "='NPL - SW EOL Timeline'!H18",
            "F2:F2": "='NPL - SW EOL Timeline'!I2",
            "F3:F3": "='NPL - SW EOL Timeline'!I3",
            "F4:F4": "='NPL - SW EOL Timeline'!I4",
            "F5:F5": "='NPL - SW EOL Timeline'!I5",
            "F6:F6": "='NPL - SW EOL Timeline'!I6",
            "F7:F7": "='NPL - SW EOL Timeline'!I7",
            "F8:F8": "='NPL - SW EOL Timeline'!I8",
            "F9:F9": "='NPL - SW EOL Timeline'!I9",
            "F10:F10": "='NPL - SW EOL Timeline'!I10",
            "F11:F11": "='NPL - SW EOL Timeline'!I11",
            "F12:F12": "='NPL - SW EOL Timeline'!I12",
            "F13:F13": "='NPL - SW EOL Timeline'!I13",
            "F14:F14": "='NPL - SW EOL Timeline'!I14",
            "F15:F15": "='NPL - SW EOL Timeline'!I15",
            "F16:F16": "='NPL - SW EOL Timeline'!I16",
            "F17:F17": "='NPL - SW EOL Timeline'!I17",
            "F18:F18": "='NPL - SW EOL Timeline'!I18"
        }
cell_format = {
            "B1:F1": {
                "font": standart_bold_font,
                "fill": standart_pattern_fill,
                "wrap_text": True,
                "border": thin_border,
            },
            "B2:F18": {
                "align_right": True,
                "border": thin_border,
            },
            "A2:A18": {
                "font": standart_bold_font,
                "fill": standart_pattern_fill,
                "border": thin_border,
            },
    }


def format_cells_percentage_combsum(work_sheet):
    # Apply percentage format to cells C2:F28
    for row in work_sheet.iter_rows(min_row=2, max_row=28, min_col=3, max_col=6):
        for cell in row:
            cell.number_format = '0.0%'


def create_combined_summaries(wb: Workbook):
    ensure_tabs_exist(
        wb,
        (
            "Combined Summaries",
        )
    )
    input_values(
        wb["Combined Summaries"],
        cell_values,
        cell_formulas,
    )

    wb["Combined Summaries"].add_chart(
        create_bar_chart(
            data_categories=Reference(
                wb["Combined Summaries"],
                min_col=1,
                min_row=2,
                max_row=18
            ),
            data_values=Reference(
                wb["Combined Summaries"],
                min_col=3,
                max_col=6,
                min_row=1,
                max_row=18
            ),
            title="By Product Owner",
            chart_width=30,
            legend_pos_b=True,
        ),
        "H2",
    )

    format_cells(
        wb["Combined Summaries"],
        cell_format,
    )
    format_cells_percentage_combsum(wb['Combined Summaries'])
    column_sizing(wb["Combined Summaries"])
