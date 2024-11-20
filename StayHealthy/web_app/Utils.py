from decimal import Decimal, ROUND_UP
import pandas as pd


def create_workbook(name):
    writer = pd.ExcelWriter(name, engine='xlsxwriter')
    # workbook = writer.book
    return writer

def save_workbook(workbook):
    workbook.close()
    return True

def round_up(number):
    return Decimal(str(number)).quantize(Decimal('.01'), rounding=ROUND_UP)