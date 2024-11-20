from copy import deepcopy
from datetime import date, datetime, timedelta

import pandas as pd


SPARING_POLICY = ["Unknown"]



def set_recommended_action(df: pd.DataFrame):
    """ Set Recommended Action column based on:
        Refresh HW, Add Coverage, Upgrade Software,
        Reload, Correct, Secure, Treat, Discover
        :param df: All Device Details
        :type df: pd.DataFrame
    """
    df.loc[(udpate_hardware_this_year), 'Update Hardware This Year'] = 'YES'
    df.loc[(udpate_hardware_next_year), 'Update Hardware Next Year'] = 'YES'
    df.loc[(udpate_hardware_next_next_year), 'Update Hardware In 2 Years'] = 'YES'
    df.loc[(udpate_software_this_year), 'Update Software This Year'] = 'YES'
    df.loc[(udpate_software_next_year), 'Update Software Next Year'] = 'YES'
    df.loc[(udpate_software_next_next_year), 'Update Software In 2 Years'] = 'YES'
    df.loc[(refresh_hardware), 'Refresh Hardware'] = 'YES'
    df.loc[(refresh_software), 'Refresh Software'] = 'YES'
    df.loc[:,'Recommended Action'] = 'No Action Required'
    df.loc[(refresh_software),'Recommended Action'] = 'Upgrade Software'
    df.loc[(refresh_hardware),'Recommended Action'] = 'Refresh Hardware'


def udpate_hardware_this_year(input: pd.DataFrame):
    end_of_current_year = datetime(date.today().year, 12, 31)
    return input['HW Last Date of Support (LDoS)'].transform(lambda x: isinstance(x, datetime) and x < end_of_current_year)

def udpate_hardware_next_year(input: pd.DataFrame):
    end_of_current_year = datetime(date.today().year+1, 12, 31)
    return input['HW Last Date of Support (LDoS)'].transform(lambda x: isinstance(x, datetime) and x < end_of_current_year)

def udpate_hardware_next_next_year(input: pd.DataFrame):
    end_of_current_year = datetime(date.today().year+2, 12, 31)
    return input['HW Last Date of Support (LDoS)'].transform(lambda x: isinstance(x, datetime) and x < end_of_current_year)

def udpate_software_this_year(input: pd.DataFrame):
    end_of_current_year = datetime(date.today().year, 12, 31)
    return input['SW End of SW Maintenance (EoSWM)'].transform(lambda x: isinstance(x, datetime) and x < end_of_current_year)

def udpate_software_next_year(input: pd.DataFrame):
    end_of_current_year = datetime(date.today().year+1, 12, 31)
    return input['SW End of SW Maintenance (EoSWM)'].transform(lambda x: isinstance(x, datetime) and x < end_of_current_year)

def udpate_software_next_next_year(input: pd.DataFrame):
    end_of_current_year = datetime(date.today().year+2, 12, 31)
    return input['SW End of SW Maintenance (EoSWM)'].transform(lambda x: isinstance(x, datetime) and x < end_of_current_year)

def refresh_hardware(input: pd.DataFrame):
    future_target = datetime.combine(date.today() + timedelta(days=365), datetime.min.time())
    return input['HW Last Date of Support (LDoS)'].transform(lambda x: isinstance(x, datetime) and x < future_target)


def refresh_software(input: pd.DataFrame):
    future_target = datetime.combine(date.today() + timedelta(days=365), datetime.min.time())
    return input['SW End of SW Maintenance (EoSWM)'].transform(lambda x: isinstance(x, datetime) and x < future_target)

