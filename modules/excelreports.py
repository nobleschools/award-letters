#!python3

import numpy as np
import pandas as pd
from modules.filework import safe2int

"""
Module for creating the Excel reports from gdoc and local data
"""
def create_excel(dfs, campus, config, debug):
    """Will create Excel reports for sharing details from Google Docs"""
    if debug:
        print('Creating Excel report for {}'.format(campus),flush=True)
    ### First, create a dataframe for the Award_data tab
    award_data_df = build_award_df(dfs, campus, config, debug)

    print(award_data_df.head(2))

def build_award_df(dfs, campus, config, debug):
    """Builds a dataframe for the award fields"""
    ## First, start the df for the items that are straight pulls from live_data
    report_award_fields = config['report_award_fields']
    all_award_fields = []
    live_award_fields = [] # to hold the excel names
    live_award_targets = [] # to hold the live names
    complex_award_fields = []
    
    for column in report_award_fields:
        # Each column will be a dict with a single element
        # The key will be the Excel column name and the value the source
        # from the live table or other (lookup) table
        this_key = list(column.keys())[0]
        this_value = list(column.values())[0]
        all_award_fields.append(this_key)
        if ':' in this_value:
            complex_award_fields.append((this_key,this_value))
        else:
            live_award_fields.append(this_key)
            live_award_targets.append(this_value)
    if live_award_targets: # any fields here will be straight pulls from live df
        award_df = dfs['live_award'][live_award_targets]
        award_df = award_df.rename(columns=dict(zip(live_award_targets,
                                               live_award_fields)))
    else:
        print('Probably an error: no report columns pulling from live data')

    ## Second, pull columns that are lookups from other tables and append
    print(dfs['college'].head(3))
    for column, target in complex_award_fields:
        # parse the target and then call the appropriate function
        # to add a column to award_df
        print(column)
        print(target)
        tokens = target.split(sep=':')
        if tokens[0] == 'ROSTER':
            award_df[column] = dfs['live_award'][tokens[1]].apply(
                    lambda x: dfs['ros'].loc[x, tokens[2]])
        elif tokens[0] == 'COLLEGE':
            award_df[column] = dfs['live_award'][tokens[1]].apply(
                    lambda x: dfs['college'][tokens[2]].get(safe2int(x),
                                                            np.nan))

    award_df.to_csv('foo.csv')
    return award_df
