#!python3
"""
Module for creating the Excel reports from gdoc and local data
"""

import numpy as np
import pandas as pd
from modules.filework import safe2int


def create_excel(dfs, campus, config, debug):
    """Will create Excel reports for sharing details from Google Docs"""
    if debug:
        print("Creating Excel report for {}".format(campus), flush=True)
    # First, create a dataframe for the "Award data" tab
    award_data_df = build_award_df(dfs, campus, config, debug)
    award_data_df.to_csv("award_table_for_excel.csv", index=False)

    # Second, create a dataframe for the "Students" tab
    # This one will have extra columns if the Decisions tab exists
    student_data_df = build_student_df(dfs, campus, config, debug)
    print(student_data_df)

    # Finally, actually create the excel
    # Initial document and hidden college lookup
    # Students tab
    # Award data tab
    # Summary tab
    # OptionsReport


def build_student_df(dfs, campus, config, debug):
    """Builds a dataframe for the student fields"""
    report_student_fields = config["report_student_fields"]
    return report_student_fields


def build_award_df(dfs, campus, config, debug):
    """Builds a dataframe for the award fields"""
    #  First, start the df for the items that are straight pulls from live_data
    report_award_fields = config["report_award_fields"]
    all_award_fields = []
    live_award_fields = []  # to hold the excel names
    live_award_targets = []  # to hold the live names
    complex_award_fields = []

    for column in report_award_fields:
        # Each column will be a dict with a single element
        # The key will be the Excel column name and the value the source
        # from the live table or other (lookup) table
        this_key = list(column.keys())[0]
        this_value = list(column.values())[0]
        all_award_fields.append(this_key)
        if ":" in this_value:
            complex_award_fields.append((this_key, this_value))
        else:
            live_award_fields.append(this_key)
            live_award_targets.append(this_value)
    if live_award_targets:  # fields here will be straight pulls from live df
        award_df = dfs["live_award"][live_award_targets]
        award_df = award_df.rename(
            columns=dict(zip(live_award_targets, live_award_fields))
        )
    else:
        print("Probably an error: no report columns pulling from live data")

    #  Second, pull columns that are lookups from other tables and append
    for column, target in complex_award_fields:
        # parse the target and then call the appropriate function
        # to add a column to award_df
        if debug:
            print(f"{column} w spec({target})")
        tokens = target.split(sep=":")
        if tokens[0] == "ROSTER":
            award_df[column] = dfs["live_award"][tokens[1]].apply(
                lambda x: dfs["ros"].loc[x, tokens[2]]
            )
        elif tokens[0] == "COLLEGE":
            award_df[column] = dfs["live_award"][tokens[1]].apply(
                lambda x: np.nan if pd.isnull(x) else
                dfs["college"][tokens[2]].get(safe2int(x), np.nan)
            )
        elif tokens[0] == "APPS":
            award_df[column] = award_df.apply(
                _do_app_field, args=(dfs["app"], tokens[1:]), axis=1)
        elif tokens[0] == "SPECIAL":
            award_df[column] = award_df.apply(
                _do_special_award, args=(column, tokens[1:]), axis=1)

    return award_df[[x for x in all_award_fields if not x.startswith("x")]]


def _do_app_field(award, app_df, keys):
    """Apply function to grab an application field given the NCES and SID"""
    nces = safe2int(award[keys[0]])
    sid = safe2int(award[keys[1]])
    matches = app_df[(app_df["NCES"] == nces) & (app_df["hs_student_id"] == sid)]
    if len(matches) >= 1:
        return matches.iloc[0][keys[-1]]
    else:
        return np.nan


def _do_special_award(award, column_name, args):
    """Apply function for special cases in award table. Custom for each column name"""
    if column_name == "Grad rate":
        response = award[args[1]] if (
            award["Race/Eth"] in ["H", "B", "M"]
            ) else award[args[0]]
        return "N/A" if pd.isnull(response) else response
    if column_name == "Grad rate for sorting":
        base_gr = award[args[0]]
        comments = award[args[1]]
        if base_gr == "N/A":
            return 0.0
        else:
            if comments == "Posse":
                return (base_gr+0.15) if base_gr < 0.7 else (1.0-(1.0-base_gr)/2)
            else:
                return base_gr
    # Catchall for errors:
    return '*'.join(args)+'+'+column_name
