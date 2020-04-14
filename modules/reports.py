#!python3
"""
Module for creating the Excel reports from gdoc and local data
"""

import numpy as np
import pandas as pd
from modules.filework import safe2int


def create_report_tables(dfs, campus, config, debug):
    # First, create a dataframe for the "Award data" tab
    dfs["award_report"] = build_award_df(dfs, campus, config, debug)
    dfs["award_report"].to_csv("award_table_for_excel.csv", index=False)

    # Second, create a dataframe for the "Students" tab
    # This one will have extra columns if the Decisions tab exists
    dfs["student_report"] = build_student_df(dfs, campus, config, debug)
    dfs["student_report"].to_csv("student_table_for_excel.csv", index=False)


def create_excel(dfs, campus, config, debug):
    """Will create Excel reports for sharing details from Google Docs"""
    if debug:
        print("Creating Excel report for {}".format(campus), flush=True)

    # Create the excel:
    # Initial document and hidden college lookup
    # Students tab
    # Award data tab
    # Summary tab
    # OptionsReport


def build_student_df(dfs, campus, config, debug):
    """Builds a dataframe for the student fields"""
    report_student_fields = config["report_student_fields"]
    report_student_sorts = config["report_student_sorts"]
    all_student_fields = []
    live_student_fields = []  # to hold the excel names
    live_student_targets = []  # to hold the live names
    complex_student_fields = []

    for column in report_student_fields:
        # Each column will be a dict with a single element
        # The key will be the Excel column name and the value the source
        # from the live (EFC) table or other (lookup) table
        this_key = list(column.keys())[0]
        this_value = list(column.values())[0]
        all_student_fields.append(this_key)
        if ":" in this_value:
            complex_student_fields.append((this_key, this_value))
        else:
            live_student_fields.append(this_key)
            live_student_targets.append(this_value)
    if live_student_targets:  # fields here will be straight pulls from live df
        """
        print(dfs["live_efc"].head())
        print(dfs["live_efc"].columns)
        print(len(dfs["live_efc"]))
        """
        # These 2 lines are necessary to handle single campus reports
        if "Campus" not in dfs["live_efc"].columns:
            dfs["live_efc"]["Campus"] = campus
        student_df = dfs["live_efc"][live_student_targets]
        student_df = student_df.rename(
            columns=dict(zip(live_student_targets, live_student_fields))
        )
    else:
        print("Probably an error: no report columns pulling from live data")

    #  Second, pull columns that are lookups from other tables and append
    #  We skip the "special" ones for now because they might calculate off lookups
    for column, target in (
        f for f in complex_student_fields if not f[1].startswith("SPECIAL")
    ):
        # parse the target and then call the appropriate function
        # to add a column to award_df
        if debug:
            print(f"{column} w spec({target})")
        tokens = target.split(sep=":")
        if tokens[0] == "INDEX":
            student_df[column] = dfs["live_efc"].index
        elif tokens[0] == "ROSTER":
            student_df[column] = dfs["live_efc"].index.map(
                lambda x: dfs["ros"].loc[x, tokens[1]]
            )

    for column, target in (
        f for f in complex_student_fields if f[1].startswith("SPECIAL")
    ):
        if debug:
            print(f"{column} w spec({target})")
        tokens = target.split(sep=":")
        student_df[column] = student_df.apply(
            _do_special_award, args=(column, tokens[1:]), axis=1
        )

    student_df = student_df[[x for x in all_student_fields if not x.startswith("x")]]
    # These generators work on a list of single pair dicts
    sort_terms = [list(item.keys())[0] for item in report_student_sorts]
    sort_order = [list(item.values())[0] for item in report_student_sorts]
    return student_df.sort_values(by=sort_terms, ascending=sort_order)


def build_award_df(dfs, campus, config, debug):
    """Builds a dataframe for the award fields"""
    #  First, start the df for the items that are straight pulls from live_data
    report_award_fields = config["report_award_fields"]
    report_award_sorts = config["report_award_sorts"]
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

    # Quick detour: make a calculated index for app table lookups:
    award_df["xAppIndex"] = (
        award_df["NCESid"].astype(str) + ":" + award_df["SID"].astype(str)
    )

    #  Second, pull columns that are lookups from other tables and append
    #  We skip the "special" ones for now because they might calculate off lookups
    for column, target in (
        f for f in complex_award_fields if not f[1].startswith("SPECIAL")
    ):
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
                lambda x: np.nan
                if pd.isnull(x)
                else dfs["college"][tokens[2]].get(safe2int(x), np.nan)
            )
        elif tokens[0] == "APPS":
            test_df = dfs["app"][["NCES", "hs_student_id", tokens[-1]]].copy()
            test_df["MergeIndex"] = (
                test_df.loc[:, "NCES"].astype(str) + ":" +
                test_df.loc[:, "hs_student_id"].astype(str)
            )
            test_df.set_index("MergeIndex", inplace=True)
            test_df.drop_duplicates(subset=["NCES", "hs_student_id"], inplace=True)
            award_df[column] = award_df["xAppIndex"].apply(
                lambda x: test_df[tokens[-1]].get(x, np.nan)
            )

    for column, target in (
        f for f in complex_award_fields if f[1].startswith("SPECIAL")
    ):
        if debug:
            print(f"{column} w spec({target})")
        tokens = target.split(sep=":")
        award_df[column] = award_df.apply(
            _do_special_award, args=(column, tokens[1:]), axis=1
        )

    award_df = award_df[[x for x in all_award_fields if not x.startswith("x")]]
    # These generators work on a list of single pair dicts
    sort_terms = [list(item.keys())[0] for item in report_award_sorts]
    sort_order = [list(item.values())[0] for item in report_award_sorts]
    return award_df.sort_values(by=sort_terms, ascending=sort_order)


def _do_special_award(award, column_name, args):
    """Apply function for special cases in award table. Custom for each column name"""
    if column_name == "Grad rate":
        overall_gr = args[0]
        aah_gr = args[1]
        response = (
            award[aah_gr]
            if (award["Race/Eth"] in ["H", "B", "M", "I"])
            else award[overall_gr]
        )
        return "N/A" if pd.isnull(response) else response

    elif column_name == "Grad rate for sorting":
        base_gr = award[args[0]]
        comments = award[args[1]]
        if base_gr == "N/A":
            return 0.0
        else:
            if comments == "Posse":
                return (
                    (base_gr + 0.15) if base_gr < 0.7 else (1.0 - (1.0 - base_gr) / 2)
                )
            else:
                return base_gr

    elif column_name == "Unique":
        return 1

    elif column_name == "Award":
        return 1

    # Catchall for errors:
    else:
        return "*".join(args) + "+" + column_name
