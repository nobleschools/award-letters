#!python3
"""
Module for creating the Excel reports from gdoc and local data
"""

import numpy as np
import pandas as pd
import os
from datetime import datetime
from modules.filework import safe2int


# ----------------------------------------------------------------------
# Some helper functions for Excel writing
def safe_write(ws, r, c, val, f=None, n_a=""):
    """calls the write method of worksheet after first screening for NaN"""
    if not pd.isnull(val):
        if f:
            ws.write(r, c, val, f)
        else:
            ws.write(r, c, val)
    elif n_a:
        if f:
            ws.write(r, c, n_a, f)
        else:
            ws.write(r, c, n_a)


def write_array(ws, r, c, val, f=None):
    """speciality function to write an array. Assumed non-null"""
    if f:
        ws.write_formula(r, c, val, f)
    else:
        ws.write_formula(r, c, val)


def create_formats(wb, cfg_fmt, f_db={}):
    """Takes a workbook and (likely empty) database to fill with formats"""
    for name, db in cfg_fmt.items():
        f_db[name] = wb.add_format(db)

    return f_db


def make_excel_indices():
    """returns an array of Excel header columns from A through ZZ"""
    alphabet = string.ascii_uppercase
    master = list(alphabet)
    for i in range(len(alphabet)):
        master.extend([alphabet[i] + x for x in alphabet])
    return master


def _do_simple_sheet(writer, df, sheet_name, na_rep, index=True, f=None):
    """Helper function to write cells and bypass the Pandas write"""
    wb = writer.book
    ws = wb.add_worksheet(sheet_name)
    if index:
        safe_write(ws, 0, 0, df.index.name, f=f, n_a=na_rep)
    for col, label in enumerate(df.columns):
        safe_write(ws, 0, col + 1 * index, label, f=f, n_a=na_rep)
    
    row = 1
    for i, data in df.iterrows():
        if index:
            safe_write(ws, row, 0, i, f=f, n_a=na_rep)
        for col_num, col_name in enumerate(df.columns):
            safe_write(ws, row, col_num, data[col_name], f=f, n_a=na_rep)
        row += 1

    # This function is incomplete--doesn't currently write the data
    return (wb, ws, sheet_name, len(df) + 1)


def _do_initial_output(writer, df, sheet_name, na_rep, index=True):
    """Helper function to push data to xlsx and return formatting handles"""
    df.to_excel(writer, sheet_name=sheet_name, na_rep=na_rep, index=index)
    wb = writer.book
    ws = writer.sheets[sheet_name]
    max_row = len(df) + 1
    return (wb, ws, sheet_name, max_row)


def create_awards_tab(writer, df, format_db):
    """Adds the Awards tab to the output"""
    df.drop(columns=["Unique", "Award", "MoneyCode"], inplace=True)
    wb, ws, sn, max_row = _do_simple_sheet(writer, df, "AwardData", "", index=False)
    ws.set_column("A:B", 8, None, {"hidden": 1})
    ws.set_row(0, 75, format_db["p_header"])

    # Add the calculated columns:
    ws.write(0, 17, "Unique")
    ws.write(0, 18, "Award")

    for r in range(1, max_row):
        ws.write(r, 17, f'=IF(OR(A{r+1}<>A{r},B{r+1}<>B{r}),1,0)', format_db["centered_integer"])
        ws.write(r, 18, f'=IF(OR(AND(R{r+1}=1,ISNUMBER(M{r+1})),AND(R{r+1}=0,ISNUMBER(M{r+1}),M{r}=""),AND(R{r+1}=1,ISNUMBER(N{r+1})),AND(R{r+1}=0,ISNUMBER(N{r+1}),N{r}=""),AND(R{r+1}=1,ISNUMBER(O{r+1})),AND(R{r+1}=0,ISNUMBER(O{r+1}),O{r}="")),1,0)', format_db["centered_integer"])

    names = {
        "Students": "A",
        "NCESs": "B",
        "Names": "G",
        "Results": "H",
        "DataA": "K",
        "DataB": "L",
        "DataC": "M",
        "DataD": "N",
        "DataF": "O",
        "DataW": "P",
        "Unique": "R",
        "Award": "S",
    }
    for name, col in names.items():
        wb.define_name(
            name, "=" + sn + "!$" + col + "$2:$" + col + "$" + str(max_row)
        )

    max_col = max(names.values())
    ws.autofilter("A1:" + max_col + "1")
    ws.freeze_panes(1, 3)


def create_students_tab(writer, df, format_db, hide_campus=False):
    """Adds the Students tab to the output"""
    wb, ws, sn, max_row = _do_initial_output(writer, df, "Students", "N/A", index=False)
    
    # Add the calculated columns:
    ws.write(0, 12, "Acceptances", format_db["p_header"])
    ws.write(0, 13, "Unique Awards", format_db["p_header"])
    ws.write(0, 14, "% of awards collected", format_db["p_header"])
    ws.write(0, 15, "College Choice", format_db["p_header"])

    for r in range(1, max_row):
        ws.write(r, 12, f'=COUNTIFS(Students,B{r+1},Results,"Accepted!",Unique,1)+COUNTIFS(Students,B{r+1},Results,"Choice!",Unique,1)', format_db["centered_integer"])
        ws.write(r, 13, f'=COUNTIFS(Students,B{r+1},Award,1)', format_db["centered_integer"])
        ws.write(r, 14, f'=IF(M{r+1}>0,N{r+1}/M{r+1},0)', format_db["centered_integer"])
        ws.write(r, 15, 'TBD', format_db["centered_integer"])

    # format data columns
    ws.set_column("A:A", 9, format_db["left_normal_text"])  # , {"hidden", 1})
    ws.set_column("B:B", 9)
    ws.set_column("C:C", 34)
    ws.set_column("E:E", 9, format_db["single_percent_centered"])
    # ws.set_column("D:L", 9)

    ws.set_row(0, 60)
    names = {
        "SIDs": "B",
        "LastFirst": "C",
        "EFCs": "D",
        "MGRs": "E",
        "GPAs": "F",
        "SATs": "G",
        "Counselors": "H",
        "Advisors": "I",
        "CollegeChoice": "M",
    }
    for name, col in names.items():
        wb.define_name(
            name, "=" + sn + "!$" + col + "$2:$" + col + "$" + str(max_row)
        )

    max_col = max(names.values())
    ws.autofilter("A1:" + max_col + "1")
    ws.freeze_panes(1, 3)


def create_college_money_tab(writer, df, format_db):
    """Creates AllColleges from static file"""
    wb, ws, sn, max_row = _do_initial_output(writer, df, "CollegeMoney", "N/A")

    ws.set_column("D:E", 7, format_db["single_percent_centered"])
    ws.set_column("B:B", 40)
    ws.set_column("C:C", 22)
    ws.set_column("F:M", 7)
    names = {
        "AllCollegeNCES": "A",
        "AllCollegeMoneyCode": "H",
        "AllCollegeLocation": "M",
    }
    for name, col in names.items():
        wb.define_name(
            name, "=" + sn + "!$" + col + "$2:$" + col + "$" + str(max_row)
        )

    max_col = max(names.values())
    ws.autofilter("A1:" + max_col + "1")
    ws.hide()


# ----------------------------------------------------------------------------


def create_report_tables(dfs, campus, config, debug):
    # First, create a dataframe for the "Award data" tab
    dfs["award_report"] = build_award_df(dfs, campus, config, debug)

    # Second, create a dataframe for the "Students" tab
    # This one will have extra columns if the Decisions tab exists
    dfs["student_report"] = build_student_df(dfs, campus, config, debug)


def create_excel(dfs, campus, config, debug):
    """Will create Excel reports for sharing details from Google Docs"""
    if debug:
        print("Creating Excel report for {}".format(campus), flush=True)
    dfs["award_report"].to_csv("award_table_for_excel.csv", index=False)
    dfs["student_report"].to_csv("student_table_for_excel.csv", index=False)

    # Create the excel:
    date_string = datetime.now().strftime("%m_%d_%Y")
    fn = config["report_filename"].replace("CAMPUS", campus).replace("DATE", date_string)
    writer = pd.ExcelWriter(os.path.join(config["report_folder"],fn), engine="xlsxwriter")
    wb = writer.book
    formats = create_formats(wb, config["excel_formats"])

    # Award data tab
    create_awards_tab(writer, dfs["award_report"], formats)

    # Students tab
    create_students_tab(writer, dfs["student_report"], formats, hide_campus=(campus=="All"))
    # Hidden college lookup
    create_college_money_tab(writer, dfs["college"], formats)
    # Summary tab
    # OptionsReport (maybe don't create in Excel?)
    writer.save()


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
