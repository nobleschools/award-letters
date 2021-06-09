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
def safe_write(ws, r, c, val, f=None, n_a="", make_float=False):
    """calls the write method of worksheet after first screening for NaN"""
    if not pd.isnull(val):
        if make_float:
            try:
                val = float(val)
            except:
                pass
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
    import string  # We're not currently using this function, so leaving import here so as not to forget
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
            safe_write(ws, row, 0, i, f=None, n_a=na_rep)
        for col_num, col_name in enumerate(df.columns):
            safe_write(ws, row, col_num, data[col_name], f=None, n_a=na_rep)
        row += 1

    return (wb, ws, sheet_name, len(df) + 1)


def _do_initial_output(writer, df, sheet_name, na_rep, index=True):
    """Helper function to push data to xlsx and return formatting handles"""
    df.to_excel(writer, sheet_name=sheet_name, na_rep=na_rep, index=index)
    wb = writer.book
    ws = writer.sheets[sheet_name]
    max_row = len(df) + 1
    return (wb, ws, sheet_name, max_row)


def create_summary_tab(writer, config, format_db, do_campus):
    """Adds the Summary tab to the output"""
    wb = writer.book
    ws = wb.add_worksheet("Summary")
    for c, column in enumerate(config["columns"]):
        for label, fmt in column.items():
            ws.write(0, c, label, format_db[fmt])

    # Summarizes by campus if this is all campuses, otherwise by strategy
    row_labels = config["campuses"] if do_campus else config["strats"]
    s_name = "Campus" if do_campus else "Strats"
    for r, label in enumerate(row_labels, start=1):
        rx = r + 1  # Excel reference is 1-indexed
        ws.write(r, 0, label)  # field to summarize by
        ws.write(r, 1, f"=COUNTIF({s_name},A{rx})")  # student column
        ws.write(r, 2, f'=IF(B{rx}>0,SUMIF({s_name},A{rx},MGRs)/B{rx},"")')  # TGR
        ws.write(r, 3, f'=IF(A{rx}>0,SUMIFS(Schol4YR,{s_name},$A{rx}),"")')  # Total 4yr
        ws.write(r, 4, f'=IF(B{rx}>0,D{rx}/B{rx},"")')  # Avg 4yr
        ws.write(r, 5, f'=IF(B{rx}>0,COUNTIFS(PGR,"<>TBD",{s_name},$A{rx})/$B{rx},"")')  # % decided
        ws.write(  # % of awards collected
            r,
            6,
            f'=IF(AND(B{rx}>0,SUMIFS(Accepts,{s_name},$A{rx})),SUMIFS(UAwards,{s_name},$A{rx})/SUMIFS(Accepts,{s_name},$A{rx}),"")',
        )
        ws.write(  # PGR
            r,
            7,
            f'=IF(AND(B{rx}>0,F{rx}>0),SUMIF({s_name},A{rx},PGR)/COUNTIFS(PGR,"<>TBD",{s_name},$A{rx}),"")',
        )
        ws.write(  # PGR-TGR
            r,
            8,
            f'=IF(AND(B{rx}>0,F{rx}>0),SUMIF({s_name},A{rx},PGRTGR)/COUNTIFS(PGR,"<>TBD",{s_name},$A{rx}),"")',
        )
        ws.write(  # % of students w/in 10% of TGR
            r,
            9,
            f'=IF(AND(B{rx}>0,F{rx}>0),COUNTIFS({s_name},A{rx},PGRin10,"Yes")/COUNTIFS(PGRin10,"<>TBD",{s_name},$A{rx}),"")',
        )
        ws.write(  # % of students w/ award at choice
            r,
            10,
            f'=IF(AND(B{rx}>0,F{rx}>0),COUNTIFS({s_name},A{rx}, OOP,"<>TBD")/COUNTIFS(PGR,"<>TBD",{s_name},$A{rx}),"")',
        )
        ws.write(  # Avg unmet need at choice college
            r,
            11,
            f'=IF(COUNTIFS({s_name},A{rx}, UMN,"<>TBD")>0,SUMIFS(UMN,{s_name},A{rx}, UMN,"<>TBD")/COUNTIFS({s_name},A{rx}, UMN,"<>TBD"),"")',
        )
    
    # Summary row
    fr = 2
    lr = len(row_labels) + 1  # This is a little tricky--it's the write location and last value row to sum
    lrx = lr + 1
    ws.write(lr, 1, f'=SUM(B{fr}:B{lr})', format_db["sum_centered_integer"])
    ws.write(lr, 2, f'=SUMPRODUCT(B{fr}:B{lr},C{fr}:C{lr})/B{lrx}', format_db["sum_percent"])
    ws.write(lr, 3, f'=SUM(D{fr}:D{lr})', format_db["sum_dollar"])
    ws.write(lr, 4, f'=IF(B{lrx}>0,D{lrx}/B{lrx},"")', format_db["sum_dollar"])
    ws.write(lr, 5, f'=SUMPRODUCT(B{fr}:B{lr},F{fr}:F{lr})/B{lrx}', format_db["sum_percent"])  # % decided
    ws.write(lr, 6, f'=SUM(UAwards)/SUM(Accepts)', format_db["sum_percent"])
    ws.write(lr, 7, f'=SUMPRODUCT(B{fr}:B{lr},H{fr}:H{lr})/B{lrx}', format_db["sum_percent"])
    ws.write(lr, 8, f'=SUMPRODUCT(B{fr}:B{lr},I{fr}:I{lr})/B{lrx}', format_db["sum_percent"])
    ws.write(lr, 9, f'=SUMPRODUCT(B{fr}:B{lr},J{fr}:J{lr})/B{lrx}', format_db["sum_percent"])
    ws.write(lr,10, f'=SUMPRODUCT(B{fr}:B{lr},K{fr}:K{lr})/B{lrx}', format_db["sum_percent"])
    ws.write(lr,11, f'=SUMPRODUCT(B{fr}:B{lr},L{fr}:L{lr})/B{lrx}', format_db["sum_dollar"])

    # Final formatting
    ws.set_column("A:A", 8.09, format_db["left_normal_text"]) 
    ws.set_column("B:B", 8.09, format_db["centered_integer"])
    ws.set_column("C:C", 9.55, format_db["single_percent_centered"])
    ws.set_column("D:D", 13.73, format_db["dollar_no_cents_fmt"])
    ws.set_column("E:E", 12.73, format_db["dollar_no_cents_fmt"])
    ws.set_column("F:F", 9.73, format_db["single_percent_centered"])
    ws.set_column("G:G", 8.09, format_db["single_percent_centered"])
    ws.set_column("H:I", 6.36, format_db["single_percent_centered"])
    ws.set_column("J:K", 8.09, format_db["single_percent_centered"])
    ws.set_column("L:L", 10.91, format_db["dollar_no_cents_fmt"])

    ws.activate()


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
        ws.write(
            r,
            17,
            f"=IF(OR(A{r+1}<>A{r},B{r+1}<>B{r}),1,0)",
            format_db["centered_integer"],
        )
        ws.write(
            r,
            18,
            f'=IF(OR(AND(R{r+1}=1,ISNUMBER(M{r+1})),AND(R{r+1}=0,ISNUMBER(M{r+1}),M{r}=""),AND(R{r+1}=1,ISNUMBER(N{r+1})),AND(R{r+1}=0,ISNUMBER(N{r+1}),N{r}=""),AND(R{r+1}=1,ISNUMBER(O{r+1})),AND(R{r+1}=0,ISNUMBER(O{r+1}),O{r}="")),1,0)',
            format_db["centered_integer"],
        )

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
        wb.define_name(name, "=" + sn + "!$" + col + "$2:$" + col + "$" + str(max_row))

    max_col = max(names.values())
    ws.autofilter("A1:" + max_col + "1")
    ws.freeze_panes(1, 3)


def create_students_tab(writer, df, format_db, hide_campus=False):
    """Adds the Students tab to the output"""
    # wb, ws, sn, max_row = _do_initial_output(writer, df, "Students", "N/A", index=False)
    wb, ws, sn, max_row = _do_simple_sheet(
        writer, df.iloc[:, :12], "Students", "N/A", index=False, f=format_db["p_header"]
    )

    # Add the calculated columns:
    ws.write(0, 12, "Acceptances", format_db["p_header_y"])
    ws.write(0, 13, "Unique Awards", format_db["p_header_y"])
    ws.write(0, 14, "% of awards collected", format_db["p_header_y"])
    ws.write(0, 15, "Total grants & scholarships (1 yr value)", format_db["p_header_y"])
    ws.write(0, 16, "Total grants & scholarships (4 yr value)", format_db["p_header_y"])
    ws.write(0, 17, "College Choice", format_db["p_header_o"])
    ws.write(0, 18, "Ambitious Postsecondary Pathway choice", format_db["p_header_o"])
    ws.write(0, 19, "Other College Choice", format_db["p_header_o"])
    ws.write(0, 20, "PGR for choice school", format_db["p_header_y"])
    ws.write(0, 21, "PGR-TGR", format_db["p_header_y"])
    ws.write(0, 22, "PGR within 10% of TGR?", format_db["p_header_y"])
    ws.write(0, 23, "Reason for not meeting TGR", format_db["p_header_o"])
    ws.write(0, 24, "Out of Pocket at Choice", format_db["p_header_o"])
    ws.write(0, 25, "Unmet need", format_db["p_header_o"])
    ws.write(
        0, 26, "Exceeds Goal? (no more than 3000 over EFC)", format_db["p_header_o"]
    )
    ws.write(
        0,
        27,
        "Comments (use for undermatching and affordability concerns)",
        format_db["p_header_o"],
    )

    for r in range(1, max_row):
        ws.write(
            r,
            12,
            f'=COUNTIFS(Students,B{r+1},Results,"Accepted!",Unique,1)+COUNTIFS(Students,B{r+1},Results,"Choice!",Unique,1)',
            format_db["centered_integer"],
        )
        ws.write(
            r, 13, f"=COUNTIFS(Students,B{r+1},Award,1)", format_db["centered_integer"]
        )
        ws.write(
            r,
            14,
            f"=IF(M{r+1}>0,N{r+1}/M{r+1},0)",
            format_db["single_percent_centered"],
        )
        ws.write(
            r, 15, f"=SUMIFS(DataC,Students,B{r+1},Award,1)", format_db["dollar_fmt"]
        )
        ws.write(r, 16, f"=P{r+1}*4", format_db["dollar_fmt"])
        safe_write(ws, r, 17, df["College Choice"].iloc[r - 1])
        safe_write(ws, r, 18, df["Ambitious Postsecondary Pathway choice"].iloc[r - 1])
        safe_write(ws, r, 19, df["Other College Choice"].iloc[r - 1])
        safe_write(
            ws,
            r,
            20,
            df["PGR for choice school"].iloc[r - 1],
            n_a="TBD",
            f=format_db["single_percent_centered"],
            make_float=True,
        )
        safe_write(
            ws,
            r,
            21,
            df["PGR-TGR"].iloc[r - 1],
            n_a="TBD",
            f=format_db["single_percent_centered"],
            make_float=True,
        )
        safe_write(
            ws,
            r,
            22,
            df["PGR within 10% of TGR?"].iloc[r - 1],
            n_a="TBD",
            f=format_db["centered"],
        )
        safe_write(ws, r, 23, df["Reason for not meeting TGR"].iloc[r - 1])
        safe_write(
            ws,
            r,
            24,
            df["Out of Pocket at Choice (pulls from Award data tab weekly)"].iloc[
                r - 1
            ],
            n_a="TBD",
            f=format_db["dollar_no_cents_fmt"],
            make_float=True,
        )
        safe_write(
            ws,
            r,
            25,
            f'=IF(AND(ISNUMBER(Y{r+1}),ISNUMBER(D{r+1})),MAX(Y{r+1}-D{r+1},0),"TBD")',
            n_a="TBD",
        )
        safe_write(
            ws,
            r,
            26,
            df["Exceeds Goal? (no more than 3000 over EFC)"].iloc[r - 1],
            n_a="TBD",
        )
        safe_write(
            ws,
            r,
            27,
            df["Comments (use for undermatching and affordability concerns)"].iloc[
                r - 1
            ],
        )

    # format data columns
    ws.set_column("A:A", 9, format_db["left_normal_text"])  # , {"hidden", 1})
    ws.set_column("B:B", 9)
    ws.set_column("C:C", 34)
    ws.set_column("E:E", 9, format_db["single_percent_centered"])
    # ws.set_column("D:L", 9)
    ws.set_column("P:Q", 13)
    ws.set_column("R:R", 35)
    ws.set_column("S:T", 22)
    ws.set_column("U:U", 9)
    ws.set_column("V:V", 7)
    ws.set_column("W:W", 10)
    ws.set_column("X:X", 23)
    ws.set_column("Y:Y", 9)
    ws.set_column("Z:Z", 8)
    ws.set_column("AA:AA", 14)
    ws.set_column("AB:AB", 33)

    ws.set_row(0, 60)
    names = {
        "Campus": "A",
        "SIDs": "B",
        "LastFirst": "C",
        "EFCs": "D",
        "MGRs": "E",
        "GPAs": "F",
        "SATs": "G",
        "Counselors": "H",
        "Advisors": "I",
        "Strats": "J",
        "Accepts": "M",
        "UAwards": "N",
        "Schol4Yr": "Q",
        "CollegeChoice": "R",
        "PGR": "U",
        "PGRTGR": "V",
        "PGRin10": "W",
        "OOP": "Y",
        "UMN": "Z",
        "Affordable": "AA",
    }
    for name, col in names.items():
        wb.define_name(name, "=" + sn + "!$" + col + "$2:$" + col + "$" + str(max_row))

    ws.autofilter("A1:AB" + "1")
    ws.freeze_panes(1, 3)


def create_college_money_tab(writer, df, format_db):
    """Creates AllColleges from static file"""
    wb, ws, sn, max_row = _do_initial_output(writer, df, "CollegeMoneyData", "N/A")

    ws.set_column("D:E", 7, format_db["single_percent_centered"])
    ws.set_column("B:B", 40)
    ws.set_column("C:C", 22)
    ws.set_column("F:L", 7)
    names = {
        "AllCollegeNCES": "A",
        "AllCollegeMoneyCode": "H",
        "AllCollegeLocation": "M",
    }
    for name, col in names.items():
        wb.define_name(name, "=" + sn + "!$" + col + "$2:$" + col + "$" + str(max_row))

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
    fn = (
        config["report_filename"].replace("CAMPUS", campus).replace("DATE", date_string)
    )
    writer = pd.ExcelWriter(
        os.path.join(config["report_folder"], fn), engine="xlsxwriter"
    )
    wb = writer.book
    formats = create_formats(wb, config["excel_formats"])

    # Award data tab
    create_awards_tab(writer, dfs["award_report"], formats)

    # Students tab
    create_students_tab(
        writer, dfs["student_report"], formats, hide_campus=(campus == "All")
    )

    # Summary tab
    create_summary_tab(
        writer, config["summary_settings"], formats, do_campus=(campus == "All")
    )

    # Hidden college lookup
    create_college_money_tab(writer, dfs["college"], formats)

    # OptionsReport (maybe don't create in Excel?)
    writer.save()


def build_student_df(dfs, campus, config, debug):
    """Builds a dataframe for the student fields"""
    report_student_fields = config["report_student_fields"]
    report_student_sorts = config["report_student_sorts"]
    all_student_fields = []
    live_student_fields = []  # to hold the excel names
    live_student_targets = []  # to hold the live names
    # live_decision_fields = []  # to hold excel names for decision tabl
    # live_decision_targets = []  # to hold the live names
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
        # if debug:
        #     print(f"{column} w spec({target})")
        tokens = target.split(sep=":")
        if tokens[0] == "INDEX":
            student_df[column] = dfs["live_efc"].index
        elif tokens[0] == "ROSTER":
            student_df[column] = dfs["live_efc"].index.map(
                lambda x: dfs["ros"].loc[x, tokens[1]]
            )
        elif tokens[0] == "DECISION":
            if "live_decision" in dfs:
                student_df[column] = dfs["live_efc"].index.map(
                    lambda x: dfs["live_decision"].loc[x, tokens[1]]
                )

    for column, target in (
        f for f in complex_student_fields if f[1].startswith("SPECIAL")
    ):
        # if debug:
        #     print(f"{column} w spec({target})")
        tokens = target.split(sep=":")
        student_df[column] = student_df.apply(
            _do_special_award, args=(column, tokens[1:]), axis=1
        )

    student_df = student_df[[x for x in all_student_fields if not x.startswith("x")]]
    # These generators work on a list of single pair dicts
    sort_terms = [list(item.keys())[0] for item in report_student_sorts]
    sort_order = [list(item.values())[0] for item in report_student_sorts]
    # recast the EFC as numbers where possible:
    student_df.EFC = student_df.EFC.apply(lambda x: pd.to_numeric(x, errors="ignore"))
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
                test_df.loc[:, "NCES"].astype(str)
                + ":"
                + test_df.loc[:, "hs_student_id"].astype(str)
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
