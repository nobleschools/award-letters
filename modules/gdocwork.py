#!python3
"""
Module for doing direct reads and writes from the google docs
"""

from time import time
import pandas as pd
import numpy as np

from modules import googleapi
from modules import filework

MAX_ROWS_ADD = 60


def safefloat(x):
    """Converts to a float if possible"""
    try:
        return float(x)
    except BaseException:
        return x


def safeint(x):
    """Converts to an integer if possible"""
    try:
        return int(x)
    except BaseException:
        return x


def _get_pgr(x, roster_df, college_df, bump_list_df):
    """Apply function to get the PGR for a college after determining student
    race; gives the "15%" bump for any sid/nces pair in the bump_list df
    """
    sid, nces = x
    nces = safeint(nces)
    race = roster_df.loc[sid, "Race/ Eth"]
    field = "Adj6yrGrad_All" if race in ["W", "A"] else "Adj6yrGrad_AA_Hisp"
    if not pd.isnull(nces):
        raw_pgr = college_df[field].get(nces, "TBD")
    else:
        raw_pgr = 0.0
    if sid in bump_list_df["SID"].values:
        student_bumps = bump_list_df[bump_list_df["SID"] == sid]
        if nces in student_bumps["NCESid"].values:
            return (raw_pgr + 0.15) if (raw_pgr <= 0.7) else ((raw_pgr + 1.0) / 2)
        else:
            return raw_pgr
    else:
        return raw_pgr


def _write_df_to_sheet(
    ws, df, key, title, na_val="", resize=True, use_index=False, use_apps_script=False
):
    """Takes a dataframe and writes to a google sheet"""
    n_rows = len(df) + 1
    n_cols = len(df.columns) + (1 if use_index else 0)
    if resize:
        ws.resize(rows=n_rows, cols=n_cols)

    # Turn data into list of lists for writing
    l_o_l = [([use_index] if use_index else []) + df.columns.tolist()] + (  # header
        df.reset_index() if use_index else df
    ).values.tolist()  # rows

    # Replace the n/a's
    l_o_l = [[na_val if pd.isnull(x) else x for x in row] for row in l_o_l]

    # Option to either use Apps Script or gspread to write the data
    if use_apps_script:
        googleapi.call_script_service(
            {"function": "writeDataTable", "parameters": [key, title, l_o_l]}
        )
    else:
        # Create a write_range for the gspread library and then pair with data
        write_range = ws.range(1, 1, n_rows, n_cols)

        # Serialize the list of lists to match the write range
        flat_data = [item for row in l_o_l for item in row]

        # Now determine which fields not to write and pop the unwritten cells
        for i in range(len(flat_data) - 1, -1, -1):
            if flat_data[i] == "":
                write_range.pop(i)
            else:
                write_range[i].value = flat_data[i]

        ws.update_cells(write_range, value_input_option="USER_ENTERED")


def read_current_doc(dfs, campus, config, debug):
    """
    Does a simple read of the two main tables and saves them as dfs.
    If the third table (Decisions) is there, it's read as well.
    """
    doc_key = dfs["key"].loc[campus, "ss_key"]

    if debug:
        print("About to read doc for {}...".format(campus), flush=True)

    sheets = ["efc", "award", "decision"]
    for sheet in sheets:
        t0 = time()

        # First do the read
        raw_data = googleapi.call_script_service(
            {
                "function": "readDataTable",
                "parameters": [doc_key, config[sheet + "_tab_name"]],
            }
        )
        if debug:
            print(
                "--{} read completed in {:.2f} seconds".format(sheet, time() - t0),
                flush=True,
            )
        if raw_data[0][0] == "NULL":
            if debug:
                print("--" + sheet + " tab has no data")
            continue

        # Then convert to DataFrame inside the df dict
        header_row_ix = int(config[sheet + "_header_row"])
        live_df = "live_" + sheet
        dfs[live_df] = pd.DataFrame(
            raw_data[header_row_ix:], columns=raw_data[(header_row_ix - 1)]
        )
        if sheet in ["efc", "decision"]:
            dfs[live_df].set_index("StudentID", inplace=True)


def _do_table_diff(current_index_set, new_index_set):
    """Utility function to perform a couple of set operations"""
    indices_to_insert = new_index_set - current_index_set
    indices_to_delete = current_index_set - new_index_set
    return (indices_to_insert, indices_to_delete)


def _do_table_diff_df(current_data, new_data, debug):
    """
    Utility function to perform similar set operations on 3 column dfs.
    Here, the 3 columns are intended to be StudentID, NCESid, and Home/Away
    """
    # First, flag any live rows with missing data
    missing_index = ((current_data.isnull()) | (current_data == "")).apply(
        sum, axis=1
    ) > 0
    current_data_clean = current_data[~missing_index]
    if debug:
        print("There are {} rows with missing indices".format(missing_index.sum()))
        print(
            "{} rows in live_data, {} after removing missing indices".format(
                len(current_data), len(current_data_clean)
            )
        )

    # Look for rows not present in both tables
    current_tuples = [
        (x[1], safeint(x[2]), x[3]) for x in current_data_clean.itertuples()
    ]
    new_tuples = [x[1:4] for x in new_data.itertuples()]

    indices_to_insert = list(set(new_tuples) - set(current_tuples))
    indices_to_delete = list(set(current_tuples) - set(new_tuples))

    # Build a record of rows present in both tables
    joint_tuples = list(set(current_tuples) & set(new_tuples))

    # Then find the conflict in app results and save the "new" values to push
    joint_current = [
        x[1:] for x in current_data_clean.itertuples() if x[1:4] in joint_tuples
    ]
    joint_new = [x[1:] for x in new_data.itertuples() if x[1:4] in joint_tuples]
    result_changes = list(set(joint_new) - set(joint_current))

    return (indices_to_insert, indices_to_delete, result_changes)


def _match_to_tuple_index(x, tuple_list):
    """Apply function to see if passed fields are in the tuple_list"""
    sid, ncesid, home_away = x
    return (sid, ncesid, home_away) in tuple_list


def _calculate_6000_out_of_pocket(x):
    """Apply function to increase out_of_pocket if loans are > 6,000"""
    loans, out_of_pocket = x
    out_of_pocket = safefloat(out_of_pocket)
    if isinstance(loans, float) and isinstance(out_of_pocket, float):
        if loans > 6000.0:
            return out_of_pocket + (loans - 6000.0)
    return out_of_pocket


def refresh_decisions(dfs, campus, config, debug):
    """
    Works with the two decisions tabs specifically to make sure they're
    updated from the Award data and EFC tabs. Creates the two tabs if they do
    not exist.
    """
    # Set local config variables
    decision_options_sheet_title = config["decision_options_tab_name"]
    decision_sheet_title = config["decision_tab_name"]
    decision_options_header_row = config["decision_options_header_row"]
    decision_header_row = int(config["decision_header_row"])
    decision_defaults = config["decision_defaults"]
    do_fds = config["decision_option_fields"]

    # #################################################################
    #  First, create starter tables for both the decision options tab and
    #  the decisions tab based on the order of info in efc tab

    # First, pair down the tables to just the columns we need
    # And add any lookup values (PGR, TGR) from local tables
    sid, nces, home, college, result_code, out_of_pocket, s_loans, cgs = do_fds
    a_df = dfs["live_award"][do_fds]
    a_df = a_df[a_df[result_code] != "Denied"].sort_values([sid, college])
    a_df[s_loans] = a_df[s_loans].fillna(0.0)
    a_df[out_of_pocket] = a_df[out_of_pocket].fillna("TBD")
    a_df["out_of_pocket6000"] = a_df[[s_loans, out_of_pocket]].apply(
        _calculate_6000_out_of_pocket, axis=1
    )
    a_df["PGR"] = a_df[[sid, nces]].apply(
        _get_pgr, args=(dfs["ros"], dfs["college"], dfs["bump_list"]), axis=1
    )
    a_df["PGR"] = a_df["PGR"].fillna("N/A")
    a_df[cgs] = a_df[cgs].fillna("N/A")
    a_df[result_code] = a_df[result_code].fillna("TBD")
    a_df = a_df[[sid, college, result_code, "PGR", "out_of_pocket6000", cgs]]

    s_df = dfs["live_efc"].copy()
    s_df["Student TGR"] = s_df.index.map(
        lambda x: dfs["ros"].loc[x, "Target Grad Rate"]
    )
    s_df = s_df[["LastFirst", "Student TGR"]]
    s_df["Student TGR"] = s_df["Student TGR"].fillna("TBD")

    # a_df.to_csv('foo_award.csv')
    # s_df.to_csv('foo_s.csv')

    # Second, create lists of lists for the actual tables
    do_table = [["student", "college", "Result", "pgr", "out_of_pocket", "cgs"]]
    d_table = [["SID", "LastFirst", "SR", "ER", "Choice", "Student TGR"]]
    app_table = [["ProgramName", "Index"]]  # This little table is for APP choices
    for index, row in dfs["ambitious_pp"].iterrows():
        app_table.append([index, list(row)[0]])
    app_table.insert(1, ["", "N/A"])
    current_row = 1 + decision_options_header_row  # index of choice table
    for index, row in s_df.iterrows():
        # First get the a_df records that match the index on sid
        last_first, student_tgr = list(row)
        these_options = a_df[a_df[sid] == index]

        # Second determine if any of them are UNIQUELY CHOICE!
        this_choice = ""
        if len(these_options):
            choice_options = these_options[these_options[result_code] == "CHOICE!"]
            if len(choice_options) == 1:
                this_choice = choice_options.iloc[0, 1]
            # If there are two "CHOICE!" schools (Home/Away),
            # Pick the Home one
            elif len(choice_options) == 2:
                this_choice = choice_options.iloc[0, 1]
                if this_choice.endswith("Campus"):
                    this_choice = choice_options.iloc[1, 1]

        # Create do rows: first blank, second all options, third standard
        do_table.append([index, "", "N/A", "TBD", "TBD", 0.0])

        if len(these_options):
            for ignore, option in these_options.iterrows():
                do_table.append(list(option))

        for label, pgr in decision_defaults.items():
            do_table.append([index, label, "N/A", pgr, 0.0, 0.0])

        # Create d row using the count from above
        num_rows = 1 + len(these_options) + len(decision_defaults)
        d_table.append(
            [
                index,
                last_first,
                current_row,
                current_row + num_rows - 1,
                this_choice,
                student_tgr,
            ]
        )
        current_row += num_rows  # ready for the next student

    filework.save_csv_from_table("temp_do.csv", ".", do_table)
    filework.save_csv_from_table("temp_d.csv", ".", d_table)

    ###################################################################
    #  Second, push the starter tables to the doc where the AppsScript
    #  will handle updating things in the decisons_options and decisions tabs
    doc_key = dfs["key"].loc[campus, "ss_key"]

    if debug:
        print(
            "DecisionOptions tab, pushing {} rows...".format(len(do_table)),
            end="",
            flush=True,
        )
    t0 = time()
    googleapi.call_script_service(
        {
            "function": "refreshDecisionOptions",
            "parameters": [doc_key, decision_options_sheet_title, do_table, app_table],
        }
    )
    if debug:
        print("done in {:.2f} seconds".format(time() - t0), flush=True)

    if debug:
        print(
            "Decisions tab, pushing {} rows...".format(len(d_table)), end="", flush=True
        )
    t0 = time()
    googleapi.call_script_service(
        {
            "function": "refreshDecisions",
            "parameters": [
                doc_key,
                decision_sheet_title,
                decision_options_sheet_title,
                d_table,
                decision_header_row,
            ],
        }
    )
    if debug:
        print("done in {:.2f} seconds".format(time() - t0), flush=True)


def sync_doc_rows(dfs, campus, config, debug):
    """
    Does all of the syncing work between the live tabs (from current Docs)
    and the new tabs (created fresh from latest Naviance downloads)
    Creates two sets of "orders" for the Apps Script functions to insert
    or delete specific rows
    """
    # Name local dfs:
    key_df = dfs["key"]
    live_award_df = dfs["live_award"]
    live_efc_df = dfs["live_efc"]
    new_award_df = dfs["award"]
    new_efc_df = dfs["efc"]

    # Set local config variables
    efc_sheet_title = config["efc_tab_name"]
    award_sheet_title = config["award_tab_name"]
    efc_header_row = config["efc_header_row"]
    award_header_row = config["award_header_row"]

    # First the EFC tab
    #  Make a comparison of new rows and rows to delete
    efc_indices_to_insert, efc_indices_to_delete = _do_table_diff(
        set(live_efc_df.index), set(new_efc_df.index)
    )

    # Get parameters for working with the Google Doc
    doc_key = key_df.loc[campus, "ss_key"]

    #  Push the new rows to the doc
    if efc_indices_to_insert:
        # Get the full rows of data to add
        efc_to_add_df = new_efc_df[new_efc_df.index.isin(efc_indices_to_insert)]

        # Now convert it to a list of headers and a list of lists for data
        efc_to_add_header = list(efc_to_add_df.columns)
        efc_to_add_header.insert(0, efc_to_add_df.index.name)
        efc_list_of_list_data = []
        for index, row in efc_to_add_df.iterrows():
            this_row = list(row)
            this_row.insert(0, index)
            efc_list_of_list_data.append(this_row)
        # Now replace the n_as:
        efc_list_of_list_data = [
            ["" if pd.isnull(x) else x for x in row] for row in efc_list_of_list_data
        ]

        if debug:
            print(
                "EFC tab, adding {} rows...".format(len(efc_list_of_list_data)),
                end="",
                flush=True,
            )
        t0 = time()
        a_response = googleapi.call_script_service(
            {
                "function": "insertEFCStudentRows",
                "parameters": [
                    doc_key,
                    efc_sheet_title,
                    "LastFirst",
                    efc_to_add_header,
                    efc_list_of_list_data,
                    efc_header_row,
                ],
            }
        )
        if debug:
            print("done in {:.2f} seconds".format(time() - t0), flush=True)

    #  Delete the rows for removal
    if efc_indices_to_delete:
        if debug:
            print(
                "EFC tab, deleting {} rows...".format(len(efc_indices_to_delete)),
                end="",
                flush=True,
            )
        t0 = time()
        d_response = googleapi.call_script_service(
            {
                "function": "deleteEFCStudentRows",
                "parameters": [
                    doc_key,
                    efc_sheet_title,
                    "StudentID",
                    list(efc_indices_to_delete),
                ],
            }
        )
        if debug:
            print("done in {:.2f} seconds".format(time() - t0), flush=True)
            # print(d_response, flush=True)

    # Second the Award tab
    #  Make a comparison of new rows and rows to delete
    # figure out the number of error indices in live data:
    # (The two lines below fix the problem of propagating N/As)
    new_award_df["NCESid"].replace(np.nan, "N/A", inplace=True)
    live_award_df["NCESid"].replace(np.nan, "N/A", inplace=True)

    award_ix_to_insert, award_ix_to_delete, result_changes = _do_table_diff_df(
        live_award_df[["SID", "NCESid", "Home/Away", "Result (from Naviance)"]],
        new_award_df[["SID", "NCESid", "Home/Away", "Result (from Naviance)"]],
        debug,
    )

    #  Push the new rows to the doc
    if award_ix_to_insert:
        # Get the full rows of data to add
        award_to_add_df = new_award_df[
            new_award_df[["SID", "NCESid", "Home/Away"]].apply(
                _match_to_tuple_index, axis=1, args=(award_ix_to_insert,)
            )
        ]

        # Now convert it to a list of headers and a list of lists for data
        award_to_add_header = list(award_to_add_df.columns)
        award_list_of_list_data = []
        for index, row in award_to_add_df.iterrows():
            award_list_of_list_data.append(list(row))

        # Now replace the n_as:
        award_list_of_list_data = [
            ["" if pd.isnull(x) else x for x in row] for row in award_list_of_list_data
        ]

        if debug:
            print(
                "Awards tab, adding {} rows...".format(len(award_list_of_list_data)),
                end="",
                flush=True,
            )
        if len(award_list_of_list_data) <= MAX_ROWS_ADD:
            t0 = time()
            a_response = googleapi.call_script_service(
                {
                    "function": "insertAwardStudentRows",
                    "parameters": [
                        doc_key,
                        award_sheet_title,
                        result_changes,
                        award_to_add_header,
                        award_list_of_list_data,
                        award_header_row,
                    ],
                }
            )
            if debug:
                print("done in {:.2f} seconds".format(time() - t0), flush=True)
                # print(a_response, flush=True)
        else:  # This data set is too large, so we're going to push twice
            alold1 = award_list_of_list_data[:MAX_ROWS_ADD]
            alold2 = award_list_of_list_data[MAX_ROWS_ADD:]
            for alold in [alold1, alold2]:
                t0 = time()
                a_response = googleapi.call_script_service(
                    {
                        "function": "insertAwardStudentRows",
                        "parameters": [
                            doc_key,
                            award_sheet_title,
                            result_changes,
                            award_to_add_header,
                            alold,
                            award_header_row,
                        ],
                    }
                )
                if debug:
                    print(
                        "done ({}) in {:.2f} seconds..".format(len(alold), time() - t0),
                        flush=True,
                        end="",
                    )
            if debug:
                print("", flush=True)

    elif result_changes:
        if debug:
            print(
                "Awards tab, changing decision on {} rows...".format(
                    len(result_changes)
                ),
                end="",
                flush=True,
            )
        t0 = time()
        a_response = googleapi.call_script_service(
            {
                "function": "updateAwardStatuses",
                "parameters": [
                    doc_key,
                    award_sheet_title,
                    result_changes,
                    award_header_row,
                ],
            }
        )
        if debug:
            print("done in {:.2f} seconds".format(time() - t0), flush=True)
            print("Total of {} actual changes".format(int(a_response)))

    #  Delete the rows for removal
    if award_ix_to_delete:
        if debug:
            print(
                "Awards tab, deleting {} rows...".format(len(award_ix_to_delete)),
                end="",
                flush=True,
            )
        t0 = time()
        d_response = googleapi.call_script_service(
            {
                "function": "deleteAwardStudentRows",
                "parameters": [
                    doc_key,
                    award_sheet_title,
                    award_header_row,
                    award_ix_to_delete,
                ],
            }
        )
        if debug:
            print(
                "done ({} deleted) in {:.2f} seconds".format(
                    (len(d_response) if isinstance(d_response, list) else "?"),
                    time() - t0,
                ),
                flush=True,
            )


def write_new_doc(dfs, campus, config, debug):
    """Creates a new doc from scratch using the passed tables"""
    # Pull out local variables:
    key_df = dfs["key"]
    efc_df = dfs["efc"]
    award_df = dfs["award"]
    folder = config["drive_folder"]
    file_stem = config["file_stem"]
    efc_sheet_title = config["efc_tab_name"]
    award_sheet_title = config["award_tab_name"]

    # Only runs if there is no current doc for the campus
    if isinstance(key_df, pd.DataFrame) and campus in key_df.index:
        if debug:
            print("Will not create a new doc:" + " doc already exists for this campus")
        return None

    if debug:
        print("About to create doc for {}...".format(campus), flush=True)

    # Create the brand new Google Doc and open permissions
    t0 = time()
    print("Calling credentials", flush=True)
    credentials = googleapi.get_credentials()
    gc = googleapi.gspread_client(credentials)
    if debug:
        print(
            "--Credentials and authorization took {:.2f} seconds".format(time() - t0),
            flush=True,
        )

    t0 = time()
    new_doc = gc.create(campus + " " + file_stem)
    new_key = new_doc.id
    if debug:
        print(
            "--Doc created with key: {} in {:.2f} seconds".format(new_key, time() - t0),
            flush=True,
        )

    t0 = time()
    googleapi.move_spreadsheet_and_share(new_key, folder, credentials)
    if debug:
        print("--Renamed and shared in {:.2f} seconds".format(time() - t0), flush=True)

    #  Make the EFC tab first
    wb = gc.open_by_key(new_key)
    ws = wb.sheet1
    t0 = time()
    ws.update_title(efc_sheet_title)
    if debug:
        print("--Title updated in {:.2f} seconds".format(time() - t0), flush=True)
    t0 = time()
    ws = wb.worksheet(efc_sheet_title)  # This line needed until gspread>=3.1.0
    _write_df_to_sheet(ws, efc_df, new_key, efc_sheet_title, use_index="StudentID")
    if debug:
        print("--EFC data written in {:.2f} seconds".format(time() - t0), flush=True)

    t0 = time()
    googleapi.call_script_service(
        {"function": "doEFCFormats", "parameters": [new_key, efc_sheet_title]}
    )
    if debug:
        print(
            "--Formatting (AppsScript) completed in {:.2f} seconds".format(time() - t0),
            flush=True,
        )

    #  Make the second, awards tab
    t0 = time()
    ws = wb.add_worksheet(title=award_sheet_title, rows=5, cols=5)
    if debug:
        print(
            "--Added Award data sheet in {:.2f} seconds".format(time() - t0), flush=True
        )
    t0 = time()
    _write_df_to_sheet(ws, award_df, new_key, award_sheet_title)
    if debug:
        print("--Award data written in {:.2f} seconds".format(time() - t0), flush=True)
    t0 = time()
    googleapi.call_script_service(
        {"function": "doAwardsFormats", "parameters": [new_key, award_sheet_title]}
    )
    if debug:
        print(
            "--Formatting (AppsScript) completed in {:.2f} seconds".format(time() - t0),
            flush=True,
        )

    #  Third, add formula columns to the efc tab that require award references
    googleapi.call_script_service(
        {"function": "doEFCSecondPass", "parameters": [new_key, efc_sheet_title]}
    )

    # Finally, return the new key for stashing to the key file
    return new_key
