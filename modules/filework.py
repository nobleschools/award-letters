#!python3

"""
Modules for working with csv inputs and manipulating them
"""

import os
import shutil
import yaml
import csv
import pandas as pd


def process_config(settings_file, campus):
    """Returns a dict of simple keyword configurations based on what
    was in the yaml file and the specific campus"""
    with open(settings_file, "r") as ymlfile:
        cfg = yaml.load(ymlfile, Loader=yaml.FullLoader)
    config = {}

    # Handle the settings based on the complex/standard switch
    if campus in cfg["use_complex"]:
        config["award_fields"] = cfg["award_fields"]["Complex"]
        config["award_sort"] = cfg["award_sort"]["Complex"]
    else:
        config["award_fields"] = cfg["award_fields"]["Standard"]
        config["award_sort"] = cfg["award_sort"]["Standard"]

    # For campus switches that modify "standard"
    for key in [
        "app_status_to_include",
        "efc_tab_name",
        "award_tab_name",
        "efc_header_row",
        "award_header_row",
        "decision_tab_name",
        "decision_options_tab_name",
        "decision_options_header_row",
        "decision_header_row",
        "decision_defaults",
        "report_award_sorts",
        "report_student_sorts",
    ]:
        if campus in cfg[key]:
            config[key] = cfg[key][campus]
        else:
            config[key] = cfg[key]["Standard"]

    # For straight reads:
    for key in [
        "output_folder",
        "report_filename",
        "report_folder",
        "excel_formats",
        "live_backup_folder",
        "live_backup_prefix",
        "drive_folder",
        "live_archive_folder",
        "campus_list",
        "live_award_fields",
        "file_stem",
        "efc_tab_fields",
        "app_fields",
        "roster_fields",
        "live_efc_fields",
        "report_award_fields",
        "report_award_formats",
        "decision_option_fields",
        "live_decision_fields",
        "report_student_fields",
        "summary_settings",
    ]:
        config[key] = cfg[key]

    for input_key in [
        "key_file",
        "current_applications",
        "current_roster",
        "strategies",
        "targets",
        "colleges",
        "acttosat",
        "bump_list",
        "ambitious_pp",
    ]:
        config[input_key] = cfg["inputs"][input_key]

    return config


def safe2int(x):
    """converts to int if possible, otherwise original"""
    try:
        return int(x)
    except BaseException:
        return x


def safe2f(x):
    """converts to float if possible, otherwise is a string"""
    try:
        return float(x)
    except BaseException:
        return x


def p2f(x):
    """converts percent string to float number"""
    return None if x == "N/A" else float(x.strip("%")) / 100


def save_live_dfs(dfs, campus, config, debug):
    """Takes the current live_ keyed dataframes and saves them to the live
    folder, backing up the current item in the live folder to the backup
    folder. If the live folder is in S3, detects that and works with that
    system
    """
    # Find all the current DataFrames with 'live_' prefix
    dfs_to_save = [x[5:] for x in dfs.keys() if x[:5] == "live_"]
    if debug and not dfs_to_save:
        print("No live dataframes to save")
        return

    for key in dfs_to_save:
        filename = config["live_backup_prefix"] + "-" + campus + "-" + key + ".csv"
        full_path = os.path.join(config["live_backup_folder"], filename)
        # If the file already exists, we'll backup to the archive directory
        if os.path.isfile(full_path):
            archive_path = os.path.join(config["live_archive_folder"], filename)
            shutil.copy(full_path, archive_path)

        # We have a special index label to preserve for the efc table only
        index_label = "StudentID" if key in ["efc", "decision"] else "DefaultIndex"
        dfs["live_" + key].to_csv(full_path, index_label=index_label)


def read_local_live_all_decision(dfs, campus, config, debug):
    """Hack to repeat the below function for the decision tab reading the "All" file"""
    filename = config["live_backup_prefix"] + "-All-decision.csv"
    full_path = os.path.join(config["live_backup_folder"], filename)
    if os.path.isfile(full_path):
        this_df = pd.read_csv(full_path, index_col=0)
        dfs["live_decision"] = this_df[this_df["Campus"]==campus]
    else:
        if debug:
            print("{} does not exist".format(full_path))


def read_local_live_data(dfs, campus, config, debug):
    """Loads the live data (recently read from the Google Doc) to the file
    into the live dataframes"""
    if debug:
        print("Reading local version of live dataframes", flush=True)
    for key in ["efc", "award", "decision"]:
        filename = config["live_backup_prefix"] + "-" + campus + "-" + key + ".csv"
        full_path = os.path.join(config["live_backup_folder"], filename)
        if os.path.isfile(full_path):
            if key in ["efc", "decision"]:
                dfs["live_" + key] = pd.read_csv(full_path, index_col=0)
            else:
                dfs["live_" + key] = pd.read_csv(full_path, index_col=False)
                dfs["live_" + key].drop(["DefaultIndex"], axis=1, inplace=True)
        else:
            if debug:
                print("{} does not exist".format(full_path))


def combine_all_local_files(dfs, config, debug):
    """Runs through the list of all campuses and combines to three
    merged csvs
    """
    if debug:
        print("About to combine the files for campuses:")
        print(config["campus_list"])

    big_df = {}
    big_df["live_efc"] = None  # This will be empty for the first pass
    big_df["live_award"] = None
    big_df["live_decision"] = None

    # Merge all the files
    for campus in config["campus_list"]:
        if debug:
            print("Reading data for {}".format(campus), flush=True)
        read_local_live_data(dfs, campus, config, debug=False)
        for key in ["live_efc", "live_award", "live_decision"]:
            if key in dfs.keys():
                dfs[key]["Campus"] = campus
                if isinstance(big_df[key], pd.DataFrame):
                    big_df[key] = pd.concat([big_df[key], dfs[key]], sort=False)
                else:
                    big_df[key] = dfs[key]
                dfs.pop(key)

    # Save
    for key in ["efc", "award", "decision"]:
        if isinstance(big_df["live_" + key], pd.DataFrame):
            # Reduce to just the columns we want
            these_fields = config["live_" + key + "_fields"]
            big_df["live_" + key] = big_df["live_" + key][these_fields]
            filename = config["live_backup_prefix"] + "-All-" + key + ".csv"
            full_path = os.path.join(config["live_backup_folder"], filename)
            # If the file already exists, we'll backup to the archive directory
            if os.path.isfile(full_path):
                archive_path = os.path.join(config["live_archive_folder"], filename)
                shutil.copy(full_path, archive_path)

            # We have a special index label to preserve for these tables
            index_label = "StudentID" if key in ["efc", "decision"] else "DefaultIndex"
            big_df["live_" + key].to_csv(full_path, index_label=index_label)


def read_standard_csv(fn):
    """
    Reads an input file and returns a DataFrame with first column as index
    """
    df = pd.read_csv(fn, index_col=0, na_values=["N/A", ""])
    return df


def read_apps(fn, cols):
    """
    Reads the applications file into a DataFrame, using the correct formatting
    for special columns; is passed a list of columns to pay attention to
    """
    df = pd.read_csv(
        fn,
        na_values=[""],
        encoding="cp1252",
        usecols=cols,
        converters={"hs_student_id": safe2int, "NCES": safe2int},
    )
    return df


def read_bumplist(fn):
    """
    Reads the bump list into a DataFrame with dummy index
    """
    df = pd.read_csv(
        fn, encoding="cp1252", converters={"SID": safe2int, "NCESid": safe2int}
    )
    return df


def read_roster(fn, cols):
    """
    Reads the roster file into a DataFrame, using the correct formatting
    for special columns; is passed a list of columns to use
    """
    df = pd.read_csv(
        fn,
        index_col="StudentID",
        na_values=["N/A", ""],
        usecols=cols,
        encoding="cp1252",
        converters={
            "EFC": safe2int,
            "ACT": safe2int,
            "InterimSAT": safe2int,
            "SAT": safe2int,
            "GPA": safe2f,
            "StudentID": safe2int,
        },
    )
    return df


def read_colleges(fn):
    """
    Reads the colleges file into a DataFrame, using the correct formatting
    for special columns
    """
    df = pd.read_csv(
        fn,
        na_values=["N/A"],
        encoding="cp1252",
        index_col=0,
        converters={
            "UNITID": safe2int,
            "Adj6yrGrad_All": p2f,
            "Adj6yrGrad_AA_Hisp": p2f,
        },
    )
    return df


def save_csv_from_table(fn, folder, list_of_lists):
    """
    Saves a csv to the given FileName and folder; creates folder if does
    not exist
    """
    if not os.path.exists(folder):
        os.mkdir(folder)
    long_fn = os.path.join(folder, fn)
    outf = open(long_fn, "wt", encoding="utf-8")
    writer = csv.writer(
        outf, delimiter=",", quoting=csv.QUOTE_MINIMAL, lineterminator="\n"
    )
    for row in list_of_lists:
        writer.writerow(row)
    outf.close()


def read_doclist(fn):
    """
    Reads the doclist and returns a dataframe with that info.
    Returns None if file doesn't exist yet
    """
    if not os.path.exists(fn):
        return None
    else:
        return pd.read_csv(fn, index_col=0)


def save_to_doclist(fn, campus, key):
    """
    Saves the passed campus and key to the doclist csv
    """
    if not os.path.exists(fn):
        df = pd.DataFrame({"ss_key": key}, index=[campus])
    else:
        df = pd.read_csv(fn, index_col=0)
        df.loc[campus, "ss_key"] = key

    df.to_csv(fn, index_label="Campus")


def give_campus(x, ref_df):
    """Apply function to lookup the name of the campus from the SchoolID"""
    return ref_df.loc[x][0]


def give_table_value(x, ref_df, field):
    """Apply function to lookup the value of a field from the student number"""
    return ref_df.loc[x, field]


def read_dfs(config, debug):
    """Master function for reading input data files based on config input.
    Returns a dict of dfs"""
    if debug:
        print("Reading configuration inputs", flush=True)

    dfs = {}
    dfs["key"] = read_doclist(config["key_file"])
    dfs["app"] = read_apps(config["current_applications"], config["app_fields"])
    dfs["ros"] = read_roster(config["current_roster"], config["roster_fields"])
    dfs["strat"] = read_standard_csv(config["strategies"])
    dfs["target"] = read_standard_csv(config["targets"])
    dfs["college"] = read_colleges(config["colleges"])
    dfs["acttosat"] = read_standard_csv(config["acttosat"])
    dfs["bump_list"] = read_bumplist(config["bump_list"])
    dfs["ambitious_pp"] = read_standard_csv(config["ambitious_pp"])
    return dfs


def create_folder_if_necessary(location):
    """
    Checks for existence of folder and creates if not there.
    Location is a list and this function runs recursively
    """
    for i in range(len(location)):
        this_location = os.path.join(*location[:(i+1)])
        if not os.path.exists(this_location):
            os.makedirs(this_location)
