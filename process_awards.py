#!python3

"""Master file for processing csv data and interacting with Google Docs
   containing award letters"""

import argparse

from modules import filework  # Works with csv and yaml inputs
from modules import basedata  # Creates "clean" tables for Google Docs
from modules import gdocwork  # Works with the Google Docs
from modules import excelreports  # creates Excel reports for a campus


def all_main(settings_file, mode, campus, debug, skip):
    """Meta function to call the below in series, looping through campuses"""
    config = filework.process_config(settings_file, campus)
    skiplist = skip.split(sep=",") if skip else []

    for local_campus in config["campus_list"]:
        if local_campus not in skiplist:
            if debug:
                print(local_campus)
            main(settings_file, mode, local_campus, debug)
        else:
            if debug:
                print("Skipping {}".format(local_campus))


def main(settings_file, mode, campus, debug):
    """Master control file for processing awards:
    1. Reads the settings file for details about other file sources
    2. Processes file sources and then pushes to Google Docs:
      a. If no Google Docs, yet, combines roster and applications to make
         a starting point
      b. If Docs exist, first reads them and then updates them based on any
         necessary changes from the roster and applications
    *3. Optionally, create Excel/PDF reports for each campus

    **Note that the * items are not yet implemented
    """
    # First read the settings file
    if mode in [
        "all",
        "save",
        "make_new",
        "push_local",
        "report",
        "combine",
        "refresh_decisions",
    ]:
        config = filework.process_config(settings_file, campus)

    # Grab csv inputs unless we're only saving the gdocs to a file
    if mode in ["all", "make_new", "push_local", "report", "refresh_decisions"]:
        dfs = filework.read_dfs(config, debug)
    else:
        dfs = {"key": filework.read_doclist(config["key_file"])}

    # Add calculated fields to roster files
    if mode in ["all", "make_new", "push_local", "report", "refresh_decisions"]:
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["sattoact"], campus, debug
        )

    # These are the "blank" tables that don't yet have any award info
    # will add award and efc to the dfs dict
    if mode in ["all", "make_new", "push_local"]:
        basedata.make_clean_gdocs(dfs, config, debug)

    # Read the Google Docs if available and save to local file
    if mode in ["all", "save"]:
        # this adds live_efc and live_award to dfs
        gdocwork.read_current_doc(dfs, campus, config, debug)
        if debug:
            print(
                "{} lines in award tab and {} lines in efc tab".format(
                    len(dfs["live_award"]), len(dfs["live_efc"])
                )
            )

        filework.save_live_dfs(dfs, campus, config, debug)

    # Merge Google Docs info and write back to Google Docs
    #  Just the presence of rows (don't overwrite values)
    if mode in ["all", "push_local"]:
        if ("live_award" not in dfs.keys()) or ("live_efc" not in dfs.keys()):
            filework.read_local_live_data(dfs, campus, config, debug)

        gdocwork.sync_doc_rows(dfs, campus, config, debug)

    # Write a blank document if completely blank (returns None if doc exists)
    if mode in ["make_new"]:
        new_key = gdocwork.write_new_doc(dfs, campus, config, debug)

        # Save output files
        if new_key:
            filework.save_to_doclist(config["key_file"], campus, new_key)

    # Create combined outputs for the two main tables:
    if mode in ["combine"]:
        filework.combine_all_local_files(dfs, config, debug)

    # Refresh live data after push_local for all
    if mode in ["all"]:
        gdocwork.read_current_doc(dfs, campus, config, debug)

    # Update the Decisions tab (do after refreshing the award data tab)
    if mode in ["all", "refresh_decisions"]:
        if ("live_award" not in dfs.keys()) or ("live_efc" not in dfs.keys()):
            filework.read_local_live_data(dfs, campus, config, debug)

        gdocwork.refresh_decisions(dfs, campus, config, debug)

    # Refresh live data after refresh_decisions for all
    if mode in ["all"]:
        gdocwork.read_current_doc(dfs, campus, config, debug)
        filework.save_live_dfs(dfs, campus, config, debug)

    # Generate reports
    if mode in ["all", "report"]:
        if ("live_award" not in dfs.keys()) or ("live_efc" not in dfs.keys()):
            filework.read_local_live_data(dfs, campus, config, debug)
        excelreports.create_excel(dfs, campus, config, debug)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Maintain and process awards")

    parser.add_argument(
        "-s",
        "--settings",
        dest="settings_file",
        action="store",
        help="Name/path of yaml file with detailed settings",
        default="settings/settings.yml",
    )

    parser.add_argument(
        "-ca",
        "--campus",
        dest="campus",
        action="store",
        help='Single campus name (default "All")',
        default="All",
    )

    parser.add_argument(
        "-k",
        "--skip",
        dest="skip",
        action="store",
        help='Campus(es) to skip for an "All" call',
        default="",
    )

    parser.add_argument(
        "-q",
        "--quiet",
        dest="debug",
        action="store_false",
        default=True,
        help="Suppress status messages during report creation",
    )

    parser.add_argument(
        "-m",
        "--mode",
        dest="mode",
        action="store",
        help="Function to execute [all/save/combine/make_new/"
        + "push_local/refresh_decisions/report]",
        default="all",
    )

    args = parser.parse_args()

    if args.campus == "All" and (args.mode not in ["combine", "report"]):
        # Special meta_function to loop through all
        all_main(args.settings_file, args.mode, args.campus, args.debug, args.skip)
    else:
        campus = "All" if args.mode == "combine" else args.campus
        main(args.settings_file, args.mode, campus, args.debug)
