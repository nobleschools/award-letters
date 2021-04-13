#!python3

"""Master file for processing csv data and interacting with Google Docs
   containing award letters"""

import argparse

from modules import filework  # Works with csv and yaml inputs
from modules import basedata  # Creates "clean" tables for Google Docs
from modules import gdocwork  # Works with the Google Docs
from modules import reports  # creates Excel reports for a campus
from modules import pdf_reports  # creates PDF reports


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
    # Note: comments in the all mode obviously apply to the subset modes
    if mode == "all":
        config = filework.process_config(settings_file, campus)
        # Grab csv inputs
        dfs = filework.read_dfs(config, debug)
        # Add calculated fields to roster files
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["acttosat"], campus, debug
        )
        # Add award and efc to the dfs dict
        # These are the "blank" tables that don't yet have any award info
        basedata.make_clean_gdocs(dfs, config, debug)
        # Read the Google Docs if available and save to local file
        # this adds live_efc and live_award to dfs
        gdocwork.read_current_doc(dfs, campus, config, debug)
        filework.save_live_dfs(dfs, campus, config, debug)
        # Merge Google Docs info and write back to Google Docs
        # Just the presence of rows (don't overwrite values)
        gdocwork.sync_doc_rows(dfs, campus, config, debug)
        # Refresh live data after syncing
        gdocwork.read_current_doc(dfs, campus, config, debug)
        # Update the Decisions tab (do after refreshing the award data tab)
        gdocwork.refresh_decisions(dfs, campus, config, debug)
        # Refresh live data after refresh_decisions for all
        gdocwork.read_current_doc(dfs, campus, config, debug)
        filework.save_live_dfs(dfs, campus, config, debug)
        # Generate reports
        reports.create_report_tables(dfs, campus, config, debug)
        reports.create_excel(dfs, campus, config, debug)

    elif mode == "save":
        config = filework.process_config(settings_file, campus)
        dfs = {"key": filework.read_doclist(config["key_file"])}
        gdocwork.read_current_doc(dfs, campus, config, debug)
        filework.save_live_dfs(dfs, campus, config, debug)

    elif mode == "make_new":
        config = filework.process_config(settings_file, campus)
        dfs = filework.read_dfs(config, debug)
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["acttosat"], campus, debug
        )
        basedata.make_clean_gdocs(dfs, config, debug)
        # Write a blank document if completely blank (returns None if doc exists)
        new_key = gdocwork.write_new_doc(dfs, campus, config, debug)
        # Save output files (if it's brand new)
        if new_key:
            filework.save_to_doclist(config["key_file"], campus, new_key)

    elif mode == "push_local":
        config = filework.process_config(settings_file, campus)
        dfs = filework.read_dfs(config, debug)
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["acttosat"], campus, debug
        )
        basedata.make_clean_gdocs(dfs, config, debug)
        filework.read_local_live_data(dfs, campus, config, debug)
        gdocwork.sync_doc_rows(dfs, campus, config, debug)

    elif mode == "combine":
        config = filework.process_config(settings_file, campus)
        dfs = {"key": filework.read_doclist(config["key_file"])}
        # Create combined outputs for the two main tables:
        filework.combine_all_local_files(dfs, config, debug)

    elif mode == "refresh_decisions":
        if debug:
            print("Refreshing decisions (make sure you refreshed award data first!)")
        config = filework.process_config(settings_file, campus)
        dfs = filework.read_dfs(config, debug)
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["acttosat"], campus, debug
        )
        filework.read_local_live_data(dfs, campus, config, debug)
        gdocwork.refresh_decisions(dfs, campus, config, debug)

    elif mode == "report":
        config = filework.process_config(settings_file, campus)
        dfs = filework.read_dfs(config, debug)
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["acttosat"], campus, debug
        )
        filework.read_local_live_data(dfs, campus, config, debug)
        reports.create_report_tables(dfs, campus, config, debug)
        reports.create_excel(dfs, campus, config, debug)

    elif mode == "report_single":
        config = filework.process_config(settings_file, campus)
        dfs = filework.read_dfs(config, debug)
        dfs["ros"] = basedata.add_strat_and_grs(
            dfs["ros"], dfs["strat"], dfs["target"], dfs["acttosat"], campus, debug
        )
        filework.read_local_live_data(dfs, campus, config, debug)
        reports.create_report_tables(dfs, campus, config, debug)
        # First line creates a combined campus file
        pdf_reports.create_pdfs(dfs, campus, config, debug, single_pdf=False)
        # This line creates one per student
        pdf_reports.create_pdfs(dfs, campus, config, debug)

    else:
        print("Invalid mode. Aborting")


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
        + "push_local/refresh_decisions/report/report_single]",
        default="all",
    )

    args = parser.parse_args()

    if args.campus == "All" and (args.mode not in
                                 ["combine", "report"]):
        # Special meta_function to loop through all
        all_main(args.settings_file, args.mode, args.campus, args.debug, args.skip)
    elif args.campus == "All" and args.mode == "report":
        # First loop through all campuses individually
        all_main(args.settings_file, args.mode, args.campus, args.debug, args.skip)
        # Then call for the entire network
        main(args.settings_file, args.mode, campus, args.debug)
    else:
        campus = "All" if args.mode == "combine" else args.campus
        main(args.settings_file, args.mode, campus, args.debug)
