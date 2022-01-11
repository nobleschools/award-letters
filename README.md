# award-letters
Tools for processing and maintaining Google Sheets about award letters.

## Overall structure
Works with Google Docs for award data entry and performs operations
to sync with local Naviance driven data

Relies on a set of AppsScript code not shown here

## Details on setting up AppsScript
(still to be written)


## Starting up for the year
1. Refresh the settings.yml file to change the drive_folder and file_stem values for the new year
2. If it exists, delete key_file.csv in the root directory (this will get created)
3. For each campus, type 'python process_awards -m make_new -ca [campus_name]' to create the new file
4. For each campus, type 'python process_awards -m save -ca [campus_name]' to save those two tabs locally
5. For each campus, type 'python process_awards -m refresh_decisions -ca [campus_name]' to add the Decisions tab

## Running the weekly process
_Each of these commands defaults to '-ca All', running for all campuses_
1. Refresh current_students.csv and current_applications.csv, which are the same files used in the college-lists process
2. python process_awards -m archive  # Saves the current values in each Google Sheet to local csvs and zips an archive. Also tries to fix header errors and deselects any filters in the Google Sheet tabs _(For now, this doesn't work. Instead, run with the save option then manually zip the files. Filters and headers won't be fixed.)_
3. python process_awards -m push_local  # Refreshes the 'Award data' tab and (if necessary) 'EFC data' tab
4. python process_awards -m save  # Saves the changes to those two tabs locally
5. python process_awards -m refresh_decisions  # Updates the Decisions and DecisionOptions tabs with local values
6. python process_awards -m save  # Saves the changes to the Decisions tab locally
7. python process_awards -m combine  # combines all campus data for the three tabs to a single file
8. python process_awards -m report  # Creates Excel reports for the network overall and for each campus
9. python process_awards -m report_single  # Creates a multi-page pdf report for each campus along with single file single page reports per student (and also a zip file with a collection of those per campus)

_For all of these options, they can be run with -kCampus1,Campus2,Campus3 to re-run for all campuses, skipping the named campuses. This is useful if an error is thrown mid-way through the process. (Normally, if an error is thrown, I try to understand what happened, roll back the most recent change to the Google Sheet, and run again starting with that campus.)_
