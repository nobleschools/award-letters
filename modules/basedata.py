#!python3
"""
Module for working with raw csv inputs and creating a 'clean' set of tables
to push Google Docs (before adding what's already in those docs
"""

import numpy as np
import pandas as pd


# The following functions are all used to add calculations to the main table
def _get_final_result(x):
    '''Apply function for providing a final status of the application'''
    result_code, attending, waitlisted, deferred, stage, app_type = x
    if result_code == 'denied':
        return 'Denied'
    elif result_code in ['accepted', 'cond. accept', 'summer admit']:
        if attending == 'yes':
            return 'CHOICE!'
        else:
            return 'Accepted!'
    elif result_code == 'guar. transfer':
        return 'Guar. Xfer'
    elif (waitlisted == 1) | (waitlisted == '1'):
        return 'Waitlist'
    elif (deferred == 1) | (deferred == '1'):
        return 'Deferred'
    elif stage == 'pending':
        return 'Pending'
    elif stage in ['initial materials submitted', 'mid-year submitted',
                   'final submitted']:
        return 'Submitted'
    elif app_type == 'interest':
        return 'Interest'
    else:
        return '?'


def _make_barrons_translation(x):
    '''Apply function for a custom mapping of a text Barron's field to
    a number'''
    bar_dict = {'Most Competitive+': 1,
                'Most Competitive': 2,
                'Highly Competitive': 3,
                'Very Competitive': 4,
                'Competitive': 5,
                'Less Competitive': 6,
                'Noncompetitive': 7,
                '2 year (Noncompetitive)': 8,
                '2 year (Competitive)': 8,
                'Not Available': 'N/A'}
    if x in bar_dict:
        return bar_dict[x]
    else:
        return '?'


def _get_sat_translation(x, lookup_df):
    '''Apply function for calculating equivalent ACT for SAT scores.
    Lookup table has index of SAT with value of ACT'''
    sat = x
    if np.isreal(sat):
        if sat in lookup_df.index:  # it's an SAT value in the table
            return lookup_df.loc[sat, 'ACT']
    return np.nan  # default if not in table or not a number


def _get_act_max(x):
    ''' Returns the max of two values if both are numbers, otherwise
    returns the numeric one or nan if neither is numeric'''
    act, sat_in_act = x
    if np.isreal(act):
        if np.isreal(sat_in_act):
            return max(act, sat_in_act)
        else:
            return act
    else:
        if np.isreal(sat_in_act):
            return sat_in_act
        else:
            return np.nan


def _get_strategies(x, lookup_df):
    '''Apply function for calculating strategies based on gpa and act using the
    lookup table (mirrors Excel equation for looking up strategy'''
    gpa, act = x
    if np.isreal(gpa) and np.isreal(act):
        lookup = '{:.1f}:{:.0f}'.format(
                max(np.floor(gpa*10)/10, 1.5), max(act, 12))
        return lookup_df['Strategy'].get(lookup, np.nan)
    else:
        return np.nan


def _safe2int(x):
    try:
        return int(x+20)
    except BaseException:
        return x


def _get_gr_target(x, lookup_strat, goal_type):
    '''Apply function to get the target or ideal grad rate for student'''
    strat, gpa, efc, race = x
    # 2 or 3 strategies are split by being above/below 3.0 GPA line
    # First we identify those and then adjust the lookup index accordingly
    special_strats = [int(x[0]) for x in lookup_strat.index if x[-1] == '+']
    if np.isreal(gpa) and np.isreal(strat):
        # First define the row in the lookup table
        strat_str = '{:.0f}'.format(strat)
        if strat in special_strats:
            lookup = strat_str + '+' if gpa >= 3.0 else strat_str + '<'
        else:
            lookup = strat_str

        # Then define the column in the lookup table
        if efc == -1:
            column = 'minus1_' + goal_type
        elif race in ['W', 'A']:
            column = 'W/A_' + goal_type
        else:
            column = 'AA/H_' + goal_type
        return lookup_strat[column].get(lookup, np.nan)
    else:
        return np.nan


def _make_final_gr(x):
    """Apply function to do graduation rates"""
    race, sixyrgr, sixyrgraah, comments = x
    first_gr = sixyrgraah if race in ['B', 'H', 'M'] else sixyrgr
    if comments == 'Posse':
        return (first_gr+0.15) if first_gr < 0.7 else (1.0-(1.0-first_gr)/2)
    else:
        return first_gr


# Finally, the main function that calls these

def add_strat_and_grs(df, strat_df, target_df, sattoact_df, campus, debug):
    """
    Adds Strategy and Target/Ideal grad rate numbers to the roster table
    """
    df = df.copy()
    if campus != 'All':
        df = df[df['Campus'] == campus]
    df['local_sat_in_act'] = df['SAT'].apply(_get_sat_translation,
                                             args=(sattoact_df,))
    df['local_act_max'] = df[['ACT', 'local_sat_in_act']].apply(
            _get_act_max, axis=1)
    df['Stra-tegy'] = df[['GPA', 'local_act_max']].apply(
        _get_strategies, axis=1, args=(strat_df,))
    df['Target Grad Rate'] = df[
            ['Stra-tegy', 'GPA', 'EFC', 'Race/ Eth']].apply(
            _get_gr_target, axis=1, args=(target_df, 'target'))
    df['Ideal Grad Rate'] = df[
            ['Stra-tegy', 'GPA', 'EFC', 'Race/ Eth']].apply(
            _get_gr_target, axis=1, args=(target_df, 'ideal'))

    if debug:
        print('Total roster length of {}.'.format(len(df)))
    return df


def make_clean_gdocs(dfs, config, debug):
    """
    Creates a set of tables for pushing to Google Docs assuming there
    is no existing award data based on the applications and roster files
    """

    # Pullout local config settings:
    ros_df = dfs['ros']
    app_df = dfs['app']
    college_df = dfs['college']
    award_fields = config['award_fields']
    efc_tab_fields = config['efc_tab_fields']
    include_statuses = config['app_status_to_include']
    award_sort = config['award_sort']

    if debug:
        print('Creating "Blank" Google Docs tables from source csvs',
              flush=True)

    # #####################################################
    # First do the (simpler) EFC tab, which is just a combination of
    # direct columns from the roster plus some blank columns
    efc_pull_fields = [field for field in efc_tab_fields if
                       field in ros_df.columns]

    # the line below skips the first column because it is assumed to be
    # the index
    efc_blank_fields = [field for field in efc_tab_fields[1:] if
                        field not in ros_df.columns]
    efc_df = ros_df[efc_pull_fields]
    efc_df = efc_df.reindex(columns=efc_df.columns.tolist() + efc_blank_fields)

    # #####################################################
    # Now do the more complicated awards tab
    current_students = list(efc_df.index)
    award_df = app_df[app_df['hs_student_id'].isin(current_students)].copy()

    # Do all of the lookups from the roster table:
    for dest, source, default in (
            ('lf', 'LastFirst', 'StudentMissing'),
            ('tgr', 'Target Grad Rate', np.nan),
            ('igr', 'Ideal Grad Rate', np.nan),
            ('race', 'Race/ Eth', 'N/A'),
            ):
        award_df[dest] = award_df['hs_student_id'].apply(
                lambda x: ros_df[source].get(x, default))

    # Now do all lookups from the college table:
    for dest, source, default in (
            ('cname', 'INSTNM', 'NotAvail'),
            ('barrons', 'SimpleBarrons', 'N/A'),
            ('local', 'Living', 'Campus'),
            ('sixyrgr', 'Adj6yrGrad_All', np.nan),
            ('sixyrgraah', 'Adj6yrGrad_AA_Hisp', np.nan),
            ):
        award_df[dest] = award_df['NCES'].apply(
                lambda x: college_df[source].get(x, default))

    # Cleanup from college table for missing values
    award_df['cname'] = award_df[['cname', 'collegename']].apply(
            lambda x: x[1] if x[0] == 'NotAvail' else x[0], axis=1)
    award_df['barrons'] = award_df['barrons'].apply(
            _make_barrons_translation)
    award_df['sixyrfinal'] = award_df[['race', 'sixyrgr', 'sixyrgraah',
                                       'comments']].apply(_make_final_gr, axis=1)

    # Other interpreted/calculated values:
    award_df['final_result'] = award_df[
        ['result_code', 'attending', 'waitlisted', 'deferred', 'stage', 'type']
        ].apply(_get_final_result, axis=1)

    # Calculated or blank columns (we'll push the calculations with AppsScript)
    for f in [
            'Award Receiv- ed?',
            'Tuition & Fees (including insurance if req.)',
            'Room & board (if not living at home)',
            'College grants & scholarships',
            'Government grants (Pell/SEOG/MAP)',
            'Net Price (before Loans) <CALCULATED>',
            'Student Loans offered (include all non-parent)',
            'Out of Pocket Cost (Direct Cost-Grants-Loans) <CALCULATED>',
            'Your EFC <DRAWN FROM OTHER TAB>',
            'Unmet need <CALCULATED>',
            'Work Study (enter for comparison if desired)',
            'Award',
            ]:
        award_df[f] = ''

    # Still need to double up home colleges for home/away rows
    both_rows = award_df[award_df['local'] == 'Both'].copy()
    # 'Both' rows become 'Home' here and the replicants (above) will be Away
    award_df['cname'] = award_df[['cname', 'local']].apply(
            lambda x: x[0] + ('' if x[1] == 'Campus' else '--At Home'), axis=1)
    award_df['local'] = award_df['local'].apply(
            lambda x: 'Home' if x == 'Both' else x)
    award_df['Unique'] = 1
    # these are the duplicate rows
    both_rows['local'] = 'Campus'
    both_rows['Unique'] = 0
    both_rows['cname'] = both_rows['cname']+'--On Campus'
    award_df = pd.concat([award_df, both_rows])

    # Keep different statuses based on config file
    award_df = award_df[award_df['final_result'].isin(include_statuses)]

    # Rename labels to match what will be in the doc
    mapper = {
            'lf': 'Student',
            'tgr': 'Target Grad Rate',
            'igr': 'Ideal Grad Rate',
            'cname': 'College/University',
            'barrons': 'Selectivity\n' +
                       '1=Most+\n' +
                       '2=Most\n' +
                       '3=Highly\n' +
                       '4=Very\n' +
                       '5=Competitive\n' +
                       '6=Less\n' +
                       '7=Non\n' +
                       '8=2 year',
            'final_result': 'Result (from Naviance)',
            'sixyrfinal': '6-Year Minority Grad Rate',
            'hs_student_id': 'SID',
            'NCES': 'NCESid',
            'local': 'Home/Away',
            }
    use_mapper = {key: value for key, value in mapper.items() if
                  value in award_fields}
    award_df.rename(columns=use_mapper, inplace=True)

    # Final reduce the table to just what's going in the Google Doc
    award_df = award_df[award_fields]

    # Sort the table based on config file
    # award_sort is a list with the right fields
    if award_sort[1] == 'College/University':
        sort_order = [True, True]
    else:
        sort_order = [True, False, True]

    award_df.sort_values(by=award_sort, ascending=sort_order, inplace=True)

    dfs['award'] = award_df
    dfs['efc'] = efc_df
