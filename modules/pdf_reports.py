#!python3
"""
Module for creating the PDF reports from reports tables
"""

# import numpy as np
# import pandas as pd
from modules import filework


def create_pdfs(dfs, campus, config, debug):
    """Will create Excel reports for sharing details from Google Docs"""
    if debug:
        print("Creating PDF report for {}".format(campus), flush=True)
    filework.create_folder_if_necessary(
        config["output_folder"]
    )

    # Create the excel:
    # Initial document and hidden college lookup
    # Students tab
    # Award data tab
