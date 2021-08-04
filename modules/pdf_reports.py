#!python3
"""
Module for creating the PDF reports from reports tables
"""

import warnings
import os
import zipfile
import pandas as pd
from fpdf import FPDF
from datetime import date
from modules import filework


TOP_MARGIN = 0.75
BOTTOM_MARGIN = 0.75
LEFT_MARGIN = 0.5
RIGHT_MARGIN = 0.5
LINE_WIDTH = 0.0075
THICK_LINE = 0.02
W = [2.77, 0.49, 0.8, 0.61, 0.63, 0.65, 0.65, 0.65, 0.66, 0.76, 0.7]
H = [0.32, 0.21, 0.15, 0.196, 0.196, 0.196, 0.196, 0.196, 0.196, 0.25]
MH = 0.22


def _set_color_name(pdf, name, type="fill"):
    """Helper function to encapsulate RGB codes for pdf"""
    colors = {
        "light_blue": (220, 230, 241),
        "light_yellow": (255, 242, 204),
        "salmon": (253, 233, 217),
        "grey": (217, 217, 217),
        "navy_blue": (0, 32, 96),
        "red": (255, 0, 0),
        "black": (0, 0, 0),
    }
    if name in colors:
        r, g, b = colors[name]
        if type == "fill":
            pdf.set_fill_color(r=r, g=g, b=b)
        elif type == "text":
            pdf.set_text_color(r=r, g=g, b=b)
    else:
        raise (RuntimeError("PDF color not specified: " + name))


def initiate_pdf_object():
    """
    Creates a pdf object for writing using defaults.
    Defaults hardcoded for now, but may push to settings in future
    """
    pdf = FPDF(orientation="L", unit="in", format="Letter")
    fonts = [
        ["font_r", "./fonts/Carlito-Regular.ttf"],
        ["font_b", "./fonts/Carlito-Bold.ttf"],
        ["font_i", "./fonts/Carlito-Italic.ttf"],
        ["font_bi", "./fonts/Carlito-BoldItalic.ttf"],
    ]
    for font_name, filename in fonts:
        pdf.add_font(font_name, "", filename, uni=True)
    pdf.set_line_width(LINE_WIDTH)
    pdf.set_margins(left=LEFT_MARGIN, top=TOP_MARGIN, right=RIGHT_MARGIN)
    return pdf


def _safe_dollar(amt, nan="N/A", blank_zeros=False):
    """Returns the EFC/amount as a dollar format if applicable"""
    if pd.isnull(amt):
        return nan
    elif amt == -1:
        return "-1"
    elif blank_zeros and amt == 0:
        return "$      -"
    else:
        try:
            return (
                f"${float(amt):,.0f}" if float(amt) >= 0 else f"$ ({float(-amt):,.0f})"
            )
        except Exception:
            return str(amt)


def add_pdf_header_row(pdf, data, campus, second=False):
    """
    Creates a new page and adds the non-college specific data to report
    """
    pdf.add_page()
    pdf.set_y(TOP_MARGIN)

    # First row
    text = "College Options Worksheet" + (" (page 2)" if second else "")
    pdf.set_font("font_b", "", 15)
    pdf.cell(w=sum(W[:2]), txt=text, h=H[0], border=0, ln=0, align="L", fill=False)

    pdf.set_font("font_b", "", 11)
    _set_color_name(pdf, "light_blue")
    text = "Expected Family Contribution:"
    pdf.cell(w=sum(W[2:5]), txt=text, h=H[0], border=1, ln=0, align="R", fill=True)

    pdf.set_font("font_r", "", 11)
    text = _safe_dollar(data["EFC"])
    pdf.cell(w=W[5], txt=text, h=H[0], border=1, ln=0, align="C", fill=False)

    text = "Counselor: " + data["Counselor"]
    if not (data["Cohort"] == "" or pd.isnull(data["Cohort"])):
        text += " (" + data["Cohort"] + ")"
    pdf.cell(w=sum(W[6:]), txt=text, h=H[0], border=0, ln=1, align="R", fill=False)

    # Second row
    pdf.set_font("font_b", "", 11)
    _set_color_name(pdf, "salmon")
    text = data["LastFirst"]
    pdf.cell(w=W[0], txt=text, h=H[1], border=1, ln=0, align="L", fill=True)

    _set_color_name(pdf, "light_blue")
    pdf.cell(w=W[1], txt="TGR", h=H[1], border=1, ln=0, align="R", fill=True)

    pdf.set_font("font_r", "", 11)
    text = f"{data['TGR']:.0%}"
    pdf.cell(w=W[2], txt=text, h=H[1], border=1, ln=0, align="C", fill=False)

    pdf.set_font("font_b", "", 11)
    _set_color_name(pdf, "light_blue")
    pdf.cell(w=W[3], txt="GPA", h=H[1], border=1, ln=0, align="R", fill=True)

    pdf.set_font("font_r", "", 11)
    text = (
        "N/A"
        if (pd.isnull(data["GPA"]) or isinstance(data["GPA"], str))
        else f"{data['GPA']:.2f}"
    )
    pdf.cell(w=W[4], txt=text, h=H[1], border=1, ln=0, align="C", fill=False)

    pdf.set_font("font_b", "", 11)
    pdf.cell(w=W[5], txt="SAT", h=H[1], border=1, ln=0, align="R", fill=True)

    pdf.set_font("font_r", "", 11)
    text = "N/A" if pd.isnull(data["SAT"]) else f"{int(data['SAT'])}"
    pdf.cell(w=W[6], txt=text, h=H[1], border=1, ln=0, align="C", fill=False)

    text = "Advisor: " + data["Advisor"]
    pdf.cell(w=sum(W[7:]), txt=text, h=H[1], border=0, ln=1, align="R", fill=False)

    # Third row
    pdf.set_x(LEFT_MARGIN + sum(W[:3]))
    pdf.set_font("font_i", "", 9)
    pdf.cell(w=W[3], txt="A", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[4], txt="B", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[5], txt="C", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[6], txt="D", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[7], txt="E=A+B-C-D", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[8], txt="F", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[9], txt="G=E-F", h=H[2], border=0, ln=0, align="C", fill=False)
    pdf.cell(w=W[10], txt="G=E-F", h=H[2], border=0, ln=1, align="C", fill=False)

    # Fourth row: this is actually split into multiple rows
    row_4 = [
        ["", "", "", "", "", "College or University"],
        ["", "", "", "6 yr", "Grad", "Rate"],
        ["", "", "", "", "Application", "Status"],
        ["", "", "", "", "Tuition &", "Fees"],
        ["", "", "Room &", "board (if", "not living", "at home"],
        ["", "", "College", "grants &", "scholar-", "ships"],
        ["", "", "", "Govern-", "ment", "grants"],
        ["", "", "", "Net Price", "(before", "Loans)"],
        ["Direct", "Loans", "offered", "(include", "all non-", "parent)"],
        ["Out of", "Pocket Cost", "(includes", "up to", "$6,000 in", "loans)"],
        ["Out of", "Pocket", "Cost", "(includes", "all direct", "loans)"],
    ]
    # As of now, only customized for Bulls--if this expands, should be in settings
    if campus == "Bulls":
        row_4[7] = ["Net", "Price/", "Out of", "Pocket", "(before", "Loans)"]
        row_4[9] = ["", "Left to Pay", "(includes", "up to", "$6,000 in", "loans)"]
        row_4[10] = ["", "Left to", "Pay", "(includes", "all direct", "loans)"]
    pdf.set_font("font_b", "", 11)
    for i in range(len(row_4[0])):
        t = [x[i] for x in row_4]
        _set_color_name(pdf, "light_blue")
        pdf.cell(w=W[0], txt=t[0], h=H[i + 3], border=0, ln=0, align="L", fill=True)
        pdf.cell(w=W[1], txt=t[1], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        pdf.cell(w=W[2], txt=t[2], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        pdf.cell(w=W[3], txt=t[3], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        pdf.cell(w=W[4], txt=t[4], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        pdf.cell(w=W[5], txt=t[5], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        pdf.cell(w=W[6], txt=t[6], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        _set_color_name(pdf, "light_yellow")
        pdf.cell(w=W[7], txt=t[7], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        _set_color_name(pdf, "light_blue")
        pdf.cell(w=W[8], txt=t[8], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        _set_color_name(pdf, "light_yellow")
        pdf.cell(w=W[9], txt=t[9], h=H[i + 3], border=0, ln=0, align="C", fill=True)
        pdf.cell(w=W[10], txt=t[10], h=H[i + 3], border=0, ln=1, align="C", fill=True)

    # End of header except for the lines separating this last tall header row
    line_top = TOP_MARGIN + sum(H[:3])
    line_bottom = TOP_MARGIN + sum(H[:9])
    for i in [1, 2, 3, 5, 7, 8, 9]:
        line_x = LEFT_MARGIN + sum(W[:i])
        pdf.line(line_x, line_top, line_x, line_bottom)


def _s_cell(pdf, w, txt, h, border, ln, align, fill, font="font_r", size=11):
    """
    Fits text in a cell by reducing the font size
    """
    new_size = size
    while (w < pdf.get_string_width(txt)) and new_size > 6:
        new_size -= 1
        pdf.set_font(font, "", new_size)
    pdf.cell(w=w, txt=txt, h=h, border=border, ln=ln, align=align, fill=fill)
    pdf.set_font(font, "", size)


def _get_net_price(data):
    """Utility to pull out net price from a data row"""
    rb = 0 if pd.isnull(data["Room & board"]) else data["Room & board"]
    cgs = (
        0
        if pd.isnull(data["College grants & scholarships"])
        else data["College grants & scholarships"]
    )
    gg = 0 if pd.isnull(data["Government grants"]) else data["Government grants"]
    sl = (
        0 if pd.isnull(data["Student Loans offered"]) else data["Student Loans offered"]
    )
    try:
        net_price = data["Tuition & Fees"] + rb - cgs - gg
        oop1 = net_price - min(sl, 6000)
        oop2 = net_price - sl
        return (
            _safe_dollar(net_price, nan="", blank_zeros=True),
            _safe_dollar(oop1, nan="", blank_zeros=True),
            _safe_dollar(oop2, nan="", blank_zeros=True),
        )
    except Exception:
        return ("", "", "")


def add_college_rows(pdf, student_award):
    """
    Adds a row for each record in the student_award dataframe.
    The dataframe size is managed by the calling function.
    """
    for i, data in student_award.iterrows():
        pdf.set_font("font_r", "", 11)
        text = data["College/University"]
        _s_cell(pdf, w=W[0], txt=text, h=H[-1], border=1, ln=0, align="L", fill=False)
        text = f'{data["Grad rate for sorting"]:.0%}'
        pdf.cell(w=W[1], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        text = data["Result"]
        pdf.cell(w=W[2], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        text = _safe_dollar(data["Tuition & Fees"], nan="", blank_zeros=True)
        pdf.cell(w=W[3], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        text = _safe_dollar(data["Room & board"], nan="", blank_zeros=True)
        pdf.cell(w=W[4], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        text = _safe_dollar(
            data["College grants & scholarships"], nan="", blank_zeros=True
        )
        pdf.cell(w=W[5], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        text = _safe_dollar(data["Government grants"], nan="", blank_zeros=True)
        pdf.cell(w=W[6], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        netp, oop1, oop2 = _get_net_price(data)
        pdf.set_font("font_b", "", 11)
        pdf.cell(w=W[7], txt=netp, h=H[-1], border=1, ln=0, align="C", fill=False)
        text = _safe_dollar(data["Student Loans offered"], nan="", blank_zeros=True)
        pdf.set_font("font_r", "", 11)
        pdf.cell(w=W[8], txt=text, h=H[-1], border=1, ln=0, align="C", fill=False)
        pdf.set_font("font_b", "", 11)
        if oop1 != "":
            pdf.cell(w=W[9], txt=oop1, h=H[-1], border=1, ln=0, align="C", fill=False)
            pdf.cell(w=W[10], txt=oop2, h=H[-1], border=1, ln=1, align="C", fill=False)
        else:
            text = "?" if pd.isnull(data["MoneyCode"]) else data["MoneyCode"]
            pdf.cell(w=W[9], txt=text, h=H[-1], border=1, ln=0, align="L")
            pdf.cell(w=W[10], txt="", h=H[-1], border=1, ln=1, align="L")


def add_money_descriptions(pdf):
    """
    Adds money code descriptions to the bottom of the page
    """
    y_start = 8.5 - BOTTOM_MARGIN - 5.5 * MH
    pdf.set_y(y_start)
    pdf.set_font("font_b", "", 10)
    text = (
        'Definition for award codes shown in "Out of Pocket Cost" column: '
        + "average unmet need for 0 EFC students BEFORE loans"
    )
    pdf.cell(w=sum(W), txt=text, h=MH, border=1, ln=1, align="L", fill=False)
    money_descriptions = [
        "+++: <$5,000 (i.e. no family contribution after Stafford loans)",
        "++: $5,000-$8,000 (i.e. no more than $2,500 in need after Stafford)",
        "++/-: Most awards ++, but some not as good",
        "+/--: Most awards are bad, but some are good",
        "+/---: Almost all awards are bad, but we had a few surprises",
        "--: $12,000-$15,000",
        "---: >$15,000",
        "?: We really don't know (but that is more likely bad than not)",
    ]
    for i in range(4):
        text = money_descriptions[i]
        pdf.cell(w=sum(W[:5]), txt=text, h=MH, border=0, ln=0, align="L", fill=False)
        text = money_descriptions[i + 4]
        pdf.cell(w=sum(W[5:]), txt=text, h=MH, border=0, ln=1, align="L", fill=False)
    pdf.set_line_width(THICK_LINE)
    pdf.rect(LEFT_MARGIN, y_start, sum(W), 5 * MH)
    pdf.set_line_width(LINE_WIDTH)


def create_pdfs(dfs, campus, config, debug, single_pdf=True):
    """Will create PDF reports for the campus"""
    if debug:
        print("Creating PDF report for {}".format(campus), flush=True)
    filework.create_folder_if_necessary([config["output_folder"], campus])
    award_df = dfs["award_report"]
    student_df = dfs["student_report"]

    if not single_pdf:
        pdf = initiate_pdf_object()
        campus_fn = (
            campus
            + "_PDF_Decision_Reports_"
            + date.today().strftime("_%m_%d_%Y")
            + ".pdf"
        )
        if debug:
            print(f"Outputing to ({campus_fn})")

    # Loop through the roster
    count = 0
    filenames = []
    for i, student_data in student_df.iterrows():
        student_award = award_df[award_df["SID"] == i]
        if single_pdf:
            student_fn = (
                campus
                + "_"
                + student_data["LastFirst"].replace(" ", "_")
                + "_"
                + str(i)
                + date.today().strftime("_on_%m_%d_%Y")
                + ".pdf"
            )
            pdf = initiate_pdf_object()
        add_pdf_header_row(pdf, student_data, campus)
        if len(student_award) <= 15:
            add_college_rows(pdf, student_award)
            add_money_descriptions(pdf)
        else:
            add_college_rows(pdf, student_award.iloc[:18, :])
            add_pdf_header_row(pdf, student_data, campus, second=True)
            add_college_rows(pdf, student_award.iloc[18:32, :])
            add_money_descriptions(pdf)
        # Add rows
        # Add extra page and extra rows

        if single_pdf:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                this_file = os.path.join(config["output_folder"], campus, student_fn)
                pdf.output(this_file, "F")
            count += 1
            filenames.append(this_file)
            if debug and (count % 10 == 0):
                print(f"{count}..", end="", flush=True)

    if not single_pdf:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            pdf.output(os.path.join(config["output_folder"], campus_fn), "F")
    else:
        # ZIPUP filenames into a single zip file!
        campus_fn = (
            campus
            + "_Single_Decision_Reports_"
            + date.today().strftime("_%m_%d_%Y")
            + ".zip"
        )
        with zipfile.ZipFile(
            os.path.join(config["output_folder"], campus_fn), "w", zipfile.ZIP_DEFLATED
        ) as myzip:
            for file in filenames:
                myzip.write(file)

        if debug:
            print("Done", flush=True)
