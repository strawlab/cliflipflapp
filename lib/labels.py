#!/usr/bin/env python

import csv
import sys
import datetime
import string
import random
import re
import itertools
import os

import xlrd

import flyboxes

def replace_latex_cmd_chars(string):
    CC = [ ('\\', '\\textbackslash '),
           ('&nbsp;', ' '), # needs to be before & replacement.
           ('&gt;', '>'),
           ('&lt;', '<'),
           ('_', '\\textunderscore '),
           ('&', '\\& '),
           ('%', '\\% '),
           ('#', '\\# '),
           # (' ', '\\_'), do not replace spaces
           ('{', '\\{ '),
           ('}', '\\} '),
           ('^', '\\textasciicircum '),
           ('$', '\\textdollar '), ]
    for c, lc in CC:
        string = string.replace(c, lc)
    return string.strip()


def crop_string(string, N):
    if len(string) > N:
        string = string[:N-2] + '...'
    return string


def label(fields):
    LIMITS = [12,5,27,80,30]
    fields = map(lambda s: crop_string(*s), zip(fields, LIMITS))
    fields = map(replace_latex_cmd_chars, fields)
    return "\\prettylabel{%s}{%s}{%s}{%s}{%s}" % tuple(fields)


def row2fields(row, select, override=(None,)*5):
    fields = []
    for i in range(5):
        if override[i] is not None:
            fields.append(override[i])
        else:
            fields.append( row[select[i]] if select.get(i) is not None else '' )
    return fields

TEMPLATE_US_START = [
    "\\documentclass[letter,10pt]{article}",
    "\\usepackage[utf8]{inputenc}",
    "\\usepackage{lmodern}",
    "\\usepackage[T1]{fontenc}",
    "\\usepackage{array}",
    "\\usepackage{seqsplit}",
    "\\renewcommand{\\tabcolsep}{1mm}",
    "\\linespread{0.4}",
    "\\usepackage[newdimens]{labels}",
    "\\LabelCols=4",
    "\\LabelRows=15",
    "\\LeftPageMargin=8mm",
    "\\RightPageMargin=8mm",
    "\\TopPageMargin=15mm",
    "\\BottomPageMargin=12mm",
    "\\InterLabelColumn=7.5mm",
    "\\InterLabelRow=0.0mm",
    "\\LeftLabelBorder=2mm",
    "\\RightLabelBorder=2mm",
    "\\TopLabelBorder=0.0mm",
    "\\BottomLabelBorder=0.0mm",
    "\\newcommand{\\prettylabel}[5]{%",
    "\\genericlabel{%",
    "\\begin{tabular}{|l r| @{}m{0mm}@{}}",
    "\\hline",
    "~ & ~ & ~ \\\\[-1.5mm] % ",
    "\\multicolumn{3}{|l|}{\\textsf{\\footnotesize{#3}}} \\\\[0.5mm] ",
    "\\multicolumn{2}{m{38.7mm}}{\\texttt{\\tiny{\\seqsplit{#4}}}} & \\rule{0pt}{5mm} \\\\",
    "\\texttt{\\textbf{\\tiny{#5}}} & %",
    "\\sffamily{\\textbf{\\large{#1}}\\textit{#2}} & ~ \\\\",
    "\\hline",
    "\\end{tabular}",
    "}",
    "}",
    "",
    "\\begin{document}\n" ]

TEMPLATE_A4_START = [
    "\\documentclass[a4paper,10pt]{article}",
    "\\usepackage[utf8]{inputenc}",
    "\\usepackage{lmodern}",
    "\\usepackage[T1]{fontenc}",
    "\\usepackage{array}",
    "\\usepackage{seqsplit}",
    "\\renewcommand{\\tabcolsep}{1mm}",
    "\\linespread{0.4}",
    "\\usepackage[newdimens]{labels}",
    "\\LabelCols=4",
    "\\LabelRows=16",
    "\\LeftPageMargin=8mm",
    "\\RightPageMargin=8mm",
    "\\TopPageMargin=15mm",
    "\\BottomPageMargin=12mm",
    "\\InterLabelColumn=0.0mm",
    "\\InterLabelRow=0.0mm",
    "\\LeftLabelBorder=2mm",
    "\\RightLabelBorder=1mm",
    "\\TopLabelBorder=0.0mm",
    "\\BottomLabelBorder=0.0mm",
    "\\newcommand{\\prettylabel}[5]{%",
    "\\genericlabel{%",
    "\\begin{tabular}{|l r| @{}m{0mm}@{}}",
    "\\hline",
    "~ & ~ & ~ \\\\[-1.5mm] % ",
    "\\multicolumn{3}{|l|}{\\textsf{\\footnotesize{#3}}} \\\\[0.5mm] ",
    "\\multicolumn{2}{m{42.7mm}}{\\texttt{\\tiny{\\seqsplit{#4}}}} & \\rule{0pt}{5mm} \\\\",
    "\\texttt{\\textbf{\\tiny{#5}}} & %",
    "\\sffamily{\\textbf{\\large{#1}}\\textit{#2}} & ~ \\\\",
    "\\hline",
    "\\end{tabular}",
    "}",
    "}",
    "",
    "\\begin{document}\n" ]
#TEMPLATE_LABEL = "\\prettylabel{%s}{%s}{%s}{%s}{%s}"
TEMPLATE_SKIP = ["\\addresslabel{}"]
TEMPLATE_STOP = ["\n\\end{document}"]



def get_tex(flies, skip=0, template='a4', repeats=1):

    if template == 'us':
        TEMPLATE_START = TEMPLATE_US_START
    else:
        TEMPLATE_START = TEMPLATE_A4_START

    SELECT = { 0 : 'Label', 1 : None, 2 : 'Short Identifier', 3 : 'Genotype', 4 : None}
    OVERRIDE = ( None, None, None, None, datetime.datetime.now().strftime('%Y-%m-%d') )

    try:
        repeats = max(min(100, repeats), 1)
    except:
        repeats = 1

    LABELS = []
    for fly in flies:
        fields = row2fields(fly, SELECT, OVERRIDE)
        LABELS.extend( [label(fields)]*repeats )

    return "\n".join( TEMPLATE_START + (TEMPLATE_SKIP*skip) + LABELS + TEMPLATE_STOP )


def create_output(flies, template='a4', skip=0, repeats=1):
    """Creates the output files requested by the user"""
    tex = get_tex(flies, skip=int(skip), template=template, repeats=repeats)
    try:
        tex = unicode(tex, 'utf-8')
    except:
        pass
    content_type, data = 'text/plain', tex
    return content_type, data


def doit(xlsx_content):
    bn=None
    ps = 'a4'
    cf = fakecellfeed_from_ssid(xlsx_content)
    boxes = flyboxes.get_boxes_from_cellfeed(cf)
    flies = []
    for box in boxes:
        if (bn is not None) and box['name'] != bn:
            continue
        flies.extend(box['flies'])

    # Give me my output!!!
    content_type, data = create_output(flies, template=ps)
    return data

class YX(object):
    def __init__(self, y, x):
        self.row = y
        self.col = x

class CT(object):
    def __init__(self, v):
        self.text = v

class Fakecell(object):
    def __init__(self, y, x, v):
        self.cell = YX(y, x)
        self.content = CT(v)

def fakecellfeed_from_ssid(xlsx_content):
    xls_spreadsheet = xlrd.open_workbook(file_contents=xlsx_content)
    sheet = xls_spreadsheet.sheet_by_index(0)
    ROWS = sheet.nrows
    COLS = sheet.ncols
    CCC = []
    for y, x in itertools.product(range(ROWS), range(COLS)):
        cell = sheet.cell(y, x)
        try:
            cv = str(cell.value)
        except:
            cv = ''
        if len(cv) > 0:
            CCC.append(Fakecell(y+1,x+1, cv))
    return CCC

if __name__=='__main__':
    xlsx_name = sys.argv[1]
    suffix = '.xlsx'
    assert xlsx_name.endswith(suffix)

    head,tail = os.path.split(xlsx_name)
    core_name = tail[:-len(suffix)]
    tex_name = os.path.join(head,core_name + '.tex')

    with open(xlsx_name,mode='rb') as fd:
        xlsx_content = fd.read()
        fd.close()

    tex_contents = doit(xlsx_content)
    print(f'Saving output to: {tex_name}')
    with open(tex_name,mode='w') as out_fd:
        out_fd.write(tex_contents)
        out_fd.close()
    # install with "sudo apt install texlive-latex-base texlive-latex-extra"
    print(f'Convert to pdf with: pdflatex {tex_name}')
