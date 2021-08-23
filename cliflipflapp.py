#!/usr/bin/env python

import csv
import sys
import datetime
import string
import re
import itertools
import os
import html as cgi
import urllib.parse

import xlrd

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
    boxes = get_boxes_from_cellfeed(cf)
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

FBLABELS = [ 'Label',
             'Short Identifier',
             'Genotype',
             'X',
             'Y',
             'C2a',
             'C2b',
             'C3a',
             'C3b',
             'Extra info',
             'Mod' ]


class FlyBoxError(Exception):
    pass


class Box(dict):
    def __init__(self):
        default = {
                    'name'    : 'N/A',
                    'urlname' : 'N/A',
                    'flipped' : 'N/A',
                    'calid'   : 'N/A',
                    'labels'  : {}, # dict idx : label
                    'flies'   : [], # list of dicts
                  }
        super(Box, self).__init__(default)



def get_boxes_from_cellfeed(cellfeed):
    #seperate boxes
    BOXES = []
    y_old = y_off = -2
    ym, xm = 0, 0

    for c in cellfeed:
        # get contents
        y = int(c.cell.row)
        x = int(c.cell.col)
        v = cgi.escape(c.content.text)
        if x == 1 and y == 1:
            if v != cgi.escape("WFF:FLYSTOCK"):
                raise FlyBoxError("""<strong>ERROR:</strong> The first cell in the Stocklist is not "WFF:FLYSTOCK". This is a required setting. Please use the template file for your stocklists. '%s'""" % v)
        # skip the first two lines
        if y < 3:
            continue
        # ignore everything after a #WFF-IGNORE
        if x == 1 and v == "WFF:IGNORE":
            break
        # seperate boxes
        if y - y_old < 2:
            y_old = y
        else:
            y_off = y_old = y
            BOXES.append(Box())
        by, bx = y-y_off, x
        lastBox = BOXES[-1]
        lastBox['_width'] = max(bx, lastBox.get('_width', 0))
        lastBox['_height'] = max(by, lastBox.get('_height', 0))
        if by == 0 and bx == 1:
            lastBox['name'] = v
            lastBox['urlname'] = urllib.parse.quote_plus(v)
        if by == 0 and bx == 2: lastBox['flipped'] = v[9:] # ugly hack
        if by == 0 and bx == 3: lastBox['calid'] = v[7:] # ugly hack
        if by == 1: lastBox.setdefault('_labels', {})[bx] = v if v else '&nbsp;'
        if by >= 2: lastBox.setdefault('_elements', {})[(by,bx)] = v if v else '&nbsp;'
    # get flies
    BEENTHERE = []
    for b in BOXES:
        if b['name'] in BEENTHERE:
            raise FlyBoxError("""<strong>Error:</strong> Boxname %s is duplicated in Stocklist""" % b['name'])
        else:
            BEENTHERE.append(b['name'])
        elements = b.pop('_elements', {})
        ylen, xlen = b.pop('_height', 0), b.pop('_width', 0)
        labels = b.pop('_labels', {})
        # get labels
        CHECKLAB = []
        for k in range(xlen):
            ktmp = labels.get(k+1, None)
            if ktmp == '':
                ktmp = None
            b['labels'][k] = ktmp
            CHECKLAB.append(ktmp)
        for i in range(11):
            try:
                if CHECKLAB[i] != FBLABELS[i]:
                    raise IndexError()
            except IndexError:
                raise FlyBoxError("""<strong>Error:</strong> The box %s does not have all Labels! Error with column-label in column #%d: should be &quot;%s&quot; but is &quot;%s&quot;. Please look at the template.""" % (b['name'], i+1, FBLABELS[i], CHECKLAB[i]))
        # get flies
        for j in range(1,ylen):
            fly = {}
            for i in range(xlen):
                if b['labels'][i] is not None:
                    fly[b['labels'][i]] = elements.get((j+1,i+1), '&nbsp;')
            b['flies'].append(fly)

    return BOXES

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
