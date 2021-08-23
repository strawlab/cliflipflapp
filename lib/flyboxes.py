
import html as cgi
import datetime
import re
import random
import string
import urllib.parse


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
