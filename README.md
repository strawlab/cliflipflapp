# cliflipflapp

This is a port of [webflipflapp](https://github.com/ap--/webflipflapp) by
Andreas Poehlmann to run entirely on the command line interface (CLI) with
minimal dependencies and to run self-contained in a single Python file.

It takes an .xlsx file as input and outputs a .tex file.

```
┌───────────┐  cliflipflapp   ┌─────────┐
│.xlsx file ├────────────────►│.tex file│
└───────────┘                 └─────────┘
```

`cliflipflapp` expects the .xlsx file to have a specific format. If this is not
true, Bad Things will happen.

## Installing prerequisites

Tested with Ubuntu 20.04:

```
sudo apt install python3-xlrd
```

You can probably also use xlrd downloaded with pip.

## Running

```
python3 cliflipflapp.py /path/to/my-fly-sheet.xlsx
```

If all goes will, this will generate the file `/path/to/my-fly-sheet.tex` which
can be converted to PDF.

## Converting to PDF

Although strictly speaking, converting to PDFs is not the task of cliflipflapp,
it is useful to do so, and here are quick instructions:

Tested with Ubuntu 20.04:

### Install `pdflatex` with appropriate packages

```
sudo apt install texlive-latex-base texlive-latex-extra
```

### Run `pdflatex`

Note that you may need to run this **from the path with the `.tex` file**. (In other words, do not specify the path to the `.tex` file as anything other than the current directory.)

```
pdflatex my-fly-sheet.tex
```
