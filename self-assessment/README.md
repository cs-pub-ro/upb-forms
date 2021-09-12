# UPB Self Assessment Form Processing

A self assessment form is used inside University POLITEHNICA of Bucharest (UPB) for every staff member.
Each staff member fills this form yearly.

These scripts extract and process data from the form.
Forms have to be in `.xlsx` format (Microsoft OOXML files).

The `extract_self_assessment.py` script extracts data from a single form file.
The `process_folder.py` script process all form files in a given folder.
Typically you would have one folder per year.

## Usage

Assuming a subfolder `2020/` storing self assessment form files (for 2020), use the command below to extract data from all files:

```
$ ./process_folder.py 2020/
Nume	Poziție	Educație	Cercetare	Management	Recunoaștere	Comunitate	Altele
[...]
```

This prints out a column header and then contents: one line per file (i.e. per staff member).
