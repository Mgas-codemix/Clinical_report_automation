# A Guide to This Project:

In this example, I will be showing how to generate a report on set of psuedo-Patients. I use an excel file as my input.

## Files and Folders explained:

  - GeneratedReports: This directory will house any of the final reports generated.
  - PsuedoData: Contains a single excel of made up patients with random mutations from gnomAD database.
  - ReportTemplate: This directory contains the simple template report template
  - script: Contains a Python script that turns the excel data into set of reports.

### NGSexcel2Doc.py Script:

The script can be viewed by opening the the python file in a text editor. If running on the command line the man can be viewed with:
```
  python3 script/NGSexcel2Doc.py --h
```

This returns:
```
  usage: NGSexcel2Doc.py [-h] -inFile INFILE -template TEMPLATE -outdir OUTPUT
  
  options:
    -h, --help          show this help message and exit
    -inFile INFILE      Path to input file containing NGS excel file
    -template TEMPLATE  Path to the word template file
    -outdir OUTPUT      Path to output directory. This is the directory where all reports will be written too
```
To run the script on the commandline:
```
  python3 NGSexcel2Doc.py -inFile <path to input excel> -template <path to template docx> -outdir <path to output dir>
```
