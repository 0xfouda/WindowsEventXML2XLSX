# Windows Event Logs XML to XLSX

---------------------------------------------------------
Converting Windows Events in XML format into XSLX format
---------------------------------------------------------

Tool to convert evtxtract, evtxexport XML output to XSLX format for further forensics analysis. xlsxwriter python library is needed, You can install it using pip 
```
pip install xlsxwriter
```

```
usage: main.py [-h] [--debug] input output

Tool to convert evtxexport, evtxtract XML output into XLSX format

positional arguments:
  input        Path to events XML file
  output       Path to XLSX output file

optional arguments:
  -h, --help   show this help message and exit
  --debug, -d  Show Debugging Information
```
