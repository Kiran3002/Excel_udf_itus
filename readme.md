# Excel Financial Data Connector (Index Constituent UDFs)

## Introduction

This project implements a robust and performant system that exposes powerful financial data retrieval capabilities directly within Microsoft Excel via User-Defined Functions (UDFs). The system is designed to allow financial analysts to query structured financial index constituent data (including weights, sector, and market capitalization category) using simple, formula-based syntax directly in an Excel cell, simulating a seamless integration experience.

## install requirements.txt
```
$ pip install -r requirements.txt
```

## install xlwings addin
xlwings lets you call Python functions directly from Excel, or manipulate Excel workbooks via Python code
```
$ pip install xlwings
```
Addin installation for Excel
```
$  xlwings addin install
```
Confirmation and shows installed path for xlwings
```
$ pip show xlwings
```
## go to excel-> file -> options -> Trust Center -> Trust center settings -> Macro settings -> give access to Trust access to the VBA project object module

## go to excel-> file -> options -> Add-ins -> manage:Excel addins -> click on go -> browse to your xlwings addin path and select xlwings.xlam -> click ok
## In Excel -> click alt-F11 -> VBA tool opens -> in toolbar click on tools -> go to references -> make sure to tick xlwings -> if not present browse to xlwings addin folder and select xlwing.xlam file and click ok
## excel -> xlwings tab -> intterpreter -> makesure path to python is present else add it
## excel -> xlwings tab -> pythonpath -> add project path

## excel -> xlwings tab -> UDF modules -> give python file name to run in project folder , dont add .py just file name

## click on import functions

# use functions in excel
