# FOSSEE_Spreadsheet
Email - ritik.sahu@adypu.edu.in
Username - ritikrks

# New Features!

  - A desktop application that will take inputs for four different categories (modules).
  - GUI shall has a spreadsheet, Load Inputs, Validate and Submit buttons, message box to display warning messages if the user gives a bad value.
  - A spreadsheet of different modules opens in different tabs of the same window.
  - Based on the selected module corresponding header row is displaying in spreadsheet GUI. Details of header rows, with sample input values for each module, are given in resources.
  - “Load Inputs” button prompts for selecting CSV/xlxs file, which will populate the spreadsheet. Also, Users can fill data manually in each row.
  - The clicking of the “Validate” button validated the data and a suitable error message for bad values is displayed in the message box.
  - All cells other than headers takes only numerical inputs.
  - Headers are not editable.
  - ID column shall be unique, i.e., ID number should not be repeated
  - Once the user submits the data by clicking on the “Submit” button, it creates a new text file for each row. This text file is a dictionary with header value as key and cell value as value.
  - Text files are saved in the working folder or you can take folder location from the user.
  - Text files are saved as Modulename_ID. For example, if the user submits fin plate inputs, the first row will be saved as FinPlate_1 automatically, i.e., the user does not have to specify the file name for each row.
  - An easy to use, user-friendly and clean looking GUI application would help the user to quickly adapt to the application.
  - An installer (Windows or Ubuntu) for our application.

### Dependencies/Requirements

FOSSEE_Spreadsheet uses following python packages that one should install to use our App.

* [OS](https://docs.python.org/2/library/os.html?highlight=os#module-os) - This module provides a portable way of using operating system dependent functionality.
* [SYS](https://docs.python.org/2/library/sys.html?highlight=sys#module-sys) - This module provides access to some variables used or maintained by the interpreter and to functions that interact strongly with the interpreter.
* [csv](https://docs.python.org/2/library/csv.html?highlight=csv#module-csv) - The so-called CSV (Comma Separated Values) format is the most common import and export format for spreadsheets and databases.
* [xlrd](https://pypi.org/project/xlrd/) - xtract data from Excel spreadsheets (.xls and .xlsx, versions 2.0 onwards) on any platform. Pure Python (2.7, 3.4+). Strong support for Excel dates. Unicode-aware.
* [PyQt5](https://pypi.org/project/PyQt5/) - Qt is set of cross-platform C++ libraries that implement high-level APIs for accessing many aspects of modern desktop and mobile systems.
* [JSON](https://docs.python.org/2/library/json.html?highlight=json#module-json) - JSON (JavaScript Object Notation), specified by RFC 7159 (which obsoletes RFC 4627) and by ECMA-404, is a lightweight data interchange format inspired by JavaScript object literal syntax (although it is not a strict subset of JavaScript 1 ).

### Installation

Install the dependencies

```sh
$ pip install python-csv
$ pip install xlrd
$ pip install PyQt5
```
### Steps To install 
- Download our project files from Github
- Run Windows/Linux installer named (FOSSEE_Ritik) to unzip the desktop application
- Proceed with installer
- Go to Installed Directory
- Go to dist
- execute the FOSSEE_Ritik.exe Application
- Enjoy
### Documentation
- Download our project files from Github
- Go to Documentation/_build/html Folder
- Start index.html
# Made by Ritik Kumar Sahu for Fossee Summer Fellowship ©

