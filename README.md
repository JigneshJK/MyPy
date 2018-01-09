#MongoDB integration with Excel using Python

+ Problem Statement:

Query MongoDB from Excel - VBA Macro & use its output for automation Validations

+ Solution:

As VBA don’t have direct connectivity option for any NoSQL database, like we have for SQL- we have to use some interface to bride between Excel & MongoDB

+ Available Options:

C# (.dll)
Java (.jar)
Python (xlwings)

+ Why Python?

Only one time configuration required with Excel using XLWings
Readymade library available – PyMongo to interact with MongoDB
Can use python for any further enhancements if not doable in vba

+ High Level Architecture:

Excel Macro : (XLWings, Panda, Pywin32) <> Python : (PyMongo) <> MongoDB

+ Technologies Used:

Excel VBA Macro: Framework to trigger overall workflow

XLWings: XLWings is a module to allow Excel to be automated with Python instead of VBA. Used for connecting Python and Excel Macro

Panda: Pandas is a Python package providing fast, flexible, and expressive data structures designed to make working with “relational” or “labeled” data both easy and intuitive. It aims to be the fundamental high-level building block for doing practical, real world data analysis in Python. Used for Integrating Xlwings with Python

PyWin32: Python extensions for Microsoft Windows Provides access to much of the Win32API (run application on all versions of windows), & ability to create and use COM Objects (allows objects to interact across process and computer boundaries as easily as within a single process) and the Pythonwin environment. Required for interacting with Excel (any windows application) & Python

Python: It is powerful scripting language when it comes to interact with analytical data, so there is two way connectivity for python & excel using XLWings & OpenPyXl. Also direct library available - PyMongo to interact with MongoDB. Used to bridge between MongoDB & Excel

PyMongo: It is a Python distribution containing tools for working with MongoDB, It is is official Python driver for MongoDB created and maintained by MongoDB Inc & includes all capability to interact with MongoDB from Python. Used to connect MongoDB & access data from it. Used to connect MongoDB & access data from it

+ Setup & Installation Steps :

Excel Macro : Excel must be available with Macro enabled

Python : Download & install python 3.6.4 from - https://www.python.org/downloads/ (Make sure to check add python to path option while installing)

Go to command prompt Type command >python to check if path added properly (should show version info)

PyMongo: Install from command prompt as > pip install pymongo

Pandas: Install from command prompt as > pip install pandas

Pypiwin32: Install from command prompt as > pip install pypiwin32

XlWings: Install from command prompt as > pip install xlwings

xlwings.xlam : go to excel & add reference for xlwings. It is one time activity required to run xlwings modules in vba & also for xlwings ribbon

+ Working :

Mongo_Qry (Excel - VBA): 
1. Call Python using Xlwings with required parameters
2. Create Data Folders
3. Get .Jason File as output

My_Pyintegration (main-Python) :
1. Connect to mongodb using con string through pymongo
2. Convert key:value string to PyDictionary
3.Call respective qry.py
4. Get cursor as output & convert it to key:value string & store in .jason file

qry.py (qry file - Python):
1. Get query parameters from dictionary passed from main
2. Connect to respective DB & Collection
3. Execute Mongoquery using pymongo
4. Return cursor to main.payhon

