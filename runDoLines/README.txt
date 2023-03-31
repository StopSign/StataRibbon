===Installation/setup===
1. Extract the contents of the zip file, and put the runDoLines folder in your C:/ drive
2. Open the installer program and run setup.exe
3. When you load excel next, you will see the Stata ribbon available
4. Click Edit Configuration, and modify statapath to match your stata exe location, and statawin to the window title when stata is opened

===How to use===

- Do File
	Change the folder path with Browse, or by entering in a new file path 
	Change the file name by entering it in. The .do will be added the filename automatically.
	Make .do File - Creates a .do file with the given name in the given folder. Takes information from the Header sheet and Code sheet

- Edit Configuration
	Opens the rundo.ini file
	
- Run Stata Lines 
	Runs the selected lines in Stata. Can only select one column, so select the code column you want.
	Opens Stata if it isn't already, otherwise selects the existing Stata window.
	
-Run Stata
	Runs all code in column C from row 2 until theres 5 blank rows in a row. 
	Opens Stata if it isn't already, otherwise selects the existing Stata window.
	
===GLD File version===
The code will look on the Header tab, from row 2 until there are multiple blank lines in a row, for the following keywords under the Title column to add to the file:
	"country" will be put into <_Country_>
	"survey name" or "survey title" will be put into <_Survey Title_>
	"year" will be put into <_Survey Year_>
	"study id" or "survey source" will be put into <_Study ID_>
	"data collection from" will be put into <_Data collection from_>
	"data collection to" will be put into <_Data collection to_>
	"source of dataset" or "unit of analysis" will be put into <_Source of dataset_>
	"sample size" and "hh" will be put into <_Sample size (HH)_>
	"sample size" and "ind" will be put into <_Sample size (IND)_>
	"sampling method" will be put into <_Sampling method_>
	"geographic coverage" <_Geographic coverage_>
	"currency"<_Currency_>
	"version" and "icls" will be put into <_ICLS Version_>
	"version" and "isced" will be put into <_ISCED Version_>
	"version" and "isco" will be put into <_ISCO Version_>
	"version" and "occup" will be put into <_OCCUP National_>
	"version" and "isic" will be put into <_ISIC Version_>
	"version" and "indus" will be put into <_INDUS National_>

===How to uninstall===
Use the Windows program Add or Uninstall Programs, and uninstall StataRibbon


===Steps to modify the program===
This project was coded in C#, using Microsoft Visual Studio 2022

I setup the project and code base using this guide: https://docs.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-a-custom-tab-by-using-ribbon-xml?view=vs-2022&tabs=csharp 
Following that, you can setup the IDE to load the code, for modification from there.

This code also uses runDo, an AutoIt program described here https://huebler.blogspot.com/2008/04/stata.html 
With a command line, you can do "start rundo.exe FILE.do" and it will load FILE in Stata. This done programmatically in C# 

Deploying:
Following this guide: https://docs.microsoft.com/en-us/visualstudio/vsto/deploying-an-office-solution?view=vs-2022 
Use the Visual Studio publish option, which creates the folders and .exe. Zip it from there and good to send it!