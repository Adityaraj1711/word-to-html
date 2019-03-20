# HTML Convertor


### HOW TO USE:- 

Place the two files (wordtohtml.ps1 and script.py) in the parent/working directory.

Open powershell in the same directory and then enter command  " ./wordtohtml.ps1 "
and press enter. 
It will start searching all the .docx files recursively and inplace stores the converted html files.


##### author - ZER0-Co0L and devilDock

#
#
#
### Explanation of Code
> wordtohtml.ps1
 --------------
 This file is used to convert word files to html
 Function Wrd-HTML is converting word files to html using Microsoft word API
 and storing html files converted using LibreOffice API in \_\_\_FILES___ directory.
 The function Wrd-HTML needs two arguments : $f , $p
 $f is the location of .docx file which needs to be converted
 $p is the location where the converted file is stored
 
 The $saveformat is using the Microsoft API in which format it should be stored.(Here: wdFormatFilteredHTML)
 
 For converting docx file to html using Libre
 command ->  soffice --headless --convert-to html:"XHTML Writer File:UTF8" $f --outdir $LibreDir
 $LibreDir is the temporary location where the converted html to be stored.
 Start-Sleep bash command delays or halts the program for -s seconds, use -m for milliseconds.
 Python script is called which takes command line arguemnt $p (location of html file(MS_office's))
 
 An iterative call is then made to a python script with command line args as location of each html file.

 > script.py
 --------
 This file takes an argument which is the location of html file converted by Microsoft office and then searches for \_\_\_FILES___ folder where the html file of libreoffice present. The file will merge the mathml tags into the html(of MSoffice) and then remove the redundant images of equations and checks if the \_\_\_FILES\_\_\_ folder is empty or not, if empty, the folder is then deleted.
 The logs of the conversion will be stored in a conversion.log file in the parent directory
 Any errors or exception will pe prefixed by '****************'.
 