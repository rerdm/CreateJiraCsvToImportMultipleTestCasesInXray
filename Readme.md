# CreateJiraCsvToImportMultipleTestCasesInXray

**UPDATE** : 2025.08.07 New version of CreateJiraCsvToImportMultipleTestCasesInXray.py -> CreateJiraCsvToImportMultipleTestCasesInXray_v002.exe

This Program can be used to generate several test-cases in one big csv file wich can be imported into XRAY. Functionality will be explained in the fallowing picture.<br>

![Program Structure](./img/Program_structur_v2.PNG)

###  PROGRAM USAGE (EXECUTABLE)
1. You have to download the executable from executable folder.
2. You have to create these folders before running  the program (Need to be placed in the same location as the executable). <br>
   <b>input</b> -  (contains the test-cases in xlsx format)</b><br>
   <b>output</b> - (empty folder - here the testcases will be stored in csv files)</b><br>
   <b>jira_import </b>- (empty folder - here the combined.csv and the multiple_testcases_file_for_jira will be saved as csv files)</b><br>
<br>The **_combined.csv_** contains all test-cases attached (including all headers). <br>
   The **_multiple_testcases_file_for_jira.csv_** contains all testcases with the correct header and with the correct format for a valid XRAY import.<br>
<br>
3. Open the CMD in the current dir and run the program (otherwise you will not see the loging and the program closes immediately after execution)

### Execute the Program

```yml
C:\xxx\xxx\xxx>CreateJiraCsvToImportMultipleTestCasesInXray_v002.exe
Programm is starting ...
SUCCESS: Konvertiere Test-Cases aus dem [input] ordner und speichere sie im [output] Ordner im csv Format
 - XLSX Test-Cases aus dem Ordner [input] :    001-Test-Case-1.xlsx
 - XLSX Test-Cases aus dem Ordner [input] :    001-Test-Case-2.xlsx
 - XLSX Test-Cases aus dem Ordner [input] :    001-Test-Case-3.xlsx
SUCCESS: Die Datei [combined.csv] wurde im Ordner [jira_import] erstellt
SUCCESS: Die Zeilen wurden erfolgreich aus der [combined.csv] gelöscht :  [4, 7]
 - Zeilen gelöscht um nur einen Unique Header in der Datei zu haben ( Zeile 1 )
 - Dies ist notwendig um ein Xray konformes csv file zu erzeugen.
 - Erstelltes file im ordner [output]:  [multiple_testcases_file_for_jira.csv]
SUCCESS: File erstellt  im ordner [jira_import]:  multiple_testcases_file_for_jira.csv - Mit diesem File können nun mehrere Test-Cases mit einem mal über den Xray-Test-Case-Importer importiert werden.
```

 

