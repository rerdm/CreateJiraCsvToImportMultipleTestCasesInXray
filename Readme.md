# CreateJiraCsvForToImportMultipleTestCasesInXray

### Purpose and explanation of teh program: 

- This Program will take you testcase(s) (xlsx) and create a new clv file wich contains all you testcases .The creation file will be accepted by the Xray-Test-importer.

<b>This program will</b>: 
1. Take your testcases (xlsx) from input folder ( needs to be present in actual dir) convert it to csv and store it in output folder ( need to be present in actual dir) <br>
<br>
3. Create a new file (combined.csv) wich contains all test-cases attached. In this file the number of the testcases will be increased ( <b>NOTE</b>: In XRAY the testcase nummer will be represented via teh Issue ID).<br> Also a new colum 'Summay' will be created and the name of the particular testcase will be saved for every test-case ( because in XRAY the Summary is the name of the Test-Case) <br>
<br>
3. Create a new file (multiple_testcases_file_for_jira.csv). This file only contain 1 headline ( other headlines are removed) to have a valid csv for XRAY.<br>
<br>
4. Take the (multiple_testcases_file_for_jira.csv) and store it in UTF-8 format. XRAY only accept csv files with UTF-8 ( otherwise some carters are not shown properly ).<br>

