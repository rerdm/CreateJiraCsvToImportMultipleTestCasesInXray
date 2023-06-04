# CreateJiraCsvToImportMultipleTestCasesInXray


- This Program can be used to attach several test-cases in one big csv file wich can be imported into XRAY. Functionality will be explained in the fallowing picture.<br> 
  <br><b> NOTES TO THE PROGRAM </b><br> <br>
  - The input files need to have this columns, see picture. <br>
  - A Summary colum will be created with the name of each the test-case. <br>
  - The first column in the created csv represents the test-case numer which need to be mapped to the column "Issue ID" in Xray( All other Columns have the same name as the equivalent columns in XRAY).
<br>
  - These 5 columns are required otherwise the csv will be rejected from XRAY.<br>


![Program Structure](./img/Program_structur.PNG)

