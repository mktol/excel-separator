# excel-separator
jar to separate big xlsx file to number of smaller.
Current version support only one sheet. 

how to use:

mvn clean install

run generated file in next way: 

java -jar excel-separator-1.0-SNAPSHOT-spring-boot.jar "E:\Users\User\Documents\largeFile.xlsx" 15000
--------------------------------------------------------------------------------------------------------- 
"C:\Users\User\Documents\largeFile.xlsx" - path to large xlsx file.
15000 - number of lines in each new generated file.

expected result for file wiht name largeFile.xlsx is 
largeFile_1.xlsx
largeFile_2.xlsx
...
largeFile_n.xlsx

