# JiraPandasExcelTool
Create Jira issues from Excel , using ~~Pandas~~ (openpyxl), with possibility to add links to existing issues in another project


Source Jira --> Target Jira issue copying:

* Excel defines custom fields for to be created Jira issue, including linked issues info
(Jira Imp/Exp plugin used to produce exported excel)


* Tool can check if existing project has excels linked issue info (issue summary) and do the linking (hardcoded link names)

* Thus copy first target project issues, then "linking project" issues using -l option

