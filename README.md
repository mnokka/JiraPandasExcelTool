# JiraPandasExcelTool
Create Jira issues from Excel , using Pandas, with possibility to add links to existing issues


Source Jira --> Target Jira issue copying:

* Excel defines custom fields for to be created Jira issue, including linked issues info

* Tool can add linked issues Summary (as reference for later usage) for excel based on links in source Jira

* When executed in create mode, issues are created for target Jira. Any found existing issues (target Jira) based
on added ( source Jira) summary field are being linked using given link type (of course these to be linked issues must have been copied earlier to target Jira)

