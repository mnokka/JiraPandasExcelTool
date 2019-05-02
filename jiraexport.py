#
#
#


from jira import JIRA
from datetime import datetime
import logging as log
import pandas
import argparse
import getpass
import time
import sys, logging
from author import Authenticate  # no need to use as external command
from author import DoJIRAStuff

start = time.clock()
__version__ = u"0.1.RISKS" 

# should pass via parameters
#ENV="demo"
ENV=u"PROD"

logging.basicConfig(level=logging.DEBUG) # IF calling from Groovy, this must be set logging level DEBUG in Groovy side order these to be written out


def main():

    
    JIRASERVICE=u""
    JIRAPROJECT=u""
    PSWD=u''
    USER=u''
  
    logging.debug (u"--Python starting checking Jira issues for attachemnt adding --") 

 
    parser = argparse.ArgumentParser(usage="""
    {1}    Version:{0}     -  mika.nokka1@gmail.com 
    
    USAGE:
    -filepath  | -p <Path to Excel file directory>
    -filename   | -n <Excel filename>

    """.format(__version__,sys.argv[0]))

    #parser.add_argument('-f','--filepath', help='<Path to attachment directory>')
    parser.add_argument('-q','--excelfilepath', help='<Path to excel directory>')
    parser.add_argument('-n','--filename', help='<Excel filename>')
    
    parser.add_argument('-v','--version', help='<Version>', action='store_true')
    
    parser.add_argument('-w','--password', help='<JIRA password>')
    parser.add_argument('-u','--user', help='<JIRA user>')
    parser.add_argument('-s','--service', help='<JIRA service>')
    #parser.add_argument('-p','--project', help='<JIRA project>')
    #parser.add_argument('-z','--rename', help='<rename files>') #adhoc operation activation
    #parser.add_argument('-x','--ascii', help='<ascii file names>') #adhoc operation activation
        
    args = parser.parse_args()
    
    if args.version:
        print 'Tool version: %s'  % __version__
        sys.exit(2)    
           
    #filepath = args.filepath or ''
    excelfilepath = args.excelfilepath or ''
    filename = args.filename or ''
    
    JIRASERVICE = args.service or ''
    #JIRAPROJECT = args.project or ''
    PSWD= args.password or ''
    USER= args.user or ''
    #RENAME= args.rename or ''
    #ASCII=args.ascii or ''
    
    # quick old-school way to check needed parameters
    if (JIRASERVICE=='' or PSWD=='' or USER==''  or excelfilepath=='' or filename==''):
        parser.print_help()
        print "args: {0}".format(args)
        sys.exit(2)
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Authenticate(JIRASERVICE,PSWD,USER)
    jira=DoJIRAStuff(USER,PSWD,JIRASERVICE)
    
    
    sys.exit(0)
    return
    
    #
    log.basicConfig(filename='update-duedate.log',level=log.INFO)

    JQL_multiple = []
    JQL_none = []

    args = parse_arguments()
    excel_file = args.excel_file
    jira_url = args.jira_url
    username = args.username
    password = getpass.getpass()

    jira = JIRA(server=jira_url, basic_auth=(username, password))

    # Local test to check notifications
    """
    issue_list = jira.search_issues("Summary ~ 'Component'")
    for issue in issue_list:
        issue.update(duedate="2019-10-10")
    """
    df = pandas.read_excel(excel_file)

    # HELPER FUNCTIONS
    #
    def update_issue_duedate(issue, new_duedate):
        issue.update(duedate=new_duedate)
        log.info("ISSUE: {0}, {1}, {2} | New due date: {3}".format(issue.key, issue.fields.summary, issue.fields.customfield_10019, new_duedate))

    # MAIN()
    #
    log.info("Starting update-duedate.py - {0}".format(datetime.now()))
    log.info("Jira: {0}".format(jira_url))
    log.info("Excel parsing:")
    
    for index, row in df.iterrows():
        document_number = row['Document Number']
        new_duedate = row['Due date'].to_pydatetime().isoformat()

        # Drawing Number == cf[10019]
        issue_list = jira.search_issues("'Drawing Number' ~ '{0}'".format(document_number))
        if len(issue_list) == 1:
            for issue in issue_list:
                update_issue_duedate(issue, new_duedate)

        elif len(issue_list) > 1:
            log.info("More than 1 issue was returned by JQL query: {0}".format(document_number))
            JQL_multiple.append(document_number)

        else:
            log.info("No issue(s) returned by JQL query: {0}".format(document_number))
            JQL_none.append(document_number)

        time.sleep(0.7)

    log.info("Count of document numbers that returned more than one issue: {0}".format(len(JQL_multiple)))
    for doc_num in JQL_multiple:
        log.info(doc_num)
    log.info("Count of document numbers that returned no issues: {0}".format(len(JQL_none)))
    for doc_num in JQL_none:
        log.info(doc_num)

    log.info("Stopped update-duedate.py - {0}".format(datetime.now()))
    

def parse_arguments():
    parser = argparse.ArgumentParser(description='')
    parser.add_argument("--jira-url", required=True)
    parser.add_argument("--excel-file", required=True)
    parser.add_argument("--username", required=True)

    return parser.parse_args()

if __name__ == '__main__':
    main()