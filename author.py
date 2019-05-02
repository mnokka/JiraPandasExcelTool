# mika.nokka1@gmail.com 2.5.2019
# 
# Jira authorization routines
# to be used via importing only


import datetime 
import time
import argparse
import sys
import netrc
import requests, os
from requests.auth import HTTPBasicAuth
# We don't want InsecureRequest warnings:
import requests
requests.packages.urllib3.disable_warnings()
import itertools, re, sys
from jira import JIRA
import random


__version__ = "0.1"
thisFile = __file__

    
def main(argv):

    print ("Main part started and ending now. Use Jira authorization via importing")
    return

    
####################################################################################################   
# POC skips .netrc usage as used mostly in WIN10 env
# 
def Authenticate(JIRASERVICE,PSWD,USER):
    host=JIRASERVICE
    #credentials = netrc.netrc()
    #auth = credentials.authenticators(host)
    #if auth:
    #    user = auth[0]
    #    PASSWORD = auth[2]
    #    print "Got .netrc OK"
    #else:
    #    print "ERROR: .netrc file problem (Server:{0} . EXITING!".format(host)
    #    sys.exit(1)
    user=USER
    PASSWORD=PSWD

    f = requests.get(host,auth=(user, PASSWORD))
         
    # CHECK WRONG AUTHENTICATION    
    header=str(f.headers)
    HeaderCheck = re.search( r"(.*?)(AUTHENTICATION_DENIED|AUTHENTICATION_FAILED)", header)
    if HeaderCheck:
        CurrentGroups=HeaderCheck.groups()    
        print ("Group 1: %s" % CurrentGroups[0]) 
        print ("Group 2: %s" % CurrentGroups[1]) 
        print ("Header: %s" % header)         
        print "Authentication FAILED - HEADER: {0}".format(header) 
        print "--> ERROR: Apparantly user authentication gone wrong. EXITING!"
        sys.exit(1)
    else:
        print "Authentication OK \nHEADER: {0}".format(header)    
    print "---------------------------------------------------------"
    return user,PASSWORD

###################################################################################    
def DoJIRAStuff(user,PASSWORD,JIRASERVICE):
 jira_server=JIRASERVICE
 try:
     print("Connecting to JIRA: %s" % jira_server)
     jira_options = {'server': jira_server}
     jira = JIRA(options=jira_options,basic_auth=(user,PASSWORD))
     print "JIRA Authorization OK"
 except Exception,e:
    print("Failed to connect to JIRA: %s" % e)
 return jira   
    
####################################################

        
if __name__ == "__main__":
        main(sys.argv[1:])
        