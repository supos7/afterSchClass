#-*- coding: utf-8 -*-
# compStuDB.py

import logging
import os.path, shutil
import glob
import sqlite3
from openpyxl import load_workbook
import getopt, sys

# delete the previous log file
logfile = 'compStuDB.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of compStuDB.exe >>>>>')

def usage():
   print("Usage: compStuDB -y year -m month -h hj_stu_DB")

try:
   opts, args = getopt.getopt(sys.argv[1:], "y:m:h:")
except getopt.GetoptError, err:
   # print help information and exit:
   print str(err) # will print something like "option -a not recognized"
   usage()
   os._exit(1)
if 3 != len(opts):
   usage()
   os._exit(1)
for o, a in opts:
   if o == "-y":
      cYear = a
   elif o == "-m":
      cMonth = a
   elif o == "-h":
      hj_dbName = a
   else:
      assert False, "unhandled option"

# open the afterSchool database
dbName = 'asClass.db'
if not os.path.isfile(dbName):
   logging.error("The database file \'asClass.db\' not found.")
   exit(1)
if not os.path.isfile(hj_dbName):
   logging.error("The database file \'" + hj_dbName + "\' not found.")
   exit(1)

try:
   dbNames = [dbName, hj_dbName]
   dbTitles = [u'방과후 행정사 파일', u'행정실 파일']
   for i in range(0,2):
      db = sqlite3.connect(dbNames[0])
      cur = db.cursor()
      dbCmp = sqlite3.connect(dbNames[1])
      curCmp = dbCmp.cursor()

      logging.info(u'')
      logging.info(dbTitles[0] + u'에는 있고, ' + dbTitles[1] + u'에는 없음')
      t = (cYear,cMonth)
      cur.execute("SELECT cname,grade,class,odr,name FROM afterSchStu WHERE year=? AND month=? ORDER BY classId,grade,class,odr", t)
      for row in cur:
         t = (cYear,cMonth,row[0],row[1],row[2],row[3],row[4])
         curCmp.execute("SELECT name FROM afterSchStu WHERE year=? AND month=? AND cname=? AND grade=? AND class=? AND odr=? AND name=?", t)
         rowCmp = curCmp.fetchone()
         if rowCmp is None:
            logging.info(u'%s: %s학년 %s반 %s번 %s', row[0],row[1],row[2],row[3],row[4])

      curCmp.close()
      cur.close()
      dbCmp.close()
      db.close()
      
      # swap
      dbNames[0],dbNames[1] = dbNames[1],dbNames[0]
      dbTitles[0],dbTitles[1] = dbTitles[1],dbTitles[0]

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info('Comparing databases done.')
