#-*- coding: utf-8 -*-
# addClassH.py

import logging
import os.path, shutil
import glob
import sqlite3
from openpyxl import load_workbook
import getopt, sys

# delete the previous log file
logfile = 'addClassH.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of addClassH.exe >>>>>')

def usage():
   print("Usage: addClassH -y year -m month -t data_folder")

try:
   opts, args = getopt.getopt(sys.argv[1:], "y:m:t:")
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
   elif o == "-t":
      tuitionDir = a
   else:
      assert False, "unhandled option"

# open the afterSchool database
dbName = 'asClassH.db'
if not os.path.isfile(dbName):
   logging.error("The database file \'asClassH.db\' not found.")
   exit(1)

try:
   db = sqlite3.connect(dbName)
   cur = db.cursor()

   xlsPath = os.path.join(tuitionDir,'*.xlsx')
   xlList = glob.glob(xlsPath)
   for xlFile in xlList:
      xlsName = xlFile[len(tuitionDir)+1:].decode('cp949')
      logging.info(u'')
      logging.info(u'<<< ' + xlsName + u' >>>')
      wb = load_workbook(filename=xlFile, read_only=True)
      for sheet in wb:
         logging.info(u'< ' + sheet.title + u' >')
         bBreak = False
         for row in sheet.rows:
            for cell in row:
               if cell.value and -1 < cell.value.find(u'학년'):
                  rowFirst = cell.row +1
                  colFirst = cell.column -1  # base 0
                  bBreak = True
                  break
            if bBreak:
               break
         if bBreak:
            # class
            className = xlsName[:xlsName.find(u'(')]
            t = (className, cYear, cMonth)
            cur.execute("SELECT id FROM afterSchoolClass WHERE cname = ? AND year = ? AND month = ?", t)
            r = cur.fetchone()
            if r is not None:
               classId = r[0]
            else:
               cur.execute("INSERT INTO afterSchoolClass(cname,year,month) values(?,?,?)", t)
               classId = cur.lastrowid
               logging.info('A class inserted: %s,%s,%s,%s', classId,className,cYear,cMonth)

            for row in sheet.rows:
               if row[0].row < rowFirst:
                  continue
               # student
               if row[colFirst].value and row[colFirst +1].value and row[colFirst +2].value and row[colFirst +3].value:
                  if type(row[colFirst].value) is unicode:
                     stuGrade = row[colFirst].value.replace(u'학년',u'').strip()
                  else:
                     stuGrade = row[colFirst].value
                  if type(row[colFirst +1].value) is unicode:
                     stuClass = row[colFirst +1].value.replace(u'반',u'').strip()
                  else:
                     stuClass = row[colFirst +1].value
                  t = (cYear,stuGrade,stuClass,row[colFirst +2].value,row[colFirst +3].value)
                  cur.execute("SELECT id FROM student WHERE year = ? AND grade = ? AND class = ? AND odr = ? AND name = ?", t)
                  r = cur.fetchone()
                  if r is not None:
                     stuId = r[0]
                  else:
                     t = (row[colFirst +3].value, cYear, stuGrade, stuClass, row[colFirst +2].value)
                     #logging.debug('Try to insert a student: %s,%s,%s,%s,%s', cYear, stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)
                     cur.execute("INSERT INTO student(name,year,grade,class,odr) values(?,?,?,?,?)", t)
                     stuId = cur.lastrowid
                     logging.info('A student inserted: %s,%s,%s,%s,%s,%s', stuId,cYear, stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)

                  # class of the student
                  t = (classId,stuId)
                  cur.execute("SELECT id,tuition,mcost FROM classStu WHERE classId = ? AND stuId = ?", t)
                  r = cur.fetchone()
                  if r is None: #insert
                     t = (classId,stuId)
                     cur.execute("INSERT INTO classStu(classId,stuId) values(?,?)", t)
                     classStuId = cur.lastrowid
                     logging.info('A student of class inserted: %s,%s,%s', classStuId,classId,stuId)

         else: # not found '학년'
            logging.info(u'Worksheet \'' + sheet.title + u'\' may not have any student.')

   cur.close()
   db.commit()
   db.close()

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info("Adding afterSchool classes done.")
