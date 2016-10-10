#-*- coding: utf-8 -*-
# checkFreePass.py

import logging
import os.path, shutil
import glob
import sqlite3
from openpyxl import load_workbook
import getopt, sys

# delete the previous log file
logfile = 'checkFreePass.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of checkFreePass.exe >>>>>')

def usage():
   print("Usage: checkFreePass -y year -m month -f file")

try:
   opts, args = getopt.getopt(sys.argv[1:], "y:m:f:")
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
   elif o == "-f":
      xlFile = a
   else:
      assert False, "unhandled option"

# open the afterSchool database
dbName = 'asClass.db'
if not os.path.isfile(dbName):
   logging.error("The database file \'asClass.db\' not found.")
   exit(1)

try:
   db = sqlite3.connect(dbName)
   cur = db.cursor()

   xlsName = xlFile.decode('cp949')
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

         for row in sheet.rows:
            if row[colFirst].row < rowFirst:
               logging.debug('this line skipped: %s,%s,%s,%s,%s', row[colFirst].value, row[colFirst +1].value, row[colFirst +2].value, \
                  row[colFirst +3].value, row[colFirst +4].value)
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
               cur.execute("SELECT id,code FROM student WHERE year=? AND grade=? AND class=? AND odr=? AND name=?", t)
               r = cur.fetchone()
               if r is not None: # update code
                  code = r[1]
                  if code == 'FP' or code == 'FPN':
                     logging.info('The student found as a free pass: %s,%s,%s,%s', stuGrade,stuClass,row[colFirst +2].value,row[colFirst +3].value)
                  #elif code == 'FPNJ':
                  #   logging.info('The student not joined: %s,%s,%s,%s', stuGrade,stuClass,row[colFirst +2].value,row[colFirst +3].value)
                  else:
                     logging.warn('Check if the student is a free pass: %s,%s,%s,%s', stuGrade,stuClass,row[colFirst +2].value,row[colFirst +3].value)
               else: # student not found
                  t = (row[colFirst +3].value, cYear, stuGrade, stuClass, row[colFirst +2].value,'FP')
                  cur.execute("INSERT INTO student(name,year,grade,class,odr,code) values(?,?,?,?,?,?)", t)
                  stuId = cur.lastrowid
                  logging.info('A student inserted as a free pass: %s,%s,%s,%s,%s,%s', stuId, cYear, stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)
            else:
               logging.debug('Invalid data: %s,%s,%s,%s,%s', row[colFirst].value, row[colFirst +1].value, row[colFirst +2].value, \
                  row[colFirst +3].value, row[colFirst +4].value)

      else: # not found '학년'
         logging.info(u'Worksheet \'' + sheet.title + u'\' may not have any student.')

   cur.close()
   db.commit()
   db.close()

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info("Checking afterSchool free pass done.")
