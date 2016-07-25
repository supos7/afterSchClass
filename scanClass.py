#-*- coding: utf-8 -*-
# scanClass.py

import logging
import os.path, shutil
import glob
import sqlite3
from openpyxl import load_workbook
import getopt, sys

# delete the previous log file
logfile = 'scanClass.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of scanClass.exe >>>>>')

def usage():
   print("Usage: scanClass -y year -m cMonth -t tuition_folder -c mcost_folder")

try:
   opts, args = getopt.getopt(sys.argv[1:], "y:m:t:c:")
except getopt.GetoptError, err:
   # print help information and exit:
   print str(err) # will print something like "option -a not recognized"
   usage()
   os._exit(1)
if 4 != len(opts):
   usage()
   os._exit(1)
for o, a in opts:
   if o == "-y":
      cYear = a
   elif o == "-m":
      cMonth = a
   elif o == "-t":
      tuitionDir = a
   elif o == "-c":
      mcostDir = a
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

   folders = (tuitionDir,mcostDir)
   for i in range(0,2):
      # find the month
      folder = folders[i].decode('cp949')
      ie = folder.find(u'월')
      if -1 == ie:
         logging.error('Month not found from the folder name')
         exit(1)
      ib = folder.rfind(u' ') +1
      month = int(folder[ib:ie])
      bPrevMonth = int(cMonth) != month
      if bPrevMonth:
         prevMonth = month -1
         if 0 == prevMonth:
            prevMonth = 12
         logging.debug('Month: %d, prev. month: %d', month, prevMonth)

      xlsPath = os.path.join(folders[i],'*.xlsx')
      xlList = glob.glob(xlsPath)
      for xlFile in xlList:
         xlsName = xlFile[len(folders[i])+1:].decode('cp949')
         logging.info(u'')
         logging.info(u'<<< ' + xlsName + u' >>>')
         wb = load_workbook(filename=xlFile, read_only=True)
         for sheet in wb:
            #logging.info(u'< ' + sheet.title + u' >')
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
               className = xlsName[:xlsName.find(u' ')]
               t = (className, cYear, cMonth)
               cur.execute("SELECT id FROM afterSchoolClass WHERE cname=? AND year=? AND month=?", t)
               r = cur.fetchone()
               if r is not None:
                  classId = r[0]
               else:
                  cur.execute("INSERT INTO afterSchoolClass(cname,year,month) values(?,?,?)", t)
                  classId = cur.lastrowid
                  #logging.info('A class inserted: %s,%s,%s,%s', classId,className,cYear,cMonth)

               for row in sheet.rows:
                  if row[0].row < rowFirst:
                     continue
                  # student
                  if row[colFirst].value and row[colFirst +1].value and row[colFirst +2].value and \
                     row[colFirst +3].value and row[colFirst +4].value:
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
                     if r is not None:
                        stuId = r[0]
                        sCode = r[1]
                     else:
                        #t = (row[colFirst +3].value, cYear, stuGrade, stuClass, row[colFirst +2].value)
                        #logging.debug('Try to insert a student: %s,%s,%s,%s,%s', cYear, stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)
                        #cur.execute("INSERT INTO student(name,year,grade,class,odr) values(?,?,?,?,?)", t)
                        #stuId = cur.lastrowid
                        #logging.info('A student inserted: %s,%s,%s,%s,%s,%s', stuId,cYear, stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)
                        t = (cYear,stuGrade,stuClass,row[colFirst +2].value)
                        cur.execute("SELECT id,code,name FROM student WHERE year=? AND grade=? AND class=? AND odr=?", t)
                        r = cur.fetchone()
                        if r is not None:
                           logging.info(u'')
                           logging.warning(u'틀린 학생 정보: %s학년 %s반 %s번 %s', stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)
                           logging.warning(u'학년반번호 같음: %s학년 %s반 %s번 %s', stuGrade, stuClass, row[colFirst +2].value, r[2])
                           t = (cYear,row[colFirst +3].value)
                           cur.execute("SELECT id,grade,class,odr FROM student WHERE year=? AND name=?", t)
                           for r in cur:
                              logging.warning(u'이름 같음: %s학년 %s반 %s번 %s', r[1],r[2],r[3],row[colFirst +3].value)
                        else:
                           logging.info(u'새로운 학생: %s학년 %s반 %s번 %s', stuGrade, stuClass, row[colFirst +2].value, row[colFirst +3].value)

            #else: # not found '학년'
            #   logging.info(u'Worksheet \'' + sheet.title + u'\' may not have any student.')

   cur.close()
   #db.commit()
   db.close()

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info("Adding afterSchool classes done.")
