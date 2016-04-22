#-*- coding: utf-8 -*-
# addClass2.py

import logging
import os.path, shutil
import glob
import sqlite3
from openpyxl import load_workbook
import getopt, sys

# delete the previous log file
logfile = 'addClass2.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of addClass2.exe >>>>>')

def usage():
   print("Usage: addClass2 -y year -m month -t tuition_folder -c mcost_folder")

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
   cur1 = db.cursor()

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
         if 0 == month:
            month = 12

      xlsPath = os.path.join(folders[i],'*.xlsx')
      xlList = glob.glob(xlsPath)
      for xlFile in xlList:
         xlsName = xlFile[len(folders[i])+1:].decode('cp949')
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
               className = xlsName[:xlsName.find(u' ')]
               t = (className, cYear, cMonth)
               cur.execute("SELECT id FROM afterSchoolClass WHERE cname=? AND year=? AND month=?", t)
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
                     cur.execute("SELECT id FROM student WHERE year=? AND grade=? AND class=? AND odr=? AND name=?", t)
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
                     code = None;
                     if -1 < row[colFirst +4].value.find(u'자유수강대상자'):
                        code = 'FP'

                     t = (classId,stuId,month)
                     cur.execute("SELECT id,tuition,mcost,code FROM classStu WHERE classId=? AND stuId=? AND month=?", t)
                     r = cur.fetchone()
                     if r is not None: # update
                        t = (row[colFirst +4].value,code,classId,stuId,month)
                        if 0 == i: # tuition
                           cur.execute("UPDATE classStu SET tuition=?,code=? WHERE classId=? AND stuId=? AND month=?", t)
                           logging.info('A student of class updated: %s,%s,%s,%s,%s,%s,%s', r[0],classId,stuId,month,row[colFirst +4].value,r[2],code)
                        elif 1 == i: # mcost
                           cur.execute("UPDATE classStu SET mcost=?,code=? WHERE classId=? AND stuId=? AND month=?", t)
                           logging.info('A student of class updated: %s,%s,%s,%s,%s,%s,%s', r[0],classId,stuId,month,r[1],row[colFirst +4].value,code)
                        # check FP
                        if code != r[3]:
                           logging.warn('Please check the student as a free pass. not the same in two docs: %s,%s,%s,%s', \
                              row[colFirst].value,row[colFirst +1].value,row[colFirst +2].value,row[colFirst +3].value)
                     else: # insert
                        t = (classId,stuId,month,row[colFirst +4].value,code)
                        if 0 == i: # tuition
                           cur.execute("INSERT INTO classStu(classId,stuId,month,tuition,code) values(?,?,?,?,?)", t)
                        elif 1 == i: # mcost
                           cur.execute("INSERT INTO classStu(classId,stuId,month,mcost,code) values(?,?,?,?,?)", t)
                        classStuId = cur.lastrowid
                        if 0 == i: # tuition
                           logging.info('A student of class inserted: %s,%s,%s,%s,%s,%s,%s', classStuId,classId,stuId,month,row[colFirst +4].value,None,code)
                        elif 1 == i: # mcost
                           logging.info('A student of class inserted: %s,%s,%s,%s,%s,%s,%s', classStuId,classId,stuId,month,None,row[colFirst +4].value,code)
                        # check FP
                        if bPrevMonth:
                           t = (classId,stuId,prevMonth)
                           cur.execute("SELECT id,tuition,mcost,code FROM classStu WHERE classId=? AND stuId=? AND month=?", t)
                           r = cur.fetchone()
                           if r is not None:
                              if code != r[3]:
                                 logging.warn('Please check the student as a free pass. not the same with previous month: %s,%s,%s,%s, this month: %s, prev. month: %s', \
                                    row[colFirst].value,row[colFirst +1].value,row[colFirst +2].value,row[colFirst +3].value, code, r[3])
                           else: # new student of class
                              logging.info('A new student of class. %s: %s,%s,%s,%s', className,\
                                 row[colFirst].value,row[colFirst +1].value,row[colFirst +2].value,row[colFirst +3].value)
                              t = ('N',classStuId)
                              cur.execute("UPDATE classStu SET quitNew=? WHERE id=?", t)
               # check students quitted
               if 1 == i and bPrevMonth:
                  t = (classId,prevMonth)
                  cur.execute("SELECT id,stuId,code FROM classStu WHERE classId=? AND month=?", t)
                  for r in cur:
                     t = (classId,r[1],month)
                     cur1.execute("SELECT id FROM classStu WHERE classId=? AND stuId=? AND month=?", t)
                     r1 = cur1.fetchone()
                     if r1 is None:
                        logging.info('A student quitted class. %s: %s,%s,%s,%s', className,\
                           row[colFirst].value,row[colFirst +1].value,row[colFirst +2].value,row[colFirst +3].value)
                        t = ('Q',r1[0])
                        cur.execute("UPDATE classStu SET quitNew=? WHERE id=?", t)

            else: # not found '학년'
               logging.info(u'Worksheet \'' + sheet.title + u'\' may not have any student.')

   cur.close()
   cur1.close()
   db.commit()
   db.close()

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info("Adding afterSchool classes done.")