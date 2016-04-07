#-*- coding: utf-8 -*-
# writeFreePassRefund.py

import logging
import os.path, shutil
import glob
import sqlite3
#from openpyxl import load_workbook
#from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import win32com.client as win32
import getopt, sys

# delete the previous log file
logfile = 'writeFreePassRefund.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of writeFreePassRefund.exe >>>>>')

def usage():
   print("Usage: writeFreePassRefund -y year -m month -f xls_folder -o out_folder")

try:
   opts, args = getopt.getopt(sys.argv[1:], "y:m:f:o:")
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
   elif o == "-f":
      xlsDir = a
   elif o == "-o":
      outDir = a
   else:
      assert False, "unhandled option"

# open the afterSchool database
dbName = 'asClass.db'
if not os.path.isfile(dbName):
   logging.error("The database file \'asClass.db\' not found.")
   exit(1)

appDir = os.path.dirname(os.path.abspath('__file__'))

try:
   db = sqlite3.connect(dbName)
   cur = db.cursor()

   excel = win32.gencache.EnsureDispatch('Excel.Application')
   #excel.Visible = True;

   xlsPath = os.path.join(xlsDir,'*.xls')
   xlList = glob.glob(xlsPath)
   for xlFile in xlList:
      xlsName = xlFile[len(xlsDir)+1:].decode('cp949')
      logging.info(u'')
      logging.info(u'<<< ' + xlsName + u' >>>')
      #wb = load_workbook(filename=xlFile)
      xlPath = os.path.join(appDir, xlFile)
      print xlPath
      #wb = excel.Workbooks.Open(xlFile)
      wb = excel.Workbooks.Open(xlPath)
      #for sheet in wb.Worksheets:
      sheet = wb.Worksheets(1)
      logging.info(u'< ' + sheet.Name + u' >')
      bBreak = False
      for i in range(1,len(sheet.UsedRange.Rows) +1):
         for j in range(1,len(sheet.UsedRange.Columns) +1):
            if sheet.Cells(i,j).Value and -1 < sheet.Cells(i,j).Value.find(u'학년'):
               rowFirst = i +1
               colFirst = j
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
            logging.info('The class not found: %s', className)
            continue
         bTuition = -1 < xlsName.find(u'강사')

         rowIdx = ()
         for i in range(rowFirst,len(sheet.UsedRange.Rows) +1):
            # student
            if sheet.Cells(i,colFirst).Value and sheet.Cells(i,colFirst +1).Value and sheet.Cells(i,colFirst +2).Value and sheet.Cells(i,colFirst +3).Value:
               if type(sheet.Cells(i,colFirst).Value) is unicode:
                  stuGrade = sheet.Cells(i,colFirst).Value.replace(u'학년',u'').strip()
               else:
                  stuGrade = sheet.Cells(i,colFirst).Value
               if type(sheet.Cells(i,colFirst +1).Value) is unicode:
                  stuClass = sheet.Cells(i,colFirst +1).Value.replace(u'반',u'').strip()
               else:
                  stuClass = sheet.Cells(i,colFirst +1).Value
               # class of the student
               t = (classId,cYear,stuGrade,stuClass,sheet.Cells(i,colFirst +2).Value,sheet.Cells(i,colFirst +3).Value,'FP','Y')
               if bTuition:
                  cur.execute("SELECT stuId,tuition FROM afterSchStu WHERE classId=? AND year=? AND grade=? AND class=? AND odr=? AND name=? AND code=? AND tuit_pay=?", t)
               else:
                  cur.execute("SELECT stuId,mcost FROM afterSchStu WHERE classId=? AND year=? AND grade=? AND class=? AND odr=? AND name=? AND code=? AND mcos_pay=?", t)
               r = cur.fetchone()
               sheet.Cells(i,colFirst +4).Value = 0
               if r is not None:
                  if r[1] is not None:
                     sheet.Cells(i,colFirst +5).Value = -r[1]
                     rowIdx = rowIdx + (i,)

         # arrange rows
         for i in range(0,len(rowIdx)):
            #r = sheet.rows[rowFirst -1 + i]
            #row = sheet.rows[rowIdx[i]]
            #for c, cell in zip(r, row):
            #   c.value = cell.value
            sheet.Range(sheet.Cells(rowFirst + i,colFirst),sheet.Cells(rowFirst + i,len(sheet.UsedRange.Columns))).Value = \
               sheet.Range(sheet.Cells(rowIdx[i],colFirst),sheet.Cells(rowIdx[i],len(sheet.UsedRange.Columns))).Value
         # delete residual
         #for r in range(rowFirst -1 + len(rowIdx), len(sheet.rows)):
         #   row = sheet.rows[r]
         #   for cell in row:
         #      cell.value = None
         #      cell.border = Border()
         #sheet.Range(sheet.Cells(rowFirst + len(rowIdx),colFirst),sheet.Cells(len(sheet.UsedRange.Rows),len(sheet.UsedRange.Columns))).Select()
         #excel.Selection.ClearContents()

      else: # not found '학년'
         logging.info(u'Worksheet \'' + sheet.title + u'\' may not have any student.')

      # copy worksheet
      #sheet.Range(sheet.Cells(1,1),sheet.Cells(len(rowIdx) +1,len(sheet.UsedRange.Columns))).Copy()
      shName = sheet.Name
      #print type(shName)
      #print shName
      wsOut = wb.Worksheets.Add()
      #wsOut.Cells(1,1).Select()
      #wsOut.Paste()
      #wsOut.Cells(1,1).Select()
      wb.Worksheets.FillAcrossSheets(sheet.Range(sheet.Cells(1,1),sheet.Cells(len(rowIdx) +1,len(sheet.UsedRange.Columns))))
      sheet.Delete()
      #wsOut.Name = shName
      outPath = os.path.join(appDir, outDir, xlFile[len(xlsDir)+1:])
      wb.SaveAs(outPath)
      wsOut.Name = shName
      wb.Save()
      
      #outPath = os.path.join(appDir, outDir, xlFile[len(xlsDir)+1:])
      #wbOut = excel.Workbooks.Add()
      #wsOut = wbOut.Worksheets(1)
      #wsOut.Name = sheet.Name
      #wsOut.Cells(1,1).Select()
      #wsOut.Paste()
      #wbOut.SaveAs(outPath)
      logging.info(u'Saved: ' + outPath.decode('cp949'))
      excel.Application.Quit()
      break

   cur.close()
   db.commit()
   db.close()

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info("Writing refund of free pass classes done.")
