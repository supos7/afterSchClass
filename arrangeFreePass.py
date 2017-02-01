#-*- coding: utf-8 -*-
# arrangeFreePass.py

import logging
import os.path, shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import getopt, sys

# delete the previous log file
logfile = 'arrangeFreePass.log'

FORMAT = "%(asctime)-15s  %(levelname)s %(message)s"
logging.basicConfig(filename=logfile, level=logging.DEBUG, format=FORMAT)

logging.info('<<<<< The log file of arrangeFreePass.exe >>>>>')

def usage():
   print("Usage: arrangeFreePass -f temp_xls -o out_xls")

try:
   opts, args = getopt.getopt(sys.argv[1:], "f:o:")
except getopt.GetoptError, err:
   # print help information and exit:
   print str(err) # will print something like "option -a not recognized"
   usage()
   os._exit(1)
if 2 != len(opts):
   usage()
   os._exit(1)
for o, a in opts:
   if o == "-f":
      tmpxls = a
   elif o == "-o":
      outxls = a
   else:
      assert False, "unhandled option"

try:
   xlsName = tmpxls.decode('cp949')
   logging.info(u'')
   logging.info(u'<<< ' + xlsName + u' >>>')
   wb = load_workbook(filename=tmpxls)
   #for sheet in wb:
   sheet = wb.active
   logging.info(u'< ' + sheet.title + u' >')
   bBreak = False
   for row in sheet.rows:
      for cell in row:
         if cell.value and -1 < cell.value.find(u'학년'):
            rowFirst = cell.row +1
            colFirst = cell.col_idx -1 # 0 based
            bBreak = True
            break
      if bBreak:
         break
   if bBreak:
      # copy
      shTitle = sheet.title
      sheet.title = 'temp'
      wsOut = wb.create_sheet()
      wsOut.title = shTitle
      iRow = 0
      i = -1
      oldName = ''
      while iRow < sheet.max_row -1:
         iRow += 1
         row = sheet.rows[iRow]
         if row[colFirst].value and row[colFirst +1].value and row[colFirst +2].value:
            if oldName != row[colFirst +3].value:
               wsOut.append(28*[0])
               i += 1
               r = wsOut.rows[i]
               for j in range(4):
                  r[j].value = row[colFirst +j].value
            off = 4 + 2 * (row[colFirst +4].value -1)
            for j in range(2):
               r[off +j].value = row[colFirst +5 + j].value
            oldName = row[colFirst +3].value
#            for c, cell in zip(r, row):
#               c.value = cell.value
#               if cell.has_style:
#                  c.font = cell.font
#                  c.border = cell.border
#                  c.fill = cell.fill
#                  c.number_format = cell.number_format
#                  c.protection = cell.protection
#                  c.alignment = cell.alignment
         else:
            break
      # delete sheet
      wb.remove_sheet(sheet)         

   else: # not found '학년'
      logging.info(u'Worksheet \'' + sheet.title + u'\' may not have any student.')

   outPath = outxls
   wb.save(outPath)
   logging.info(u'Saved: ' + outPath.decode('cp949'))

except:
   logging.exception('Got an exception on database handler')
   raise

logging.info("Arranging free pass total done.")
