'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.12

Library of tools used to process XLS/XLSX files
'''

import openpyxl  # xlsx (read/write)
import xlrd      # xls (read)
import xlwt      # xls (write)

import kvutil
import kvmatch
import datetime

# logging
import logging
logger = logging.getLogger(__name__)

# global variables
AppVersion = '1.12'

#----- OPTIONS ---------------------------------------
# debug
# dupkeyfail
# data_only
# req_cols
# col_aref
# xlat_dict
# dictkeys
#
# optiondict:
# col_header
# no_header
# aref_result
# saverow
# start_row
# sheetname
# sheetrow
# dateflds



# ---- UTILITY FUNCTIONS ------------------------------

# utility used to convert an xls date number into a datetime object
def xldate_to_datetime(xldate, skipblank=False):
    if isinstance(xldate,str):
        logger.debug('converting xldate string to date using kvutil.datetime_from_str:%s', xldate)
        return kvutil.datetime_from_str( xldate, skipblank )
    else:
        logger.debug('converting xldate float to date:%s', xldate)
        temp = datetime.datetime(1899, 12, 30)
        delta = datetime.timedelta(days=xldate)
        return temp+delta
    
# routine extracts a row from the excel file and passes back as a list
def _extract_excel_row_into_list(xlsxfiletype, s, row, colstart, colmax, debug=False):

    # debugging
    if debug:
        print('_extract_excel_row_into_list:row:', row)
        print('_extract_excel_row_into_list:xlsxfiletype:', xlsxfiletype)
    logger.debug('row: %s', row)
    logger.debug('xlsxfiletype:%s', xlsxfiletype)
    
    # clear the row
    rowdata = []

    # pull each column out of XLS and build the row array
    for col in range(colstart,colmax):
        # get cell value
        if xlsxfiletype:
            cValue = s.cell(row=row+1, column=col+1).value
        else:
            cValue = s.cell(row, col).value
            
        # debugging
        if debug:  print('row:', row, ':col:', col, ':cValue:', cValue)
        logger.debug('row:%s:col:%s:cValue:%s', row, col, cValue)
            
        # add this value to the array that will be used to determine if this is header
        rowdata.append(cValue)

    # return the row
    return rowdata

# routine to get a cell
def getExcelCellValue( excelDict, row, col_name, debug=False ):
    if debug:
        print('getExcelCellValue:excelDict:', excelDict)
        print('getExcelCellValue:row:', row)
        print('getExcelCellValue:col_name:', col_name)
    logger.debug('excelDict:%s', excelDict)
    logger.debug('row:%s', row)
    logger.debug('col_name:%s', col_name)

    # determine the col # we are using but doing a header lookup
    col = excelDict['header'].index(col_name) + excelDict['sheetmincol']
    
    # get cell value
    if excelDict['xlsxfiletype']:
        return excelDict['s'].cell(row=row+1, column=col+1).value
    else:
        return excelDict['s'].cell(row, col).value
        
# routine to set a cell value
def setExcelCellValue( excelDict, row, col_name, value, debug=False ):
    if debug:
        print('setExcelCellValue:excelDict:', excelDict)
        print('setExcelCellValue:row:', row)
        print('setExcelCellValue:col_name:', col_name)
    logger.debug('excelDict:%s', excelDict)
    logger.debug('row:%s', row)
    logger.debug('col_name:%s', col_name)

    # determine the col # we are using but doing a header lookup
    col = excelDict['header'].index(col_name) + excelDict['sheetmincol']
    
    # get cell value
    if excelDict['xlsxfiletype']:
        excelDict['s'].cell(row=row+1, column=col+1, value=value)
    else:
        logger.error('feature not supported on xls file - only XLSX')
        print('kvxls:setExcelCellValue:feature not supported on xls file - only XLSX')
        raise

# routine to get a cell fill pattern - returns the (solid,rgb) values
def getExcelCellPatternFill( excelDict, row, col_name, debug=False ):
    if debug:
        print('setExcelCellPatternFill:excelDict:', excelDict)
        print('setExcelCellPatternFill:row:', row)
        print('setExcelCellPatternFill:col_name:', col_name)
    logger.debug('excelDict:%s', excelDict)
    logger.debug('row:%s', row)
    logger.debug('col_name:%s', col_name)

    # determine the col # we are using but doing a header lookup
    col = excelDict['header'].index(col_name) + excelDict['sheetmincol']
    
    # get cell value
    if excelDict['xlsxfiletype']:
        cellFill = excelDict['s'].cell(row=row+1, column=col+1).fill
        return cellFill.solid, cellFill.fgColor.rgb
    else:
        logger.error('feature not supported on xls file - only XLSX')
        print('kvxls:setExcelCellValue:feature not supported on xls file - only XLSX')
        raise

# routine to set a cell fill pattern
def setExcelCellPatternFill( excelDict, row, col_name, fgColor, fill_type="solid", debug=False ):
    if debug:
        print('setExcelCellPatternFill:excelDict:', excelDict)
        print('setExcelCellPatternFill:row:', row)
        print('setExcelCellPatternFill:col_name:', col_name)
    logger.debug('excelDict:%s', excelDict)
    logger.debug('row:%s', row)
    logger.debug('col_name:%s', col_name)

    # determine the col # we are using but doing a header lookup
    col = excelDict['header'].index(col_name) + excelDict['sheetmincol']
    
    # get cell value
    if excelDict['xlsxfiletype']:
        if not fill_type:
            excelDict['s'].cell(row=row+1, column=col+1).fill = openpyxl.styles.PatternFill(fill_type=None)
        else:
            excelDict['s'].cell(row=row+1, column=col+1).fill = openpyxl.styles.PatternFill(fill_type, fgColor=fgColor)
    else:
        logger.error('feature not supported on xls file - only XLSX')
        print('kvxls:setExcelCellValue:feature not supported on xls file - only XLSX')
        raise
        
# -------- READ FILES -------------------------

# read in the XLS and create a dictionary to the records
# assumes the first line of the XLS file is the header/defintion of the XLS
def readxls2list( xlsfile, debug=False ):
    return readxls2list_findheader( xlsfile, [], optiondict={'col_header' : True}, debug=debug )

# read in the XLS and create a dictionary to the records
# based on one or more key fields
# assumes the first line of the CSV file is the header/defintion of the CSV
def readxls2dict( xlsfile, dictkeys, dupkeyfail=False, debug=False ):
    return readxls2dict_findheader( xlsfile, dictkeys, [], optiondict={'col_header' : True}, debug=debug, dupkeyfail=dupkeyfail )

# ---------- GENERIC OPEN EXCEL to enable EDIT ----------------------
#
# or passed on to other routines to extract the data for processing
#
# Open to edit and save:
# # example how to use:  open file for editting
# xls = kvxls.readxls_findheader( 'Wine Collection 20-05-07-v02.xlsm', [], optiondict={'col_header' : True}, data_only=False )
# # change a cell
# kvxls.setExcelCellValue( xls, 2, 'Rating', 'Changed')
# # save the file
# kvxls.writexls( xls, 'newfile.xlsm' )
#
#
# generic routine that reads in the XLS and returns back a dictionary for that xls
# that is either used to interact with that XLS object, or is passed to other routines
# that then create the dictionary/list of that xls and then close out that XLS.
#    data_only - when set to FALSE - will allow you to read macro enable file and update directly
#                and save the updated file
def readxls_findheader( xlsfile, req_cols, xlatdict={}, optiondict={}, col_aref=None, data_only=True, debug=False ):

    # local variables
    header = None
    
    # debugging
    if debug:
        print('req_cols:', req_cols)
        print('xlatdict:', xlatdict)
        print('optiondict:', optiondict)
        print('col_aref:', col_aref)
    logger.debug('req_cols:%s', req_cols)
    logger.debug('xlatdict:%s', xlatdict)
    logger.debug('optiondict:%s', optiondict)
    logger.debug('col_aref:%s', col_aref)


    # set flags
    col_header  = False  # if true - we take the first row of the file as the header
    no_header   = False  # if true - there are no headers read - we either return 
    aref_result = False  # if true - we don't return dicts, we return a list
    save_row    = False  # if true - then we append/save the XLSRow with the record
    
    start_row   = 0      # if passed in - we start the search at this row (starts at 1 or greater)

    # create the list of misconfigured solutions
    badoptiondict = {
        'startrow'       : 'start_row',
        'startrows'      : 'start_row',
        'start_rows'     : 'start_row',
        'colheaders'     : 'col_header',
        'col_headers'    : 'col_header',
        'noheader'       : 'no_header',
        'noheaders'      : 'no_header',
        'no_headers'     : 'no_header',
        'arefresult'     : 'aref_result',
        'arefresults'    : 'aref_result',
        'aref_results'   : 'aref_result',
        'saverow'        : 'saverow',
        'saverows'       : 'saverow',
        'save_rows'      : 'saverow',
    }

    # check what got passed in
    kvmatch.badoptiondict_check( 'kvxls.readxls_findheader', optiondict, badoptiondict, True )
        
    
    # pull in passed values from optiondict
    if 'col_header'  in optiondict: col_header  = optiondict['col_header']
    if 'aref_result' in optiondict: aref_result = optiondict['aref_result']
    if 'no_header'   in optiondict: no_header   = optiondict['no_header']
    if 'start_row'   in optiondict: start_row   = optiondict['start_row'] - 1 # because we are not ZERO based in the users mind
    if 'save_row'    in optiondict: save_row    = optiondict['save_row']
    

    # debugging
    if debug:
        print('col_header:', col_header)
        print('aref_result:', aref_result)
        print('no_header:', no_header)
        print('start_row:', start_row)
        print('save_row:', save_row)
        print('optiondict:', optiondict)
    logger.debug('col_header:%s', col_header)
    logger.debug('aref_result:%s', aref_result)
    logger.debug('no_header:%s', no_header)
    logger.debug('start_row:%s', start_row)
    logger.debug('save_row:%s', save_row)
    logger.debug('optiondict:%s', optiondict)

    # build object that will be used for record matching
    p = kvmatch.MatchRow( req_cols, xlatdict, optiondict )

    # determine what filetype we have here
    xlsxfiletype = xlsfile.endswith('.xlsx') or xlsfile.endswith('.xlsm')

    # debugging
    logger.debug('xlsxfiletype:%s', xlsxfiletype)
    
    # Load in the workbook (set the data_only=True flag to get the value on the formula)
    if xlsxfiletype:
        # XLSX file
        if data_only:
            wb = openpyxl.load_workbook(xlsfile, data_only=True)
        else:
            wb = openpyxl.load_workbook(xlsfile, read_only=False, keep_vba=True)
        sheetNames = wb.sheetnames
    else:
        # XLS file
        wb = xlrd.open_workbook(xlsfile)
        sheetNames = wb.sheet_names()

    # debugging
    if debug:  print('sheetNames:', sheetNames)
    logger.debug('sheetNames:%s', sheetNames)

    # get the sheet we are going to work with
    if 'sheetname' in optiondict:
        sheetName = optiondict['sheetname']
    elif 'sheetrow' in optiondict:
        sheetName = sheetNames[optiondict['sheetrow']]
    else:
        sheetName = sheetNames[0]

    # debugging
    if debug:  print('sheetName:', sheetName)
    logger.debug('sheetName:%s', sheetName)

    # create a workbook sheet object - using the name to get to the right sheet
    if xlsxfiletype:
        s = wb[sheetName]
        sheettitle = s.title
        sheetmaxrow = s.max_row
        sheetmaxcol = s.max_column
        sheetminrow = 0
        sheetmincol = 0
    else:
        s = wb.sheet_by_name(sheetName)
        sheettitle = s.name
        sheetmaxrow = s.nrows
        sheetmaxcol = s.ncols
        sheetminrow = 0
        sheetmincol = 0

    # debugging
    if debug:
        print('sheettitle:', sheettitle)
        print('sheetmaxrow:', sheetmaxrow)
        print('sheetmaxcol:', sheetmaxcol)
    logger.debug('sheettitle:%s', sheettitle)
    logger.debug('sheetmaxrow:%s', sheetmaxrow)
    logger.debug('sheetmaxcol:%s', sheetmaxcol)
        

    # ------------------------------- HEADER START ------------------------------

    # define the header for the records being read in
    if no_header:
        # user said we are not to look for the header in this file
        # we need to subtract 1 here because we are going to increment PAST the header
        # in the next section - so if there is no header - we need to start at zero ( -1 + 1 later)
        row_header = start_row - 1

        # if no col_aref - then we must force this to aref_result
        if not col_aref:
            aref_result = True
            if debug:  print('no_header:no col_aref:set aref_result to true')
            logger.debug('no_header:no col_aref:set aref_result to true')
            
        # debug
        if debug:  print('no_header:start_row:', start_row)
        logger.debug('no_header:start_row:%d', start_row)
        
    else:
        # debug
        if debug: print('find_header:start_row:', start_row)
        logger.debug('find_header:start_row:%d', start_row)
        
        # look for the header in the file
        for row in range(start_row, sheetmaxrow):
            # read in a row of data
            rowdata = _extract_excel_row_into_list(xlsxfiletype, s, row, sheetmincol, sheetmaxcol, debug)

            # user may have specified that the first row read is the header
            if col_header:
                # first row read is header - set the values
                header = rowdata
                row_header = row
                # debugging
                if debug: print('header_1strow:',header)
                logger.debug('header_1strow:%s',header)
                # break out of this loop we are done
                break

            # have not found the header yet - so look
            if debug:  print('looking for header at row:', row)
            logger.debug('looking for header at row:%d', row)

            # Search to see if this row is the header
            if p.matchRowList( rowdata, debug=debug ) or p.search_exceeded:
                # determine if we found the header
                if p.search_exceeded:
                    # debugging
                    if debug: print('maxrows_search_exceeded:',p.error_msg)
                    logger.debug('maxrows in search exceeded:%s',p.error_msg)
                    # did not find the header
                    raise Exception(p.error_msg)
                elif p.search_failed:
                    # debugging
                    if debug: print('search_failed:',p.error_msg)
                    logger.debug('search_failed:%s',p.error_msg)
                    # did not find the header
                    raise Exception(p.error_msg)
                else:
                    # set the row_header
                    row_header = row
                    # found the header grab the output
                    header = p._data_mapped
                    # debugging
                    if debug: print('header_found:',header)
                    logger.debug('header_found:%s',header)
                    # break out of the loop
                    break
            elif debug:
                print('no match found loop again')


    # ------------------------------- HEADER END ------------------------------

    # debug
    if debug:  print('exitted header find loop')
    logger.debug('exitted header find loop')
    
    # user wants to define/override the column headers rather than read them in
    if col_aref:
        # debugging
        if debug:  print('copying col_aref into header')
        logger.debug('copying col_aref into header')
        # copy over the values - and determine if we need to fill in more header values
        header = col_aref[:]
        # user defined the row definiton - make sure they passed in enough values to fill the row
        if len(col_aref) < sheetmaxcol - sheetmincol:
            # not enough entries - so we add more to the end
            for colcnt in range(1, sheetmaxcol - sheetmincol - len(col_aref) + 1 ):
                header.append('')

        # now pass the final information through remapped
        header = p.remappedRow(header)
        # debug
        if debug: print('col_aref:header:', header)
        logger.debug('col_aref:header:%s', header)

    # ------------------------------- OBJECT DEFINITION ------------------------------
    excelDict = {
        'xlsfile' : xlsfile,
        'xlsxfiletype' : xlsxfiletype,
        'wb' : wb,
        'sheetNames' : sheetName,
        'sheetName' : sheetName,
        's' : s,
        'sheettitle' : sheettitle,
        'sheetmaxrow' : sheetmaxrow,
        'sheetmaxcol' : sheetmaxcol,
        'sheetminrow' : sheetminrow,
        'sheetmincol' : sheetmincol,
        'row_header' : row_header,
        'header' : header,
        'start_row' : start_row,
    }

    return excelDict

# or passed on to other routines to extract the data for processing
#
# Open to edit and save:
# # example how to use:  open file for editting
# xls = kvxls.readxls_findheader( 'Wine Collection 20-05-07-v02.xlsm', [], optiondict={'col_header' : True}, data_only=False )
# # change a cell
# kvxls.setExcelCellValue( xls, 2, 'Rating', 'Changed')
# # save the file
# kvxls.writexls( xls, 'newfile.xlsm' )
#
#
# generic routine that reads in the XLS and returns back a dictionary for that xls
# that is either used to interact with that XLS object, or is passed to other routines
# that then create the dictionary/list of that xls and then close out that XLS.
#    data_only - when set to FALSE - will allow you to read macro enable file and update directly
#                and save the updated file
def chgsheet_findheader( excelDict, req_cols, xlatdict={}, optiondict={}, col_aref=None, data_only=True, debug=False ):

    # local variables
    header = None
    
    # debugging
    if debug:
        print('req_cols:', req_cols)
        print('xlatdict:', xlatdict)
        print('optiondict:', optiondict)
        print('col_aref:', col_aref)
    logger.debug('req_cols:%s', req_cols)
    logger.debug('xlatdict:%s', xlatdict)
    logger.debug('optiondict:%s', optiondict)
    logger.debug('col_aref:%s', col_aref)


    # check to see if we are actually changing anyting - if not return back what was sent in
    if 'sheetname' in optiondict and excelDict['sheetName'] == optiondict['sheetname']:
        logger.debug('nothing changed - return what was sent in')
        return excelDict

    # set flags
    col_header  = False  # if true - we take the first row of the file as the header
    no_header   = False  # if true - there are no headers read - we either return 
    aref_result = False  # if true - we don't return dicts, we return a list
    save_row    = False  # if true - then we append/save the XLSRow with the record
    
    start_row   = 0      # if passed in - we start the search at this row (starts at 1 or greater)

    # create the list of misconfigured solutions
    badoptiondict = {
        'startrow'       : 'start_row',
        'startrows'      : 'start_row',
        'start_rows'     : 'start_row',
        'colheaders'     : 'col_header',
        'col_headers'    : 'col_header',
        'noheader'       : 'no_header',
        'noheaders'      : 'no_header',
        'no_headers'     : 'no_header',
        'arefresult'     : 'aref_result',
        'arefresults'    : 'aref_result',
        'aref_results'   : 'aref_result',
        'saverow'        : 'saverow',
        'saverows'       : 'saverow',
        'save_rows'      : 'saverow',
    }

    # check what got passed in
    kvmatch.badoptiondict_check( 'kvxls.readxls_findheader', optiondict, badoptiondict, True )
        
    
    # pull in passed values from optiondict
    if 'col_header'  in optiondict: col_header  = optiondict['col_header']
    if 'aref_result' in optiondict: aref_result = optiondict['aref_result']
    if 'no_header'   in optiondict: no_header   = optiondict['no_header']
    if 'start_row'   in optiondict: start_row   = optiondict['start_row'] - 1 # because we are not ZERO based in the users mind
    if 'save_row'    in optiondict: save_row    = optiondict['save_row']
    

    # debugging
    if debug:
        print('col_header:', col_header)
        print('aref_result:', aref_result)
        print('no_header:', no_header)
        print('start_row:', start_row)
        print('save_row:', save_row)
        print('optiondict:', optiondict)
    logger.debug('col_header:%s', col_header)
    logger.debug('aref_result:%s', aref_result)
    logger.debug('no_header:%s', no_header)
    logger.debug('start_row:%s', start_row)
    logger.debug('save_row:%s', save_row)
    logger.debug('optiondict:%s', optiondict)
        
    # build object that will be used for record matching
    p = kvmatch.MatchRow( req_cols, xlatdict, optiondict )

    # read in values from excelDict
    # determine what filetype we have here
    xlsfile = excelDict['xlsfile']
    xlsxfiletype = excelDict['xlsxfiletype']
    wb = excelDict['wb']
    sheetNames = excelDict['sheetNames']

    # debugging
    if debug:  print('sheetNames:', sheetNames)
    logger.debug('sheetNames:%s', sheetNames)

    # get the sheet we are going to work with
    if 'sheetname' in optiondict:
        sheetName = optiondict['sheetname']
    elif 'sheetrow' in optiondict:
        sheetName = sheetNames[optiondict['sheetrow']]
    else:
        sheetName = sheetNames[0]

    # debugging
    if debug:  print('sheetName:', sheetName)
    logger.debug('sheetName:%s', sheetName)

    # create a workbook sheet object - using the name to get to the right sheet
    if xlsxfiletype:
        s = wb[sheetName]
        sheettitle = s.title
        sheetmaxrow = s.max_row
        sheetmaxcol = s.max_column
        sheetminrow = 0
        sheetmincol = 0
    else:
        s = wb.sheet_by_name(sheetName)
        sheettitle = s.name
        sheetmaxrow = s.nrows
        sheetmaxcol = s.ncols
        sheetminrow = 0
        sheetmincol = 0

    # debugging
    if debug:
        print('sheettitle:', sheettitle)
        print('sheetmaxrow:', sheetmaxrow)
        print('sheetmaxcol:', sheetmaxcol)
    logger.debug('sheettitle:%s', sheettitle)
    logger.debug('sheetmaxrow:%s', sheetmaxrow)
    logger.debug('sheetmaxcol:%s', sheetmaxcol)
        

    # ------------------------------- HEADER START ------------------------------

    # define the header for the records being read in
    if no_header:
        # user said we are not to look for the header in this file
        # we need to subtract 1 here because we are going to increment PAST the header
        # in the next section - so if there is no header - we need to start at zero ( -1 + 1 later)
        row_header = start_row - 1

        # if no col_aref - then we must force this to aref_result
        if not col_aref:
            aref_result = True
            if debug:  print('no_header:no col_aref:set aref_result to true')
            logger.debug('no_header:no col_aref:set aref_result to true')
            
        # debug
        if debug:  print('no_header:start_row:', start_row)
        logger.debug('no_header:start_row:%d', start_row)
        
    else:
        # debug
        if debug: print('find_header:start_row:', start_row)
        logger.debug('find_header:start_row:%d', start_row)
        
        # look for the header in the file
        for row in range(start_row, sheetmaxrow):
            # read in a row of data
            rowdata = _extract_excel_row_into_list(xlsxfiletype, s, row, sheetmincol, sheetmaxcol, debug)

            # user may have specified that the first row read is the header
            if col_header:
                # first row read is header - set the values
                header = rowdata
                row_header = row
                # debugging
                if debug: print('header_1strow:',header)
                logger.debug('header_1strow:%s',header)
                # break out of this loop we are done
                break

            # have not found the header yet - so look
            if debug:  print('looking for header at row:', row)
            logger.debug('looking for header at row:%d', row)

            # Search to see if this row is the header
            if p.matchRowList( rowdata, debug=debug ) or p.search_exceeded:
                # determine if we found the header
                if p.search_exceeded:
                    # did not find the header
                    raise
                else:
                    # set the row_header
                    row_header = row
                    # found the header grab the output
                    header = p._data_mapped
                    # debugging
                    if debug: print('header_found:',header)
                    logger.debug('header_found:%s',header)
                    # break out of the loop
                    break


    # ------------------------------- HEADER END ------------------------------

    # debug
    if debug:  print('exitted header find loop')
    logger.debug('exitted header find loop')
    
    # user wants to define/override the column headers rather than read them in
    if col_aref:
        # debugging
        if debug:  print('copying col_aref into header')
        logger.debug('copying col_aref into header')
        # copy over the values - and determine if we need to fill in more header values
        header = col_aref[:]
        # user defined the row definiton - make sure they passed in enough values to fill the row
        if len(col_aref) < sheetmaxcol - sheetmincol:
            # not enough entries - so we add more to the end
            for colcnt in range(1, sheetmaxcol - sheetmincol - len(col_aref) + 1 ):
                header.append('')

        # now pass the final information through remapped
        header = p.remappedRow(header)
        # debug
        if debug: print('col_aref:header:', header)
        logger.debug('col_aref:header:%s', header)

    # ------------------------------- OBJECT DEFINITION ------------------------------
    excelDict = {
        'xlsfile' : xlsfile,
        'xlsxfiletype' : xlsxfiletype,
        'wb' : wb,
        'sheetNames' : sheetName,
        'sheetName' : sheetName,
        's' : s,
        'sheettitle' : sheettitle,
        'sheetmaxrow' : sheetmaxrow,
        'sheetmaxcol' : sheetmaxcol,
        'sheetminrow' : sheetminrow,
        'sheetmincol' : sheetmincol,
        'row_header' : row_header,
        'header' : header,
        'start_row' : start_row,
    }

    return excelDict


# ---------- EXTRACT DATA FROM EXCEL  ----------------------
#
# coding structure - build one generic (INTERNAL) function that does all the various things
# with passed in variables that are all optional
# and based on variable settings - executes the behavior being asked
#
# then create external functions - with clear passed in parameters that calls this internal function with teh right settings
#

# read in the CSV and create a dictionary to the records
# based on one or more key fields
# assumes the first line of the CSV file is the header/defintion of the CSV
# features to add
#
#   noheader - flag means we pass in the array that is the header
#   col_header - flag means we get the header from the very first line (or start_row) in the file
#   aref_result - flag that tells us to return the array of array (otherwise we pass back the array of dict)
#
#   col_aref - array that defines the columns we read into dict
#
#   start_row - user entered values start at row 1 (not zero)
#
#   unique_column - if enabled, we must validate that our column names are unique
#   ignore_blank_row - if enabled, if the row read has no data - we don't put the data into the extracted list
#   ignore_not_fill - not sure what this one is
#
#   dateflds - array of fields in xls that convert to a date
#   save_row - flag defines if we should save the row number
#
#   required_fields_populated - checking logic to assure that all required fields have data with optional
#     required_fld_swap - a dict that says if key is not populated - check the value tied to that key to see if it is populated
#

def readxls2list_findheader( xlsfile, req_cols, xlatdict={}, optiondict={}, col_aref=None, debug=False ):

    # local variables
    results = []
    header = None
    
    # debugging
    if debug: print('req_cols:', req_cols)
    if debug: print('xlatdict:', xlatdict)
    if debug: print('optiondict:', optiondict)
    if debug: print('col_aref:', col_aref)
    logger.debug('req_cols:%s', req_cols)
    logger.debug('xlatdict:%s', xlatdict)
    logger.debug('optiondict:%s', optiondict)
    logger.debug('col_aref:%s', col_aref)
    
    # set flags
    col_header  = False  # if true - we take the first row of the file as the header
    no_header   = False  # if true - there are no headers read - we either return 
    aref_result = False  # if true - we don't return dicts, we return a list
    save_row    = False  # if true - then we append/save the XLSRow with the record
    
    start_row   = 0      # if passed in - we start the search at this row (starts at 1 or greater)

    # call the routine that opens the XLS and returns back the excelDict
    # (missing data_only attribute between optiondict and debug)
    excelDict = readxls_findheader( xlsfile, req_cols, xlatdict, optiondict, col_aref, debug=debug )
    
    # pull in passed values from optiondict
    if 'col_header'  in optiondict: col_header  = optiondict['col_header']
    if 'aref_result' in optiondict: aref_result = optiondict['aref_result']
    if 'no_header'   in optiondict: no_header   = optiondict['no_header']
    if 'start_row'   in optiondict: start_row   = optiondict['start_row'] - 1 # because we are not ZERO based in the users mind
    if 'save_row'    in optiondict: save_row    = optiondict['save_row']
    

    # debugging
    if debug:
        print('col_header:', col_header)
        print('aref_result:', aref_result)
        print('no_header:', no_header)
        print('start_row:', start_row)
        print('save_row:', save_row)
        print('optiondict:', optiondict)
    logger.debug('col_header:%s', col_header)
    logger.debug('aref_result:%s', aref_result)
    logger.debug('no_header:%s', no_header)
    logger.debug('start_row:%s', start_row)
    logger.debug('save_row:%s', save_row)
    logger.debug('optiondict:%s', optiondict)

    # expand out all the values that came from excelDict
    xlsxfiletype = excelDict['xlsxfiletype']
    wb = excelDict['wb']
    sheetNames = excelDict['sheetNames']
    sheetName = excelDict['sheetName']
    s = excelDict['s']
    sheettitle = excelDict['sheettitle']
    sheetmaxrow = excelDict['sheetmaxrow']
    sheetmaxcol = excelDict['sheetmaxcol']
    sheetminrow = excelDict['sheetminrow']
    sheetmincol = excelDict['sheetmincol']
    row_header = excelDict['row_header']
    header = excelDict['header']
    start_row = excelDict['start_row']

    # if we don't have a header we must set the aref_result flag
    if not header and not aref_result:
        if debug: print('setting aref_results because there is no header')
        logger.debug('setting aref_results becaus there is no header')
        
        aref_result = True
        
    # debugging
    if debug:
        print('sheettitle:', sheettitle)
        print('sheetmaxrow:', sheetmaxrow)
        print('sheetmaxcol:', sheetmaxcol)
        

    # ------------------------------- RECORDS START ------------------------------

    for row in range( row_header + 1, sheetmaxrow ):
        # read in a row of data
        rowdata = _extract_excel_row_into_list(xlsxfiletype, s, row, sheetmincol, sheetmaxcol, debug)

        # determine what we are returning
        if aref_result:

            # we want to return the data we read
            rowdict = rowdata
            if debug:  print('saving as array')
            logger.debug('saving as array')
            
            # optionally add the XLSRow attribute to this dictionary (not here right now
            if save_row:
                rowdict.append( row + 1 )
                if debug: print('append row to record')
                logger.debug('append row to record')

        else:
            if debug:
                print('saving as dict')
                print('header:',header)
                print('rowdata:', rowdata)
            logger.debug('saving as dict:header:%s:rowdata:%s', header,rowdata)

            # we found the header so now build up the records
            rowdict = dict(zip(header,rowdata))

            # optionally add the XLSRow attribute to this dictionary (not here right now
            if save_row:
                rowdict['XLSRow'] = row + 1
                if debug: print('add column XLSRow with row to record')
                logger.debug('add column XLSRow with row to record')
                
            # do field manipulations here - date - but only on XLS not XLSX files
            if not xlsxfiletype:
                if 'dateflds' in optiondict:
                    for fld in optiondict['dateflds']:
                        if fld in rowdict:
                            rowdict[fld] = xldate_to_datetime(rowdict[fld])
                            if debug: print('xldate conversion on:', fld)
                            logger.debug('xldate conversion on:%s', fld)

        # add this dictionary to the results
        results.append(rowdict)
        if debug:  print('append rowdict to results')
        logger.debug('append rowdict to results')

    # ------------------------------- RECORDS END ------------------------------

    # debugging
    # if debug: print('results:', results)
    
    # return the results
    return results

# read in the XLS and create a dictionary to the records
# based on one or more key fields
def readxls2dict_findheader( xlsfile, dictkeys, req_cols=[], xlatdict={}, optiondict={}, col_aref=None, debug=False, dupkeyfail=False ):

    # validate we have proper input
    if not dictkeys:
        logger.error('kvxls:readxls2dict_findheader:dictkeys not populated - program error')
        print('kvxls:readxls2dict_findheader:dictkeys not populated - program error')
        raise

    
    # check for duplicate keys
    dupkeys = []

    # results defined as a dicut
    results = {}

    # debugging
    logger.debug('dictkeys:%s', dictkeys)
    if debug:
        print('readxls2dict_findheader:dictkeys:', dictkeys)
        input('press enter')
    
    # test how dictkeys was passed in
    if isinstance(dictkeys, str):
        dictkeys = [dictkeys]
        if debug:  print('readxls2dict_findheader:converted dictkeys from string to list')
        logger.debug('converted dictkeys from string to list')
        
    # debugging
    if debug:  print('readxls2dict_findheader:reading in xls as a list first')
    logger.debug('reading in xls as a list first')
    
    # read in the data from the file
    resultslist = readxls2list_findheader( xlsfile, req_cols, xlatdict=xlatdict, optiondict=optiondict, col_aref=col_aref, debug=debug )

    # debugging
    if debug:
        print('readxls2dict_findheader:xls data is in an array - now convert to a dictionary')
        print('readxls2dict_findheader:dictkeys:', dictkeys)
    logger.debug('xls data is in an array - now convert to a dictionary')
    logger.debug('dictkeys:%s', dictkeys)

    
    # convert to a dictionary based on keys provided
    for rowdict in resultslist:
        #rowdict = dict(zip(header,row))
        if debug:
            print('rowdict:', rowdict)
            print('dictkeys:', dictkeys)
        logger.debug('rowdict:%s', rowdict)
        logger.debug('dictkeys:%s', dictkeys)
        reckey = kvmatch.build_multifield_key(rowdict, dictkeys)
        # do we fail if we see the same key multiple times?
        if dupkeyfail:
            if reckey in results.keys():
                # capture this key
                dupkeys.append(reckey)

        # create/update the dictionary
        results[reckey] = rowdict

    # fail if we found dupkeys
    if dupkeys:
        logger.error('duplicate key failure:%s', ','.join(dupkeys))
        print('readxls2dict:duplicate key failure:', ','.join(dupkeys))
        raise

    # return the results
    return results


# -------- WRITE FILES -------------------------

# write out a dict of (dict or aref) to an XLS/XLSX based on the filename passed in
def writedict2xls( xlsfile, data, col_aref=None, optiondict={}, debug=False ):

    # convert dict to array and then call writelist2xls
    data2 = [data[key] for key in sorted(data.keys())]

    # call the other library
    return writelist2xls( xlsfile, data2, col_aref=None, optiondict={}, debug=debug )


# write out a list of (dict or aref) to an XLS/XLSX based on the filename passed in
def writelist2xls( xlsfile, data, col_aref=None, optiondict={}, debug=False ):

    # local variables
    sheet_name = 'Sheet1'
    no_header = False
    aref_result = False
    
    # determine what filetype we have here
    xlsxfiletype = xlsfile.endswith('.xlsx') or xlsfile.endswith('.xlsm')

    # change settings based on user input
    if 'sheet_name' in optiondict:   sheet_name  = optiondict['sheet_name']
    if 'no_header' in optiondict:    no_header   = optiondict['no_header']
    if 'aref_result' in optiondict:  aref_result = optiondict['aref_result']

    # set this value if the record we get is a list not a dictionary
    if isinstance(data[0], list):    aref_result = True

    # debugging
    if debug:
        print('sheet_name:', sheet_name)
        print('no_header:', no_header)
        print('aref_result:', aref_result)
        print('xlsxfiletype:', xlsxfiletype)
        
    # validate we have columns defined - or create one if we can
    if not col_aref:
        if aref_result:
            # this is a list passed in - we don't need header
            no_header = True
        else:
            # we can pull the keys from this record to create the col_aref
            col_aref = list(data[0].keys())

    # debuging
    if debug: print('col_aref:', col_aref)
    logger.debug('col_aref:%s', col_aref)

    # Create a new workbook
    if xlsxfiletype:
        # XLSX file
        wb = openpyxl.Workbook()
        ws = wb.active
        
        # set the title if one is specified
        if sheet_name != 'Sheet1':
            ws.title = sheet_name

    else:
        # XLS file
        wb = xlwt.Workbook() # None # xlrd.open_workbook(xlsfile)
        ws = wb.add_sheet(sheet_name, cell_overwrite_ok=True)

    
    # set the output row
    xlsrow = 0
    
    # get the header created
    if not no_header:
        for xlscol in range(0,len(col_aref)):
            if xlsxfiletype:
                d = ws.cell(row=xlsrow+1, column=xlscol+1, value=col_aref[xlscol])
            else:
                d = ws.write(xlsrow, xlscol, col_aref[xlscol])

    # increment the row
    xlsrow += 1

    # now step through the data itself
    for record in data:
        # output this row of data
        if col_aref and len(col_aref):
            for xlscol in range(0,len(col_aref)):
                # determine the value - based on how the records are structured
                if aref_result:
                    value = record[xlscol]
                else:
                    value = record[col_aref[xlscol]]
                
                # could put a feature in here to convert the value to a string before storing
                if xlsxfiletype:
                    d = ws.cell(row=xlsrow+1, column=xlscol+1, value=value)
                else:
                    d = ws.write(xlsrow, xlscol, value)
        elif aref_result:
            for xlscol in range(0,len(record)):
                if xlsxfiletype:
                    d = ws.cell(row=xlsrow+1, column=xlscol+1, value=record[xlscol])
                else:
                    d = ws.write(xlsrow, xlscol, record[xlscol])
                
        # done with this row - increment counter
        xlsrow += 1
        
    # now save this object
    return wb.save(xlsfile)

# write out a XLSX object in memory
def writexls( excelDict, xlsfile, debug=False ):

    # check to see that we can do this
    if not excelDict['xlsxfiletype']:
        print('kvxls:writexls:feature supported only for XLSX files')
        raise

    # if the user did not pass in a filename
    # us the same filename we read in
    if not xlsfile:
        xlsfile = excelDict['xlsfile']

    # get the workbook
    wb = excelDict['wb']
    
    # now save this object
    return wb.save(xlsfile)


if __name__ == '__main__':

    # put some quick test code here
    pass

#eof
