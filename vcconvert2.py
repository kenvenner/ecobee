'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.18

Read information from Beautiful Places XLS files,
extract out occupancy data, build a new
output file that has an entry per day of stay and stay type

'''
# we are reusing features from another applicatoin but we wnat
# the log files to be tied to this applicatoin - so we call tihs
# application/library late in order to have the logger ocnfigured to THIS app
import villaecobee
import villacalendar

import kvutil
import kvxls
import kvcsv

import os

import datetime
import sys

# for sorting a list of dicts
from operator import itemgetter

# working with Excel files
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
from openpyxl.utils import get_column_letter

# Excel formatting strings
bold = Font(bold=True, name="Arial", size=10)
regular = Font(name="Arial", size=10)
fit_centered = Alignment(shrink_to_fit=True, horizontal="center")
fit = Alignment(shrink_to_fit=True)

# this utility is used to convert the Beautiful Places Villa bookings XLS
# file into the flattened text file used by the "villaecobee.py" file
#
# General process:
# 1) get a new xls from beautiful places
# 2) save it to the directory/filename where this tool is stored
# 3) run this tool:  python vcconvert.py


import logging
# Logging Setup
# logging.basicConfig(level=logging.INFO)
logging.basicConfig(filename=os.path.splitext(kvutil.scriptinfo()['name'])[0]+'.log',
                    level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(name)s:%(lineno)d %(funcName)s %(message)s')
logger = logging.getLogger(__name__)


# application variables
optiondictconfig = {
    'AppVersion' : {
        'value' : '1.18',
        'description' : 'defines the version number for the app',
    },
    'debug' : {
        'value' : False,
        'type'  : 'bool',
        'description' : 'defines if we are running in debug mode',
    },
    'verbose' : {
        'value' : 1,
        'type'  : 'int',
        'description' : 'defines the display level for print messages',
    },
    'xls_filename' : {
        'value' : 'Attune_Estate_2022_Bookings.xlsx',
        'description' : 'defines the name of the BP xls filename',
    },
    'occupy_filename' : {
        'value' : 'stays.txt',
        'description' : 'defines the name of the file holding the villa occupancy',
    },
    'occupy_history_filename' : {
        'value' : 'stays_history.txt',
        'description' : 'defines the name of the file holding the historical villa occupancy',
    },
    'xlsdateflds'    : {
        'value'       : ['First Night','Last Night'],
        'type'        : 'liststr',
        'description' : 'defines the list of date fields inside the xls',
    },
    'fldFirstNight' : {
        'value' : 'First Night',
        'description' : 'defines the name of the field holding the first night date',
    },
    'fldNights' : {
        'value' : 'Nights',
        'description' : 'defines the name of the field holding the int of nights stay',
    },
    'fldType' : {
        'value' : 'Type',
        'description' : 'defines the name of the field holding the type of the stay',
    },
    'fldDate' : {
        'value' : 'date',
        'description' : 'defines the name of the field that holds the date in the occupancy file',
    },
    'calendarsync' : {
        'type' : 'bool',
        'value' : True,
        'description' : 'defines if we are going to sync XLS data with calendar',
    },
    'startback' : {
        'type' : 'int',
        'description' : 'defines number of days added to today that we update the calendar (negative numbers are in the past)',
    },
    'startdate' : {
        'value': None,
        'type' : 'date',
        'description' : 'defines the start date, use this field or startback field to change the startdate',
    },
}


### GLOBAL VARIABLES ####

# date information
ADD_ONE_DAY =  datetime.timedelta(days=1)
DATE_FMT='%m/%d/%Y'

# xls file occtype conversion to
# an array that is:  new code and # of days to add to stay for temp control
OCC_TYPE_CONV = {
    'Hold-Deep Clean' : ['C', 0],
    'Hold - Clean' : ['C', 0],
    'Hold - Maint.' : ['M', 0],
    'Hold- Mainten' : ['M', 0],
    'Hold - Maint' : ['M', 0],
    'Hold - Other' : ['M', 0],
    'Hold - Owner' : ['O', 1],
    'Hold - Renter' : ['R', 1],
    'Hold - Winery Business' : ['O',0],
    'Hold - Fall Mbr Event' : ['O',1],
    'Hold - Ken Venner' : ['O',1],
    'Res. - Renter' : ['R', 1],
    'Res. - Owner' : ['O', 1],
    'Res.-Renter' : ['R', 1],
    'Res.-Owner' : ['O', 1],
    'Res - Renter' : ['R', 1],
    'Res - Owner' : ['O', 1],
}

# xls header definition
COL_REQUIRED=['Booking','First Night','Last Night','Nights','Type', 'Rent', 'Source', 'Managing', 'Confirmed', 'BookedOn']

COL_CENTERED = ['Nights', 'Source', 'Managing', 'Confirmed']

COL_WIDTH = {
    'Booking': 12,
    'First Night': 12,
    'Last Night': 12,
    'Type': 17,
    'Source': 17,
    'Managing': 15,
    'Confirmed': 12,
    'BookedOn': 12,
}
    
MON_STRING = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

SHEET_LISTING = 'Listing'

def rewrite_file( xlsfile, xlsaref, fldFirstNight, fldNights, xlsdateflds):
    """
    Take a list of dicts defined in xlsaref
    Take the list, sort by start date and filter out blank records

    Populate a new xlsx spread sheet that is properly formatted
    - First sheet called "Listing" will be the full list
    - 12 sheets after that will the the records per month

    Save the input file to .BAK and then create a new version of the file

    :param xlsfile: (str) input filename
    :param xlsaref: (list of dicts) recodrds from that file from sheet "Listing"
    :param fldFirstNight: (str) column header of the first night date column
    :param fldNights: (str) column header for the field holding the int # of nights
    :param xlsdateflds: (list) of field names that are date fields

    :return bak_filename: (str) - filename we converted the input filename into
    """

    # move the file so we can output with the same filename
    fname, fext = os.path.splitext( xlsfile)
    new_fname = fname + '.bak'
    if os.path.exists(new_fname):
        os.remove(new_fname)
    os.rename(xlsfile, new_fname)
    
    # find records where first date and number of nights are filled in
    xlsaref = [x for x in xlsaref if x[fldFirstNight] and x[fldNights]]

    # sort these records so they are in date order
    xlsaref = sorted(xlsaref, key=itemgetter(fldFirstNight))

    # create workbook for output
    wb = openpyxl.Workbook()

    # set the first sheet name
    ws = wb.active
    ws.title = SHEET_LISTING

    # pull out the keys
    header_keys = ['' if 'blank' in x else x for x in list(xlsaref[0].keys())]
    dict_keys = [x for x in list(xlsaref[0].keys())]
    
    # build out the listing sheet
    # header
    for colidx, key in enumerate(header_keys, start=1):
        a1 = ws.cell(row=1, column=colidx, value=key)
        a1.font = bold
        if key in COL_CENTERED:
            a1.alignment = fit_centered
        else:
            a1.alignment = fit
        if key in COL_WIDTH:
            ws.column_dimensions[get_column_letter(colidx)].width = COL_WIDTH[key]
        
    # data records
    for recidx, rec in enumerate(xlsaref, start=2):
        for colidx, key in enumerate(dict_keys, start=1):
            a1 = ws.cell(row=recidx, column=colidx, value=rec[key])
            a1.font = regular
            if key in COL_CENTERED:
                a1.alignment = fit_centered
            else:
                a1.alignment = fit
            if key in xlsdateflds:
                a1.number_format="MM/DD/YYYY"
            
    # build out month oriented tabs
    for mon in range(1,13):
        # create the new sheet
        ws = wb.create_sheet(title=MON_STRING[mon-1])
        # header
        for colidx, key in enumerate(header_keys, start=1):
            a1 = ws.cell(row=1, column=colidx, value=key)
            a1.font = bold
            if key in COL_CENTERED:
                a1.alignment = fit_centered
            else:
                a1.alignment = fit
            if key in COL_WIDTH:
                ws.column_dimensions[get_column_letter(colidx)].width = COL_WIDTH[key]

        # extract out data for this sheet
        monaref = [x for x in xlsaref if x[fldFirstNight].month == mon]

        # data records
        for recidx, rec in enumerate(monaref, start=2):
            for colidx, key in enumerate(dict_keys, start=1):
                a1 = ws.cell(row=recidx, column=colidx, value=rec[key])
                a1.font = regular
                if key in COL_CENTERED:
                    a1.alignment = fit_centered
                else:
                    a1.alignment = fit
                if key in xlsdateflds:
                    a1.number_format="MM/DD/YYYY"

    wb.save(xlsfile)

    return new_fname, xlsaref

def find_and_remove_dup_start_dates(xlsaref, fldFirstNight, fldNights):
    overlap = list()
    new_xlsaref = list()
    max_start_nights = {}
    for rec in xlsaref:
        if rec[fldFirstNight] in max_start_nights:
            max_start_nights[rec[fldFirstNight]] = max(int(rec[fldNights]), max_start_nights[rec[fldFirstNight]])
            overlap.append(rec[fldFirstNight])
        else:
            max_start_nights[rec[fldFirstNight]] = int(rec[fldNights])

    for rec in xlsaref:
        if int(rec[fldNights]) == max_start_nights[rec[fldFirstNight]]:
            new_xlsaref.append(rec)

    return new_xlsaref, overlap
               
# routine that reads the XLS, converts the data, and saves it to the output file
# Global variables used:
#    DATE_FMT - format string for the date string
#    ADD_ONE_DAY - delta date value that adds one day
#    OCC_TYPE_CONV - conversion of the occupytype 
#
def load_convert_save_file( xlsfile, req_cols, occupy_filename, fldFirstNight, fldNights, fldType, xlsdateflds, debug=False ):

    # logging
    logger.info('Read in XLS:%s', xlsfile)
    
    # read in the XLS
    xlsaref=kvxls.readxls2list_findheader(xlsfile,req_cols=req_cols,optiondict={'dateflds' : xlsdateflds, 'sheetname': 'Listing'},debug=False)

    # reformat the file and get back the filename of the original file renamed
    # and return the list of records sorted by firstNight
    bak_fname, xlsaref = rewrite_file( xlsfile, xlsaref, fldFirstNight, fldNights, xlsdateflds)

    logger.info("Saved orig file, reformatted file: %s",
                {'orig_file': xlsfile,
                 'bak_file': bak_fname})
    
    # remove duplicates if any exist
    xlsaref, overlap = find_and_remove_dup_start_dates(xlsaref, fldFirstNight, fldNights)

    if overlap:
        logger.info('Multiple records with same start night: %s',
                    {'overlap': overlap})
    
    # capture if the current renter is still there - their start date
    current_guest_start = None

    # get the current date
    now = datetime.datetime.now()
    
    # logging
    logger.info('Create occupancy file:%s', occupy_filename)
    
    # create the output file and start the conversion/output process
    with open( occupy_filename, 'w' ) as t:
        # create the header to the file
        t.write('date,occtype\n')

        # set the first exit date to something that will NOT match
        exitdate = datetime.datetime(2019,1,1)
        
        # run through the records read in
        for rec in xlsaref:
            # process the record from the file - convert first night into a date variable
            # eventdate = datetime.datetime.strptime(rec[fldFirstNight], DATE_FMT)
            eventdate = rec[fldFirstNight]
            
            # add to the dictionary the datetime value just calculated
            rec['startdate'] = eventdate

            # for each night of stay - plus occ_type days to model the day the guest exits
            for cnt in range(int(rec[fldNights]) + OCC_TYPE_CONV[rec[fldType]][1]):
                # skip the date if this date matches the exit date of the prior
                if eventdate == exitdate:
                    # skip the date if this date matches the exit date of the prior
                    logger.info('Start date is same as the last guests exit date:skip record creation:%s', exitdate)
                else:
                    # convert the eventdate to a string
                    eventdate_str = datetime.datetime.strftime(eventdate,DATE_FMT)
                    # output this value
                    t.write('%s,%s\n'%(eventdate_str,OCC_TYPE_CONV[rec[fldType]][0]))
                    # set the exit date to the last date written out
                    exitdate = eventdate

                # capture the date prior to looping
                rec['exitdate'] = exitdate

                # add one day to this date and loop
                eventdate += ADD_ONE_DAY

            # capture current guest if it exists
            if rec['startdate'] < now and rec['exitdate']> now:
                current_guest_start = rec['startdate']
                
        # return the BP file with modifications
        return xlsaref, current_guest_start
    
# routine that read in the history stays file and the current stays files and addes to
# the history file any dates from the current stays file that are in the past
#
def migrate_stays_to_history( occupy_filename, occupy_history_filename, fldDate, debug=False ):
    # log that we are doing this work
    # load stay history
    if os.path.isfile( occupy_history_filename ):
        logger.info('migrate_stays_to_history:load file:%s', occupy_history_filename)
        stays_history = villaecobee.load_villa_calendar( occupy_history_filename, fldDate, debug=False )
    else:
        logger.info('migrate_stays_to_history:file does not exist:%s', occupy_history_filename)
        stays_history = dict()

    # load current stay information
    stays = villaecobee.load_villa_calendar( occupy_filename, fldDate, debug=False )
    logger.info('migrate_stays_to_history:load file:%s', occupy_filename)

    # capture today
    today = datetime.datetime.today()

    # and capture the number of records add to the stay history
    records_added=0
    
    # step through the stays file and look for past due dates
    for staydate in stays:
        if not staydate:
            # skip blanks
            continue
        if datetime.datetime.strptime(staydate, DATE_FMT) < today:
            # this date is in the past - see if this date is in the history already
            if staydate not in stays_history:
                # add this date to stays history
                stays_history[staydate] = stays[staydate]
                # increment the counter
                records_added += 1
                # debugging message
                logger.debug('migrate_stays_to_history:date added:%s', staydate)
            else:
                # debugging message
                logger.debug('migrate_stays_to_history:date skipped:%s', staydate)

    # loop through - now if we added records we need to save stay history data
    if records_added:
        logger.info('migrate_stays_to_history:records added to history:%d', records_added)
        kvcsv.writedict2csv( occupy_history_filename, stays_history )
    else:
        logger.info('migrate_stays_to_history:no records added to history')
        
# ---------------------------------------------------------------------------
if __name__ == '__main__':

    # capture the command line
    optiondict = kvutil.kv_parse_command_line( optiondictconfig, debug=False )

    # set variables based on what came form command line
    debug = optiondict['debug']

    # logging
    kvutil.loggingAppStart( logger, optiondict, kvutil.scriptinfo()['name'] )

    # migrate stays to history
    migrate_stays_to_history(optiondict['occupy_filename'], optiondict['occupy_history_filename'], optiondict['fldDate'], debug=False)
    
    # load and convert the XLS to create the TXT
    xlsaref,current_guest_start = load_convert_save_file( optiondict['xls_filename'],
                                                          COL_REQUIRED,
                                                          optiondict['occupy_filename'],
                                                          optiondict['fldFirstNight'],
                                                          optiondict['fldNights'],
                                                          optiondict['fldType'],
                                                          optiondict['xlsdateflds'],
                                                          debug=optiondict['debug'] )

    # if the google calendar sync flag is set - sync
    if optiondict['calendarsync']:
        # determine the starting date for the run
        if optiondict['startback']:
            now = datetime.datetime.now()+datetime.timedelta(days=optiondict['startback'])
        elif optiondict['startdate']:
            now = optiondict['startdate']
        elif current_guest_start:
            now = current_guest_start
        else:
            now = datetime.datetime.now()
        # now update the calendar
        villacalendar.sync_villa_cal_with_bp_xls( xlsaref, now=now, debug=optiondict['debug'] )

    # validate we can load this file after we created it
    logger.info('Load newly created file looking for problems')
    try:
        villaecobee.load_villa_calendar( optiondict['occupy_filename'], optiondict['fldDate'], debug=False )
    except:
        # logging
        logger.error('Error in newly created file:%s', optiondict['occupy_filename'])

        # display message
        print("ERROR:unable to load the newly created file:see error in line above")
        print('Please correct file:', optiondict['occupy_filename'] )
        