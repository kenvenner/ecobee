'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.16

Read information from Beautiful Places XLS files,
extract out occupancy data, build a new
output file that has an entry per day of stay and stay type

'''
import kvutil
import kvxls
import kvcsv

import datetime
import sys

# this utility is used to convert the Beautiful Places Villa bookings XLS
# file into the flattened text file used by the "villaecobee.py" file
#
# General process:
# 1) get a new xls from beautiful places
# 2) save it to the directory/filename where this tool is stored
# 3) run this tool:  python vcconvert.py


import os
import logging
# Logging Setup
# logging.basicConfig(level=logging.INFO)
logging.basicConfig(filename=os.path.splitext(kvutil.scriptinfo()['name'])[0]+'.log',
                    level=logging.INFO,
                    format='%(asctime)s %(levelname)s %(name)s:%(lineno)d %(funcName)s %(message)s')
logger = logging.getLogger(__name__)


# we are reusing features from another applicatoin but we wnat
# the log files to be tied to this applicatoin - so we call tihs
# application/library late in order to have the logger ocnfigured to THIS app
import villaecobee
import villacalendar

# application variables
optiondictconfig = {
    'AppVersion' : {
        'value' : '1.16',
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
        'value' : 'VillaCarnerosCalendar.xls',
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
        'value'       : ['First Night','Last Night','Created'],
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
addoneday =  datetime.timedelta(days=1)
datefmt='%m/%d/%Y'

# xls header definition
req_cols=['Booking','Unit Code','First Night','Last Night','Nights','Type']

# xls file occtype conversion to
# an array that is:  new code and # of days to add to stay for temp control
occtype_conv = {
    'Hold - Clean' : ['C', 0],
    'Hold - Maint.' : ['M', 0],
    'Hold- Mainten' : ['M', 0],
    'Hold - Other' : ['M', 0],
    'Hold - Owner' : ['O', 1],
    'Hold - Renter' : ['R', 1],
    'Hold - Winery Business' : ['O',0],
    'Hold - Fall Mbr Event' : ['O',1],
    'Hold - Ken Venner' : ['O',1],
    'Res. - Renter' : ['R', 1],
    'Res. - Owner' : ['O', 1],
    'Res - Renter' : ['R', 1],
    'Res - Owner' : ['O', 1],
}

# routine that reads the XLS, converts the data, and saves it to the output file
# Global variables used:
#    datefmt - format string for the date string
#    addoneday - delta date value that adds one day
#    occtype_conv - conversion of the occupytype 
#
def load_convert_save_file( xlsfile, req_cols, occupy_filename, fldFirstNight, fldNights, fldType, xlsdateflds, debug=False ):

    # logging
    logger.info('Read in XLS:%s', xlsfile)
    
    # read in the XLS
    xlsaref=kvxls.readxls2list_findheader(xlsfile,req_cols=req_cols,optiondict={'dateflds' : xlsdateflds},debug=False)

    #print(xlsaref)
    #sys.exit(1)

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
            # eventdate = datetime.datetime.strptime(rec[fldFirstNight], datefmt)
            eventdate = rec[fldFirstNight]
            
            # add to the dictionary the datetime value just calculated
            rec['startdate'] = eventdate

            # for each night of stay - plus occ_type days to model the day the guest exits
            for cnt in range(int(rec[fldNights]) + occtype_conv[rec[fldType]][1]):
                # skip the date if this date matches the exit date of the prior
                if eventdate == exitdate:
                    # skip the date if this date matches the exit date of the prior
                    logger.info('Start date is same as the last guests exit date:skip record creation:%s', exitdate)
                else:
                    # convert the eventdate to a string
                    eventdate_str = datetime.datetime.strftime(eventdate,datefmt)
                    # output this value
                    t.write('%s,%s\n'%(eventdate_str,occtype_conv[rec[fldType]][0]))
                    # set the exit date to the last date written out
                    exitdate = eventdate

                # capture the date prior to looping
                rec['exitdate'] = exitdate

                # add one day to this date and loop
                eventdate += addoneday

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
        if datetime.datetime.strptime(staydate, datefmt) < today:
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
    xlsaref,current_guest_start = load_convert_save_file( optiondict['xls_filename'], req_cols, optiondict['occupy_filename'], optiondict['fldFirstNight'], optiondict['fldNights'], optiondict['fldType'], optiondict['xlsdateflds'], debug=optiondict['debug'] )

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
        
