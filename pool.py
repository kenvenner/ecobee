'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.13

Take the output from "screenlogic > output.txt" 
and parse that data and create append the output
to the output filename

'''
import os.path
import os
import logging
import re
import datetime
import kvutil
import kvgmailsendsimple
import kvdate
# import poolapi

# CONSTANTS
DAY_SECONDS = 60 * 60 * 24
FIFTEEN_MIN_SECONDS = 60 * 15
FOUR_HOUR_SECONDS = 60 * 60 * 4
MAX_POOL_TEMP = 85.0


def modification_date(filename):
    '''
    return the last modified date on a file in datetime format
    '''
    t = os.path.getmtime(filename)
    return datetime.datetime.fromtimestamp(t)

def modification_days_and_seconds(filename):
    '''
    return an array that has the (days, seconds) of the file modification date
    '''
    file_datetime = modification_date(filename)
    file_duration = datetime.datetime.now() - file_datetime
    return divmod(file_duration.total_seconds(), DAY_SECONDS)

               
# Logging Setup
# logging.basicConfig(level=logging.INFO)
logging.basicConfig(filename=os.path.splitext(kvutil.scriptinfo()['name'])[0]+'.log',
                    level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(threadName)s -  %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


# application variables
optiondictconfig = {
    'AppVersion' : {
        'value': '1.13',
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
    'conf_json' : {
        'value' : ['pool.json'],
        'description' : 'defines the json configuration file to be read',
    },
    'direct_connect' : {
        'value' : False,
        'type': 'bool',
        'description' : 'when true - we talk directly to the pool controller',
    },
    'pool_ip' : {
        'value' : '192.168.8.141',
        'description' : 'IP address of the pool controller',
    },
    'input_filename' : {
        'value' : 'output.txt',
        'description' : 'defines the name of the file generated from screenlogic',
    },
    'pool_filename' : {
        'value' : 'pool_temps.csv',
        'description' : 'defines the name of the file that holds the temperature readings',
    },
    'pool_heater_filename' : {
        'value' : 'pool_heater.lck',
        'description' : 'defines the name of the file that says we sent a message about pool heater being on',
    },
    'pool_missing_filename' : {
#        'value' : 'pool_missing.lck',
        'value' : None,  # not set - we want to alert spa team not pool team
        'description' : 'defines the name of the file that says we sent a message about pool settings not being read',
    },
    'pool_heater_off_filename' : {
        'value' : 'pool_heater_off.lck',
        'description' : 'defines the name of the file that says we need to turn off the pool heater',
    },
    'pool_heater_allowed_filename' : {
        'value' : 'pool_heater_allowed.txt',
        'description' : 'defines the name of the file that hold the list of dates we enable the pool heater to work',
    },
    'pool_email_from' : {
        'value' : '210608thSt@gmail.com',
        'description' : 'who sends out the email about pool heater on',
    },
    'pool_email_to' : {
        'value' : 'ken@vennerllc.com, mike.kmsdev2@outlook.com, reservations@michelleleighvacationrentals.com',
#        'value' : 'ken@vennerllc.com',  # uncomment for testing purposes
        'description' : 'defines the name of the file that says we sent a message about pool heater being on',
    },
    'pool_email_subject' : {
        'value' : 'Villa Carneros Pool Heater is ',
        'description' : 'defines the name of the file that says we sent a message about pool heater being on',
    },
    'pool_email_body' : {
        'value' : 'We have just detected that the pool heater is ',
        'description' : 'defines the name of the file that says we sent a message about pool heater being on',
    },
    'spa_heater_filename' : {
        'value' : 'spa_heater.lck',
        'description' : 'defines the name of the file that says we sent a message about pool heater being on',
    },
    'spa_missing_filename' : {
        'value' : 'spa_missing.lck',
        'description' : 'defines the name of the file that says we sent a message about pool settings not being read',
    },
    'spa_heater_off_filename' : {
        'value' : 'spa_heater_off.lck',
        'description' : 'defines the name of the file that says we need to turn off the spa heater',
    },
    'spa_heater_off_hours' : {
        'value' : 3.0,
        'description' : 'defines the number of hours we allow the SPA to remain on until we automatically turn it off',
    },
    'spa_email_from' : {
        'value' : '210608thSt@gmail.com',
        'description' : 'who sends out the email about spa heater on',
    },
    'spa_email_to' : {
        'value' : 'ken@vennerllc.com, mike.kmsdev2@outlook.com, reservations@michelleleighvacationrentals.com',
#        'value' : 'ken@vennerllc.com',
        'description' : 'defines the name of the file that says we sent a message about spa heater being on',
    },
    'spa_email_subject' : {
        'value' : 'Villa Carneros SPA Heater is ',
        'description' : 'defines the name of the file that says we sent a message about spa heater being on',
    },
    'spa_email_body' : {
        'value' : 'We have just detected that the spa heater is ',
        'description' : 'defines the name of the file that says we sent a message about spa heater being on',
    },
    'scopes' : {
        'value' : None,
        'description' : 'defines the gmail scopes used to generate and send emails - see kvgmailsendsimple.py',
    },
    'file_token_json' : {
        'value' : None,
        'description' : 'defines the gmail filename of the json file that contains the account token (access and refresh) ',
    },
    'file_credentials_json' : {
        'value' : None,
        'description' : 'defines the gmail filename of the json file that contains the account credentials ',
    },
}

### GLOBAL VARIABLES AND CONVERSIONS ###


# set the time for the run
now = datetime.datetime.now()
now_str = now.strftime('%Y-%m-%d:%H:%M:%S')



def check_file_writable(fnm):
    if os.path.exists(fnm):
        # path exists
        if os.path.isfile(fnm): # is it a file or a dir?
            # also works when file is a link and the target is writable
            return os.access(fnm, os.W_OK)
        else:
            return False # path is a dir, so cannot write as a file
    # return False - file is not there
    return False

    # target does not exist, check perms on parent dir
    pdir = os.path.dirname(fnm)
    if not pdir: pdir = '.'
    # target is creatable if parent dir is writable
    return os.access(pdir, os.W_OK)


def read_pool_heater_allowable_file(input_file):
    '''
    if file exists, read in the file and convert each line to a date and build a list of dates
    that we will not flag the pool is enabled and attempt to turn it off
    '''
    pool_heater_allowed = []
    pool_heater_invalid_dates = []

    # no file - so no inputs
    if not os.path.exists(input_file):
        logger.info(input_file + ' not found')
        return pool_heater_allowed, pool_heater_invalid_dates

    # get the file read in the lines and convert the string to date
    with open(input_file, 'r') as file:
        # Read each line in the file
        for idx, line in enumerate(file):
            try:
                pool_heater_allowed.append(kvdate.datetime_from_str(line.strip()).date())
            except Exception as e:
                pool_heater_invalid_dates.append(f'{idx+1}|{line.strip()}|{e}')

    logger.info(str(len(pool_heater_allowed)) + ' dates allowed to have pool enabled')
    return pool_heater_allowed, pool_heater_invalid_dates

def save_data_to_output(now_str, output_file, result_values, result_keys):

    out_header = ['now_str'] + result_keys
    out_data = [now_str] + result_values

    # append results
    file_writeable = check_file_writable( output_file )


    # debugging
    # print('file_writeable: ', file_writeable)
    
    
    # open file for output
    with open(output_file, "a") as file1:
        # create header if file does not exist
        if not file_writeable:
            # write header if it doe snot exist
            file1.write(','.join(out_header)
                        +'\n')
        
        # Writing data to a file
        file1.write(','.join(out_data)
                    +'\n')

        # logging
        logger.info('Appended record to: %s ', output_file)
        
    
def read_parse_output_pool(input_file, output_file):

    # Using readlines()
    file1 = open(input_file, 'r')
    Lines = file1.readlines()

    # logging
    logger.info('Read in pool data from:  %s', input_file)
    logger.info('Lines in this file:  %d', len(Lines))

    # if the lines is not greater than 20 we did not get valid run
    if len(Lines) < 20:
        logger.info('Insufficient lines created - unable to parse file - EXITTING')
        return
    
    # Debugging
    # print(Lines)
    
    # count on lines
    count = 0
    
    # Strips the newline character
    for line in Lines:
        count += 1

        # debugging
        # print("Line{}: {}".format(count, line.strip()))

        m = re.search('Pool temperature is last\s+(\d+)', line)
        
        # print( count, m, line )
        if m:
            pool_temp_last = m.group(1)
            # print( pool_temp_last  )
            
            
        m = re.search('Pool Heat Set Point:\s+(\d+)', line)

        # print( count, m, line )
        if m:
            pool_temp_set = m.group(1)
            # print( pool_temp_set )

        m = re.search('Pool Heat:\s+(.+)', line)

        # print( count, m, line )
        if m:
            pool_heat_set = m.group(1)
            # print( pool_heat_set )


        m = re.search('Pool Heat Mode:\s+(.+)', line)

        # print( count, m, line )
        if m:
            pool_heat_mode = m.group(1)
            # print( pool_heat_mode)


        m = re.search('Spa temperature is last\s+(\d+)', line)

        # print( count, m, line )
        if m:
            spa_temp_last = m.group(1)
            # print( spa_temp_last )


        m = re.search('Spa Heat Set Point:\s+(\d+)', line)

        # print( count, m, line )
        if m:
            spa_temp_set = m.group(1)
            # print( spa_temp_set )

        m = re.search('Spa Heat:\s+(.+)', line)

        # print( count, m, line )
        if m:
            spa_heat_set = m.group(1)
            # print( spa_heat_set )
        

        m = re.search('Spa Heat Mode:\s+(.+)', line)

        # print( count, m, line )
        if m:
            spa_heat_mode = m.group(1)
            # print( spa_heat_mode )

    
    # call routine to save out the results
    result_keys = ['pool_temp_last','pool_temp_set','pool_heat_set','pool_heat_mode',
                   'spa_temp_last','spa_temp_set','spa_heat_set','spa_heat_mode']
    result_values = [pool_temp_last, pool_temp_set, pool_heat_set, pool_heat_mode,
                     spa_temp_last, spa_temp_set, spa_heat_set, spa_heat_mode]
    save_data_to_output(now_str, output_file, result_values, result_keys)
    

    # remove the file if it exists
    if os.path.isfile(input_file):
        # remove the file
        os.remove(input_file)
        # logging
        logger.info('Removed input file:  %s', input_file)

    # return what we just read in
    return {
        'pool_temp_last': pool_temp_last,
        'pool_temp_set': pool_temp_set,
        'pool_heat_set': pool_heat_set,
        'pool_heat_mode': pool_heat_mode,
        'spa_temp_last': spa_temp_last,
        'spa_temp_set': spa_temp_set,
        'spa_heat_set': spa_heat_set,
        'spa_heat_mode': spa_heat_mode
    }


def read_direct_output_pool(ip, output_file):
    '''
    Read pool settings directly and save to output file
    '''

    # read using the library
    try:
        # result_values, result_keys = poolapi.main(ip)
        print('ken')
    except Exception as e:
        logger.info('Could not read values - EXITTING')
        logger.error(e)
        
    # save the results to the output file
    save_data_to_output(now_str, output_file, result_values, result_keys)

    # return dictoinary
    return dict(zip(result_keys, result_values))

    
def message_on_pool_state_change(pool_settings, optiondict, pool_heater_allowed):
    ''' create an email when the state changes on pool heater
    using a lock file to capture what the state currently is

    pool_settings - dict of values read in 
    optiondict - the options dictionary
    pool_heater_invalid_dates - list of dates the pool is enabled to be on

    '''

    msgid = None

    ### NO DATA READ - POOL
    
    # special processing when we get no pool_settings
    if not pool_settings:
        if not optiondict['pool_missing_filename']:
            # we are not alerting on pool issues
            return
        
        # check to see if thereis a pool not woking lock file
        if os.path.isfile(optiondict['pool_missing_filename']):
            # check the age/duraction of this file to see if we should send again
            pool_days, pool_seconds = modification_days_and_seconds(optiondict['pool_missing_filename'])        

            # if we have not met the next notification window - skip
            if pool_seconds < FOUR_HOUR_SECONDS:
                # take no action yet
                return

            # we got here - so we need to send another reminder
            logger.info('Time threshold exceed - sending another reminder')
            
        # create message that we are not currently reading pool settings
        # send message that heater is ON
        msgid = kvgmailsendsimple.gmail_send_simple_message(
            optiondict['pool_email_from'],
            optiondict['pool_email_to'],
            optiondict['pool_email_subject']+'Not Reading Pool Settings',
            optiondict['pool_email_body']+'Not Reading Pool Settings',
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
        )

        # create the lock file
        with open(optiondict['pool_missing_filename'], 'w') as lock_file:
            lock_file.write('Not Reading Pool Settings')

        # log message
        logger.info('Not reading pool settings - sent message: %s and created file: %s', msgid['id'], optiondict['pool_missing_filename'])

        # return - we have nothing to process
        return
    elif optiondict['pool_missing_filename'] and os.path.exists(optiondict['pool_missing_filename']):
        # we are getting data and we have a lock file that must now be removed

        # send message that heater is off
        msgid = kvgmailsendsimple.gmail_send_simple_message(
            optiondict['pool_email_from'],
            optiondict['pool_email_to'],
            optiondict['pool_email_subject']+'NOW Reading Pool Settings',
            optiondict['pool_email_body']+'NOW Reading Pool Settings',
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
        )

        # remove the lock file
        os.remove(optiondict['pool_missing_filename'])
        
        # log message
        logger.info('NOW reading pool settings - sent message: %s and removed file: %s', msgid['id'], optiondict['pool_missing_filename'])
        
    # is it ok to have the pool on
    if datetime.datetime.now().date() in pool_heater_allowed:
        pool_ok = '\nThe pool is allowed to be on'
    else:
        pool_ok = '\nThe pool is NOT supposed to be on - automation should shut it off'

    # POOL
    if os.path.isfile(optiondict['pool_heater_filename']):
        # if there is a lock file - capture the informatoin about this lock file
        pool_days, pool_seconds = modification_days_and_seconds(optiondict['pool_heater_filename'])        

        # and the pool heater is not ON message
        # that the pool heater turned off and remove the lock file
        if pool_settings['pool_heat_mode'] == 'Off':
            # send message that heater is off
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['pool_email_from'],
                optiondict['pool_email_to'],
                optiondict['pool_email_subject']+'OFF',
                optiondict['pool_email_body']+'OFF',
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # remove the lock file
            os.remove(optiondict['pool_heater_filename'])

            # log message
            logger.info('Pool heater off - sent message: %s and removed file: %s', msgid['id'], optiondict['pool_heater_filename'])

        elif pool_days and pool_seconds < FIFTEEN_MIN_SECONDS:
            # we have a lock file and we have had this lock file exist for more than a day
            # we are greater than a day and less then the first 15 minutes of that next day
            # we should send another message about the duratoin of this being on
            # send message that heater is off
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['pool_email_from'],
                optiondict['pool_email_to'],
                optiondict['pool_email_subject']+'STILL ON - DAY ' + str(pool_days),
                'Pool Heater continues to be on' + pool_ok,
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # log message
            logger.info('Pool heater still ON [%s] days - sent message: %s and removed file: %s',
                        pool_days, msgid['id'], optiondict['pool_heater_filename'])
        elif pool_settings['pool_temp_set'] and float(pool_settings['pool_temp_set']) > MAX_POOL_TEMP:
            # SETTING GREATER THAN MAX
            # check to see if the pool setting exceeds our max
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['pool_email_from'],
                optiondict['pool_email_to'],
                optiondict['pool_email_subject']+'SET OVER THE MAX SETTING:  ' + str(MAX_POOL_TEMP),
                'Pool Heater set to a temp ' + pool_settings['pool_temp_set'] + ' that is over MAX SETTING:  ' + str(MAX_POOL_TEMP) + pool_ok,
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # log message
            logger.info('Pool heater set over max [%s/%s] days - sent message: %s and removed file: %s',
                        pool_settings['pool_temp_set'], str(MAX_POOL_TEMP), msgid['id'], optiondict['pool_heater_filename'])
            
    else:

        # if there is NO lock file

        # and the pool heater is ON message
        # that the pool heater is now ON and create a lock file.
        if pool_settings['pool_heat_mode'] != 'Off':
            # send message that heater is ON
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['pool_email_from'],
                optiondict['pool_email_to'],
                optiondict['pool_email_subject']+'ON',
                optiondict['pool_email_body']+'ON' + pool_ok,
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # create the lock file
            with open(optiondict['pool_heater_filename'], 'w') as lock_file:
                lock_file.write('Pool ON')

            # log message
            logger.info('Pool heater ON - sent message: %s and created file: %s', msgid['id'], optiondict['pool_heater_filename'])
            
            # SETTING GREATER THAN MAX - only check when we just turned on the heat
            if pool_settings['pool_temp_set'] and float(pool_settings['pool_temp_set']) > MAX_POOL_TEMP:
                # check to see if the pool setting exceeds our max
                msgid = kvgmailsendsimple.gmail_send_simple_message(
                    optiondict['pool_email_from'],
                    optiondict['pool_email_to'],
                    optiondict['pool_email_subject']+'SET OVER THE MAX SETTING:  ' + str(MAX_POOL_TEMP),
                    'Pool Heater set to a temp ' + pool_settings['pool_temp_set'] + ' that is over MAX SETTING:  ' + str(MAX_POOL_TEMP) + pool_ok,
                    optiondict['scopes'],
                    optiondict['file_token_json'],
                    optiondict['file_credentials_json']
                )
                
                # log message
                logger.info('Pool heater set over max [%s/%s] days - sent message: %s and removed file: %s',
                            pool_settings['pool_temp_set'], str(MAX_POOL_TEMP), msgid['id'], optiondict['pool_heater_filename'])
                
                # if we have the ability turn off the heat
                if optiondict['pool_ip']:
                    print('ken')
                    # poolapi.set_pool_temp(ip=optiondict['pool_ip'], desired_temp=MAX_POOL_TEMP)

    # return back the message id or none
    return msgid
    
def message_on_pool_turn_off(pool_settings, pool_heater_allowed, pool_heater_invalid_dates, optiondict):
    ''' create an email when we are creating a file that will turn off the pool

    pool_settings - dict of values read in
    pool_heater_allowed - list of datetime values where the pool can be on
    pool_heater_invalid_dates - list of strings and row numbers where we could not convert the string to a date
    optiondict - the options dictionary

    check to see if we have invalid date in the import file and message
    check to see if the list of valid dates for the pool to be on is today and if so - don't turn it off

    '''

    msgid = None

    ### invalid dates in read file
    if pool_heater_invalid_dates:
        # send message that heater is ON
        msgid = kvgmailsendsimple.gmail_send_simple_message(
            optiondict['pool_email_from'],
            optiondict['pool_email_to'],
            optiondict['pool_email_subject']+'Invalid date lines in file',
            'Unable to convert following lines in file to datetime strings:\n' + '\n'.join(pool_heater_invalid_dates),
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
        )
       
    ### NO DATA READ - POOL
    
    # special processing when we get no pool_settings
    if not pool_settings:
        # we are not alerting on pool issues
        return

    # if there is no filename we do not turn off the heat
    if not optiondict['pool_heater_off_filename']:
        return

    # test to see if this is a non-alerting datetime
    if datetime.datetime.now().date() in pool_heater_allowed:
        # log message
        logger.info('Pool heater ON and pool_heater_allowed is enabled - no action taken')
        return
    
    # the pool heater is ON message
    # create the file to have it turned off and message
    # that we are turning off the pool heater
    if pool_settings['pool_heat_mode'] != 'Off':
        # send message that heater is ON
        msgid = kvgmailsendsimple.gmail_send_simple_message(
            optiondict['pool_email_from'],
            optiondict['pool_email_to'],
            optiondict['pool_email_subject']+'Being Turned OFF',
            optiondict['pool_email_body']+'Being Turned OFF',
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
        )

        # create the lock file
        with open(optiondict['pool_heater_off_filename'], 'w') as lock_file:
            lock_file.write('Pool ON being turned OFF')

        # if we hvae the ability turn off the heat
        if optiondict['pool_ip']:
            print('ken')
            # poolapi.set_heat_mode_pool(ip=optiondict['pool_ip'], mode=0)

            
        # log message
        logger.info('Pool heater ON - being turned OFF sent message: %s and created file: %s', msgid['id'], optiondict['pool_heater_off_filename'])
            

    # return back the message id or none
    return msgid
    
def message_on_spa_state_change(pool_settings, optiondict):
    ''' create an email when the state changes on spa heater
    using a lock file to capture what the state currently is

    pool_settings - dict of values read in 
    optiondict - the options dictionary

    '''

    msgid = None

    # special processing when we get no spa_settings
    if not pool_settings:
        if not optiondict['spa_missing_filename']:
            # we are not alerting on spa issues
            return

        # check to see if thereis a pool not woking lock file
        if os.path.isfile(optiondict['spa_missing_filename']):
            # check the age/duraction of this file to see if we should send again
            spa_days, spa_seconds = modification_days_and_seconds(optiondict['spa_missing_filename'])        

            # if we have not met the next notification window - skip
            if spa_seconds < FOUR_HOUR_SECONDS:
                # take no action yet
                return

            # we got here - so we need to send another reminder
            logger.info('Time threshold exceed - sending another reminder')


        # create message that we are not currently reading pool settings
        # send message that heater is ON
        msgid = kvgmailsendsimple.gmail_send_simple_message(
            optiondict['spa_email_from'],
            optiondict['spa_email_to'],
            optiondict['spa_email_subject']+'Not Reading Pool Settings',
            optiondict['spa_email_body']+'Not Reading Pool Settings',
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
        )
        
        # create the lock file
        with open(optiondict['spa_missing_filename'], 'w') as lock_file:
            lock_file.write('Not Reading Pool Settings')
            
        # log message
        logger.info('Not reading pool settings - sent message: %s and created file: %s', msgid['id'], optiondict['spa_missing_filename'])

        # return - we have nothing to process
        return
    elif optiondict['spa_missing_filename'] and os.path.exists(optiondict['spa_missing_filename']):
        # we are getting data and we have a lock file that must now be removed

        # send message that heater is off
        msgid = kvgmailsendsimple.gmail_send_simple_message(
            optiondict['spa_email_from'],
            optiondict['spa_email_to'],
            optiondict['spa_email_subject']+'NOW Reading Pool Settings',
            optiondict['spa_email_body']+'NOW Reading Pool Settings',
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
        )

        # remove the lock file
        os.remove(optiondict['spa_missing_filename'])

        # log message
        logger.info('NOW reading pool settings - sent message: %s and removed file: %s', msgid['id'], optiondict['spa_missing_filename'])
        
            
    # SPA
    if os.path.isfile(optiondict['spa_heater_filename']):
        # if there is a lock file - capture the informatoin about this lock file
        spa_days, spa_seconds = modification_days_and_seconds(optiondict['spa_heater_filename'])
        
        # and the spa heater is not ON message
        # that the spa heater turned off and remove the lock file
        if pool_settings['spa_heat_mode'] == 'Off':
            # send message that heater is off
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['spa_email_from'],
                optiondict['spa_email_to'],
                optiondict['spa_email_subject']+'OFF',
                optiondict['spa_email_body']+'OFF',
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # remove the lock file
            os.remove(optiondict['spa_heater_filename'])

            # log message
            logger.info('SPA heater off - sent message: %s and removed file: %s', msgid['id'], optiondict['spa_heater_filename'])

        elif optiondict['spa_heater_off_filename'] and optiondict['spa_heater_off_hours'] and spa_seconds > optiondict['spa_heater_off_hours'] * 60 * 60:
            # we have a desire to turn off the spa because we defined two variable - filename and hours
            # and the spa has been on longer than that defined max time
            # so generate the file that causes the spa to be disabled
            # and send message that we are going to turn off the spa
            
            # send message that heater is ON
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['spa_email_from'],
                optiondict['spa_email_to'],
                optiondict['spa_email_subject']+'Being Turned OFF',
                optiondict['spa_email_body']+'Being Turned OFF',
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # create the lock file
            with open(optiondict['spa_heater_off_filename'], 'w') as lock_file:
                lock_file.write('SPA ON being turned OFF')

            # if we hvae the ability turn off the heat
            if optiondict['pool_ip']:
                print('ken')
                # poolapi.set_heat_mode_spa(ip=optiondict['pool_ip'], mode=0)

            # log message
            logger.info('SPA heater on too long turning off SPA - sent message: %s and created file: %s', msgid['id'], optiondict['spa_heater_off_filename'])
            
        elif spa_days and spa_seconds < FIFTEEN_MIN_SECONDS:
            # we have a lock file and we have had this lock file exist for more than a day
            # we are greater than a day and less then the first 15 minutes of that next day
            # we should send another message about the duratoin of this being on
            # send message that heater is off
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['spa_email_from'],
                optiondict['spa_email_to'],
                optiondict['spa_email_subject']+'STILL ON - DAY ' + str(spa_days),
                'SPA Heater continues to be on',
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # log message
            logger.info('SPA heater still ON [%s] days - sent message: %s and removed file: %s',
                        spa_days, msgid['id'], optiondict['spa_heater_filename'])
            
    else:

        # if there is NO lock file

        # and the spa heater is ON message
        # that the spa heater is now ON and create a lock file.
        if pool_settings['spa_heat_mode'] != 'Off':
            # send message that heater is ON
            msgid = kvgmailsendsimple.gmail_send_simple_message(
                optiondict['spa_email_from'],
                optiondict['spa_email_to'],
                optiondict['spa_email_subject']+'ON',
                optiondict['spa_email_body']+'ON',
                optiondict['scopes'],
                optiondict['file_token_json'],
                optiondict['file_credentials_json']
            )

            # create the lock file
            with open(optiondict['spa_heater_filename'], 'w') as lock_file:
                lock_file.write('SPA ON')

            # log message
            logger.info('SPA heater ON - sent message: %s and created file: %s', msgid['id'], optiondict['spa_heater_filename'])


    # return back the message id or none
    return msgid
    
# ---------------------------------------------------------------------------
if __name__ == '__main__':

    # capture the command line
    optiondict = kvutil.kv_parse_command_line( optiondictconfig, debug=False )

    # set variables based on what came form command line
    debug = optiondict['debug']

    # print header to show what is going on (convert this to a kvutil function:  kvutil.loggingStart(logger,optiondict))
    kvutil.loggingAppStart( logger, optiondict, kvutil.scriptinfo()['name'] )

    # refresh the token - always do this as we don't always send an email
    kvgmailsendsimple.gmail_refresh_token_take_no_action(
            optiondict['pool_email_from'],
            optiondict['scopes'],
            optiondict['file_token_json'],
            optiondict['file_credentials_json']
    )
    # log message
    logger.info('Refreshed the gmail token')
        
    # process the pool file
    logger.info( "Call read and save pool data function" )
    if optiondict['direct_connect']:
        if not optiondict['pool_ip']:
            logger.error('Direct connect enabled - no pool_ip defined')
            sys.exit(1)

        pool_settings = read_direct_output_pool(optiondict['pool_ip'], optiondict['pool_filename'])
    else:
        pool_settings = read_parse_output_pool(optiondict['input_filename'], optiondict['pool_filename'])

    # POOL - capture valid dates for pool to be enabled
    pool_heater_allowed, pool_heater_invalid_dates = read_pool_heater_allowable_file(optiondict['pool_heater_allowed_filename'])

    # POOL - determine if we need to message people
    message_on_pool_state_change(pool_settings, optiondict, pool_heater_allowed)
    
    # SPA determine if we need to message people
    message_on_spa_state_change(pool_settings, optiondict)
    
    # POOL - generate file to turn off pool
    message_on_pool_turn_off(pool_settings, pool_heater_allowed, pool_heater_invalid_dates, optiondict)

# eof

