'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.14

Read data from ecobee thermostats, and store to file
Read occupancy from flat file
Set or release temp holds base on occupancy data

'''
import pyecobee
import pprint
import sys
import kvcsv
import kvutil
import datetime
import os
import sys

# this utility queries the villa thermostats and sensor
# saves the information to a text file for review in the future
#
# it also reads the villa usage/scheduling file (stays.txt)
# to determine if the villa is
#   currently occupied - do nothing
#   soon to be occupied - remove the temperature holds (if they exist)
#   vacated - set a temperature hold if none exists
#
# the temperature hold set is defined by the mode of the thermostat
# defined in global variable (holdSetting)
#   when heat is enabled - set the temperature to a low value
#   when AC is enabled - set the temperaure to a high value
#
# this utility is scheduled through a windows schedule task
# to run on a regular basis (every 4 hours) - executed through
# villaecobee.bat BATCH file

import os
import logging
# Logging Setup
# logging.basicConfig(level=logging.INFO)
logging.basicConfig(filename=os.path.splitext(kvutil.scriptinfo()['name'])[0]+'.log',
                    level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(threadName)s -  %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)


# application variables
optiondictconfig = {
    'AppVersion' : {
        'value' : '1.14',
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
        'value' : ['villaecobee.json'],
        'description' : 'defines the json configuration file to be read',
    },
    'api_key' : {
        'value' : None,
        'description' : 'defines the api_key to reference the application - comes from cmd line or json config file',
    },
    'config_filename' : {
        'value' : 'ecobee.conf',
        'description' : 'defines the name of the input file to updated',
    },
    'occupy_filename' : {
        'value' : 'stays.txt',
        'description' : 'defines the name of the file holding the villa occupancy',
    },
    'fldDate' : {
        'value' : 'date',
        'description' : 'defines the name of the field that holds the date in the occupancy file',
    },
    'temperature_filename' : {
        'value' : 'villatemps.txt',
        'description' : 'defines the name of the file that holds the temperature readings',
    },
    'holdType' : {
        'value' : 'indefinite',
        'description' : 'defines the hold string used when placing temperature holds',
    },
    'sethold_hour' : {
        'value' : 17,
        'type' : 'int',
        'description' : 'defines the hour after which to set holds on day the guest depart or the day before arrival',
    },
    ### This section is used to connect the app to a thermostat
    ### to use this:
    ###   1) you set the flag on the command line
    ###   2) display message to go to consumer app and get ready to enable app
    ###   3) user presses enter when they are ready
    ###   4) tool will connect and return a pin code (and wait for user action)
    ###   5) go to consumer portal and put pin in
    ###   6) come back to this app, and press enter to proceed
    ###   7) call request_token and save infomration to config file
    ### done
    'connect' : {
        'type' : 'bool',
        'description' : 'ONLY used to connect app and thermostat - this is the pin code used to enable a thermostat - then the first configuration file is created',
    },
}


### GLOBAL VARIABLES AND CONVERSIONS ###
pp = pprint.PrettyPrinter(indent=4)
datefmt='%m/%d/%Y'
today=datetime.date.today()
today_str=datetime.datetime.strftime(today,datefmt)
today_hour=datetime.datetime.now().hour
tomorrow=datetime.date.today() + datetime.timedelta(days=1)
tomorrow_str=datetime.datetime.strftime(tomorrow,datefmt)

# set the temps we want to compare to based on thermo setting
#
holdSetting = { 'heat' : 55.0, 'cool' : 80.0 }



# read in the file of dates the villa is booked into a dictionary keyed on date
#
def load_villa_calendar( occupy_filename, fldDate, debug=False ):
    # read in the file as a dictionary
    return kvcsv.readcsv2dict(occupy_filename, [fldDate], True)


# read the current thermostat readings, save them to a file (if filename is provided),
# and return the thermometer values of interest in a list of lists each
# thermostat record has the following array of values:
#     'thermo,hvacMode,currentTemp,holdName,holdCool,holdHeat'
#
def readSave_thermoSensor_rtn_therms( ecobee, temperature_filename, debug=False ):
    # local variables - list of therm values
    thermvals = []

    # read in the current settings
    thermos=ecobee.get_thermostats()

    # debugging
    if debug:
        print('readSave_thermoSensor_rtn_therms:thermos:')
        pp.pprint(thermos)

    # validate we got something back
    if not thermos:
        return thermvals
    
    # loop through the thermometers
    for thermo in thermos:
        # get the name
        name = thermo['name']
        # get the hvacMode
        hvacMode = thermo['settings']['hvacMode']
        # thermTemp
        thermTemp = float(thermo['runtime']['actualTemperature'])/10
        # get if there is a hold on
        holdName = holdHeat = holdCool = None
        if thermo['events']:
            event1 = thermo['events'][0]
            holdName = event1['name']
            holdHeat = str(event1['heatHoldTemp'])
            holdCool = str(event1['coolHoldTemp'])
            
        # save the values
        thermvals.append( (name,hvacMode,thermTemp,holdName,holdCool,holdHeat) )
    
        # debugging - list of sensors
        if debug:
            print('readSave_thermoSensor_rtn_therms:remote sensors:')
            pp.pprint(thermo['remoteSensors'])

        # capture values of interest
        sensors = dict()
        for rSensor in thermo['remoteSensors']:
            sensors[rSensor['name']] = dict()
            for capability in rSensor['capability']:
                sensors[rSensor['name']][capability['type']] = capability['value']


        # Save results to file if a filename is provided
        if temperature_filename:
            # check to see if the file exists
            exists = os.path.isfile( temperature_filename )

            # determine the open type
            open_type = 'a' if exists else 'w'
            
            # now open a file and dump the values collected
            with open( temperature_filename, open_type ) as t:
                # check to see if the file exists - and if it did not - create header in the file
                if not exists:
                    # create the header for the file as we created the file
                    t.write('datetime,thermo,hvacMode,desiredCool,desiredHeat,sensor,temp,occupied,holdName,holdCool,holdHeat\n')
                # now output the sensor data
                for sensor in sensors:
                    t.write("%s,%s,%s,%3.1f,%3.1f,%s,%3.1f,%s,%s,%s,%s\n" % (thermo['thermostatTime'],name, hvacMode, thermo['runtime']['desiredCool']/10, thermo['runtime']['desiredHeat']/10, sensor, float(sensors[sensor]['temperature'])/10, sensors[sensor]['occupancy'],holdName,holdCool,holdHeat))

    # return the array of thermos values
    return thermvals

# set the thermostat to auto if it is set to off
def set_therm_to_auto_if_off( ecobee, idx, thermos, debug=False ):
    # get the hvacMode
    hvacMode = thermos[idx]['settings']['hvacMode']
    # if it is not off then return
    if hvacMode != 'off':
        return
    # show the result of the thermostat call
    logger.info('set_therm_to_auto_if_off:thermo-is-off:idx:%d:set-to:%s', idx, 'auto')
    # the mode is off - so we are going to set it to 'auto'
    result=ecobee.set_hvac_mode(idx,'auto')
    # validate and communicate if it does not work
    if result.status_code != 200:
        # show the result of the thermostat call
        logger.info('set_therm_to_auto_if_off:status-not-changed:idx:%d:returned-status:%s', idx, result)

# remove any temperature holds that exist on all thermostats
#
def remove_temp_holds_all( ecobee, debug=False ):
    # read in the current set of thermostats - so we have a list to work
    thermos=ecobee.get_thermostats()

    # loop through the thermometers
    for idx in range(len(thermos)):
        # make sure the thermostat is not off
        set_therm_to_auto_if_off( ecobee, idx, thermos, debug=debug )
        # remove all holds on this thermostat
        result = ecobee.resume_program(idx, resume_all=True)
        # show the result of the thermostat call
        logger.info('remove_temp_holds_all:idx:%d:result:%s', idx,result)

# set temperature holds on all thermostats
#
def set_temp_holds_all( ecobee, therms, holdSetting, hold_type, debug=False ):
    # read in the current set of thermostats - so we have a list to work
    thermos=ecobee.get_thermostats()

    # Get values from the array that was passed in (therms)
    (thermo,hvacMode,currentTemp,holdName,holdCool,holdHeat) = therms[0]
    
    # loop through the thermometers
    for idx in range(len(thermos)):
        # make sure the thermostat is not off
        set_therm_to_auto_if_off( ecobee, idx, thermos, debug=debug )
        # The prescribed way to do this
        #  1) if hvacMode is NOT 'auto', then set the cool and heat to the same value
        #  2) if hvacMode is 'auto', then set the cool and heat differently.
        #
        # set a temperature hold - temperatue defined by therm mod and holdSetting value
        # set_hold_temp(index, cool_temp, heat_temp, hold_type= nextTransition )
        # result = ecobee.set_hold_temp(idx,holdSetting[hvacMode],holdSetting[hvacMode],hold_type=hold_type)
        # decided to fill this in with the cool/heat settings uniquely - not as the same value to deal with auto.
        result = ecobee.set_hold_temp(idx,holdSetting['cool'],holdSetting['heat'],hold_type=hold_type)
        # print out the results of setting the hold
        logger.info('set_temp_holds_all:therm:%s:result:%s', therms[idx][0], result)
        # debugging
        if debug:
            # pull the thermos data again
            thermos_check =readSave_thermoSensor_rtn_therms( ecobee, None, debug=debug )
            # check if the change took place
            #print('set_temp_holds_all:thermos_check:')
            #pp.pprint(thermos_check)
            # split up this attribute
            (thermo1,hvacMode1,currentTemp1,holdName1,holdCool1,holdHeat1) = thermos_check[idx]
            # display it.
            print('set_temp_holds_all:thermos:', thermo1,':holdName:', holdName1, ':holdCool:', holdCool1, ':holdHeat:', holdHeat1)
            

# determine if there is a house wide hold on temperature settings enabled
# all thermostats are set to the same hold temperature
#
def systemHoldOn_chk( therms, holdSetting, debug=False ):
    # debugging
    if debug:
        print('systemHoldOn_chk:therms:', therms)
        print('systemHoldOn_chk:holdSetting:', holdSetting)

    # loop through the thermometers
    for idx in range(len(therms)):
        # get first readings from the thermometer
        (thermo,hvacMode,currentTemp,holdName,holdCool,holdHeat) = therms[idx]

        # debugging
        if debug:
            print('systemHoldOn_chk:hvacMode:', hvacMode, ':holdName:', holdName)
            print('systemHoldOn_chk:holdCool:', holdCool, ':holdHeat:', holdHeat)

        # check the hold setting and temp
        if holdName == None:
            # there is no hold on - so hold is NOT on
            logger.info('systemHoldOn_chk:thermo:%s:holdName:%s:return false', thermo, holdName)
            return False
        elif hvacMode == 'off':
            # thermostat is not enabled so the hold is not set
            logger.info('systemHoldOn_chk:thermo:%s:hvacMode:%s:return false', thermo, hvacMode)
            return False
        elif int(holdCool) != int(10*holdSetting['cool']):
            # we should just compare cool and heat values and if either does not match - the return false.
            # hold enabled, but temperature set is not same as our hold setting - so hold is NOT on
            logger.info('systemHoldOn_chk:thermo:%s:holdCool:%s:holdSetting:%s:return false', thermo, holdCool,int(10*holdSetting['cool']))
            return False
        elif int(holdHeat) != int(10*holdSetting['heat']):
            # we should just compare cool and heat values and if either does not match - the return false.
            # hold enabled, but temperature set is not same as our hold setting - so hold is NOT on
            logger.info('systemHoldOn_chk:thermo:%s:holdHeat:%s:holdSetting:%s:return false', thermo, holdHeat,int(10*holdSetting['heat']))
            return False

    # checked everything - the hold must be on
    return True

    
# routine used to connect the application to a thermostat/account
# used then the command line has "connect=True" on it
#
# goes through the steps of connecting the app to an account
#   1) Check the configuratoin file if passed in and build a stub configuration file if file does not exist
#   2) Prompt user to get the customer portal ready to go
#   3) Create ecobee object using configuration file - this will generate a PIN
#   4) Check to see if we authenticated vs generated a PIN - if auth - there is NO work to do - exit
#   5) Prompt user to take PIN and put it into the consumer portal
#   6) Now call request_tokens to update data and save to configuration file
#   7) Validate request_tokens response and display the proper message
#   
def connect_app_and_account(optiondict):
    # first create logging for what we are donig
    logger.info("Connecting application to account")
    logger.info("Have user get prepared by logging into the consumerportal")

    # STEP1 - check to see if the configuration file exists that the user passed in
    if optiondict['config_filename']:
        config_filename = optiondict['config_filename']
        if not os.path.isfile(config_filename):
            logger.info('Config_filename specified but does not exist, creating it')
            jsonconfig = {"API_KEY": optiondict['api_key']}
            pyecobee.config_from_file(config_filename, jsonconfig)
                
    # STEP2 - display message to the user
    print("Connecting this application to an account")
    print("Please login to the account that we will be connecting to by going to:")
    print("https://www.ecobee.com/consumerportal/index.html")

    # get the input when the user is ready for us to move on
    name = input("\nPress [Enter] when this is done ==> ")

    # logging
    logger.info("Consumer has the portal ready to go")

    # STEP3 - create the ecobee object with no configuration file - we are starting new here
    logger.info( "Building ecobee object" )
    ecobee = pyecobee.Ecobee(api_key=optiondict['api_key'],config_filename=optiondict['config_filename'])
#    ecobee = pyecobee.Ecobee(config_filename=optiondict['config_filename'], config={'API_KEY': optiondict['api_key']})

    
    # STEP4 - should have failed, but if it did not - then no work to do - exit
    if ecobee.authenticated:
        # logging
        logger.info("App already connected to account - no additional action required - terminating program")
        print("App already connected to account - no additional action required - terminating program")
        sys.exit(1)

    # STEP 4 - did not fail - we need to request a pin
    logger.info('App is now requesting a pin be generated')
    print('App requesting a pin')
    ecobee.request_pin()
    logger.info('Delivered pin:'+ecobee.pin)
    if not ecobee.pin:
        logger.info('A pin was not created - failing')
        print('A pin was not generated - there is a problem - EXITTING')
        sys.exit(1)
    
    # logging
    logger.info("Request user to enter PIN into consumer portal My Apps, Add application" )
    
    # STEP5 - we have the PIN code that the user needs to authenticate - have them go do it.
    print("Go to the consumerportal web page and click")
    print('My Apps, Add application, and when prompted enter PIN: [', ecobee.pin, '] and click Authorize.')
    print('Once authorized, click Add application')

    # get the input when the user is ready for us to move on
    name = input("\nPress [Enter] when this is done ==> ")
    
    # logging
    logger.info("User entered PIN [ %s ] - application will now call request_tokens", ecobee.pin)

    # STEP6 - call the request_tokens
    if ecobee.request_tokens():
        print('Request_tokens delivered true and autenticated is:', ecobee.authenticated)

    # STEP7 - validate we got what we wanted
    if ecobee.pin is not None:
        logger.error("Request_tokens call did not work - you are still not connected - please try again")
        print("Request_tokens call did not work - you are still not connected - please try again")
        sys.exit(1)
    else:
        logger.info('Request_tokens now has the thermostat connected and a configuration file generated for future runs')
        print('Request_tokens now has the thermostat connected and a configuration file generated for future runs')

    # STEP8 - make sure this delivers a working solution - get thermostats
    logger.info('Validating we can authenticate')
    ecobee.get_thermostats()
    print('We should be authenticated:', ecobee.authenticated)
    if not ecobee.authenticated:
        logger.info('failed to authenticate')
    else:
        logger.info('All worked as expected - we are authenticated')
    sys.exit()
        
    
# ---------------------------------------------------------------------------
if __name__ == '__main__':

    # capture the command line
    optiondict = kvutil.kv_parse_command_line( optiondictconfig, debug=False )

    # set variables based on what came form command line
    debug = optiondict['debug']

    # print header to show what is going on (convert this to a kvutil function:  kvutil.loggingStart(logger,optiondict))
    kvutil.loggingAppStart( logger, optiondict, kvutil.scriptinfo()['name'] )

    # OPTIONAL ACTION - determine if we are connecting and if so do that - the called routine will exit - no other work to do.
    if optiondict['connect']:
        connect_app_and_account(optiondict)
        
    # create the ecobee object
    logger.info( "Building ecobee object - it may refresh the tokens and update config file:%s",optiondict['config_filename'] )
    ecobee = pyecobee.Ecobee(api_key=optiondict['api_key'],config_filename=optiondict['config_filename'])
#    ecobee = pyecobee.Ecobee(config_filename=optiondict['config_filename'], config={'API_KEY': optiondict['api_key']})

    # read in therms, save data, and get therm values
    logger.info( 'Fetch ecobee thermostat data - save temp readings to file:%s', optiondict['temperature_filename'])
    therms = readSave_thermoSensor_rtn_therms(ecobee, optiondict['temperature_filename'], debug=debug )

    # validate that we were authenticated after we called in
    if not ecobee.authenticated:
        logger.error( "Did not authenticate - program terminated!")
        print( "Did not authenticate - program terminated!")
        sys.exit(1)
        
    # check to see if system holds are on
    systemHoldOn = systemHoldOn_chk( therms, holdSetting, debug=debug )

    # addition debug data
    logger.info('systemHoldOn:%s', systemHoldOn)
    logger.info('therms:%s',therms)
    
    # read in the villa occupancy information
    logger.info('Read in villa occupancy data from file:%s', optiondict['occupy_filename'])
    villacal = load_villa_calendar( optiondict['occupy_filename'], optiondict['fldDate'], debug=debug )

    # check to see if we are in an occupied day
    if today_str in villacal and tomorrow_str in villacal:
        # today and tomorrow are villa days no action
        logger.info('Villa occupied:no action to take')
    elif today_str in villacal:
        # today but not tomorrow - check the time - late enough - set the hold
        logger.info('Villa vacated today:current_hour:%d:update_after:%d:temp holds set:%s', today_hour, optiondict['sethold_hour'], systemHoldOn)
        if not systemHoldOn and today_hour > optiondict['sethold_hour']:
            logger.info('Villa vacated:set temperature holds')
            set_temp_holds_all( ecobee, therms, holdSetting, optiondict['holdType'], debug=debug )
    elif tomorrow_str in villacal:
        # if tomorrow is a villa day and today is not, remove holds if the occupant is owner or renter
        logger.info('Villa occupied tomorrow:current_hour:%d:update_after:%d:temp holds set:%s', today_hour, optiondict['sethold_hour'], systemHoldOn)
        logger.info('Villa occupied tomorrow')
        if systemHoldOn and today_hour > optiondict['sethold_hour']:
            # renter or owner - remove temp holds
            logger.info('Villa occupied tomorrow:remove holds')
            remove_temp_holds_all( ecobee, debug=debug )
    elif not systemHoldOn:
        # today is not occupied and no holds on - so set a hold
        logger.info('Villa not occupied:set temp holds')
        set_temp_holds_all( ecobee, therms, holdSetting, optiondict['holdType'], debug=debug )
    else:
        logger.info('No action taken on this run')

# eof

