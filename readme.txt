this suite of tools is used to automate the management and tracking
of the thermostats at the Villa.

For each run, we capture the temp and humidty settings for each sensor
and save them to a temperature tracking file.  This allows us to track/trend
the temperature in the villa.

When unoccupied, the thermostats will be placed in a tempeature hold that
minimizes the energy required to heat/cool the house (55 in cold, 80 in hot).
In a timely basis, just before guests arrive, the temperature holds will
be removed, allowing the house to return to normal occupied temperature ranges.
Just after occupants leave the temperature holds will be "reimplemented" in 
order operate efficiently.

Each run - we will validate holds exist, if they should, and if they are not
set as required, they will re-implemented.  (This happens when unplanned
visitors come ot the house and change the thermostat settings.

this utility is called on a scheduled basis through a scheduled task (windows) 
or cron job (linux/mac).  Currently scheduled to run every 

this tool suite consists of two core tools:
- vcconvert.py - Beautiful places file conversion tool, creates the file 'stays.txt'
- villaecobee.py - Regularly scheduled tool to manage thermostat settings

VCCONVERT.PY

this tool reads in the XLS provided by the BP staff, and creates the 'stays.txt'
which captures the days the property is occupied and what type of occupant we
are planning on (renter, owner, housekeeping, maintenance)

this tool takes each line in the XLS file, extracts the start date, and
the number of nights stay, and the type of occupancy.
It then creates an output line for each day that the villa will be
occupied, starting with teh start date, and adding an additional record
for each additoinal night stay, incrementing the date by one.
For each record created, it saves the date of the stay and the stay type.

this file is an input file to the other tool.

Input:   'VillaCarnerosCalendar.xls' (emailed to us from BP)
Output:  'stays.txt'

Process:
1) receive new file from BP
2) save to the directory where tools are executed from, saving it
   as the filename listed above in "input"
3) execute the tool:  python vcconvert.py


VILLAECOBEE.PY

this is the work horse tool.  this tool performs the following tasks:
1) connet with ecobee
2) refresh authorizatoin tokens and save to local config file
3) read data from all thermostats and sensors
4) save temp/humidity readings from all thermostats/sensors to temphistory file
5) determine if a systemically set temperature hold exists for the villa
6) read in the villa occupancy information (from stays.txt)
7) determine if an action is required on the thermostats
   a) Occupied - No action - today and tomorrow are occupied days
   b) Vacating - set temp holds at right time - today occupied - tomorrow not
   c) Arriving - release temp holds at right time - today NOT occupied - tomorro occupied
   d) Vacant - today and tomorrow not occupied - assure temp hold is set

When systemically changing the holds, we have set a time of day in
which we must be greater than in order to cause this action to
take place.  For setting the hold, we want to provide sufficient
time for the guests to exit.  For removing the holds, we want
to do it late enough the day before, to provide suffient time
for the house to reach desired temperatures while minmizing
the total time we are operating at this temp range without
the hose being occupied.  We curently set this time to 5pm
or later for actions to be taken.


Input files:
    ecobee.conf - authorizatoin token file - maintained by the tool
    occupy_filename - file that defines the dates the villa is occupied (stays.txt)
Output files:
    ecobee.conf - updated authorizatoin data
    temperature_filename - file that houses the temp/humidity readings over time


Command line variables:
    debug = Bool - defines when we are running in debugging mode
    api_key = Str - defines the ecobee API Key for this app - set by developer
    config_filename - Str - name of the authorizatoin config file (ecobee.conf)
    occupy_filename - Str - name of the file housing the dates villa is occupied
                            file is records in "Date,Type" format.
			    where format is:  O,R,M,H
    fldDate - Str - field name of the date field in the occupy_filename file
    holdType - Str - string passed to ecobee to define the type of temperature
                     hold we create when we create a temp hold ('indefinate')
    sethold_hour - int - hour in 24 format, that it must be greater than in order
                         to set/remove a system temperature hold (17 - 5pm)


Hardcoded values in the program:
    holdSetting - define the temperature settings to be set when in
                  heat mode (55) and in AC mode (80)
    datefmt - string defining the expected date formats

Processing:
1) go to the directory where tools and files are installed
2) run:  python villaecobee.py


To Reset the PIN/Access Code (Manually)

If the tokens fall out of sync, you need to manually reconnect the app to the
thermostats

Go to this page:
https://www.ecobee.com/home/developer/api/examples/ex1.shtml

In the API Key field enter (taken from ecobee.conf) and press submit:
AbdHocetV5yJ7iirDoun7EOLSR2341eq

Copy to clipboard the ecobeePin value

Login to ecobee as the user (ken@attunewines.com) for the thermos
Go to the My Apps section
Add an application
Paste in the pin from above
Click validate

Go back to the original web page - step 2 section - press [submit]

Obtain the:
access_token
refresh_token

And from the earlier dialogue the authorization code

and put these into the ecobee.conf file



#eof
