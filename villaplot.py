'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.02

Read in the time series data created by villaecobee.py
and generate temperature plots from these time series

'''

import kvcsv
import kvutil
import datetime
import csv
import sys


# application variables
optiondictconfig = {
    'AppVersion' : {
        'value' : '1.02',
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
    'temperature_filename' : {
        'value' : 'villatemps.txt',
        'description' : 'defines the name of the file that holds the temperature readings',
    },
    'datefmt' : {
        'value' : '%Y-%m-%d %H:%M:%S',
        'description' : 'defines the date-time format',
    },
    'timedelta_minutes' : {
        'type' : 'int',
        'value' : 15,
        'description' : 'defines the time rounding delta',
    },
    'plot_date_start' : {
        'type' : 'date',
        'description' : 'defines the starting date for plotting (default: earliest date)',
    },
    'plot_date_end' : {
        'type' : 'date',
        'description' : 'defines the ending date for plotting (default: latest date)',
    },
    'ylimit_low' : {
        'type' : 'int',
        'value' : 55,
        'description' : 'defines the y axis lower limit for plotting',
    },
    'ylimit_high' : {
        'type' : 'int',
        'value' : 80,
        'description' : 'defines the y axis lower limit for plotting',
    },
}


# utility function to round time to a defined time window
def roundTime(dt=None, dateDelta=datetime.timedelta(minutes=1)):
    """Round a datetime object to a multiple of a timedelta
    dt : datetime.datetime object, default now.
    dateDelta : timedelta object, we round to a multiple of this, default 1 minute.
    Author: Thierry Husson 2012 - Use it as you want but don't blame me.
            Stijn Nevens 2014 - Changed to use only datetime objects as variables
    """
    roundTo = dateDelta.total_seconds()

    if dt == None : dt = datetime.datetime.now()
    seconds = (dt - dt.min).seconds
    # // is a floor division, not a comment on following line:
    rounding = (seconds+roundTo/2) // roundTo * roundTo
    return dt + datetime.timedelta(0,rounding-seconds,-dt.microsecond)



def read_plot_data(temperature_filename, datefmt, timedelta_minutes):
    # read in the data from the txt file
    villadata = kvcsv.readcsv2list(temperature_filename)

    # create a plot dictionary keyed by "rounded time", with a dictionary of sensor:temp key value pairs
    plotdata=dict()

    # not used at this time
    sensors=[]

    # step through each record read from the TXT file
    for rec in villadata:
        # create date/time conversions
        rec['dt_datetime_raw'] = datetime.datetime.strptime(rec['datetime'], datefmt)
        rec['dt_datetime'] = roundTime(rec['dt_datetime_raw'],datetime.timedelta(minutes=timedelta_minutes))

        # stuff this value into the plotdata (either create entry or update it)
        if rec['dt_datetime'] not in plotdata:
            plotdata[rec['dt_datetime']] = {rec['sensor'] : rec['temp']}
        else:
            plotdata[rec['dt_datetime']][rec['sensor']]= rec['temp']
        # keep a list of sensors we have seen
        if rec['sensor'] not in sensors:
            sensors.append(rec['sensor'])

    # return the plot data we just read
    return plotdata
        

# convert plot data for plotting
def create_plotting_data(plotdata, start_date=None, end_date=None):
    xaxis=[]
    y1Main=[]
    y2Bed=[]
    gooddata = True

    # for each plottable entry
    for ptime in sorted(plotdata):
        # grab the time of this plottable event
        plottime=plotdata[ptime]
        # skip dates that are not of interest
        if start_date and ptime < start_date:
            continue
        if end_date and ptime > end_date:
            break
        # validate we have both plottble data points
        if 'Villa Main' not in plottime:
            print('Missing Main:', ptime, plottime)
            gooddata = False
        elif 'Villa Bedrooms' not in plottime:
            print('Missing Bedrooms:', ptime, plottime)
            gooddata = False
        else:
            # if we have both - then add data to the arrays used for plotting
            xaxis.append(ptime)
            y1Main.append( float(plotdata[ptime]['Villa Main']) )
            y2Bed.append( float(plotdata[ptime]['Villa Bedrooms']) )

    # pass back what we determined
    return xaxis,y1Main,y2Bed,gooddata

# plot this data
def plot_temp_data(xaxis,y1Main,y2Bed,ylimit_low=55,ylimit_high=80):
    import matplotlib.pyplot as plt
    plt.plot(xaxis,y1Main,'r--',label='Main')
    plt.plot(xaxis,y2Bed,'b',label='Bed')
    plt.legend(loc='upper left')
    plt.gcf().autofmt_xdate()
    plt.ylabel('Temperature')
    plt.ylim(ylimit_low,ylimit_high)
    plt.show()


# ---------------------------------------------------------------------------
if __name__ == '__main__':
    # capture the command line
    optiondict = kvutil.kv_parse_command_line( optiondictconfig, debug=False )

    # set variables based on what came form command line
    debug = optiondict['debug']
    

    # get the plot data
    plotdata = read_plot_data(optiondict['temperature_filename'], optiondict['datefmt'], optiondict['timedelta_minutes'])

    # convert plot data into plottable data
    xaxis,y1Main,y2Bed,gooddata = create_plotting_data(plotdata, start_date=optiondict['plot_date_start'], end_date=optiondict['plot_date_end'])

    # plot the results
    plot_temp_data(xaxis,y1Main,y2Bed,optiondict['ylimit_low'],optiondict['ylimit_high'])
    
#eof
