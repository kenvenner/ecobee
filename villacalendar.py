'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.06

set of functions used to parse BP xls and update the appropriate google calendar
'''

from __future__ import print_function
import datetime
import pytz
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pytz
import time

# pretty printing
import pprint
pp = pprint.PrettyPrinter(indent=4)

# setup the logger
import kvutil
import logging
logger = logging.getLogger(__name__)

# set the module version number
AppVersion = '1.06'


# If modifying these scopes, delete the file token.pickle.
# SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
# changed access to read/write so we could create events as defined by the 2nd function
SCOPES = ['https://www.googleapis.com/auth/calendar']

# connect to google services, based on data stored in the credential.json or
# what has been created in the token.pickle file
def get_cal_service():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        logger.info('Load credentials from pickle file')
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            logger.info('Refreshing credentials with google')
            creds.refresh(Request())
        else:
            logger.info('Using credentials.json to create a set of secrets')
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        logger.info('Saving current credentials to pickle file')
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # connect with the calendar services
    logger.info('Build calendar service')
    service = build('calendar', 'v3', credentials=creds)

    return service

# read in all future events for this user
def read_future_calendar_events(service, now, debug=False):
    # capture the current time - we want events AFTER now
    utcnow = now.astimezone(pytz.UTC).isoformat()#  + 'Z' # 'Z' indicates UTC time

    if debug:
        print('now:', now)
        print('utcnow:', utcnow)
        print('gen utcnow:', datetime.datetime.utcnow().isoformat() + 'Z')

    if 0:
        # this is a single call to the routine
        if debug:
            print('read_future_calendar_events:Getting the upcoming 10 events:utcnow:', utcnow)
        events_result = service.events().list(calendarId='primary', timeMin=utcnow,
                                              maxResults=10, singleEvents=True,
                                              orderBy='startTime').execute()
        ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
        logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
        time.sleep(2)

        events = events_result.get('items', [])
    else:
        # this loop get all future events and pulls them in
        if debug:
            print('read_future_calendar_events:Getting all upcoming events:utcnow:', utcnow)
        events=list()
        loopcnt=0
        cal_request = service.events().list(calendarId='primary', timeMin=utcnow,
                                              maxResults=10, singleEvents=True,
                                              orderBy='startTime')
        while cal_request is not None:
            # make the call
            events_result=cal_request.execute()
            # add the events to the list
            events.extend(events_result.get('items', []))
            # debugging
            if debug:
                print('read_future_calendar_events:event after run:', loopcnt, ':event_count:', len(events))
                pp.pprint(events)
                print('------------------------------------------------')
            # increment the loop counter
            loopcnt += 1
            # create a new call to get the next list
            cal_request = service.events().list_next(cal_request, events_result)
            
            ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
            logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
            time.sleep(2)


    # debugging
    if debug:
        if not events:
            print('read_future_calendar_events:No upcoming events found.')
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print('read_future_calendar_events:', start, event['summary'])
            print('read_future_calendar_events:event:')
            pp.pprint(event)
            print('------------------------------------------------')
        print('=================================================================')

    # log this
    logger.info('Number of future calendar events read in:%s', len(events))
    
    # return the list of calendar entries we got back
    return events


# remove an event from the calendar base on the id
def delete_cal_event(service, id, debug=False):
    event = service.events().delete(calendarId='primary', eventId=id).execute()

    ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
    logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
    time.sleep(2)

    
    logger.info('Delete calendar event:%s', id)
    if debug:
        print('delete_cal_event:delete event:', id)
        pp.pprint(event)
        print('------------------------------------------------')



### NOT USED #####
# we are getting rid of this one at some time - this was just a starter
def create_cal_event(service):
    event = {
        'summary': 'Create test appt on calendar',
        'location': 'Kens house',
        'description': 'Testing the google API for creating an event',
        'start': {
            'dateTime': '2019-02-22T17:00:00-07:00',
            'timeZone': 'America/Los_Angeles',
        },
        'end': {
            'dateTime': '2019-02-22T18:00:00-07:00',
            'timeZone': 'America/Los_Angeles',
        },
#        'recurrence': [
#            'RRULE:FREQ=DAILY;COUNT=2'
#        ],
        'attendees': [
            {'email': 'ken@vennerllc.com'},
            {'email': 'ken_venner@yahoo.com'},
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60},
                {'method': 'popup', 'minutes': 10},
            ],
        },
    }
    
    event = service.events().insert(calendarId='primary', body=event).execute()

    ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
    logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
    time.sleep(2)

    # debugging
    print('create_cal_event:insert event:')
    pp.pprint(event)
    print('------------------------------------------------')

    # return the id
    return event.get('id')

# create a date/time string that can be compared with the data returned from google calendar
def google_time_convert( dt ):
    pt=pytz.timezone('America/Los_Angeles')
    hoursoff=pt.utcoffset(dt).total_seconds()/3600
    return dt.isoformat()+'{:+03.0f}:00'.format(hoursoff)

# generic tool used to see if an event exists, and if so, marks it as so
def event_already_exists( starttime, event, calexist_dict, debug=False ):
    # create the google start time
    gstarttime =  google_time_convert(starttime)
    
    # debugging print
    if debug:
        print('event_already_exists:create_cal_event_start:gstarttime:', gstarttime)
        pp.pprint(event)
        print('------------------------------------------------')

    # check to see if this event already exists
    if gstarttime in calexist_dict:
        if debug:
            if debug:
                print('event_already_exists:Cal entry time/date match:', event['summary'])
        if calexist_dict[gstarttime]['summary'] == event['summary']:
            calexist_dict[gstarttime]['eventmatch'] = True
            logger.info('Calendar entry exists:marked and skipping:%s', event['summary'])
            # debugging
            if debug:
                print('event_already_exists:Cal entry exists - marked and skipping:', event['summary'])
            return True

    # did not get a match
    return False

    
# generic tool to take the dict that is a stay and convert it into the event information to be saved
def create_event_dict( eType, starttime, endtime, stay, debug=False ):
    event = dict()
    event['summary'] = ':'.join((eType,stay['Booking'],stay['Type'],str(stay['First Night']),'Nights',str(stay['Nights'])))
    event['description'] = 'Booking...:' + stay['Booking'] + '\nEvent Type:' + stay['Type'] + '\nStart Date:' + str(stay['First Night']) + '\nNights....:' + str(stay['Nights']) + '\nCreated...:' + str(stay['Created'])
    event['start'] = {
        'dateTime' : starttime.isoformat(),
        'timeZone' : 'America/Los_Angeles',
    }
    event['end'] = {
        'dateTime' : endtime.isoformat(),
        'timeZone' : 'America/Los_Angeles',
    }

    # debbugingg
    if debug:
        print('create_event_dict:event:')
        pp.pprint(event)
        print('------------------------------------------------')
        
    return event
    

# defines for the appt what type of label to put on it
def who_by_type( eventType ):
    who = 'Guest'
    for newwho in ('Clean', 'Owner', 'Maint'):
        if newwho in eventType:
            who=newwho
            break
    return who


# stay starts - set the times and create the event
def create_cal_event_start( service, calexist_dict, caldate, stay, debug=False ):
    # starts at noon
    starttime = datetime.datetime.combine(caldate.date(), datetime.time(hour=12))
    # goes to midnight (11:59pm)
    endtime = datetime.datetime.combine(caldate.date(), datetime.time(hour=23, minute=59))

    # build the event dictionary
    eventbody = create_event_dict( who_by_type(stay['Type']) + '-start', starttime, endtime, stay, debug=False )

    # checkt to see if the event exists - if so - we are done
    if event_already_exists( starttime, eventbody, calexist_dict, debug=False ):
        return
    
    # create the event - it does not exist
    event = service.events().insert(calendarId='primary', body=eventbody).execute()

    ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
    logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
    time.sleep(2)

    # debugging
    if debug:
        print('create_cal_event_start:insert event:')
        pp.pprint(event)
        print('------------------------------------------------')

    # logging
    logger.info('Calendar event created:start:%s:%s', event.get('id'),eventbody['summary'])

    # return the id that was created
    return event.get('id')

# stay is on going - full day
def create_cal_event_stay( service, calexist_dict, caldate, stay, debug=False ):
    # all day event - as this is a stay date
    # starts at midnight
    starttime = datetime.datetime.combine(caldate.date(), datetime.time(hour=0))
    # goes to just before noon
    endtime = datetime.datetime.combine(caldate.date(), datetime.time(hour=23, minute=59))

    # build the event dictionary
    eventbody = create_event_dict( who_by_type(stay['Type']) + '-stay', starttime, endtime, stay, debug=False )

    # checkt to see if the event exists - if so - we are done
    if event_already_exists( starttime, eventbody, calexist_dict, debug=False ):
        return
    
    # create the event - it does not exist
    event = service.events().insert(calendarId='primary', body=eventbody).execute()

    ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
    logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
    time.sleep(2)

    # debugging
    if debug:
        print('create_cal_event_stay:insert event:\n')
        pp.pprint(event)
        print('------------------------------------------------')

    # logging
    logger.info('Calendar event created:stay:%s:%s', event.get('id'),eventbody['summary'])

    # return the id that was created
    return event.get('id')

# stay is completing - exit
def create_cal_event_exit( service, calexist_dict, caldate, stay, debug=False ):
    # starts at midnight
    starttime = datetime.datetime.combine(caldate.date(), datetime.time(hour=0))
    # goes to just before noon
    endtime = datetime.datetime.combine(caldate.date(), datetime.time(hour=12, minute=0))

    # build the event dictionary
    eventbody = create_event_dict( who_by_type(stay['Type']) + '-exit', starttime, endtime, stay, debug=False )

    # checkt to see if the event exists - if so - we are done
    if event_already_exists( starttime, eventbody, calexist_dict, debug=False ):
        return
    
    # create the event - it does not exist
    event = service.events().insert(calendarId='primary', body=eventbody).execute()

    ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
    logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
    time.sleep(2)

    # debugging
    if debug:
        print('create_cal_event_exit:insert event:\n')
        pp.pprint(event)
        print('------------------------------------------------')

    # logging
    logger.info('Calendar event created:exit:%s:%s', event.get('id'), eventbody['summary'])

    # return the id that was created
    return event.get('id')


# not used - it was the other option of just one event per stay
def create_cal_event_start_exit( service, startdate, enddate, stay, debug=False ):
    # starts at noon
    starttime = datetime.datetime.combine(startdate.date(), datetime.time(hour=12))
    # exit at noon on the final day
    endtime = datetime.datetime.combine(enddate.date(), datetime.time(hour=12))

    # build the event dictionary
    eventbody = create_event_dict( who_by_type(stay['Type']) + '-stay', starttime, endtime, stay, debug=False )

    # checkt to see if the event exists - if so - we are done
    if event_already_exists( starttime, eventbody, calexist_dict, debug=False ):
        return
    
    # create the event - it does not exist
    event = service.events().insert(calendarId='primary', body=eventbody).execute()

    ### 2020-05-23 we hit rate limiting so we are going to sleep 2 seconds between each event
    logger.info('called service.events - sleep for 2 seconds to avoid rate limiting')
    time.sleep(2)
    
    # debugging
    if debug:
        print('create_cal_event_start_exit:insert event:\n')
        pp.pprint(event)
        print('------------------------------------------------')

    # logging
    logger.info('Calendar event created:start_exit:%s:%s', event.get('id'),eventbody['summary'])

    # return the id that was created
    return event.get('id')


# utility to convert a list of calendar events into a dictionary keyed by start datetime
def cal_events_dict_on_startdatetime(cal_events, service, delDupes=False, debug=False ):
    # debugging            
    if debug:
        # what the user sent in
        print('----------------------------------------')
        print('cal_events_dict_on_startdatetime:cal_events:')
        pp.pprint(cal_events)
        # find out type type this is
        print('----------------------------------------')
        print('cal_events_dict_on_startdatetime:type of cal_events:', type(cal_events))

    # utcnow build a dictionary based on the start datetime of the calendar event
    cal_events_start_dict = dict()
    for cal_event in cal_events:
        if cal_event['start']['dateTime'] in cal_events_start_dict:
            # if we have a duplicate, log and replace
            logger.info('Duplicate events at time:%s',  cal_event['start']['dateTime'])
            # debugging message
            if debug:
                print('cal_events_dict_on_startdatetime:duplicate calendar entry on start datetime:', cal_event['start']['dateTime'] )
            # and if we flagged it - we delete the current record and keep the prior
            if delDupes:
                logger.info('deleting the first duplicate record now id:%s:%s', cal_events_start_dict[cal_event['start']['dateTime']]['id'], cal_events_start_dict[cal_event['start']['dateTime']]['summary'])
                # debugging
                if debug:
                    print('cal_events_dict_on_startdatetime:deleting the first duplicate record now id:', cal_events_start_dict[cal_event['start']['dateTime']]['id'])
                # delete this event
                delete_cal_event(service, cal_events_start_dict[cal_event['start']['dateTime']]['id'])

        # set the entry to the current event at this datetime
        cal_events_start_dict[ cal_event['start']['dateTime'] ] = cal_event

    # debugging
    if debug:
        print('cal_events_dict_on_startdatetime:cal_events_start_dict:')
        pp.pprint(cal_events_start_dict)
        print('------------------------------------------------')

    # return what we created
    return cal_events_start_dict

# Create Events if they don't exists for stays at the villa
def create_cal_events_for_villa_stays( service, xlsaref, cal_events_start_dict, ignorebefore, debug=False ):
    
    # constant
    addoneday =  datetime.timedelta(days=1)

    # debugging
    if debug:
        print('ignorebefore:', ignorebefore)
        
    # loop through the BP records
    for stay in xlsaref:
        # check to see if this stay is in the past
        if stay['startdate'] < ignorebefore:
            logger.info('Skipping this stay it is in the past:%s', stay['startdate'])
            # debugging
            if debug:
                print('create_cal_events_for_villa_stays:Stay already started get next stay:', stay['startdate'])
            continue

        # debugging
        if debug:
            print('create_cal_events_for_villa_stays:stay:', stay)

        # set the variable to start
        caldate = stay['startdate']
        
        # create the start entry
        id = create_cal_event_start( service, cal_events_start_dict, caldate, stay, debug=debug )

        # now get the stay days but stop when we get to the exit date
        caldate += addoneday

        # debugging
        if debug:
            print('create_cal_events_for_villa_stays:before loop')
            print('create_cal_events_for_villa_stays:caldate:', caldate)
            print('create_cal_events_for_villa_stays:exitdate:', stay['exitdate'])

        # loop through dates until we hit the exit date
        while caldate < stay['exitdate']:
            # debugging
            if debug:
                print('create_cal_events_for_villa_stays:in loop\ncreate_cal_events_for_villa_stays:caldate:', caldate)
            # create the stay event
            id = create_cal_event_stay( service, cal_events_start_dict, caldate, stay, debug=debug )
            # increment the date
            caldate += addoneday
        
        # we have exited because we are on the exit date
        id = create_cal_event_exit( service, cal_events_start_dict, caldate, stay, debug=debug )
        

#loop through the events we pulled and if we find any that don't have the eventmatch set - then we need to move that entry
def remove_nonmatch_events( service, cal_events_start_dict, debug=False ):
    for caldatetime in cal_events_start_dict:
        if 'eventmatch' not in cal_events_start_dict[caldatetime]:
            logger.info('No match on event-removing event:%s:%s', cal_events_start_dict[caldatetime]['id'], cal_events_start_dict[caldatetime]['summary'])
            # this did not have a match - show the ID (in the future just delete this entry
            if debug:
                print('remove_nonmatch_events:no match on event:', caldatetime, ':id:', cal_events_start_dict[caldatetime]['id'])
            # now just delete it
            delete_cal_event(service, cal_events_start_dict[caldatetime]['id'], debug=debug)            


# core routine - takes in the list of records from the BP xls and updates the google calendar to match
def sync_villa_cal_with_bp_xls( xlsaref, now=None, debug=False ):
    # logger
    logger.info('Synching XLS with calendar events:XLS event count:%s', len(xlsaref))
    
    # connect to the account and get the service up and running
    service =  get_cal_service()

    # capture the current time
    if not now:
        now = datetime.datetime.now()
    
    # read in the existing calendar informaton
    cal_events = read_future_calendar_events(service, now, debug=debug)

    # convert to a dictionary for comparison base on start datetime
    cal_events_start_dict = cal_events_dict_on_startdatetime( cal_events, service, delDupes=True, debug=debug )

    # debugging
    if debug:
        print('-----------------------------------------------------------')
        print('sync_villa_cal_with_bp_xls:cal_events_start_dict:')
        pp.pprint(cal_events_start_dict)
        print('-----------------------------------------------------------')
    
    # now create the various calendar entries
    create_cal_events_for_villa_stays( service, xlsaref, cal_events_start_dict, now, debug=debug )

    # then remove all calendar events that did not have a match
    remove_nonmatch_events( service, cal_events_start_dict, debug=debug )


# copied in from the output of vcconvert.py program
def seed_xlsaref():
    xlsaref = [
    {   'Booking': 'HLD-17324',
        'Created': '1/15/2019',
        'First Night': '2/22/2019',
        'Last Night': '2/23/2019',
        'Nights': '2',
        'Type': 'Hold - Owner',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 2, 24, 0, 0),
        'startdate': datetime.datetime(2019, 2, 22, 0, 0)},
    {   'Booking': 'BKG-09559',
        'Created': '2/2/2019',
        'First Night': '3/1/2019',
        'Last Night': '3/2/2019',
        'Nights': '2',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 3, 3, 0, 0),
        'startdate': datetime.datetime(2019, 3, 1, 0, 0)},
    {   'Booking': 'BKG-09084',
        'Created': '6/28/2018',
        'First Night': '3/6/2019',
        'Last Night': '3/9/2019',
        'Nights': '4',
        'Type': 'Res. - Owner',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 3, 10, 0, 0),
        'startdate': datetime.datetime(2019, 3, 6, 0, 0)},
    {   'Booking': 'BKG-09222',
        'Created': '8/28/2018',
        'First Night': '4/1/2019',
        'Last Night': '4/5/2019',
        'Nights': '5',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 4, 6, 0, 0),
        'startdate': datetime.datetime(2019, 4, 1, 0, 0)},
    {   'Booking': 'HLD-17444',
        'Created': '2/2/2019',
        'First Night': '4/11/2019',
        'Last Night': '4/14/2019',
        'Nights': '4',
        'Type': 'Hold - Owner',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 4, 15, 0, 0),
        'startdate': datetime.datetime(2019, 4, 11, 0, 0)},
    {   'Booking': 'BKG-09587',
        'Created': '2/14/2019',
        'First Night': '4/18/2019',
        'Last Night': '4/20/2019',
        'Nights': '3',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 4, 21, 0, 0),
        'startdate': datetime.datetime(2019, 4, 18, 0, 0)},
    {   'Booking': 'HLD-17310',
        'Created': '1/12/2019',
        'First Night': '4/26/2019',
        'Last Night': '4/28/2019',
        'Nights': '3',
        'Type': 'Hold - Owner',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 4, 29, 0, 0),
        'startdate': datetime.datetime(2019, 4, 26, 0, 0)},
    {   'Booking': 'BKG-09250',
        'Created': '9/20/2018',
        'First Night': '5/4/2019',
        'Last Night': '5/10/2019',
        'Nights': '7',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 5, 11, 0, 0),
        'startdate': datetime.datetime(2019, 5, 4, 0, 0)},
    {   'Booking': 'BKG-09248',
        'Created': '9/17/2018',
        'First Night': '5/13/2019',
        'Last Night': '5/16/2019',
        'Nights': '4',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 5, 17, 0, 0),
        'startdate': datetime.datetime(2019, 5, 13, 0, 0)},
    {   'Booking': 'BKG-09028',
        'Created': '6/14/2018',
        'First Night': '5/17/2019',
        'Last Night': '5/19/2019',
        'Nights': '3',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 5, 20, 0, 0),
        'startdate': datetime.datetime(2019, 5, 17, 0, 0)},
    {   'Booking': 'BKG-09268',
        'Created': '10/2/2018',
        'First Night': '5/25/2019',
        'Last Night': '5/31/2019',
        'Nights': '7',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 6, 1, 0, 0),
        'startdate': datetime.datetime(2019, 5, 25, 0, 0)},
    {   'Booking': 'BKG-09491',
        'Created': '1/7/2019',
        'First Night': '6/9/2019',
        'Last Night': '6/13/2019',
        'Nights': '5',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 6, 14, 0, 0),
        'startdate': datetime.datetime(2019, 6, 9, 0, 0)},
    {   'Booking': 'BKG-09297',
        'Created': '10/19/2018',
        'First Night': '6/21/2019',
        'Last Night': '6/25/2019',
        'Nights': '5',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 6, 26, 0, 0),
        'startdate': datetime.datetime(2019, 6, 21, 0, 0)},
    {   'Booking': 'BKG-09572',
        'Created': '2/11/2019',
        'First Night': '6/29/2019',
        'Last Night': '7/5/2019',
        'Nights': '7',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 7, 6, 0, 0),
        'startdate': datetime.datetime(2019, 6, 29, 0, 0)},
    {   'Booking': 'BKG-09182',
        'Created': '8/10/2018',
        'First Night': '8/3/2019',
        'Last Night': '8/9/2019',
        'Nights': '7',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 8, 10, 0, 0),
        'startdate': datetime.datetime(2019, 8, 3, 0, 0)},
    {   'Booking': 'BKG-09368',
        'Created': '11/29/2018',
        'First Night': '10/10/2019',
        'Last Night': '10/14/2019',
        'Nights': '5',
        'Type': 'Res. - Renter',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 10, 15, 0, 0),
        'startdate': datetime.datetime(2019, 10, 10, 0, 0)},
    {   'Booking': 'HLD-17491',
        'Created': '2/11/2019',
        'First Night': '10/18/2019',
        'Last Night': '10/19/2019',
        'Nights': '2',
        'Type': 'Hold - Owner',
        'Unit Code': 'CASVC',
        'blank001': '',
        'blank002': '',
        'blank003': '',
        'blank004': '',
        'exitdate': datetime.datetime(2019, 10, 20, 0, 0),
        'startdate': datetime.datetime(2019, 10, 18, 0, 0)
    }
    ]

    return xlsaref

########################################

if __name__ == '__main__':

    logging.basicConfig(filename=os.path.splitext(kvutil.scriptinfo()['name'])[0]+'.log',
                    level=logging.INFO,
                    format='%(asctime)s - %(name)s - %(threadName)s -  %(levelname)s - %(message)s')
    logger = logging.getLogger(__name__)

    # read in the xlsaref
    xlsaref = seed_xlsaref()
    print('main:xlsaref:')
    pp.pprint(xlsaref)
    print('------------------------------------------------')

    # and then sync it all up
    sync_villa_cal_with_bp_xls( xlsaref, debug=False )

