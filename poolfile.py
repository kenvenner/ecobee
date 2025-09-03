'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.13

Utility used to readn and write files from 

'''
import os.path
import os
import logging
import datetime
import kvdate

### GLOBAL VARIABLES AND CONVERSIONS ###


def read_pool_heater_allowable_file(input_file, logger):
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


# eof

