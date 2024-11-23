from __future__ import print_function

'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.75

Library of tools used in general by KV
'''

import glob
import os
import datetime
import pprint

# moved datetime processing to its own module
import kvdate

# these were pulled out and put in kvdate.py
# from dateutil import tz
# from dateutil.zoneinfo import get_zonefile_instance

import sys
import errno
import json
from distutils.util import strtobool

# setup the logger
import logging

logger = logging.getLogger(__name__)

# set the module version number
AppVersion = '1.75'
__version__ = '1.75'
HELP_KEYS = ('help', 'helpall',)
HELP_VALUE_TABLE = ('tbl', 'table', 'helptbl', 'fmt',)


# import ast
#   and call bool(ast.literal_eval(value)) 

# ken's command line processor (UT)
#   expects options defined as key=value pair strings on the command line
# input:
#   optiondictconfig - key = variable, value = dict with keys ( value, type, descr, required )
#   raise_error - bool flag - if true and we get a command line setting we don't know raise exception
#   keymapdict - dictionary of misspelled command line values that are mapped to the official values
#
# return:
#   optiondict - dictionary of values from config and command line
#
# example:
# optiondictconfig = {
#     'AppVersion' : {
#         'value' : '1.01',
#     },
#     
#     'debug' : {
#         'value' : False,
#         'type' : 'bool',
#     },
#     'workingdir' : {
#         'required' : True,
#     }
# }
#
# keymapdict = {
#     'working_dir' : 'workingdir',
#     'dbg' : 'debug',
# }
#
# optiondict = kv_parse_command_line( optiondictconfig, keymapdict=keymapdict )
#
# -- Special behavior
#  help=<value>
#
#  will cause the system to generate a file help file when this is passed in on the command line
#
#  if <value> in list ('tbl','table','helptbl','fmt'), then the output is mark down table
#
def kv_parse_command_line(optiondictconfig, raise_error=False, keymapdict=None, debug=False):
    # debug
    if debug: print('kv_parse_command_line:sys.argv:', sys.argv)
    if debug: print('kv_parse_command_line:optiondictconfig:', optiondictconfig)
    # debugging
    logger.debug('LOAD(v%s)%s', AppVersion, '-' * 40)
    logger.debug('sys.argv: %s', sys.argv)
    logger.debug('optiondictconfig: %s', optiondictconfig)

    # default a set of basic config values - so we don't need to put them in each app
    defaultdictconfig = {
        'debug': {
            'value': False,
            'type': 'bool',
            'description': 'defines if we are running in debug mode',
        },
        'verbose': {
            'value': 1,
            'type': 'int',
            'description': 'defines the display level for print messages',
        },
        'help': {
            'value': None,
            'description': 'when used we output program options.<br>If set to True, then we '
                           'display in human readable format.<br> If value set to:  tbl,table,helptbl,fmt'
                           ' - then we output in markdown format to be added to readme.md files',
        },
        'helpall': {
            'value': None,
            'description': 'when used we output program options and defaultoptions.<br>'
                           'If set to True, then we display in human readable format.<br>If value set '
                           'to:  tbl,table,helptbl,fmt - then we output in markdown format to be added '
                           'to readme.md files',
        },
        'dumpconfig': {
            'value': False,
            'type': 'bool',
            'description': 'defines if we will dump the final optiondict and exit',
        },
        'dumpconfigfile': {
            'value': None,
            'description': 'defines the filename we dump the populated optiondict dictionary to as json',
        },
        'conf_json': {
            'value': None,
            'type': 'liststr',
            'description': 'defines the list of json file(s) that houses configuration information',
        },
        'conf_mustload': {
            'value': False,
            'type': 'bool',
            'description': 'defines if we are required to load defined configuration files (default: False)',
        },
        'log_level': {
            'value': 'INFO',
            'type': 'inlist',
            'valid': ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
            'description': 'defines the overall logging level for all handlers',
        },
        'log_level_console': {
            'value': 'INFO',
            'type': 'inlist',
            'valid': ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
            'description': 'defines the logging level for console handlers',
        },
        'log_level_file': {
            'value': 'INFO',
            'type': 'inlist',
            'valid': ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
            'description': 'defines the logging level for file handlers',
        },
        'log_file': {
            'value': None,
            'description': 'defines the name of the log file',
        },
    }

    # create the dictionary - and populate values from configuration passed in
    optiondict = {}
    for key in optiondictconfig:
        if 'value' in optiondictconfig[key]:
            # the user specified a value value
            optiondict[key] = optiondictconfig[key]['value']
            # debugging
            logger.debug('Assigning [%s] value from optiondictconfig:%s', key, optiondict[key])
        else:
            # no value option - set to None
            optiondict[key] = None

    # read in the command line options that we care about and create dictionary
    cmdlineargs = {}
    for argpos in range(1, len(sys.argv)):
        # check to see if they have an equal in the string
        if '=' not in sys.argv[argpos]:
            logger.error('Command line arguments must be key=value - there is no equal:%s', sys.argv[argpos])
            raise Exception(
                u'Command line arguments must be key=value - there is no equal:{}'.format(sys.argv[argpos]))

        # get the argument and split it into key and value
        (key, value) = sys.argv[argpos].split('=')

        # debug
        if debug: print('kv_parse_command_line:sys.argv[', argpos, ']:', sys.argv[argpos])
        logger.debug('sys.argv[%s]:%s', argpos, sys.argv[argpos])

        # skip this if the key is not populated
        if not key:
            if debug: print('kv_parse_command_line:key-not-populated-skipping-arg')
            logger.debug('Key-not-populated-with-value-skipping-arg')
            continue

        # check to see if we should use the keymapping
        if keymapdict:
            if key in keymapdict and key not in optiondict and key not in defaultdictconfig:
                logger.debug('Remapping:%s:to:%s', key, keymapdict[key])
                key = keymapdict[key]

        # put this into cmdlineargs dictionary
        cmdlineargs[key] = value

    # read in configuration from json files housing configuration data
    conf_json_files = []
    if 'conf_json' in cmdlineargs:
        # config files defined in the command line
        conf_json_files = cmdlineargs['conf_json'].split(',')
        logger.debug('Config files defined on command line:%s', conf_json_files)
    elif 'conf_json' in optiondict and optiondict['conf_json']:
        # value passed in via optiondictconfig
        if isinstance(optiondict['conf_json'], list):
            # configured correctly - as a list in the optionconfigdict
            conf_json_files = optiondict['conf_json']
        else:
            # need to make sure this setting is of the proper format
            # it was not structured correctly in the json file
            logger.warning('conf_json entered as a string vs list - format converted')
            conf_json_files = [optiondict['conf_json']]
            optiondict['conf_json'] = conf_json_files
        logger.debug('Config files defined on optiondictconfig:%s', conf_json_files)

    # step through all the configuration files reading in the settings
    # and flatten them out into a final configuratin file based dictionary
    confargs = {}
    conf_files_read = list()
    for conf_json_file in conf_json_files:
        logger.debug('conf_json_file:%s', conf_json_file)
        if os.path.exists(conf_json_file):
            with open(conf_json_file, 'r') as json_conf:
                fileargs = json.load(json_conf)
                conf_files_read.append(conf_json_file)
            for key, value in fileargs.items():
                confargs[key] = value
        else:
            if ('conf_mustload' in optiondict and optiondict['conf_mustload']) or (
                    'conf_mustload' in cmdlineargs and cmdlineargs['conf_mustload']):
                raise Exception(u'Missing config file: {}'.format(conf_json_file))
            else:
                logger.warning('Skipped missing config file:%s', conf_json_file)
    if conf_files_read:
        # populate the list of config files with the actual files read in
        optiondict['conf_json'] = conf_files_read
    else:
        if 'conf_json' in optiondict:
            # no files read in - remove this config setting if it exists
            del optiondict['conf_json']
        if 'conf_json' in confargs:
            # no files read in - remove from here also
            del confargs['conf_json']
        if 'conf_json' in cmdlineargs:
            # no files read in - remove from here also
            del cmdlineargs['conf_json']

    # now that we have loaded and flattened out all file based settings
    # move these settings to the final proper destination
    for key, value in confargs.items():
        if key not in cmdlineargs:
            # this file value has no associated command line override
            # value not overridden by value on command line
            if isinstance(value, str):
                # what we have is a string - which is the only thing we can read from the command line
                # stuff this into command line args
                cmdlineargs[key] = value
                logger.debug('conf_json key put into cmdlineargs:%s', key)
            else:
                # this is other than a string - just set the optiondict value with it
                optiondict[key] = value
                logger.debug('conf_json key put into optiondict:%s', key)
        else:
            logger.debug('conf_json ignored because command line overrides it:%s:%s', key, value)

    # now step through the configuration settings we have received
    for key, value in cmdlineargs.items():
        # logic to bring in "default/implied optiondict values if key passed is not part of app definition
        if key not in optiondictconfig and key in defaultdictconfig:
            if debug: print('kv_parse_command_line:key-not-in-optiondictconfig-but-in-defaultoptiondictconfig:', key)
            logger.debug('Key-not-in-optiondictconfig-but-in-defaultoptiondictconfig:%s', key)
            # copy over this default into optiondict
            optiondictconfig[key] = defaultdictconfig[key].copy()
            # tag the defaultdictconfig that we used this key
            defaultdictconfig[key]['applied'] = True
            # set the value
            if 'value' in defaultdictconfig[key]:
                optiondict[key] = defaultdictconfig[key]['value']
            else:
                optiondict[key] = None

        # action on this command line
        if key in optiondict:
            # debug message on type
            if 'type' in optiondictconfig[key]:
                if debug: print('type:', optiondictconfig[key]['type'])
                logger.debug('Key:%stype:%s', key, optiondictconfig[key]['type'])

            if 'type' not in optiondictconfig[key]:
                # user did not specify the type of this option
                optiondict[key] = value
                if debug: print('type not in optiondictconfig[key]')
                logger.debug('Type not in optiondictconfig[key] for key:%s', key)
            elif optiondictconfig[key]['type'] == 'bool':
                optiondict[key] = bool(strtobool(value))
            elif optiondictconfig[key]['type'] == 'int':
                optiondict[key] = int(value)
            elif optiondictconfig[key]['type'] == 'float':
                optiondict[key] = float(value)
            elif optiondictconfig[key]['type'] == 'dir':
                optiondict[key] = os.path.normpath(value)
            elif optiondictconfig[key]['type'] == 'liststr':
                optiondict[key] = value.split(',')
            elif optiondictconfig[key]['type'] == 'date':
                optiondict[key] = kvdate.datetime_from_str(value)
            elif optiondictconfig[key]['type'] == 'datetimezone':
                optiondict[key] = kvdate.datetimezone_from_str(value)
            elif optiondictconfig[key]['type'] == 'inlist':
                # value must be from a predefined list of acceptable values
                if 'valid' not in optiondictconfig[key]:
                    if debug: print('missing optiondictconfig setting [valid] for key:', key)
                    logger.error('Missing optiondictconfig setting [valid] for key:%s', key)
                    raise Exception(u'Missing optiondictconfig setting [valid] for key:{}'.format(key))
                if value not in optiondictconfig[key]['valid']:
                    if debug: print('value:', value, ':not in defined list of valid values:',
                                    optiondictconfig[key]['valid'])
                    logger.error('Invalid value passed in for [%s]:%s', key, value)
                    logger.error('List of valid values are:%s', optiondictconfig[key]['valid'])
                    raise Exception(u'Invalid value passed in for [{}]:{}'.format(key, value))
                optiondict[key] = value
            else:
                # user set a type but we don't know what to do with this type
                optiondict[key] = value
                if debug: print('type not known:', type)
                logger.debug('Type unknown:%s', type)
        elif raise_error:
            logger.error('Unknown command line option:%s', key)
            raise Exception(u'Unknown command line option:{}'.format(key))
        else:
            if debug: print('kv_parse_command_line:unknown-option:', key)
            logger.warning('Unknown option:%s', key)

        # special processing if we are asking for help
        if key in HELP_KEYS:
            # user asked for help - display help and then exit
            tblfmt = False
            if value in HELP_VALUE_TABLE:
                tblfmt = True
            # determine if we are also display the additional options
            defaultoptions = {}
            if key == 'helpall':
                defaultoptions = defaultdictconfig
            kv_parse_command_line_display(optiondictconfig, defaultoptions, tblfmt=tblfmt, debug=False)
            sys.exit()
    # test for required fields being populated
    missingoption = []
    for key in optiondictconfig:
        if 'required' in optiondictconfig[key]:
            if optiondictconfig[key]['required'] and optiondict[key] is None:
                # required field but is populated with None
                missingoption.append('%s:required field not populated' % key)
                optiondictconfig[key]['error'] = 'required value not populated'

    # raise error if we should
    if missingoption:
        kv_parse_command_line_display(optiondictconfig, debug=False)
        errmsg = 'System exitted - missing required option(s):\n    ' + '\n    '.join(missingoption)
        # print('\n'.join(missingoption))
        if debug:
            print('-' * 80)
            print(errmsg)
            print('')
        logger.error(errmsg)
        raise Exception(errmsg)
        # sys.exit(1)

    # debug when we are done
    if debug: print('kv_parse_command_line:optiondict:', optiondict)
    logger.debug('optiondict:%s', optiondict)

    # check to see if we want to dump the optiondict out to a file
    if 'dumpconfigfile' in optiondict and optiondict['dumpconfigfile']:
        dump_dict_to_json_file(optiondict['dumpconfigfile'], optiondict)

    # check to see if they set the dumpconfig setting if so display and exit
    if 'dumpconfig' in optiondict and optiondict['dumpconfig']:
        print('kv_parse_command_line:Dump configuration requested:')
        for (key, val) in optiondict.items():
            print('{}{}:{}'.format(key, '.' * (30 - len(key)), val))
        sys.exit()

    # return what we created
    return optiondict


# update the value of a two level deep key if it is not already set
def set_when_not_set(input_dict, key1, key2, value):
    if key1 in input_dict:
        if key2 not in input_dict[key1]:
            input_dict[key1][key2] = value
            return True
    return False


# display the optiondictconfig information in human readable format
def kv_parse_command_line_display(optiondictconfig, defaultoptions=None, optiondict=None, tblfmt=False, debug=False):
    if defaultoptions is None:
        defaultoptions = {}
    if optiondict is None:
        optiondict = {}

    # set the sortorder for a known set of keys
    set_when_not_set(optiondictconfig, 'AppVersion', 'sortorder', 1)
    set_when_not_set(optiondictconfig, 'debug', 'sortorder', 9997)
    set_when_not_set(optiondictconfig, 'help', 'sortorder', 9998)
    set_when_not_set(optiondictconfig, 'helpall', 'sortorder', 9999)

    # predefined number ranges by type
    nextcounter = {
        'None': 2,
        'dir': 100,
        'int': 200,
        'float': 300,
        'bool': 400,
        'date': 500,
        'datetimezone': 600,
        'liststr': 700,
        'inlist': 800,
    }

    opt2sort = []

    # step through the optional keys
    for opt in sorted(optiondictconfig.keys()):
        if 'type' in optiondictconfig[opt]:
            # type set - use it
            typeupdate = optiondictconfig[opt]['type']
        else:
            # type not set - make it 'None'
            typeupdate = 'None'

        if set_when_not_set(optiondictconfig, opt, 'sortorder', nextcounter[typeupdate]):
            # we updated the sort order for this record - so we must update the counter
            nextcounter[typeupdate] += 1

        # now build sort string
        opt2sort.append([optiondictconfig[opt]['sortorder'], opt])

    # add in the default options if we have them populated
    if defaultoptions:
        sortcnt = 9996
        opt = '-----'
        optiondictconfig[opt] = {'value': opt, 'description': opt, 'type': opt}
        opt2sort.append([sortcnt, opt])
        sortcnt = 10000

        for opt in defaultoptions.keys():
            if opt not in optiondictconfig:
                if opt == 'help':
                    opt2sort.append([9998, opt])
                elif opt == 'helpall':
                    opt2sort.append([9999, opt])
                else:
                    opt2sort.append([sortcnt, opt])
                optiondictconfig[opt] = defaultoptions[opt]
                sortcnt += 1

    # header if we are doing table output
    if tblfmt:
        print('| option | type | value | description |')
        print('| ------ | ---- | ----- | ----------- |')

    # define the string format for each cell in the table
    tbl_fmt = ' {} |'

    # step through the sorted list and display things
    for row in sorted(opt2sort):
        opt = row[1]
        if opt in optiondict:
            optiondictconfig[opt]['value'] = optiondict[opt]

        # output style
        if tblfmt:
            # user wanted to output in table format - each line with no <newline>
            print('| {} |'.format(opt), end="")
            # output the type - may not be populated
            fld = 'type'
            fldout = ''
            if fld in optiondictconfig[opt]:
                fldout = optiondictconfig[opt][fld]
            print(tbl_fmt.format(fldout), end="")
            # output the value - may not be populated
            fld = 'value'
            fldout = ''
            if fld in optiondictconfig[opt]:
                fldout = optiondictconfig[opt][fld]
            if opt in optiondict and fld in optiondict[opt]:
                fldout = optiondict[opt][fld]
            print(tbl_fmt.format(fldout), end="")
            # output the type - may not be populated
            fld = 'description'
            fldout = ''
            if fld in optiondictconfig[opt]:
                fldout = optiondictconfig[opt][fld]
            # add in valid, error values if they exist
            for fld in ('valid', 'error'):
                if fld in optiondictconfig[opt]:
                    if fldout:
                        fldout += '<br>'
                    fldout += 'valid:{}'.format(optiondictconfig[opt][fld])
            # output this field - but this time with a <newline>
            print(tbl_fmt.format(fldout))
        else:
            # linear output 
            if 'type' in optiondictconfig[opt]:
                print('option.:', opt, ' (type:', optiondictconfig[opt]['type'], ')')
            else:
                print('option.:', opt)

                for fld in ('value', 'required', 'description', 'valid', 'error'):
                    if fld in optiondictconfig[opt]:
                        print('  ' + fld + '.' * (12 - len(fld)) + ':', optiondictconfig[opt][fld])


# define the filename used to create log files
# that are based on the "day" the program starts running
# generally used for short running tools
# not used with tools that start and stay running
def filename_log_day_of_month(filename, ext_override=None, path_override=None):
    file_path, base_filename, file_ext = filename_split(filename, path_blank=True)
    if ext_override:
        file_ext = ext_override
    if file_ext[:1] != '.':
        file_ext = '.' + file_ext
    if path_override:
        file_path = path_override
    day_filename = '{}{:02d}'.format(base_filename, datetime.datetime.today().day)
    logfilename = os.path.join(file_path, day_filename + file_ext)
    if os.path.exists(logfilename):
        if os.path.getmtime(logfilename) < (datetime.datetime.today() - datetime.timedelta(days=1)).timestamp():
            # remove the file if it exists but has not been modified within the past 24 hours
            remove_filename(logfilename)
    return logfilename


# return the filename that is max or min for a given query (UT)
# default is to return the MIN filematch
def filename_maxmin(file_glob, reverse=False):
    # pull the list of files
    filelist = glob.glob(file_glob)
    # debugging
    logger.debug('filelist:%s', filelist)
    # if we got no files - return none
    if not filelist:
        logger.debug('Return none')
        return None
    logger.debug('File:%s', sorted(filelist, reverse=reverse)[0])
    # sort this list - and return the desired value
    return sorted(filelist, reverse=reverse)[0]


# create a filename from part of a filename
#   pull apart the filename passed in (if passed in) and then fill in the various file parts based
#   on the other attributes passed in
def filename_create(filename=None, filename_path=None, filename_base=None, filename_ext=None, path_blank=False,
                    filename_base_append=None, filename_base_prepend=None, use_input_filename=None,
                    filename_unique=None):
    # pull apart the filename passed in
    if filename:
        file_path, base_filename, file_ext = filename_split(filename, path_blank=path_blank)
    else:
        file_path = base_filename = file_ext = ''
    if filename_ext:
        file_ext = filename_ext
    if file_ext and file_ext[:1] != '.':
        # put the dot into the extension
        file_ext = '.' + file_ext
    if filename_path:
        file_path = filename_path
    if filename_base:
        base_filename = filename_base
    if filename_base_prepend:
        base_filename = filename_base_prepend + base_filename
    if filename_base_append:
        base_filename += filename_base_append
    if filename_path:
        file_path = filename_path
    elif path_blank:
        file_path = ''
    return os.path.normpath(os.path.join(file_path, base_filename + file_ext))


# split up a filename into parts (path, basename, extension) (UT)
def filename_split(filename, path_blank=False):
    filename2, file_ext = os.path.splitext(filename)
    base_filename = os.path.basename(filename2)
    if path_blank:
        file_path = os.path.dirname(filename2)
    else:
        file_path = os.path.normpath(os.path.dirname(filename2))
    return file_path, base_filename, file_ext


# function to get back a full list of broken up file path
def filename_splitall(path):
    allparts = []
    while 1:
        parts = os.path.split(path)
        if parts[0] == path:  # sentinel for absolute paths
            allparts.insert(0, parts[0])
            break
        elif parts[1] == path:  # sentinel for relative paths
            allparts.insert(0, parts[1])
            break
        else:
            path = parts[0]
            allparts.insert(0, parts[1])
    return allparts


# create a list of filenames given a name, a list of names, file glob,
# list of include files in a file, list of exclue files in a file
def filename_list(filename=None, filenamelist=None, fileglob=None, strippath=False, includelist_filename=None,
                  excludefilenamelist=None, excludelist_filename=None, glob_filename=None):
    # local variable
    flist = []
    exclude_list = []
    # read list from files provide
    if includelist_filename:
        flist = read_list_from_file_lines(includelist_filename, trim=True)
    if excludelist_filename:
        exclude_list = read_list_from_file_lines(excludelist_filename, trim=True)
    if excludefilenamelist:
        exclude_list.extend(excludefilenamelist)
    # read list from records provided
    if fileglob:
        flist.extend(glob.glob(fileglob))
    if filenamelist:
        flist.extend(filenamelist)
    if filename:
        if glob_filename:
            if isinstance(filename, list):
                for fname in filename:
                    flist.extend(glob.glob(fname))
            else:
                flist.append(glob.glob(filename))
        else:
            if isinstance(filename, list):
                flist.extend(filename)
            else:
                flist.append(filename)
            
    # remove records if exclude definitions provided
    if exclude_list:
        for excludefile in exclude_list:
            if excludefile in flist:
                flist.remove(excludefile)

    # strip path from filename if flag is set
    if strippath:
        for ndx in range(len(flist)):
            flist[ndx] = os.path.basename(flist[ndx])

    # create the unique list of filenames and return them
    return sorted(list(set(flist)))


# create a full filename and optionally validate directory exists and is writeabile (UT)
def filename_proper(filename_full, file_dir=None, create_dir=False, write_check=False, debug=False):
    filename = os.path.basename(filename_full)
    if not file_dir:
        file_dir = os.path.dirname(filename_full)

    # if there is no directory then make it the current directory
    if not file_dir:
        file_dir = './'

    # wondering if we need to extract directory and compare if set (future feature)
    # and if they are different - what action should we take?

    # check the directory and determine if we need it to be created
    if not os.path.exists(file_dir):
        # directory needs to be created
        if create_dir:
            # needs to be created and we have enabled this option
            try:
                os.makedirs(file_dir)
            except Exception as e:
                if debug: print('kvutil:filename_proper:makedirs:%s' % e)
                logger.error('makedirs:%s' % e)
                raise Exception(u'kvutil:filename_proper:makedirs:{}'.format(e))
        else:
            # needs to be created - option not enabled - raise an error
            if debug: print('kvutil:filename_proper:directory does not exist:%s' % file_dir)
            logger.error('Directory does not exist:%s', file_dir)
            raise Exception(u'kvutil:filename_proper:directory does not exist:{}'.format(file_dir))

    # check to see if the directory is writeable if the flag is set
    if write_check:
        if not os.access(file_dir, os.W_OK):
            if debug: print('kvutil:filename_proper:directory is not writeable:%s' % file_dir)
            logger.error('Directory is not writeable:%s', file_dir)
            raise Exception(u'kvutil:filename_proper:directory is not writeable:{}'.format(file_dir))

    # build a full filename
    full_filename = os.path.join(file_dir, filename)

    # return the calculated filename
    return os.path.normpath(full_filename)


# create a unique filename
def filename_unique(filename=None, filename_href=None, debug=False):
    if filename_href is None:
        filename_href = {}

    # check input
    if isinstance(filename, dict):
        filename_href = filename
        filename = None

    # default options for the filename_href
    default_options = {
        'file_ext': '.html',  # file extension
        'full_filename': '',
        'file_path': './',  # path to where to put the file
        'filename': '',
        'tmp_file_path': '',
        'base_filename': 'tmpfile',  # basefilename
        'ov_ext': '.bak',  # overwritten saved file extension
        'uniqtype': 'cnt',  # defines how we make this filename uniq
        'cntfmt': 'v%02d',  # format string for converting count
        'datefmt': '-%Y%m%d',  # format string for converting date
        'maxcnt': 100,  # maximum count to search for unique filename
        'forceuniq': False,  # do not force unique filename creation
        'overwrite': False,  # 1=overwrite an existing file
        'create_dir': False,  # if true - we will create the directory specified if it does not exist
        'write_check': True,  # validate we can write in the specified directory
        'verbose_uf': 0,
    }
    # list of required fields to be populated
    required_values = ['file_ext', 'file_path', 'base_filename', 'uniqtype']

    # list of valid values for inputs
    validate_values = {
        'uniqtype': ['cnt', 'datecnt']
    }
    # force the value of this field if the value is blank
    force_if_blank = {
        'file_path': './',
    }

    # bring in the values that were passed in
    for key in default_options:
        if key in filename_href:
            default_options[key] = filename_href[key]

    # if filename is provided split it up
    if filename:
        default_options['file_path'], default_options['base_filename'], default_options['file_ext'] = filename_split(
            filename)
    else:
        # parse up the full_filename if passed in
        if default_options['full_filename']:
            default_options['file_path'], default_options['base_filename'], default_options[
                'file_ext'] = filename_split(default_options['full_filename'])
        elif default_options['filename']:
            default_options['file_path'], default_options['base_filename'], default_options[
                'file_ext'] = filename_split(default_options['filename'])
        else:
            # make sure base_filename is only a filename
            default_options['base_filename'] = os.path.basename(default_options['base_filename'])
            default_options['file_path'] = os.path.dirname(default_options['file_path'])

    # force if blank fields
    for key in force_if_blank:
        if not default_options[key]:
            default_options[key] = force_if_blank[key]

    # check that required fields are populated
    field_issues = []
    for key in required_values:
        if not default_options[key]:
            field_issues.append(key)

    # check to see if we have and field issues
    if field_issues:
        if debug: print('kvutil:filename_unique:missing values for: {}'.format(','.join(field_issues)))
        logger.error('Missing values for:%s', ','.join(field_issues))
        raise Exception(u'kvutil:filename_unique:missing values for: {}'.format(','.join(field_issues)))

    # check that we have valid values
    for key in validate_values:
        if not default_options[key] in validate_values[key]:
            field_issues.append(key)

    # check to see if we have and field issues
    if field_issues:
        if debug: print('kvutil:filename_unique:invalid values for: {}'.format(','.join(field_issues)))
        logger.error('Invalid values for:%s', ','.join(field_issues))
        raise Exception(u'kvutil:filename_unique:invalid values for: {}'.format(','.join(field_issues)))

    # create a filename if it does not exist
    default_options['filename'] = os.path.normpath(
        os.path.join(default_options['base_filename'] + default_options['file_ext']))

    # check the directory to see if it exists
    default_options['file_path'] = filename_proper(default_options['file_path'],
                                                   create_dir=default_options['create_dir'],
                                                   write_check=default_options['write_check'])

    # if we are NOT doing datecnt - then clear the date_file
    if default_options['uniqtype'] == 'cnt':
        date_file = ''
    else:
        date_file = datetime.datetime.now().strftime(default_options['datefmt'])

    # start the counter for file version number
    unique_counter = 1

    # set the starting filename
    if default_options['forceuniq']:
        # want a unique filename - create a filename base on the filename options
        filename = default_options['base_filename'] + date_file + (default_options['cntfmt'] % unique_counter) + \
                   default_options['file_ext']
    else:
        # not a unique - try the filename passed infirst
        filename = default_options['filename']

    # debugging
    # print('file_unique:filename:', filename)
    # print('file_unique:default_options:', default_options)

    # take action if we are not going to overwrite the filename
    if not default_options['overwrite']:

        # look for the filename that works
        while (os.path.exists(os.path.join(default_options['file_path'], filename)) and unique_counter <
               default_options['maxcnt']):
            # create a new filename
            filename = default_options['base_filename'] + date_file + (default_options['cntfmt'] % unique_counter) + \
                       default_options['file_ext']
            # increment the counter
            unique_counter += 1

        # test to see if we exceeded the max count and if so error out.
        if unique_counter >= default_options['maxcnt']:
            if debug: print('kvutil:filename_unique:reached maximum count and not unique filename:', filename)
            logger.error('Reached maximum count and not unique filename:%d:%s', unique_counter, filename)
            raise Exception(
                u'kvutil:filename_unique:reached maximum count and not unique filename: {}'.format(filename))

    # debugging
    # print('file_unique:filename:final:', filename)

    # return the final filename
    return filename_proper(filename, file_dir=default_options['file_path'])


# , \
#           filename_proper( filename +  default_options['ov_ext'], dir=default_options['file_path'])


# cloudpath - create an absolute path to a folder that is local for cloud drive
def cloudpath(filepath, filename=''):
    userdir = ''
    if filepath is None:
        filepath = ''
    if filename is None:
        filename = ''
    # determine if the path is a cloud path
    for cloudprovider in ('Box Sync', 'Dropbox', 'OneDrive'):
        index = filepath.find(cloudprovider)
        if index != -1:
            filepath = filepath[index:]
            userdir = os.path.expanduser('~')
            break

    return os.path.abspath(os.path.join(userdir, filepath, filename))


# read a text file into a string (UT)
def slurp(filename):
    with open(filename, 'r') as t:
        return t.read()


# read in a file and create a list of each populated line (UT)
def read_list_from_file_lines(filename, stripblank=False, trim=False, encoding=None):
    # read in the file as a list of strings
    if encoding:
        with open(filename, 'r', encoding=encoding) as t:
            filelist = t.readlines()
    else:
        with open(filename, 'r') as t:
            filelist = t.readlines()

    # strip the trailing \n
    filelist = [line.strip('\n') for line in filelist]

    # strip the trailing \n
    if trim:
        filelist = [line.strip() for line in filelist]

    # if they want to strip blank lines
    if stripblank:
        filelist = [line for line in filelist if line and line.strip()]

    # return the list of lines
    return filelist


# utility used to remove a filename - in windows sometimes we have a delay
# in releasing the filehandle - this routine will loop a few times giving
# time for the OS to release the blocking issue and then delete
#
# optional input:
#    calledfrom - string used to display - usually the name of module.function()
#    debug - bool defines if we display duggging print statements
#    maxretry - int - number of times we try to delete and then give up (default: 20)
#
def remove_filename(filename, calledfrom='', debug=False, maxretry=20):
    logger.debug('Remove:%s:calledfrom:%s:maxretry:%d', filename, calledfrom, maxretry)
    cnt = 0
    if calledfrom:  calledfrom += ':'
    while os.path.exists(filename):
        cnt += 1
        if debug: print(calledfrom, filename, ':exists:try to remove:cnt:', cnt)
        logger.debug('%s:%s:exists:try to remove:cnt:%d', calledfrom, filename, cnt)
        try:
            os.remove(filename)  # try to remove it directly
            logger.debug('%s:%s:removed on count:%d', calledfrom, filename, cnt)
        except Exception as e:
            if debug: print(calledfrom, 'errno:', e.errno, ':ENOENT:', errno.ENOENT)
            logger.debug('%s:errno:%d:ENOENT:%d', calledfrom, e.errno, errno.ENOENT)
            if e.errno == errno.ENOENT:  # file doesn't exist
                return
            if debug: print(calledfrom, filename, ':', str(e))
            if cnt > maxretry:
                if debug: print(calledfrom, filename, ':raise error - exceed maxretry attempts:', maxretry)
                logger.error('%s:%s:exceeded maxretry attempts:%d:raise error', calledfrom, filename, maxretry)
                raise e
        except WinError as f:
            if debug: print('Catch WinError:', str(f))
            logger.warning('Catch WinError:%s', str(f))


# utility used to remove a folder - in windows sometimes we have a delay
# in releasing the filehandle - this routine will loop a few times giving
# time for the OS to release the blocking issue and then delete
#
# optional input:
#    calledfrom - string used to display - usually the name of module.function()
#    debug - bool defines if we display duggging print statements
#    maxretry - int - number of times we try to delete and then give up (default: 20)
#
def remove_dir(dirname, calledfrom='', debug=False, maxretry=20):
    cnt = 0
    if calledfrom:  calledfrom += ':'
    while os.path.exists(dirname):
        cnt += 1
        if debug: print(calledfrom, dirname, ':exists:try to remove:cnt:', cnt)
        try:
            os.rmdir(dirname)  # try to remove it directly
        #        except OSError as e: # originally just checked for OSError - we now check for all exceptions`
        except Exception as e:
            if debug: print(calledfrom, 'errno:', e.errno, ':ENOENT:', errno.ENOENT)
            logger.debug('%s:errno:%s:ENOENT:%s', calledfrom, e.errno, errno.ENOENT)
            if e.errno == errno.ENOENT:  # file doesn't exist
                return
            if debug: print(calledfrom, dirname, ':', str(e))
            logger.debug('%s:%s:%s', calledfrom, dirname, str(e))
            if cnt > maxretry:
                if debug: print(calledfrom, dirname, ':raise error - exceed maxretry attempts:', maxretry)
                logger.error('%s:%s:maxretry attempts:%d', calledfrom, dirname, maxretry)
                raise e
        except WinError as f:
            if debug: print('Catch WinError:', str(f))
            logger.warning('Catch WinError:%s', str(f))


# return the function name of the function that called this
def functionName(callBackNumber=1):
    return sys._getframe(callBackNumber).f_code.co_name


# create the starting logger header that we want to show the separation
# between runs - this utility is just to enable logging standardization.
#
# In your program put:  kvutil.loggingAppStart( logger, optiondict, kvutil.scriptinfo()['name'] )
#
def loggingAppStart(logger, optiondict, pgm=None):
    logger.info('-----------------------------------------------------')
    if pgm:
        logger.info('%s:AppVersion:v%s', pgm, optiondict['AppVersion'])
    else:
        logger.info('AppVersion:v%s', optiondict['AppVersion'])


def scriptinfo():
    """
    Returns a dictionary with information about the running top level Python
    script:
    ---------------------------------------------------------------------------
    dir:    directory containing script or compiled executable
    name:   name of script or executable
    source: name of source code file
    ---------------------------------------------------------------------------
    "name" and "source" are identical if and only if running interpreted code.
    When running code compiled by py2exe or cx_freeze, "source" contains
    the name of the originating Python script.
    If compiled by PyInstaller, "source" contains no meaningful information.
    """

    import os
    import sys
    import inspect

    # ---------------------------------------------------------------------------
    # scan through call stack for caller information
    # ---------------------------------------------------------------------------
    trc = ''
    for teil in inspect.stack():
        # skip system calls
        if teil[1].startswith("<"):
            continue
        if teil[1].upper().startswith(sys.exec_prefix.upper()):
            continue
        trc = teil[1]

    # trc contains highest level calling script name
    # check if we have been compiled
    if getattr(sys, 'frozen', False):
        scriptdir, scriptname = os.path.split(sys.executable)
        return {"dir": scriptdir,
                "name": scriptname,
                "source": trc}

    # from here on, we are in the interpreted case
    scriptdir, trc = os.path.split(trc)
    # if trc did not contain directory information,
    # the current working directory is what we need
    if not scriptdir:
        scriptdir = os.getcwd()

    scr_dict = {"name": trc,
                "source": trc,
                "dir": scriptdir}
    return scr_dict


# utility used to dump a dictionary to a file in json format
def load_json_file_to_dict(filename):
    import json
    with open(filename, 'r') as json_in:
        try:
            json_dict = json.load(json_in)
        except json.decoder.JSONDecodeError as e:
            import re
            with open(filename, 'r') as json_error:
                json_lines = json_error.readlines()
            err_line = re.search(r'line\s+(\d+)\s+', str(e))
            print('-'*40)
            if err_line:
                err_line_int = int(err_line.group(1))
                if err_line_int < len(json_lines):
                    print('Error on line: ', err_line_int)
                    print(json_lines[err_line_int-1])
            print('-'*40)
            raise
    return json_dict


# utility used to dump a dictionary to a file in json format
def dump_dict_to_json_file(filename, optiondict):
    import json
    with open(filename, 'w') as json_out:
        json.dump(optiondict, json_out, indent=4)


# utility to convert a dict to a list of dicts that are key, value and new value
def dict2update_list(in_dict, sorted_flds=None, col_names=None):
    # colnames is a dictionary with entries tied to the desired output columname
    #  {'Field': header_col1, 'CurrentValue': header_col2, 'NewValue': header_col3}

    default_column_names = ['Field', 'CurrentValue', 'NewValue']
    output_col_names = []

    # make sure they passed the right type
    if type(in_dict) != dict:
        raise TypeError('in_dict must be a dictionary')

    # the user can pass in the fields to be generated in a sorted order
    if not sorted_flds:
        sorted_flds = list(in_dict.keys())

        
    # make sure they passed the right type
    if type(sorted_flds) != list:
        raise TypeError('sort_flds must be a list')

    
    # if they want to set the column headers
    if col_names and type(col_names) == dict:
        for hdr in default_column_names:
            if hdr in col_names:
                output_col_names.append(col_names[hdr])
            else:
                output_col_names.append(hdr)
    else:
        output_col_names = default_column_names
            
    # now flip the dictionary to the desired output
    outlist = []
    for k in sorted_flds:
        # make sure the field is a valid key
        if k in in_dict:
            outlist.append({output_col_names[0]: k, output_col_names[1]: in_dict[k], output_col_names[2]: ''})
        else:
            print('warning: dict2update_list passed in valid key: {k}')

    return outlist

# return true if one of the copy_fields values is populated
def any_field_is_populated(rec, copy_fields):
    '''
    Return a TRUE if any of the 'copy_fields' elements in rec is populated
    '''
    for fld in copy_fields:
        # current conditions - if it returns true or has a length
        if rec[fld]:
            # print('rec populated')
            return True
        elif not isinstance(rec[fld], str):
            # print('type not string')
            return True
    return False

# for a list of records and a dictionary with defaults - set columns if blank
def set_blank_field_values(src_data, set_blank_fields):
    '''
    For each record in src_data
    For each column defined in set_blank_fields dictionary (if it is spaces it will not overwrite/update)
    Check the record column value and if not set, then set it to the value from set_blank__fields

    src_data - list of dictionaries
    set_blank_fields - dictionary with key and defined value
    '''
    records_updated = 0 
    for rec in src_data:
        record_updated = False
        for k,v in set_blank_fields.items():
            # if key in record and this column has no data
            if k in rec and not rec[k]:
                rec[k] = v
                record_updated = True
        # increment count if we updated the record
        if record_updated:
            records_updated += 1
    # return the number of records update
    return records_updated


# for a list of records and a dictionary with defaults - set columns if blank
def convert_hyperlink_field_values(src_data, hyperlink_fields):
    '''
    For each record in src_data
    For each column defined in hyperlink_fields list
    Check the record column value and if not set, then set it to the value from set_blank__fields

    src_data - list of dictionaries
    hyperlink_fields - list of columns to check and update
    '''
    records_updated = 0 
    for rec in src_data:
        record_updated = False
        for fld in hyperlink_fields:
            # if key in record and this column has no data
            if fld in rec and rec[fld] and rec[fld].startswith('=HYPERLINK'):
                rec[fld] = rec[fld][11:-1]
                record_updated = True
        # increment count if we updated the record
        if record_updated:
            records_updated += 1
    # return the number of records update
    return records_updated


# create a multi-key dictionary from a list of dictionaries
def create_multi_key_lookup(src_data, fldlist, copy_fields=None):
    '''
    Create a multi key dictionary that gets to the record based on the
    keys in the record

    if user sets the copy_fields with the list of fields that can have values
    then we check the record
    to determine if any of the fields has a value, and if none have a value we skip
    that record
    '''
    if type(fldlist) is not list:
        print('fldlist must be type - list - but is: ', type(fldlist))
        raise TypeError()
    # check that the fldlist keys are in the first record
    for fld in fldlist:
        if fld not in src_data[0]:
            print('ERROR:  Unable to find key field: ', fld)
            print('in first record:')
            pprint.pprint(src_data[0])
            print('This routine will fail')
    # check that the copy_fields keys are in the first record
    if copy_fields:
        if type(copy_fields) is not list:
            print('copy_fields must be type - list - but is: ', type(copy_fields))
            raise TypeError()
        for fld in copy_fields:
            if fld not in src_data[0]:
                print('ERROR:  Unable to find copy field: ', fld)
                print('in first record:')
                pprint.pprint(src_data[0])
                print('This routine will fail')
    #
    # set up the dictionary to be populated
    src_lookup = {}
    # step through each record
    for rec in src_data:
        # test that this record has values in the copy_fields attributes
        if copy_fields and not any_field_is_populated(rec, copy_fields):
            # no values set in copy_fields has a value so we don't convert this record
            continue
        # get the first key
        if rec[fldlist[0]] not in src_lookup:
            if len(fldlist) > 1:
                # multi key
                src_lookup[rec[fldlist[0]]] = {}
            else:
                # single key - set the value
                src_lookup[rec[fldlist[0]]] = rec
        # now create the changing key
        ptr = src_lookup[rec[fldlist[0]]]
        # now work through other keys
        for fld in fldlist[1:]:
            # check to see this level is working
            if rec[fld] not in ptr:
                ptr[rec[fld]] = {}
            # if we are on the last fld then set to rec
            if fld == fldlist[-1]:
                ptr[rec[fld]] = rec
            else:
                # update the ptr
                ptr = ptr[rec[fld]]
    #
    return src_lookup


# create a multi-key dictionary from a list of dictionaries
def create_multi_key_lookup_excel(excel_dict, fldlist, copy_fields=None):
    '''
    Create a multi key dictionary that gets to the record based on the
    keys in the record

    if user sets the copy_fields with the list of fields that can have values
    then we check the record
    to determine if any of the fields has a value, and if none have a value we skip
    that record
    '''
    if type(fldlist) is not list:
        print('fldlist must be type - list - but is: ', type(fldlist))
        raise TypeError()
    # check that the fldlist keys are in the first record
    for fld in fldlist:
        if fld not in excel_dict['header']:
            print('ERROR:  Unable to find key field: ', fld)
            print('in the header:')
            pprint.pprint(excel_dict['header'])
            print('This routine will fail')
    # check that the copy_fields keys are in the first record
    if copy_fields:
        if type(copy_fields) is not list:
            print('copy_fields must be type - list - but is: ', type(copy_fields))
            raise TypeError()
        for fld in copy_fields:
            if fld not in excel_dict['header']:
                print('ERROR:  Unable to find copy field: ', fld)
                print('in the header:')
                pprint.pprint(excel_dict['header'])
                print('This routine will fail')
    #
    # set up the dictionary to be populated
    src_lookup = {}
    # step through each record
    for rec in src_data:
        # test that this record has values in the copy_fields attributes
        if copy_fields and not any_field_is_populated(rec, copy_fields):
            # no values set in copy_fields has a value so we don't convert this record
            continue
        # get the first key
        if rec[fldlist[0]] not in src_lookup:
            if len(fldlist) > 1:
                # multi key
                src_lookup[rec[fldlist[0]]] = {}
            else:
                # single key - set the value
                src_lookup[rec[fldlist[0]]] = rec
        # now create the changing key
        ptr = src_lookup[rec[fldlist[0]]]
        # now work through other keys
        for fld in fldlist[1:]:
            # check to see this level is working
            if rec[fld] not in ptr:
                ptr[rec[fld]] = {}
            # if we are on the last fld then set to rec
            if fld == fldlist[-1]:
                ptr[rec[fld]] = rec
            else:
                # update the ptr
                ptr = ptr[rec[fld]]
    #
    return src_lookup


def copy_matched_data(dst_data, src_lookup, key_fields, copy_fields):
    '''
    copy into dst_data from src_lookup, copy_fields when there is a match
    on key_fields
    '''
    # make sure we passed in a list
    if type(key_fields) is not list:
        print('key_fields must be type - list - but is: ', type(key_fields))
        raise TypeError()
    # check that the key_fields keys are in the first record
    for fld in key_fields:
        if fld not in dst_data[0]:
            print('ERROR:  Unable to find key_field field: ', fld)
            print('in first record:')
            pprint.pprint(dst_data[0])
            print('This routine will fail')
    # make sure we passed in a list
    if type(copy_fields) is not list:
        print('copy_fields must be type - list - but is: ', type(copy_fields))
        raise TypeError()
    # check that the copy_fields keys are in the first record
    for fld in copy_fields:
        if fld not in dst_data[0]:
            print('ERROR:  Unable to find copy_field field: ', fld)
            print('in first record:')
            pprint.pprint(dst_data[0])
            print('This routine will fail')
    #
    # capture the count of matched records
    matched_recs = 0
    # step through the dst_data
    for rec in dst_data:
        # cpature if we have a match
        matched = True
        # capture the pointer
        ptr = src_lookup
        # step through the key_fields and see if we find a matching record
        for fld in key_fields:
            # there is a match
            if rec[fld] in ptr:
                ptr = ptr[rec[fld]]
            else:
                matched = False
                # stop looking for match on this record
                break
        # check to see if we did match get next record
        if not matched:
            continue
        # increment the matched out
        matched_recs += 1
        # we did match so copy over the fields
        # ptr should point at the record of interest from src_lookup
        for cfld in copy_fields:
            rec[cfld] = ptr[cfld]
    # return the number of records that matched
    return matched_recs


def extract_unmatched_data(src_data, dst_lookup, key_fields):
    '''
    return the list of records in src_data that are no longer in dst_lookup
    '''
    # make sure we passed in a list
    if type(key_fields) is not list:
        print('key_fields must be type - list - but is: ', type(key_fields))
        raise TypeError()
    # check that the key_fields keys are in the first record
    for fld in key_fields:
        if fld not in src_data[0]:
            print('ERROR:  Unable to find key_field field: ', fld)
            print('in first record:')
            pprint.pprint(src_data[0])
            print('This routine will fail')
    #
    # capture the count of matched records
    unmatched_recs = []
    # step through the src_data
    for rec in src_data:
        # cpature if we have a match
        matched = True
        # capture the pointer
        ptr = dst_lookup
        # step through the key_fields and see if we find a matching record
        for fld in key_fields:
            # there is a match
            if rec[fld] in ptr:
                ptr = ptr[rec[fld]]
            else:
                matched = False
                # stop looking for match on this record
                break
        # check to see if we did match get next record
        if not matched:
            # there was not a match - that is what we are looking for
            unmatched_recs.append(rec)
    # return the number of records that matched
    return unmatched_recs

# eof
