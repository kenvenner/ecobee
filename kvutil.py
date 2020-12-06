'''
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.50

Library of tools used in general by KV
'''

import glob
import os
import datetime
import sys
import errno
import json
from distutils.util import strtobool

# setup the logger
import logging
logger = logging.getLogger(__name__)

# set the module version number
AppVersion = '1.50'

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
## Special behavior
#  help=<value>
#
#  will cause the system to generate a file help file when this is passed in on the command line
#
#  if <value> in list ('tbl','table','helptbl','fmt'), then the output is mark down table
#
def kv_parse_command_line( optiondictconfig, raise_error=False, keymapdict=None, debug=False ):
    # debug
    if debug:  print('kv_parse_command_line:sys.argv:', sys.argv)
    if debug:  print('kv_parse_command_line:optiondictconfig:', optiondictconfig)
    # debugging
    logger.debug('LOAD(v%s)%s', AppVersion, '-'*40)
    logger.debug('sys.argv: %s', sys.argv)
    logger.debug('optiondictconfig: %s', optiondictconfig)

    # default a set of basic config values - so we don't need to put them in each app
    defaultdictconfig = {
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
        'help' : {
            'value' : None,
            'description' : 'when used we output program options.<br>If set to True, then we display in human readable format.<br> If value set to:  tbl,table,helptbl,fmt - then we output in markdown format to be added to readme.md files',
        },
        'helpall' : {
            'value' : None,
            'description' : 'when used we output program options and defaultoptions.<br>If set to True, then we display in human readable format.<br>If value set to:  tbl,table,helptbl,fmt - then we output in markdown format to be added to readme.md files',
        },
        'dumpconfig' : {
            'value' : False,
            'type'  : 'bool',
            'description' : 'defines if we will dump the final optiondict and exit',
        },
        'dumpconfigfile' : {
            'value' : None,
            'description' : 'defines the filename we dump the populated optiondict dictionary to as json',
        },
        'conf_json' : {
            'value' : None,
            'type'  : 'liststr',
            'description' : 'defines the list of json file(s) that houses configuration information',
        },
        'conf_mustload' : {
            'value' : False,
            'type'  : 'bool',
            'description' : 'defines if we are required to load defined configuration files (default: False)',
        },
        'log_level' : {
            'value' : 'INFO',
            'type'  : 'inlist',
            'valid' : ['DEBUG','INFO','WARNING','ERROR','CRITICAL'],
            'description' : 'defines the overall logging level for all handlers',
        },
        'log_level_console' : {
            'value' : 'INFO',
            'type'  : 'inlist',
            'valid' : ['DEBUG','INFO','WARNING','ERROR','CRITICAL'],
            'description' : 'defines the logging level for console handlers',
        },
        'log_level_file' : {
            'value' : 'INFO',
            'type'  : 'inlist',
            'valid' : ['DEBUG','INFO','WARNING','ERROR','CRITICAL'],
            'description' : 'defines the logging level for file handlers',
        },
        'log_file'  : {
            'value' : None,
            'description' : 'defines the name of the log file',
        },
    }
                
        
    # create the dictionary - and populate values from configuration passed in
    optiondict = {}
    for key in optiondictconfig:
        if 'value' in optiondictconfig[key]:
            # the user specified a value value
            optiondict[key] = optiondictconfig[key]['value']
            # debugging
            logger.debug('assigning [%s] value from optiondictconfig:%s', key, optiondict[key])
        else:
            # no value option - set to None
            optiondict[key] = None

    # read in the command line options that we care about and create dictionary
    cmdlineargs = {}
    for argpos in range(1,len(sys.argv)):
        # get the argument and split it into key and value
        (key,value) = sys.argv[argpos].split('=')

        # debug
        if debug:  print('kv_parse_command_line:sys.argv[',argpos,']:',sys.argv[argpos])
        logger.debug('sys.argv[%s]:%s',argpos,sys.argv[argpos])

        # skip this if the key is not populated
        if not key:
            if debug:  print('kv_parse_command_line:key-not-populated-skipping-arg')
            logger.debug('key-not-populated-with-value-skipping-arg')
            continue

        # check to see if we should use the keymapping
        if keymapdict:
            if key in keymapdict and key not in optiondict and key not in defaultdictconfig:
                logger.debug('remapping:%s:to:%s', key, keymapdict[key])
                key = keymapdict[key]

        # put this into cmdlineargs dictionary
        cmdlineargs[key] = value

    # read in configuration from json files housing configuration data
    conf_json_files = []
    if 'conf_json' in cmdlineargs:
        # config files defined in the command line
        conf_json_files = cmdlineargs['conf_json'].split(',')
        logger.debug('config files defined on command line:%s', conf_json_files)
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
        logger.debug('config files defined on optiondictconfig:%s', conf_json_files)

    # step through all the configuration files reading in the settings
    # and flatten them out into a final configuratin file based dictionary
    confargs = {}
    conf_files_read=list()
    for conf_json_file in conf_json_files:
        logger.debug('conf_json_file:%s', conf_json_file)
        if os.path.exists( conf_json_file ):
            with open( conf_json_file, 'r' ) as json_conf:
                fileargs = json.load(json_conf)
                conf_files_read.append(conf_json_file)
            for key,value in fileargs.items():
                confargs[key] = value
        else:
            if ('conf_mustload' in optiondict and optiondict['conf_mustload']) or ('conf_mustload' in cmdlineargs and cmdlineargs['conf_mustload']):
                raise Exception('missing config file:'+conf_json_file)
            else:
                logger.warning('skipped missing config file:%s', conf_json_file)
    if conf_files_read:
        # populate the list of config files with the actual files read in
        optiondict['conf_json']=conf_files_read
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
    for key,value in confargs.items():
        if not key in cmdlineargs:
            # this file value has no associated command line override
            # value not overridden by value on commmand line
            if isinstance( value, str ):
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
    for key,value in cmdlineargs.items():
        # logic to bring in "default/implied optiondict values if key passed is not part of app definition
        if key not in optiondictconfig and key in defaultdictconfig:
            if debug:  print('kv_parse_command_line:key-not-in-optiondictconfig-but-in-defaultoptiondictconfig:', key)
            logger.debug('key-not-in-optiondictconfig-but-in-defaultoptiondictconfig:%s', key)
            # copy over this default into optiondict
            optiondictconfig[key]= defaultdictconfig[key].copy()
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
                if debug:  print('type:', optiondictconfig[key]['type'])
                logger.debug('key:%stype:%s', key,optiondictconfig[key]['type'])
                
            if 'type' not in optiondictconfig[key]:
                # user did not specify the type of this option
                optiondict[key] = value
                if debug: print('type not in optiondictconfig[key]')
                logger.debug('type not in optiondictconfig[key] for key:%s', key)
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
                optiondict[key] = datetime_from_str( value )
            elif optiondictconfig[key]['type'] == 'datetimezone':
                optiondict[key] = datetimezone_from_str( value )
            elif optiondictconfig[key]['type'] == 'inlist':
                # value must be from a predefined list of acceptable values
                if not 'valid' in optiondictconfig[key]:
                    if debug: print('missing optiondictconfig setting [valid] for key:', key)
                    logger.error('missing optiondictconfig setting [valid] for key:%s', key)
                    raise Exception('missing optiondictconfig setting [valid] for key:{}'.format(key))
                if value not in optiondictconfig[key]['valid']:
                    if debug:  print('value:', value, ':not in defined list of valid values:', optiondictconfig[key]['valid'])
                    logger.error('invalid value passed in for [%s]:%s',key,value)
                    logger.error('list of valid values are:%s',  optiondictconfig[key]['valid'])
                    raise Exception('invalid value passed in for [{}]:{}'.format(key,value))
                optiondict[key] = value
            else:
                # user set a type but we don't know what to do with this type
                optiondict[key] = value
                if debug: print('type not known:', type)
                logger.debug('type unknown:%s', type)
        elif raise_error:
            logger.error('unknown command line option:%s', key)
            raise Exception('unknown command line option:{}'.format(key))
        else:
            if debug:  print('kv_parse_command_line:unknown-option:', key)
            logger.warning('unknown option:%s', key)

        # special processing if we are asking for help
        if key in ('help', 'helpall'):
            # user asked for help - display help and then exit
            tblfmt = False
            if value in ('tbl','table','helptbl','fmt'):
                tblfmt = True
            # determine if we are also display the additional options
            defaultoptions={}
            if key == 'helpall':
                defaultoptions = defaultdictconfig
            kv_parse_command_line_display( optiondictconfig, defaultoptions, tblfmt=tblfmt, debug=False )
            sys.exit()
    # test for required fields being populated
    missingoption = []
    for key in optiondictconfig:
        if 'required' in optiondictconfig[key]:
            if optiondictconfig[key]['required'] and optiondict[key] == None:
                # required field but is populated with None
                missingoption.append('%s:required field not populated' % key)
                optiondictconfig[key]['error'] = 'required value not populated'
    
    # raise error if we should
    if missingoption:
        kv_parse_command_line_display( optiondictconfig, debug=False )
        errmsg = 'System exitted - missing required option(s):\n    ' + '\n    '.join(missingoption)
        # print('\n'.join(missingoption))
        if debug:
            print('-'*80)
            print(errmsg)
            print('')
        logger.error(errmsg)
        raise Exception(errmsg)
        # sys.exit(1)
    
    # debug when we are done
    if debug:  print('kv_parse_command_line:optiondict:', optiondict)
    logger.debug('optiondict:%s', optiondict)

    # check to see if we want to dump the optiondict out to a file
    if 'dumpconfigfile' in optiondict and optiondict['dumpconfigfile']:
        dump_dict_to_json_file( optiondict['dumpconfigfile'], optiondict )
        
    # check to see if they set the dumpconfig setting if so display and exit
    if 'dumpconfig' in optiondict and optiondict['dumpconfig']:
        print('kv_parse_command_line:Dump configuration requested:')
        for (key,val) in optiondict.items():
            print('{}{}:{}'.format(key, '.'*(30-len(key)), val))
        sys.exit()
        
    # return what we created
    return optiondict

# update the value of a two level deep key if it is not already set
def set_when_not_set( dict, key1, key2, value ):
    if key1 in dict:
        if key2 not in dict[key1]:
            dict[key1][key2] = value
            return True
    return False

# display the optiondictconfig information in human readable format
def kv_parse_command_line_display( optiondictconfig, defaultoptions={}, optiondict={}, tblfmt=False, debug=False ):
    # set the sortorder for a known set of keys
    set_when_not_set( optiondictconfig, 'AppVersion', 'sortorder', 1 )
    set_when_not_set( optiondictconfig, 'debug', 'sortorder', 9997 )
    set_when_not_set( optiondictconfig, 'help', 'sortorder', 9998 )
    set_when_not_set( optiondictconfig, 'helpall', 'sortorder', 9999 )

    # predefined number ranges by type
    nextcounter = {
        'None'         : 2,
        'dir'          : 100,
        'int'          : 200,
        'float'        : 300, 
        'bool'         : 400,
        'date'         : 500,
        'datetimezone' : 600,
        'liststr'      : 700,
        'inlist'       : 800,
    }        

    opt2sort = []
    
    # step through the optional keys
    for opt in sorted( optiondictconfig.keys() ):
        if 'type' in optiondictconfig[opt]:
            # type set - use it
            typeupdate = optiondictconfig[opt]['type']
        else:
            # type not set - make it 'None'
            typeupdate = 'None'
        
        if set_when_not_set( optiondictconfig, opt, 'sortorder', nextcounter[typeupdate] ):
            # we updated the sort order for this record - so we must update the counter
            nextcounter[typeupdate] += 1

        # now build sort string
        opt2sort.append([optiondictconfig[opt]['sortorder'], opt])

    # add in the default options if we have them populated
    if defaultoptions:
        sortcnt = 9996
        opt = '-----'
        optiondictconfig[opt] = { 'value' : opt, 'description' : opt, 'type' : opt }
        opt2sort.append([sortcnt, opt])
        sortcnt = 10000
        
        for opt in defaultoptions.keys():
            if opt not in optiondictconfig:
                if opt == 'help':
                    opt2sort.append( [9998, opt] )
                elif opt == 'helpall':
                    opt2sort.append( [9999, opt] )
                else:
                    opt2sort.append( [sortcnt, opt] )
                optiondictconfig[opt] = defaultoptions[opt]
                sortcnt += 1


    # header if we are doing table output
    if tblfmt:
        print('| option | type | value | description |')
        print('| ------ | ---- | ----- | ----------- |')

    # define the string format for each cell in the table
    tblFmt = ' {} |'
    
    # step through the sorted list and display things
    for row in sorted(opt2sort):
        opt = row[1]
        if opt in optiondict:
            optiondictconfig[opt]['value'] = optiondict[opt]

        # output style
        if tblfmt:
            # user wanted to output in table format - each line with no <newline>
            print('| {} |'.format(opt), end ="")
            # output the type - may not be populated
            fld = 'type'
            fldout = ''
            if fld in optiondictconfig[opt]:
                fldout = optiondictconfig[opt][fld]
            print(tblFmt.format(fldout), end="")
            # output the value - may not be populated
            fld = 'value'
            fldout = ''
            if fld in optiondictconfig[opt]:
                fldout = optiondictconfig[opt][fld]
            if opt in optiondict and fld in optiondict[opt]:
                fldout = optiondict[opt][fld]
            print(tblFmt.format(fldout), end="")
            # output the type - may not be populated
            fld = 'description'
            fldout = ''
            if fld in optiondictconfig[opt]:
                fldout = optiondictconfig[opt][fld]
            # add in valid, error values if they exist
            for fld in ('valid','error'):
                if fld in optiondictconfig[opt]:
                    if fldout:
                        fldout += '<br>'
                    fldout += 'valid:{}'.format(optiondictconfig[opt][fld])
            # output this field - but this time with a <newline>
            print(tblFmt.format(fldout))
        else:
            # linear output 
            if 'type' in optiondictconfig[opt]:
                print('option.:', opt, ' (type:',optiondictconfig[opt]['type'], ')' )
            else:
                print('option.:', opt)

                for fld in ('value','required','description', 'valid', 'error'):
                    if fld in optiondictconfig[opt]:
                        print('  ' + fld + '.'*(12-len(fld)) + ':', optiondictconfig[opt][fld])
        

# define the filename used to create log files
# that are based on the "day" the program starts running
# generally used for short running tools
# not used with tools that start and stay running
def filename_log_day_of_month( filename, ext_override=None, path_override=None ):
    file_path, base_filename, file_ext = filename_split( filename, path_blank=True )
    if ext_override:
            file_ext=ext_override
    if file_ext[:1] != '.':
        file_ext = '.' + file_ext
    if path_override:
        file_path = path_override
    day_filename = '{}{:02d}'. format(base_filename, datetime.datetime.today().day)
    logfilename = os.path.join(file_path, day_filename+file_ext)
    if os.path.exists(logfilename):
        if os.path.getmtime(logfilename) < (datetime.datetime.today() - datetime.timedelta(days=1)).timestamp():
            # remove the file if it exists but has not been modified within the past 24 hours
            remove_filename(logfilename)
    return logfilename

# return the filename that is max or min for a given query (UT)
# default is to return the MIN filematch
def filename_maxmin( file_glob, reverse=False ):
    # pull the list of files
    filelist = glob.glob( file_glob )
    # debugging
    logger.debug('filelist:%s', filelist)
    # if we got no files - return none
    if not filelist:
        logger.debug('return none')
        return None
    logger.debug('file:%s', sorted(filelist, reverse=reverse )[0])
    # sort this list - and return the desired value
    return sorted(filelist, reverse=reverse )[0]

# create a filename from part of a filename
#   pull apart the filenaem passed in (if passed in) and then fill in the various file parts based
#   on the other attributes passed in
def filename_create( filename=None, filename_path=None, filename_base=None, filename_ext=None, path_blank=False, filename_base_append=None, filename_base_prepend=None, use_input_filename=None, filename_unique=None ):
    # pull apart the filename passed in
    if filename:
        file_path, base_filename, file_ext = filename_split( filename, path_blank=path_blank )
    else:
        file_path = base_filename = file_ext = ''
    if filename_ext:
            file_ext=filename_ext
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
    return os.path.normpath(os.path.join(file_path, base_filename+file_ext))

# split up a filename into parts (path, basename, extension) (UT)
def filename_split( filename, path_blank=False ):
    filename2, file_ext = os.path.splitext(filename)
    base_filename = os.path.basename(filename2)
    if path_blank:
        file_path = os.path.dirname(filename2)
    else:
        file_path     = os.path.normpath(os.path.dirname(filename2))
    return (file_path, base_filename, file_ext)


# function to get back a full list of broken up file path
def filename_splitall(path):
    allparts = []
    while 1:
        parts = os.path.split(path)
        if parts[0] == path:  # sentinel for absolute paths
            allparts.insert(0, parts[0])
            break
        elif parts[1] == path: # sentinel for relative paths
            allparts.insert(0, parts[1])
            break
        else:
            path = parts[0]
            allparts.insert(0, parts[1])
    return allparts



# create a list of filenames given a name, a list of names, file glob,
# list of include files in a file, list of exclue files in a file
def filename_list( filename=None, filenamelist=None, fileglob=None, strippath=False, includelist_filename=None, excludefilenamelist=None, excludelist_filename=None  ):
    # local variable
    flist=[]
    xlist=[]
    # read list from files provide
    if includelist_filename:
        flist = read_list_from_file_lines(includelist_filename,trim=True)
    if excludelist_filename:
        xlist = read_list_from_file_lines(excludelist_filename,trim=True)
    if excludefilenamelist:
        xlist.extend(excludefilenamelist)
    # read list from records provided
    if fileglob:
        flist.extend(glob.glob(fileglob))
    if filenamelist:
        flist.extend(filenamelist)
    if filename:
        flist.append(filename)

    # remove records if exclude definitions provided
    if xlist:
        for excludefile in xlist:
            if excludefile in flist:
                flist.remove(excludefile)

    # strip path from filename if flag is set
    if strippath:
        for ndx in range(len(flist)):
            flist[ndx] = os.path.basename(flist[ndx])

    # create the unique list of filenames and return them
    return sorted(list(set(flist)))



# create a full filename and optionally validate directory exists and is writeabile (UT)
def filename_proper( filename_full, dir=None, create_dir=False, write_check=False ):
    filename = os.path.basename( filename_full )
    if not dir:
        dir = os.path.dirname( filename_full )

    # if there is no directory then make it the current directory
    if not dir:
        dir = './'

    # wondering if we need to extract directory and compare if set (future feature) - and if they are different - what action should we take?

    # check the directory and determine if we need it to be created
    if not os.path.exists( dir ):
        # directory needs to be created
        if create_dir:
            # needs to be created and we have enabled this option
            try:
                os.makedirs( dir )
            except Exception as e:
                if debug: print('kvutil:filename_proper:makedirs:%s' % e)
                logger.error('makedirs:%s' % e)
                raise Exception('kvutil:filename_proper:makedirs:{}'.format(e))
        else:
            # needs to be created - option not enabled - raise an error
            if debug: print('kvutil:filename_proper:directory does not exist:%s' % dir )
            logger.error('directory does not exist:%s', dir )
            raise Exception('kvutil:filename_proper:directory does not exist:{}'.format(dir))

    # check to see if the directory is writeable if the flag is set
    if write_check:
        if not os.access( dir, os.W_OK ):
            if debug: print('kvutil:filename_proper:directory is not writeable:%s' % dir )
            logger.error('directory is not writeable:%s', dir )
            raise Exception('kvutil:filename_proper:directory is not writeable:{}'.format(dir) )
    
    # build a full filename
    full_filename = os.path.join( dir, filename )
    
    # return the calculated filename
    return os.path.normpath(full_filename)



# create a unique filename
def filename_unique( filename=None, filename_href={} ):
    # check input
    if isinstance( filename, dict):
        filename_href = filename
        filename = None
        
    # default options for the filename_href
    default_options = {
	'file_ext'          : '.html',         # file extension
	'full_filename'     : '',
	'file_path'         : './',            # path to where to put the file
	'filename'          : '',
	'tmp_file_path'     : '',
	'base_filename'     : 'tmpfile',       # basefilename
	'ov_ext'            : '.bak',          # overwritten saved file extension
	'uniqtype'          : 'cnt',           # defines how we make this filename uniq
	'cntfmt'            : 'v%02d',         # format string for converting count
	'datefmt'           : '-%Y%m%d',       # format string for converting date
	'maxcnt'            : 100,             # maximum count to search for unique filename
	'forceuniq'         : False,           # do not force unique filename creation
	'overwrite'         : False,           # 1=overwrite an existing file
        'create_dir'        : False,           # if true - we will create the directory specified if it does not exist
        'write_check'       : True,            # validate we can write in the specified directory
	'verbose_uf'        : 0,
    }
    # list of required fields to be populated
    required_values = ['file_ext', 'file_path', 'base_filename', 'uniqtype']
    
    # list of valid values for inputs
    validate_values = {
        'uniqtype' : ['cnt', 'datecnt']
    }
    # force the value of this field if the value is blank
    force_if_blank = {
        'file_path' : './',
    }
                    

    # bring in the values that were passed in
    for key in default_options:
        if key in filename_href:
            default_options[key] = filename_href[key]

    # if filename is provided split it up
    if filename:
        default_options['file_path'], default_options['base_filename'], default_options['file_ext'] = filename_split(filename)
    else:
        # parse up the full_filename if passed in
        if default_options['full_filename']:
            default_options['file_path'], default_options['base_filename'], default_options['file_ext'] = filename_split(default_options['full_filename'])
        elif default_options['filename']:
            default_options['file_path'], default_options['base_filename'], default_options['file_ext'] = filename_split(default_options['filename'])
        else:
            # make sure base_filename is only a filename
            default_options['base_filename'] = os.path.basename(default_options['base_filename'])
            default_options['file_path']     = os.path.dirname(default_options['file_path'])
        

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
        if debug:  print('kvutil:filename_unique:missing values for:', ','.join(field_issues))
        logger.error('missing values for:%s', ','.join(field_issues))
        raise Exception('kvutil:filename_unique:missing values for:', ','.join(field_issues))
    
    # check that we have valid values
    for key in validate_values:
        if not default_options[key] in validate_values[key]:
            field_issues.append(key)

    # check to see if we have and field issues
    if field_issues:
        if debug: print('kvutil:filename_unique:invalid values for:', ','.join(field_issues))
        logger.error('invalid values for:%s', ','.join(field_issues))
        raise Exception('kvutil:filename_unique:invalid values for:', ','.join(field_issues))

    # create a filename if it does not exist
    default_options['filename'] = os.path.normpath(os.path.join(default_options['base_filename'] + default_options['file_ext']))

    # check the directory to see if it exists
    default_options['file_path'] = filename_proper( default_options['file_path'], create_dir=default_options['create_dir'], write_check=default_options['write_check'] )

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
        filename = default_options['base_filename'] + date_file + (default_options['cntfmt'] % unique_counter) + default_options['file_ext']
    else:
        # not a unique - try the filename passed infirst
        filename = default_options['filename']

    # debugging
    #print('file_unique:filename:', filename)
    #print('file_unique:default_options:', default_options)

    # take action if we are not going to overwrite the filename
    if not default_options['overwrite']:

        # look for the filename that works
        while( os.path.exists( os.path.join( default_options['file_path'], filename ) ) and unique_counter < default_options['maxcnt'] ):
            # create a new filename
            filename = default_options['base_filename'] + date_file + (default_options['cntfmt'] % unique_counter) + default_options['file_ext']
            # increment the counter
            unique_counter += 1

        # test to see if we exceeded the max count and if so error out.
        if unique_counter >= default_options['maxcnt']:
            if debug: print('kvutil:filename_unique:reached maximum count and not unique filename:', filename)
            logger.error('reached maximum count and not unique filename:%d:%s', unique_counter, filename)
            raise Exception('kvutil:filename_unique:reached maximum count and not unique filename:', filename)

    # debugging
    #print('file_unique:filename:final:', filename)

    # return the final filename
    return filename_proper( filename, dir=default_options['file_path'])
#, \
#           filename_proper( filename +  default_options['ov_ext'], dir=default_options['file_path'])
                                                                                          

# cloudpath - create an absolute path to a folder that is local for cloud drive
def cloudpath( filepath, filename='' ):
    userdir = ''
    if filepath == None:
        filepath = ''
    if filename == None:
        filename = ''
    # determine if the path is a cloud path
    for cloudprovider in ('Box Sync','Dropbox','OneDrive'):
        index = filepath.find(cloudprovider)
        if index != -1:
            filepath = filepath[index:]
            userdir = os.path.expanduser('~')
            break

    return os.path.abspath(os.path.join(userdir,filepath,filename)) 
            

# read a text file into a string (UT)
def slurp( filename ):
    with open( filename, 'r') as t:
        return t.read()

# read in a file and create a list of each populated line (UT)
def read_list_from_file_lines( filename, stripblank=False, trim=False, encoding=None ):
    # read in the file as a list of strings
    if encoding:
        with open( filename, 'r', encoding=encoding) as t:
            filelist = t.readlines()
    else:
        with open( filename, 'r') as t:
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
def remove_filename(filename,calledfrom='',debug=False,maxretry=20):
    logger.debug('remove:%s:calledfrom:%s:maxretry:%d',filename,calledfrom,maxretry)
    cnt=0
    if calledfrom:  calledfrom += ':'
    while os.path.exists(filename):
        cnt += 1
        if debug: print(calledfrom, filename, ':exists:try to remove:cnt:', cnt)
        logger.debug('%s:%s:exists:try to remove:cnt:%d', calledfrom, filename, cnt)
        try:
            os.remove(filename) # try to remove it directly
            logger.debug('%s:%s:removed on count:%d',calledfrom, filename, cnt)
        except Exception as e:
            if debug: print(calledfrom, 'errno:', e.errno, ':ENOENT:', errno.ENOENT)
            logger.debug('%s:errno:%d:ENOENT:%d', calledfrom, e.errno, errno.ENOENT)
            if e.errno == errno.ENOENT: # file doesn't exist
                return
            if debug: print(calledfrom, filename,':', str(e))
            if cnt > maxretry:
                if debug: print(calledfrom, filename, ':raise error - exceed maxretry attempts:', maxretry)
                logger.error('%s:%s:exceeded maxretry attempts:%d:raise error', calledfrom, filename, maxretry)
                raise e
        except WinError as f:
            if debug: print('catch WinError:', str(f))
            logger.warning('catch WinError:%s', str(f))

# utility used to remove a folder - in windows sometimes we have a delay
# in releasing the filehandle - this routine will loop a few times giving
# time for the OS to release the blocking issue and then delete
#
# optional input:
#    calledfrom - string used to display - usually the name of module.function()
#    debug - bool defines if we display duggging print statements
#    maxretry - int - number of times we try to delete and then give up (default: 20)
#
def remove_dir(dirname,calledfrom='',debug=False,maxretry=20):
    cnt=0
    if calledfrom:  calledfrom += ':'
    while os.path.exists(dirname):
        cnt += 1
        if debug: print(calledfrom, dirname, ':exists:try to remove:cnt:', cnt)
        try:
            os.rmdir(dirname) # try to remove it directly
#        except OSError as e: # originally just checked for OSError - we now check for all exceptions`
        except Exception as e:
            if debug: print(calledfrom, 'errno:', e.errno, ':ENOENT:', errno.ENOENT)
            logger.debug('%s:errno:%s:ENOENT:%s', calledfrom, e.errno, errno.ENOENT)
            if e.errno == errno.ENOENT: # file doesn't exist
                return
            if debug: print(calledfrom, dirname,':', str(e))
            logger.debug('%s:%s:%s', calledfrom, dirname, str(e))
            if cnt > maxretry:
                if debug: print(calledfrom, dirname, ':raise error - exceed maxretry attempts:', maxretry)
                logger.error('%s:%s:maxretry attempts:%d', calledfrom, dirname, maxretry)
                raise e
        except WinError as f:
            if debug: print('catch WinError:', str(f))
            logger.warning('catch WinError:%s', str(f))


# extract out a datetime value from a string if possible
# formats currently supported:
#  m/d/y
#  m-d-y
#  YYYY-MM-DD
#  YYYYMMDD
#
def datetime_from_str( value, skipblank=False ):
    import re
    datefmts = (
        ( re.compile('\d{1,2}\/\d{1,2}\/\d\d$'), '%m/%d/%y' ),
        ( re.compile('\d{1,2}\/\d{1,2}\/\d\d\d\d$'), '%m/%d/%Y' ),
        ( re.compile('\d{1,2}-\d{1,2}-\d\d$'), '%m-%d-%y' ),
        ( re.compile('\d{1,2}-\d{1,2}-\d\d\d\d$'), '%m-%d-%Y' ),
        ( re.compile('\d{4}-\d{1,2}-\d{1,2}$'), '%Y-%m-%d' ),
        ( re.compile('^\d{8}$'), '%Y%m%d' ),
    )

    if skipblank and not value:
        return value
    
    for (redate, datefmt) in datefmts:
        if redate.match(value):
            return datetime.datetime.strptime(value, datefmt)

    raise Exception('Unable to convert to date time:{}'.format(value))


# extract out a datetime value with timezone from a string if possible
# formats currently supported:
#     YYYY-MM-DD HH:MM:SS[+-]HHHH
#     YYYY-MM-DDTHH:MM:SS[+-]HHHH
#     YYYY-MM-DD HH:MM:SS.mmmm[+-]HHHH
#     YYYY-MM-DDTHH:MM:SS.mmmm[+-]HHHH
#
#     YYYY-MM-DD HH:MM:SS[+-]HH:HH
#     YYYY-MM-DDTHH:MM:SS[+-]HH:HH
#     YYYY-MM-DD HH:MM:SS.mmmm[+-]HH:HH
#     YYYY-MM-DDTHH:MM:SS.mmmm[+-]HH:HH
#
def datetimezone_from_str( value, skipblank=False ):
    import re
    datefmtscleanup = (
        ( re.compile('(.*[+-])(\d{2}):(\d{2})$'), 'remove-colon-2' ),
    )
    datefmts = (
        ( re.compile('\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}\.\d+[+-]\d{4}$'), '%Y-%m-%dT%H:%M:%S.%f%z' ),
        ( re.compile('\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}\.\d+[+-]\d{4}$'), '%Y-%m-%d %H:%M:%S.%f%z' ),
        ( re.compile('\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}[+-]\d{4}$'), '%Y-%m-%dT%H:%M:%S%z' ),
        ( re.compile('\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}[+-]\d{4}$'), '%Y-%m-%d %H:%M:%S%z' ),
    )

    if skipblank and not value:
        return value

    # see if we need to change the format of the data we got in
    for (redate, action) in datefmtscleanup:
        if redate.match(value):
            m = redate.match(value)
            if action == 'remove-colon-2':
                value = m.group(1) + m.group(2) + m.group(3)


    # convert date into date/time/zone
    for (redate, datefmt) in datefmts:
        if redate.match(value):
            return datetime.datetime.strptime(value, datefmt)

    # error out because we could not convert
    raise Exception('Unable to convert to date time:{}'.format(value))



# return the function name of the function that called this
def functionName(callBackNumber=1):
    return sys._getframe(callBackNumber).f_code.co_name



# create the starting logger header that we want to show the separation
# between runs - this utility is just to enable logging standardization.
#
# In your program put:  kvutil.loggingAppStart( logger, optiondict, kvutil.scriptinfo()['name'] )
#
def loggingAppStart(logger,optiondict,pgm=None):
    logger.info('-----------------------------------------------------')
    if pgm:
        logger.info('%s:AppVersion:v%s', pgm, optiondict['AppVersion'])
    else:
        logger.info('AppVersion:v%s', optiondict['AppVersion'])


def scriptinfo():
    '''
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
    '''

    import os, sys, inspect
    #---------------------------------------------------------------------------
    # scan through call stack for caller information
    #---------------------------------------------------------------------------
    trc=''
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

    scr_dict ={"name": trc,
               "source": trc,
               "dir": scriptdir}
    return scr_dict


#utility used to dump a dictionary to a file in json format
def dump_dict_to_json_file( filename, optiondict ):
    import json
    with open( filename, 'w' ) as json_out:
        json.dump( optiondict, json_out, indent=4 )
        
# eof
