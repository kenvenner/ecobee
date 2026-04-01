"""
@author:   Ken Venner
@contact:  ken@venerllc.com
@version: 1.16

Library of tools used to read and write CSV files
"""

import csv
import kvmatch

# logging
import logging

logger = logging.getLogger(__name__)

# version number
AppVersion = "1.16"

################################ HELPER  #############################################


def max_column_list(csvlist: list[dict]) -> list:
    """
    Read all records in the list of dictionaries and get the unique set of keys
    across all records.
    This assumes that the keys in each record are not the same.
    Used to assure that we do not drop any data when creating an output file - we
    know the max columns to generate

    Inputs
        csvlist - list[dict] - list of dictionaries

    Returns
        fieldlist - list of columns that must be created to not lose any data
    """

    # test inputs
    if not isinstance(csvlist, list):
        raise TypeError(f"csvlist must be [list] but is: {type(csvlist)}")
    if not isinstance(csvlist[0], dict):
        raise TypeError(f"csvlist[0] must be [dict] but is: {type(csvlist[0])}")

    fieldlist = []
    for rec in csvlist:
        for key in rec.keys():
            if key not in fieldlist:
                fieldlist.append(key)
    return fieldlist


################################ WRITE  #############################################


def writelist2csv(
    csvfile: str,
    csvlist: list[dict],
    csvfields: list[str] | None = None,
    mode: str = "w",
    header: bool = True,
    encoding: str = "windows-1252",
    maxcolumns: bool = False,
    col_aref: list | None = None,
    debug: bool = False,
) -> None:
    """
    write out a list of dicts => records into a CSV
    you can output a defined set of columns (subset or superset) by defining csvfields or col_aref
    if csvfields/col_aref are not defined, we get the list of columns to generate based on the keys
    of first line of the list of records to process.
    A header is placed in this file if header is enabled, which it is by default.

    Inputs:
        csvfile - str - filename/path of the CSV file to generate
        csvlist - list[dict] list of records to send into this file
        csvfields - list[str] the column names to put in the header, if None we take the keys from the first record
        mode - create the file or append to the file (default: create) if you want ot append - send in "a"
        header - bool - when true, we create a header as the first row, when false we do not generate the header line
        encoding - str - the character set used to generate the file
        maxcolumns - bool - when set to true - we go through all records and find the full set of keys across all records and output using that
        col_aref - list[str] - if you don't set csvfields, the other way to speciffy the column headers for this file
        debug - bool - wehn true, display messages while running

    """
    # if there is no data to output - just return
    if not csvlist:
        return

    # set values if not populated
    if csvfields is None:
        csvfields = list()
    if col_aref is None:
        col_aref = list()

    # test inputs
    if not csvfile:
        raise ValueError("csvfile must be populated")
    if not isinstance(csvlist, list):
        raise TypeError(f"csvlist must be [list] but is: {type(csvlist)}")
    if not isinstance(csvlist[0], dict):
        raise TypeError(f"csvlist[0] must be [dict] but is: {type(csvlist[0])}")
    if csvfields and not isinstance(csvfields, list):
        raise TypeError(f"csvfields must be [list] but is: {type(csvfields)}")
    if col_aref and not isinstance(col_aref, list):
        raise TypeError(f"col_aref must be [list] but is: {type(col_aref)}")
    if mode and mode not in ("a", "w"):
        raise ValueError(f"mode can only be [a, w] but is: {mode}")

    # calculate the header when flag is enabled
    if maxcolumns:
        csvfields = max_column_list(csvlist)

    # check to see if we passed in col_ref
    if col_aref and not csvfields:
        # and we set the fields we wanted to output
        csvfields = col_aref

    # get the keys from the dictionary keys in the first value itself
    if not csvfields:
        csvfields = list(csvlist[0].keys())

    # debugging:
    if debug:
        print(f"{csvfields=}")

    # open the output file and write out the dictionary
    with open(csvfile, mode=mode, newline="", encoding=encoding) as csv_file:
        writer = csv.DictWriter(
            csv_file, fieldnames=csvfields, extrasaction="ignore"
        )

        if header:
            writer.writeheader()
        for row in csvlist:
            writer.writerow(row)


def writedict2csv(
    csvfile,
    csvdict,
    csvfields=None,
    mode="w",
    header=True,
    encoding="windows-1252",
    maxcolumns=False,
    col_aref=None,
    debug=False,
) -> None:
    """
    write out a list of dicts => records into a CSV
    you can output a defined set of columns (subset or superset) by defining csvfields or col_aref
    if csvfields/col_aref are not defined, we get the list of columns to generate based on the keys
    of first line of the list of records to process.
    A header is placed in this file if header is enabled, which it is by default.

    Inputs:
        csvfile - str - filename/path of the CSV file to generate
        csvdict - dict[dict] dict with values of dict send into this file
        csvfields - list[str] the column names to put in the header, if None we take the keys from the first record
        mode - create the file or append to the file (default: create) if you want ot append - send in "a"
        header - bool - when true, we create a header as the first row, when false we do not generate the header line
        encoding - str - the character set used to generate the file
        maxcolumns - bool - when true, we go through all values and create the max csvfields to create
        col_aref - list[str] - the other way to speciffy the column headers for this file
        debug - bool - wehn true, display messages while running

    """
    # if there is no data to output - just return
    if not csvdict:
        return

    # test inputs
    if not csvfile:
        raise ValueError("csvfile must be populated")
    if not isinstance(csvdict, dict):
        raise TypeError(f"csvdict must be [dict] but is: {type(csvdict)}")
    if csvfields and not isinstance(csvfields, list):
        raise TypeError(f"csvfields must be [list] but is: {type(csvfields)}")
    if col_aref and not isinstance(col_aref, list):
        raise TypeError(f"col_aref must be [list] but is: {type(col_aref)}")
    if mode and mode not in ("a", "w"):
        raise ValueError(f"mode can only be [a, w] but is: {mode}")

    # if we are maxcolumns - then we need to calculate this
    if maxcolumns:
        csvfields = max_column_list(list(csvdict.values()))

    # check to see if we passed in col_ref
    if col_aref and not csvfields:
        # and we set the fields we wanted to output
        csvfields = col_aref

    # get the keys from the dictionary keys in the first value itself
    if not csvfields:
        csvfields = list(csvdict[list(csvdict.keys())[0]].keys())

    # open the output file and write out the dictionary
    with open(csvfile, mode=mode, newline="", encoding=encoding) as csv_file:
        writer = csv.DictWriter(
            csv_file, fieldnames=csvfields, extrasaction="ignore"
        )

        if header:
            writer.writeheader()
        for row in csvdict.values():
            writer.writerow(row)


################################ READ #############################################

## LISTS ##


def readcsv2list_with_header(
    csvfile: str,
    headerlc: bool = False,
    encoding: str = "windows-1252",
    debug: bool = False,
) -> tuple[list[dict], list[str]]:
    """
    read in the CSV and create a dictionary to the records
    assumes the first line of the CSV file is the header/defintion of the CSV

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        headerlc: bool - when enabled, force the header values to lower case, otherwise use the string as defined in the file
        encoding: str - string that defines character type to read in with
        debug: bool - when enabled, display messages while processing

    Returns
        results - list[dict] list of records with dictionary of key/value settings
        header - list[str]  list of header values read in
    """

    results = []
    with open(csvfile, mode="r", encoding=encoding) as csv_file:
        reader = csv.reader(csv_file)
        header = reader.__next__()
        if debug:
            print("header-before:", header)
        logger.debug("header-before:%s", header)
        if headerlc:
            header = [x.lower() for x in header]
            if debug:
                print("header-after:", header)
            logger.debug("header-after:%s", header)
        for row in reader:
            rowdict = dict(zip(header, row))
            # create/update the dictionary
            results.append(rowdict)
    # return the results
    return results, header


def readcsv2list(
    csvfile: str,
    headerlc: bool = False,
    encoding: str = "windows-1252",
    debug: bool = False,
) -> tuple[list[dict], list[str]]:
    """
    read in the CSV and create a dictionary to the records
    assumes the first line of the CSV file is the header/defintion of the CSV

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        headerlc: bool - when enabled, force the header values to lower case, otherwise use the string as defined in the file
        encoding: str - string that defines character type to read in with
        debug: bool - when enabled, display messages while processing

    Returns
        results - list[dict] list of records with dictionary of key/value settings

    """

    results, header = readcsv2list_with_header(
        csvfile, headerlc, encoding, debug
    )
    return results


def readcsv2list_with_noheader(
    csvfile: str,
    header: list,
    encoding: str = "windows-1252",
    debug: bool = False,
) -> tuple[list[dict], list[str]]:
    """
    read in the CSV and create a dictionary to the records
    no header in this file so we pass in the header defintion

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        header: list of header column names
        encoding: str - string that defines character type to read in with
        debug: bool - when enabled, display messages while processing

    Returns
        results - list[dict] list of records with dictionary of key/value settings
        header - list[str]  list of header values read in
    """

    # test inputs
    if not header:
        raise ValueError("header must be populated and is not")
    if not isinstance(header, list):
        raise ValueError(f"header must be list but is: {type(header)}")

    results = []
    with open(csvfile, mode="r", encoding=encoding) as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            rowdict = dict(zip(header, row))
            # create/update the dictionary
            results.append(rowdict)
    # return the results
    return results, header


## DICT ##


def readcsv2dict_with_header(
    csvfile: str,
    dictkeys: list,
    dupkeyfail: bool = False,
    noshowwarning: bool = False,
    headerlc: bool = False,
    encoding: str = "windows-1252",
    debug: bool = False,
) -> tuple[dict, list, int]:
    """
    read in the CSV and create a dictionary to the records, and create a dict unique on business key
    assumes the first line of the CSV file is the header/defintion of the CSV

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        dictkeys: list of keys that make up the unqiue business key and the key to the resulting dictionary
        dupkeyfail: bool - when true, if we find recrods that are duplicates we raise an error
        noshowwarning: bool - when false, if we find records that are duplicates we print out a message about this
        headerlc: bool - when enabled, force the header values to lower case, otherwise use the string as defined in the file
        encoding: str - string that defines character type to read in with
        debug: bool - when enabled, display messages while processing

    Returns
        results - dict of records on unique buiness key with value of a dict that is the record
        header - list[str]  list of header values read in
        dupcount - number of records encountered that were duplicate business keys
    """

    # debugging
    if debug:
        print("with_header")
        print(f"{csvfile=}")
        print(f"{type(csvfile)=}")

        print(f"{dictkeys=}")
        print(f"{type(dictkeys)=}")

    # test inputs
    if not dictkeys:
        raise ValueError("dictkeys must be populated and is not")
    if not isinstance(dictkeys, list):
        raise TypeError("dictkeys must be a list but is: {type(dictkeys)}")

    # read the records into a list
    results_list, header = readcsv2list_with_header(
        csvfile, headerlc=headerlc, encoding=encoding, debug=debug
    )

    # push the keys to lower if we set that flag on
    if headerlc:
        dictkeys = [x.lower() for x in dictkeys]

    # convert list to dict
    results = {}
    dupkeys = []
    dupcount = 0
    for rowdict in results_list:
        reckey = kvmatch.build_multifield_key(rowdict, dictkeys)
        # do we fail if we see the same key multiple times?
        if reckey in results:
            dupcount += 1
            # capture this key
            dupkeys.append(reckey)
        # create/update the dictionary
        results[reckey] = rowdict
    # fail if we found dupkeys
    if dupkeys:
        # log this issue
        logger.warning(
            "readcsv2dict:v%s:file:%s:duplicate key failure:keys:%s",
            AppVersion,
            csvfile,
            ",".join(dupkeys),
        )
        # display message if the user wants this displayed
        if not noshowwarning:
            print("readcsv2dict:duplicate key failure:", ",".join(dupkeys))
        # if we want to fail on dupkey then do so
        if dupkeyfail:
            raise ValueError("Duplicate key failure")

    # return the results
    return results, header, dupcount


def readcsv2dict(
    csvfile: str,
    dictkeys: list,
    dupkeyfail: bool = False,
    noshowwarning: bool = False,
    headerlc: bool = False,
    encoding: str = "windows-1252",
    debug: bool = False,
) -> dict:
    """
    read in the CSV and create a dictionary to the records, and create a dict unique on business key
    assumes the first line of the CSV file is the header/defintion of the CSV

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        dictkeys: list of keys that make up the unqiue business key and the key to the resulting dictionary
        dupkeyfail: bool - when true, if we find recrods that are duplicates we raise an error
        noshowwarning: bool - when false, if we find records that are duplicates we print out a message about this
        headerlc: bool - when enabled, force the header values to lower case, otherwise use the string as defined in the file
        encoding: str - string that defines character type to read in with
        debug: bool - when enabled, display messages while processing

    Returns
        results - dict of records on unique buiness key with value of a dict that is the record

    """
    if not dictkeys:
        raise ValueError("dictkeys must be populated and is not")
    if not isinstance(dictkeys, list):
        raise TypeError(f"dictkeys must be a list but is: {type(dictkeys)}")

    results, header, dupcount = readcsv2dict_with_header(
        csvfile,
        dictkeys,
        dupkeyfail=dupkeyfail,
        noshowwarning=noshowwarning,
        headerlc=headerlc,
        encoding=encoding,
        debug=debug,
    )
    return results


def readcsv2dict_with_noheader(
    csvfile: str,
    dictkeys: list,
    header: list,
    dupkeyfail: bool = False,
    noshowwarning: bool = False,
    encoding: str = "windows-1252",
    debug: bool = False,
) -> tuple[dict, list, int]:
    """
    read in the CSV and create a dictionary to the records, and create a dict unique on business key
    no header, so we pass it in

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        dictkeys: list of keys that make up the unqiue business key and the key to the resulting dictionary
        header: list of strings defining the header for each column
        dupkeyfail: bool - when true, if we find recrods that are duplicates we raise an error
        noshowwarning: bool - when false, if we find records that are duplicates we print out a message about this
        encoding: str - string that defines character type to read in with
        debug: bool - when enabled, display messages while processing

    Returns
        results - dict of records on unique buiness key with value of a dict that is the record
        header - list[str]  list of header values read in
        dupcount - number of records encountered that were duplicate business keys

    """
    if not dictkeys:
        raise ValueError("dictkeys must be populated and is not")
    if not isinstance(dictkeys, list):
        raise TypeError(f"dictkeys must be a list but is: {type(dictkeys)}")
    if not header:
        raise ValueError("header must be populated and is not")
    if not isinstance(header, list):
        raise TypeError(f"header must be a list but is: {type(header)}")
    bad_dictkeys = [x for x in dictkeys if x not in header]
    if bad_dictkeys:
        raise ValueError(
            f"dictkeys that are not in header: {','.join(bad_dictkeys)}"
        )

    # read the records into a list
    results_list, header = readcsv2list_with_noheader(
        csvfile, header=header, encoding=encoding, debug=debug
    )

    # convert list to dict
    results = {}
    dupkeys = []
    dupcount = 0
    for rowdict in results_list:
        reckey = kvmatch.build_multifield_key(rowdict, dictkeys)
        # do we fail if we see the same key multiple times?
        if reckey in results:
            dupcount += 1
            # capture this key
            dupkeys.append(reckey)
        # create/update the dictionary
        results[reckey] = rowdict
    # fail if we found dupkeys
    if dupkeys:
        # log this issue
        logger.warning(
            "readcsv2dict:v%s:file:%s:duplicate key failure:keys:%s",
            AppVersion,
            csvfile,
            ",".join(dupkeys),
        )
        # display message if the user wants this displayed
        if not noshowwarning:
            print("readcsv2dict:duplicate key failure:", ",".join(dupkeys))
        # if we want to fail on dupkey then do so
        if dupkeyfail:
            raise ValueError("Duplicate key failure")

    # return the results
    return results, header, dupcount


################ FINDHEADER ############################

# coding structure - build one generic (INTERNAL) function that does all the various things
# with passed in variables that are all optional
# and based on variable settings - executes the behavior being asked
#
# then create external functions - with clear passed in parameters that calls this internal function with teh right settings
#

# read in the CSV and create a dictionary to the records
# based on one or more key fields
# assumes the first line of the CSV file is the header/defintion of the CSV
# features to add
#
#   noheader - flag means we pass in the array that is the header
#   col_header - flag means we get the header from the very first line (or start_row) in the file
#   aref_result - flag that tells us to return the array of array (otherwise we pass back the array of dict)
#
#   col_aref - array that defines the columns we read into dict
#
#   start_row - user entered values start at row 1 (not zero)
#
#   unique_column - if enabled, we must validate that our column names are unique
#   ignore_blank_row - if enabled, if the row read has no data - we don't put the data into the extracted list
#   ignore_not_fill - not sure what this one is
#
#   required_fields_populated - checking logic to assure that all required fields have data with optional
#     required_fld_swap - a dict that says if key is not populated - check the value tied to that key to see if it is populated
#


def readcsv2list_findheader(
    csvfile: str,
    req_cols: list,
    xlatdict: dict | None = None,
    optiondict: dict | None = None,
    col_aref: list | None = None,
    debug: bool = False,
) -> tuple[list, list]:
    """
    read in the CSV and create a dictionary to the records
    assumes the first line of the CSV file is the header/defintion of the CSV

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        req_cols: list - list of column names that tells us we have located the header record
        xlatdict: dict - take one or more header column names definitons and map them to teh desired output header name
                         this dict key is the column header we might find, and the value is the header column name we want
        col_aref: list - user defined header defintion - don't use the values we find - use the ones the user passed in
        debug: bool - when enabled, display messages while processing

    Returns
        results - list[dict] list of records with dictionary of key/value settings
        header - list[str]  list of header values read in


    optiondict options:
        start_row - start the search starting at this row number in the file
        no_header - there is no header on this file - so we must pass the header in and use it (similar to readcsv2list_with_noheader)
        aref_result - returns each row as a list not as a dictionary
        save_row - captures the row # in the file that this line was taken from in XLSRow key
        col_header - bool - get header from start_row or first line if start_row is not set

    """
    # set values if not populated
    if xlatdict is None:
        xlatdict = dict()
    if optiondict is None:
        optiondict = dict()
    if col_aref is None:
        col_aref = list()

    # test inputs
    if not req_cols:
        raise ValueError("req_cols must be populated and is not")
    if not isinstance(req_cols, list):
        raise TypeError(f"req_cols must be list but is: {type(req_cols)}")
    if not isinstance(xlatdict, dict):
        raise TypeError(f"xlatdict must be dict but is: {type(xlatdict)}")
    if not isinstance(optiondict, dict):
        raise TypeError(f"optiondict must be dict but is: {type(optiondict)}")
    if not isinstance(col_aref, list):
        raise TypeError(f"col_aref must be list but is: {type(col_aref)}")

    # special tests
    if optiondict.get("no_header") and not col_aref:
        raise ValueError("optiondict[no_header] set and col_aref not populated")

    # local variables
    results = []
    header = None

    # debugging
    if debug:
        print("req_cols:", req_cols)
        print("xlatdict:", xlatdict)
        print("optiondict:", optiondict)
        print("col_aref:", col_aref)
    logger.debug("req_cols:%s", req_cols)
    logger.debug("xlatdict:%s", xlatdict)
    logger.debug("optiondict:%s", optiondict)
    logger.debug("col_aref:%s", col_aref)

    # set flags
    col_header = (
        False  # if true - we take the first row of the file as the header
    )
    no_header = False  # if true - there are no headers read - we either return
    aref_result = False  # if true - we don't return dicts, we return a list
    save_row = False  # if true - then we append/save the XLSRow with the record

    start_row = 0  # if passed in - we start the search at this row (starts at 1 or greater)

    # create the list of misconfigured solutions
    badoptiondict = {
        "startrow": "start_row",
        "startrows": "start_row",
        "start_rows": "start_row",
        "colheaders": "col_header",
        "col_headers": "col_header",
        "noheader": "no_header",
        "noheaders": "no_header",
        "no_headers": "no_header",
        "arefresult": "aref_result",
        "arefresults": "aref_result",
        "aref_results": "aref_result",
        "saverow": "save_row",
        "saverows": "save_row",
        "save_rows": "save_row",
    }

    # check what got passed in
    kvmatch.badoptiondict_check(
        "kvcsv.readcsv2list_findheader", optiondict, badoptiondict, True
    )

    # pull in passed values from optiondict
    if "col_header" in optiondict:
        col_header = optiondict["col_header"]
    if "aref_result" in optiondict:
        aref_result = optiondict["aref_result"]
    if "no_header" in optiondict:
        no_header = optiondict["no_header"]
    if "start_row" in optiondict:
        start_row = (
            optiondict["start_row"] - 1
        )  # because we are not ZERO based in the users mind
    if "save_row" in optiondict:
        save_row = optiondict["save_row"]

    # build object that will be used for record matching
    p = kvmatch.MatchRow(req_cols, xlatdict, optiondict)

    # get the file opened
    csv_file = open(csvfile, mode="r")
    reader = csv.reader(csv_file)

    # ------------------------------- HEADER START ------------------------------

    # define the header for the records being read in
    if no_header:
        # user said we are not to look for the header in this file
        # we need to subtract 1 here because we are going to increment PAST the header
        # in the next section - so if there is no header - we need to start at zero ( -1 + 1 later)
        # row_header = start_row - 1

        # if no col_aref - then we must force this to aref_result
        if not col_aref:
            aref_result = True
            if debug:
                print("no_header:no col_aref:set aref_result to true")
            logger.debug("no_header:no col_aref:set aref_result to true")

        # debug
        if debug:
            print("no_header:start_row:", start_row)
        logger.debug("no_header:start_row:%d", start_row)

    elif col_header:
        # extract the header as the first line in the file
        header = reader.__next__()
        # row_header = 0
        if debug:
            print("col_header:header_1strow:", header)
        logger.debug("col_header:header_1strow:%s", header)
    else:
        # debug
        if debug:
            print("find_header:start_row:", start_row)
        logger.debug("find_header:start_row:%d", start_row)

        # get to the start_row record
        for next_row in range(0, start_row):
            line = reader.__next__()
            if debug:
                print("skipping line:", line)
            logger.debug("skipping line:%s", line)

        # counting row just to provide feedback
        row = start_row

        # now start the search for the header
        for rowdata in reader:
            # increment row
            row += 1

            # have not found the header yet - so look
            if debug:
                print("looking for header at row:", row)
            logger.debug("looking for header at row:%d", row)

            # Search to see if this row is the header
            if p.matchRowList(rowdata, debug=debug) or p.search_exceeded:
                # determine if we found the header
                # set the row_header
                # row_header = row
                # found the header grab the output
                header = p._data_mapped
                # debugging
                if debug:
                    print("header_found:", header)
                logger.debug("header_found:%s", header)
                # break out of the loop
                break
            elif p.search_exceeded:
                # close the file we opened
                csv_file.close()
                # did not find the header - raise error
                raise Exception("header not found")

    # ------------------------------- HEADER END ------------------------------

    # debug
    if debug:
        print("exitted header find loop")
    logger.debug("exitted header find loop")

    # user wants to define/override the column headers rather than read them in
    if col_aref:
        # debugging
        if debug:
            print("copying col_aref into header")
        logger.debug("copying col_aref into header")
        # copy over the values - and determine if we need to fill in more header values
        header = col_aref[:]
        # user defined the row definiton - make sure they passed in enough values to fill the row
        sheetmaxcol = 0
        sheetmincol = 0
        if False:
            if len(col_aref) < sheetmaxcol - sheetmincol:
                # not enough entries - so we add more to the end
                for colcnt in range(
                    1, sheetmaxcol - sheetmincol - len(col_aref) + 1
                ):
                    header.append("")

        # now pass the final information through remapped
        header = p.remappedRow(header)
        # debug
        if debug:
            print("col_aref:header:", header)
        logger.debug("col_aref:header:%s", header)

    # ------------------------------- RECORDS START ------------------------------

    # continue processing this file
    for rowdata in reader:
        if debug:
            print("rowdata:", rowdata)
        logger.debug("rowdata:%s", rowdata)

        # determine what we are returning
        if aref_result:
            # we want to return the data we read
            rowdict = rowdata
            if debug:
                print("saving as array")
            logger.debug("saving as array")

            # optionally add the XLSRow attribute to this dictionary (not here right now
            if save_row:
                rowdict.append(row + 1)
                if debug:
                    print("append row to record")
                logger.debug("append row to record")

        else:
            # we found the header so now build up the records
            rowdict = dict(zip(header, rowdata))
            if debug:
                print("saving as dict")
            logger.debug("saving as dict")

            # optionally add the XLSRow attribute to this dictionary (not here right now
            if save_row:
                rowdict["XLSRow"] = row + 1
                if debug:
                    print("add column XLSRow with row to record")
                logger.debug("add column XLSRow with row to record")

        # add this dictionary to the results
        results.append(rowdict)
        if debug:
            print("append rowdict to results")
        logger.debug("append rowdict to results")

    # ------------------------------- RECORDS END ------------------------------

    # close the file we are reading
    csv_file.close()

    # debugging
    # if debug: print('results:', results)

    # return the results
    return results, header


def readcsv2dict_findheader(
    csvfile: str,
    req_cols: list,
    dictkeys: list | None = None,
    xlatdict: dict | None = None,
    optiondict: dict | None = None,
    col_aref: list | None = None,
    dupkeyfail: bool = False,
    debug: bool = False,
) -> tuple[dict, list, int]:
    """
    read in the CSV and create a dictionary to the records, the list of fields
    passed in dictkeys defines the unique business key that the dictionary we create
    this looks through teh records to find the row that has the values tha tmatch req_cols and
    this row is defined as the header row

    Inputs:
        csvfile: str, - filename/path to the CSV file to be read in
        req_cols: list - list of column names that tells us we have located the header record
        dictkeys: list - list of columns when concatenated defines a unique business key
        xlatdict: dict - take one or more header column names definitons and map them to teh desired output header name
                         this dict key is the column header we might find, and the value is the header column name we want
        col_aref: list - user defined header defintion - don't use the values we find - use the ones the user passed in
        dupkeyfail: bool - when true, if we find duplicate business keys across multiple records we will error out,
                           otherwise we will just ignore duplicate records based on business keys and keep only the last occurence found
        debug: bool - when enabled, display messages while processing

    Returns
        results - list[dict] list of records with dictionary of key/value settings
        header - list[str]  list of header values read in
        dupcount - number of records lost because there as a duplicate business key


    optiondict options:
        start_row
        no_header
        aref_result
        save_row
        col_header

    """

    # debugging
    if debug:
        print("findheader")
        print(f"{csvfile=}")
        print(f"{type(csvfile)=}")

        print(f"{dictkeys=}")
        print(f"{type(dictkeys)=}")

    # user did not set these values - so we must set them
    if xlatdict is None:
        xlatdict = dict()
    if optiondict is None:
        optiondict = dict()
    if col_aref is None:
        col_aref = list()

    # test inputs
    if not dictkeys:
        raise ValueError("dictkeys must be populated and is not")
    if not isinstance(dictkeys, list):
        raise TypeError("dictkeys must be a list but is: {type(dictkeys)}")

    # check processing
    if "no_header" in optiondict and optiondict["no_header"] and not col_aref:
        raise ValueError(
            "invalid setting optiondict[no_header] and no col_aref"
        )
    if "aref_result" in optiondict and optiondict["aref_result"]:
        raise ValueError("invalid setting optiondict[aref_result]")

    # read in the data from the file
    results, header = readcsv2list_findheader(
        csvfile,
        req_cols,
        xlatdict=xlatdict,
        optiondict=optiondict,
        col_aref=col_aref,
        debug=debug,
    )

    # debugging
    if debug:
        print("results from readcsvlist_findheader")
        print(f"{type(results)=}")
        print("dictkeys:", dictkeys)
        print("results:", results)

    # local variables
    dupkeys = []
    dictresults = {}
    dupcount = 0

    # convert to a dictionary based on keys provided
    for rowdict in results:
        if debug:
            print("dictkeys:", dictkeys)
            print("rowdict:", rowdict)

        # get the business key for this record
        reckey = kvmatch.build_multifield_key(rowdict, dictkeys)

        # do we fail if we see the same key multiple times?
        if reckey in dictresults:
            # incremente counter
            dupcount += 1
            # capture this key
            if reckey not in dupkeys:
                dupkeys.append(reckey)

        # create/update the dictionary
        dictresults[reckey] = rowdict

    # fail if we found dupkeys
    if dupkeys:
        print("readcsv2dict:duplicate key failure:", ",".join(dupkeys))
        if dupkeyfail:
            raise ValueError("duplicate key failure:%s", dupkeys)

    if debug:
        print("dictresults:", dictresults)

    # return the results
    return dictresults, header, dupcount


if __name__ == "__main__":
    inputfile = "wine_xlat.csv"
    inputkeys = ["Company", "Wine"]
    outputfile = "wine_xlat_test.csv"

    results = readcsv2dict(inputfile, inputkeys)
    # print( results )
    [print(row, ",") for row in results.values()]

    writedict2csv(outputfile, results)
    print("review file:", outputfile)

# eof
