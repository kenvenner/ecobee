"""
@author:   Ken Venner
@contact:  ken@venerllc.com
@version:  1.16

Library of tools for date time processing used in general by KV

Update:  2024-06-06;kv - added try/except on datetime_from_str

"""

from __future__ import print_function

# import os
import datetime
from dateutil import tz  ## python-dateutil
from dateutil.zoneinfo import get_zonefile_instance
# import sys
# import errno

from typing import Any

# setup the logger
import logging

logger = logging.getLogger(__name__)

# set the module version number
AppVersion = "1.16"
__version__ = "1.16"


def current_timezone_string():
    """
    Get the timezone setting of the current machine and return it back as a string
    """
    now = datetime.datetime.now()
    local_now = now.astimezone()
    local_tz = local_now.tzinfo
    local_tzname = local_tz.tzname(local_now)
    return local_tzname


def datetime2utcdatetime(
    dt: datetime, default_tz: str | None = None, no_tz: bool = False
) -> datetime:
    """
    Take a naive datetime object, convert to a local timezone aware object
    and convert that to a datetime UTC timezone aware object

    Inputs:
        dt - datetime - timezone unaware object (naive)
        default_tz - the timezone to convert into - if not set read from the current system timezone
        no_tz - bool - when set, we should return a naive datetime object set to UTC time, not timezone aware
    Returns:
        the timezone aware utc_datetime equivalent of the dt object

    default_tz:
    Here are some valid values that can be passed to tz.gettz:
        IANA time zone names: These are the standard names used for time zones,
            such as 'America/New_York' or 'Europe/Berlin'.
        TZ environment string: This is a string representation of the time zone,
            which can be in various formats, including GNU TZ style strings.
        Fixed offset timezone: This represents a timezone that is a fixed number
            of seconds behind or ahead of UTC.
        UTC timezone: This represents the UTC time zone, which is the time zone
            for the International Date Line.

    """

    # check to see if this current object has a timezone or is it naive (returns a blank string)
    dt_tz = dt.tzinfo

    # Define the timezone we are converting into, passed in or use current system timezone
    if default_tz is None:
        # if we did not set the timezone in the parameters,
        # then use the timezone of the object or system timezone if the object is naive
        default_tz = tz.gettz() if not dt_tz else dt_tz
    else:
        default_tz = tz.gettz(default_tz)

    # check to see we have a valid timezone
    if not default_tz:
        raise ValueError(
            f"Unable to convert timezone string to timezone: {default_tz}"
        )

    # convert the naive date to localize date
    local_dt = dt.replace(tzinfo=default_tz)

    # convert the local time to UTC time
    utc_datetime = local_dt.astimezone(tz.UTC)

    # strip the timezone from datetime
    if no_tz:
        utc_datetime = utc_datetime.replace(tzinfo=None)

    return utc_datetime


# extract out a datetime value from a string if possible
# formats currently supported:
#  mm-dd-yy
#  mm-dd-yyyy
#  mm/dd/yy
#  mm/dd/yyyy
#  YYYY-MM-DDTHH:MM:SS
#  YYYY-MM-DDTHH:MM:SS.mmmmm
#  YYYY-MM-DD HH:MM:SS
#  YYYY-MM-DD HH:MM
#  YYYY-MM-DD
#  YYYYMMDD
#  DD-MMM-YYYY HH:MM
#  DD-MMM-YYYY
#  MM/DD/YYYY  HH:MM:SS AM/PM - 12/10/2025  11:31:00 PM
#  MM/DD/YYYY HH:MM:SS AM/PM - 12/10/2025 11:31:00 PM
#
#
# and allow a Z to be on the end of this string that we will strip out
#
def datetime_from_str(
    value: Any,
    skipblank: bool = False,
    force_conversion: bool = False,
    disp_msg: bool = False,
):
    """
    Take in Any value (but process only strings) and attempt to convert it to a datetime object if possible

    Inputs:
        value - Any - the value passed in
            if datetime already - just return it
            if a string - try to convert it to datetime
            anything else - return the value or error out based on the flags below
        skipblank - bool - if enabed, and the value is blank/empty - then skip the conversion and do not error out
        force_conversion - bool - if enabled, if we are unasble to convert - fail
        disp_msg - bool -  if enabled, display messages as processing - to enable interactive debugging
    Returns:
        value - the value or the datetime equivalent of the value
    """
    # bring in the regex routine it is needed
    import re

    # list of different dateformats we convert and the proper data conversion string for that match
    datefmts = (
        (re.compile(r"\d{1,2}/\d{1,2}/\d{2}$"), "%m/%d/%y"),
        (re.compile(r"\d{1,2}/\d{1,2}/\d{4}$"), "%m/%d/%Y"),
        (re.compile(r"\d{1,2}-\d{1,2}-\d{2}$"), "%m-%d-%y"),
        (re.compile(r"\d{1,2}-\d{1,2}-\d{4}$"), "%m-%d-%Y"),
        (
            re.compile(r"\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}$"),
            "%Y-%m-%dT%H:%M:%S",
        ),
        (
            re.compile(r"\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}\.\d+$"),
            "%Y-%m-%dT%H:%M:%S.%f",
        ),
        (
            re.compile(r"\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}:\d{1,2}$"),
            "%Y-%m-%d %H:%M:%S",
        ),
        (
            re.compile(r"\d{4}-\d{1,2}-\d{1,2}\s\d{1,2}:\d{1,2}$"),
            "%Y-%m-%d %H:%M",
        ),
        (re.compile(r"\d{4}-\d{1,2}-\d{1,2}$"), "%Y-%m-%d"),
        (re.compile(r"^\d{8}$"), "%Y%m%d"),
        (re.compile(r"\d{1,2}-.{3}-\d{4}\s\d{2}:\d{2}"), "%d-%b-%Y %H:%M"),
        (re.compile(r"\d{1,2}-.{3}-\d{4}"), "%d-%b-%Y"),
        (re.compile(r"\d{1,2}-.{3}-\d{2}\s\d{2}:\d{2}"), "%d-%b-%y %H:%M"),
        (re.compile(r"\d{1,2}-.{3}-\d{2}"), "%d-%b-%y"),
        (
            re.compile(r"\d{1,2}/\d{1,2}/\d{4}\s\s\d{2}:\d{2}:\d{2} [A|P]M"),
            "%m/%d/%Y  %I:%M:%S %p",
        ),
        (
            re.compile(r"\d{1,2}/\d{1,2}/\d{4}\s\d{2}:\d{2}:\d{2} [A|P]M"),
            "%m/%d/%Y %I:%M:%S %p",
        ),
        (
            re.compile(r"\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{1,2}$"),
            "%m/%d/%Y %H:%M",
        ),
    )

    # if we passed in datetime we are done already
    if isinstance(value, datetime.datetime):
        return value

    # if we enabled skip blank - check for empty value
    if skipblank:
        if not value:
            return value
        elif isinstance(value, str) and not value.strip():
            return value

    # we only convert strings so we return the value when it is not a string
    if not isinstance(value, str):
        if not force_conversion:
            return value
        else:
            raise Exception("Unable to convert to date time:[{}]".format(value))

    # save the original value
    orig_value = value

    # strip the Z on the end before processing
    if value and value[-1].upper() == "Z":
        value = value[:-1]

    # print('value:', value)

    # step through each RE and if we find a match see if we can convert the string
    for redate, datefmt in datefmts:
        if redate.match(value):
            try:
                return datetime.datetime.strptime(value, datefmt)
            except Exception as e:
                if disp_msg:
                    print("-" * 40)
                    print("datetime_from_str - conversion error:")
                    print(f"    value..:  {value}")
                    print(f"    datefmt:  {datefmt}")
                raise e

    raise Exception("Unable to convert to date time:[{}]".format(orig_value))


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
#     2025-02-03T21:37:55Z
#
def datetimezone_from_str(
    value,
    skipblank=False,
    force_conversion: bool = False,
    disp_msg: bool = False,
):
    """
    Take in Any value (but process only strings) and attempt to convert it to a
    datetime timezone aware object if possible

    Inputs:
        value - Any - the value passed in
            if datetime already - just return it
            if a string - try to convert it to datetime
            anything else - return the value or error out based on the flags below
        skipblank - bool - if enabed, and the value is blank/empty - then skip the conversion and do not error out
        force_conversion - bool - if enabled, if we are unasble to convert - fail
        disp_msg - bool -  if enabled, display messages as processing - to enable interactive debugging
    Returns:
        value - the value or the datetime equivalent of the value
    """
    # bring in the regex routine it is needed
    import re

    # routines that clean up the data prior to processing
    datefmtscleanup = (
        (re.compile(r"(.*[+-])(\d{2}):(\d{2})$"), "remove-colon-2"),
    )
    # list of different dateformats we convert and the proper data conversion string for that match
    datefmts = (
        (
            re.compile(
                r"\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}\.\d+[+-]\d{4}$"
            ),
            "%Y-%m-%dT%H:%M:%S.%f%z",
        ),
        (
            re.compile(
                r"\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}\.\d+[+-]\d{4}$"
            ),
            "%Y-%m-%d %H:%M:%S.%f%z",
        ),
        (
            re.compile(
                r"\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}[+-]\d{4}$"
            ),
            "%Y-%m-%dT%H:%M:%S%z",
        ),
        (
            re.compile(
                r"\d{4}-\d{1,2}-\d{1,2} \d{1,2}:\d{1,2}:\d{1,2}[+-]\d{4}$"
            ),
            "%Y-%m-%d %H:%M:%S%z",
        ),
        (
            re.compile(r"\d{4}-\d{1,2}-\d{1,2}T\d{1,2}:\d{1,2}:\d{1,2}Z$"),
            "%Y-%m-%dT%H:%M:%SZ",
        ),
    )

    # if we passed in datetime we are done already
    if isinstance(value, datetime.datetime):
        # we got a date time, now see if it has timezone tied to it.
        if value.tzinfo:
            return value
        else:
            raise Exception(
                "Unable to convert to date time with timezone:[{}]".format(
                    value
                )
            )

    # if we enabled skip blank - check for empty value
    if skipblank:
        if not value:
            return value
        elif isinstance(value, str) and not value.strip():
            return value

    # we only convert strings so we return the value when it is not a string
    if not isinstance(value, str):
        if not force_conversion:
            return value
        else:
            raise Exception(
                "Unable to convert to date time with timezone:[{}]".format(
                    value
                )
            )

    # save the original value
    orig_value = value

    # trim the value to remove white space
    value = value.strip()

    # see if we need to change the format of the data we got in
    for redate, action in datefmtscleanup:
        if redate.match(value):
            m = redate.match(value)
            # each action tells us how to process and reformat the data
            if action == "remove-colon-2":
                value = m.group(1) + m.group(2) + m.group(3)

    # convert date into date/time/zone
    for redate, datefmt in datefmts:
        if redate.match(value):
            try:
                return datetime.datetime.strptime(value, datefmt)
            except Exception as e:
                if disp_msg:
                    print("-" * 40)
                    print("datetime_from_str - conversion error:")
                    print(f"    value..:  {value}")
                    print(f"    datefmt:  {datefmt}")
                raise e

    # error out because we could not convert
    raise Exception(
        "Unable to convert to date time with timezone:{}".format(orig_value)
    )


def valid_tz_string(tzstr: str) -> bool:
    """
    Ability to test if a tzstring is a valid tz string
    """
    if tz.gettz(tzstr):
        return True
    return False


def show_timezones(sublist: str | None = None, disp_msg: bool = False) -> list:
    """
    Show the list of timezone strings that you can work with

    Inputs:
        sublist - str -
        disp_msg - bool -  if enabled, display messages as processing - to enable interactive debugging
    Returns:
        tznames - list of timezone strings
    """
    # extract out the timezone strings
    sorted_zonenames = sorted(list(get_zonefile_instance().zones))
    # break it up int sections
    sections = set([x.split("/")[0] for x in sorted_zonenames if "/" in x])

    if sublist.capitalize() in sections:
        # passed in a sublist that when capitalized matches a section
        display_zonenames = [
            x
            for x in sorted_zonenames
            if x.startswith(str(sublist.capitalize()) + "/")
        ]
    elif sublist.upper() in ("US", "USA"):
        # passed in that we are looking for US values
        display_zonenames = [x for x in sorted_zonenames if x.startswith("US/")]
    elif sublist.upper() in ("SHORT", "ABBR"):
        # passed in we are looking for short or abbreviated names
        display_zonenames = [x for x in sorted_zonenames if "/" not in x]
    else:
        # otherwise - we are taking the full list
        display_zonenames = sorted_zonenames

    # display the results if we got asked to
    if disp_msg:
        print("Timezone Names:")
        for tzname in display_zonenames:
            print(tzname)

    return display_zonenames


# eof
