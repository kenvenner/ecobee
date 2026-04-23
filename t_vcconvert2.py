import unittest
import vcconvert2 as vc
import kvxls
import os
import datetime
import copy

import pprint
pp = pprint.PrettyPrinter(indent=4)

"""
"""

valid_recs = [
    {   'Booking': 'OWN-00400',
        'Checkout Day': datetime.datetime(2026, 2, 9, 0, 0),
        'Comment': None,
        'First Night': datetime.datetime(2026, 2, 6, 0, 0),
        'HoldUntil': None,
        'Nights': 3,
        'PoolEnabled': None,
        'Rent': None,
        'RevPerDay': None,
        'Source': None,
        'Type': 'Owners Hold - Scribner',
        'XLSRowAbs': 2,
        'blank001': None,
        'blank002': None,
        'blank003': None,
        'blank004': None,
        'blank005': None,
        'blank006': None,
        'blank007': None,
        'blank008': None},
    {   'Booking': 'OWN-00390',
        'Checkout Day': datetime.datetime(2026, 2, 16, 0, 0),
        'Comment': None,
        'First Night': datetime.datetime(2026, 2, 11, 0, 0),
        'HoldUntil': None,
        'Nights': 5,
        'PoolEnabled': None,
        'Rent': None,
        'RevPerDay': None,
        'Source': 'Owners Hold',
        'Type': 'Owners Hold - Venner',
        'XLSRowAbs': 3,
        'blank001': None,
        'blank002': None,
        'blank003': None,
        'blank004': None,
        'blank005': None,
        'blank006': None,
        'blank007': None,
        'blank008': None}
]

def load_vc_file():
    """
    Read in the raw content of the villa reservation file
    """
    
    xlsfile = "Attune_Estate_2026_Bookings.xlsx"
    req_cols = vc.COL_REQUIRED
    xlsdateflds = [vc.FIRST_NIGHT_FLD, vc.CHECKOUT_FLD, "BookedOn", "HoldUntil"]

    # read in the XLS
    xlsaref = kvxls.readxls2list_findheader(
        xlsfile,
        req_cols=req_cols,
        optiondict={
            "dateflds": xlsdateflds,
            "sheetname": "Listing",
            "save_row_abs": True,
        },
        debug=False,
    )

    return xlsaref


class TestVCconvert2(unittest.TestCase):
    """Unit tests for vcconvert2 villa load xlsx and related utilities."""

    def test_validate_res_records_p01_pass(self):
        """ Pass a valid set of records """
        errors = vc.validate_res_records(valid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD, vc.CHECKOUT_FLD, vc.TYPE_FLD)
        self.assertEqual(errors, [])
    def test_validate_res_records_p02_not_datetime(self):
        """ date fields not filled with datetime values errors """
        expected_errors = ["Field [First Night] not of type datetime - xlsrow [2]:\n{'recidx': 0, 'type': <class 'str'>, 'rec': {'Booking': 'OWN-00400', 'Checkout Day': 'not-date-type', 'Comment': None, 'First Night': 'not-date-type', 'HoldUntil': None, 'Nights': 3, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': None, 'Type': 'Owners Hold - Scribner', 'XLSRowAbs': 2, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n", "Field [Checkout Day] not of type datetime - xlsrow [2]:\n{'recidx': 0, 'type': <class 'str'>, 'rec': {'Booking': 'OWN-00400', 'Checkout Day': 'not-date-type', 'Comment': None, 'First Night': 'not-date-type', 'HoldUntil': None, 'Nights': 3, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': None, 'Type': 'Owners Hold - Scribner', 'XLSRowAbs': 2, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n", "Field [First Night] not of type datetime - xlsrow [3]:\n{'recidx': 1, 'type': <class 'str'>, 'rec': {'Booking': 'OWN-00390', 'Checkout Day': 'not-date-type', 'Comment': None, 'First Night': 'not-date-type', 'HoldUntil': None, 'Nights': 5, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': 'Owners Hold', 'Type': 'Owners Hold - Venner', 'XLSRowAbs': 3, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n", "Field [Checkout Day] not of type datetime - xlsrow [3]:\n{'recidx': 1, 'type': <class 'str'>, 'rec': {'Booking': 'OWN-00390', 'Checkout Day': 'not-date-type', 'Comment': None, 'First Night': 'not-date-type', 'HoldUntil': None, 'Nights': 5, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': 'Owners Hold', 'Type': 'Owners Hold - Venner', 'XLSRowAbs': 3, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n"]
        invalid_recs =copy.deepcopy( valid_recs)
        for x in invalid_recs:
            x[vc.FIRST_NIGHT_FLD] = 'not-date-type'
            x[vc.CHECKOUT_FLD] = 'not-date-type'
        errors = vc.validate_res_records(invalid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD, vc.CHECKOUT_FLD, vc.TYPE_FLD)
        self.assertEqual(errors, expected_errors)
    def test_validate_res_records_p03_nights_not_calcd(self):
        """ end-start date != number of nights errors """
        expected_errors = ["Field [Nights] not calc as date difference - xlsrow [2]:\n{'recidx': 0, 'dt_diff': 3, 'num_nights': 99, 'rec': {'Booking': 'OWN-00400', 'Checkout Day': datetime.datetime(2026, 2, 9, 0, 0), 'Comment': None, 'First Night': datetime.datetime(2026, 2, 6, 0, 0), 'HoldUntil': None, 'Nights': 99, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': None, 'Type': 'Owners Hold - Scribner', 'XLSRowAbs': 2, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n", "Field [Nights] not calc as date difference - xlsrow [3]:\n{'recidx': 1, 'dt_diff': 5, 'num_nights': 99, 'rec': {'Booking': 'OWN-00390', 'Checkout Day': datetime.datetime(2026, 2, 16, 0, 0), 'Comment': None, 'First Night': datetime.datetime(2026, 2, 11, 0, 0), 'HoldUntil': None, 'Nights': 99, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': 'Owners Hold', 'Type': 'Owners Hold - Venner', 'XLSRowAbs': 3, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n"]
        invalid_recs =copy.deepcopy( valid_recs)
        for x in invalid_recs:
            x[vc.NIGHTS_FLD] = 99
        errors = vc.validate_res_records(invalid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD, vc.CHECKOUT_FLD, vc.TYPE_FLD)
        self.assertEqual(errors, expected_errors)
    def test_validate_res_records_p04_invalid_occ_type(self):
        """ invalide occ_type errors """
        expected_errors = ["Field [Type] not in OCC_TYPE_CONV - xlsrow [2]:\n{'recidx': 0, 'rec_fld_type': 'invalid-occ-type', 'rec': {'Booking': 'OWN-00400', 'Checkout Day': datetime.datetime(2026, 2, 9, 0, 0), 'Comment': None, 'First Night': datetime.datetime(2026, 2, 6, 0, 0), 'HoldUntil': None, 'Nights': 3, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': None, 'Type': 'invalid-occ-type', 'XLSRowAbs': 2, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}, 'OCC_TYPE_CONV': {'Friends': ['R', 1], 'Hold': ['R', 1], 'Hold - Clean': ['C', 0], 'Hold - Construction': ['O', 1], 'Hold - Fall Mbr Event': ['O', 1], 'Hold - Harvest Party': ['O', 1], 'Hold - Maint': ['M', 0], 'Hold - Maint.': ['M', 0], 'Hold - Other': ['M', 0], 'Hold - Owner Scribner': ['O', 1], 'Hold - Owner': ['O', 1], 'Hold - Release Party': ['O', 1], 'Hold - Renter': ['R', 1], 'Hold - Scribner': ['O', 1], 'Hold - Venner': ['O', 1], 'Hold - Winery Business': ['O', 0], 'Hold- Mainten': ['M', 0], 'Hold-Deep Clean': ['M', 0], 'Hold-Owner': ['O', 1], 'Hold-Renter': ['R', 1], 'Karli Vendor Tour': ['O', 1], 'Owners Hold - Harvest': ['O', 1], 'Owners Hold - Venner': ['O', 1], 'Owners Hold - Scribner': ['O', 1], 'Owners Hold': ['O', 1], 'Owners Hold- Venner': ['O', 1], 'Res - Auteur': ['O', 1], 'Res - Owner': ['O', 1], 'Res - Renter': ['R', 1], 'Res- Renter': ['R', 1], 'Res-Renter': ['R', 1], 'Res. - Owner': ['O', 1], 'Res. - Renter': ['R', 1], 'Rest - Renter': ['R', 1], 'Res.-Owner': ['O', 1], 'Res.-Renter': ['R', 1], 'TA/Villa Specalist': ['R', 1], 'VRBO/Repeat guest': ['R', 1]}}\n", "Field [Type] not in OCC_TYPE_CONV - xlsrow [3]:\n{'recidx': 1, 'rec_fld_type': 'invalid-occ-type', 'rec': {'Booking': 'OWN-00390', 'Checkout Day': datetime.datetime(2026, 2, 16, 0, 0), 'Comment': None, 'First Night': datetime.datetime(2026, 2, 11, 0, 0), 'HoldUntil': None, 'Nights': 5, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': 'Owners Hold', 'Type': 'invalid-occ-type', 'XLSRowAbs': 3, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}, 'OCC_TYPE_CONV': {'Friends': ['R', 1], 'Hold': ['R', 1], 'Hold - Clean': ['C', 0], 'Hold - Construction': ['O', 1], 'Hold - Fall Mbr Event': ['O', 1], 'Hold - Harvest Party': ['O', 1], 'Hold - Maint': ['M', 0], 'Hold - Maint.': ['M', 0], 'Hold - Other': ['M', 0], 'Hold - Owner Scribner': ['O', 1], 'Hold - Owner': ['O', 1], 'Hold - Release Party': ['O', 1], 'Hold - Renter': ['R', 1], 'Hold - Scribner': ['O', 1], 'Hold - Venner': ['O', 1], 'Hold - Winery Business': ['O', 0], 'Hold- Mainten': ['M', 0], 'Hold-Deep Clean': ['M', 0], 'Hold-Owner': ['O', 1], 'Hold-Renter': ['R', 1], 'Karli Vendor Tour': ['O', 1], 'Owners Hold - Harvest': ['O', 1], 'Owners Hold - Venner': ['O', 1], 'Owners Hold - Scribner': ['O', 1], 'Owners Hold': ['O', 1], 'Owners Hold- Venner': ['O', 1], 'Res - Auteur': ['O', 1], 'Res - Owner': ['O', 1], 'Res - Renter': ['R', 1], 'Res- Renter': ['R', 1], 'Res-Renter': ['R', 1], 'Res. - Owner': ['O', 1], 'Res. - Renter': ['R', 1], 'Rest - Renter': ['R', 1], 'Res.-Owner': ['O', 1], 'Res.-Renter': ['R', 1], 'TA/Villa Specalist': ['R', 1], 'VRBO/Repeat guest': ['R', 1]}}\n"]
        invalid_recs =copy.deepcopy( valid_recs)
        for x in invalid_recs:
            x[vc.TYPE_FLD] = 'invalid-occ-type'
        errors = vc.validate_res_records(invalid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD, vc.CHECKOUT_FLD, vc.TYPE_FLD)
        self.assertEqual(errors, expected_errors)
    def test_validate_res_records_p05_dupe_booking(self):
        """ dupe booking errors """
        expected_errors = ["Field [Type] booking already exists - xlsrow [3]:\n{'recidx': 1, 'orig_recidx': 0, 'booking': 'dupe_booking', 'rec': {'Booking': 'dupe_booking', 'Checkout Day': datetime.datetime(2026, 2, 16, 0, 0), 'Comment': None, 'First Night': datetime.datetime(2026, 2, 11, 0, 0), 'HoldUntil': None, 'Nights': 5, 'PoolEnabled': None, 'Rent': None, 'RevPerDay': None, 'Source': 'Owners Hold', 'Type': 'Owners Hold - Venner', 'XLSRowAbs': 3, 'blank001': None, 'blank002': None, 'blank003': None, 'blank004': None, 'blank005': None, 'blank006': None, 'blank007': None, 'blank008': None}}\n"]
        invalid_recs =copy.deepcopy( valid_recs)
        for x in invalid_recs:
            x[vc.BOOKING_FLD] = 'dupe_booking'
        errors = vc.validate_res_records(invalid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD, vc.CHECKOUT_FLD, vc.TYPE_FLD)
        self.assertEqual(errors, expected_errors)

    def test_filtered_sorted_xlsaref_p01_pass(self):
        " filter out records missing first_night or # of nights"
        invalid_recs =copy.deepcopy( valid_recs)
        for x in valid_recs:
            # blank first night fld
            newx = copy.deepcopy(x)
            newx[vc.FIRST_NIGHT_FLD] = ''
            invalid_recs.append(newx)
            # blank # nights fld
            newx = copy.deepcopy(x)
            newx[vc.NIGHTS_FLD] = ''
            invalid_recs.append(newx)
        self.assertEqual(len(invalid_recs), len(valid_recs)*3)
        filtered_recs = vc.filtered_sorted_xlsaref(invalid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD)
        self.assertEqual(len(filtered_recs), len(valid_recs))
        self.assertEqual(filtered_recs, valid_recs)
    def test_filtered_sorted_xlsaref_p01_num_night_zero_not_filtered(self):
        " filter out records missing first_night or # of nights"
        invalid_recs =copy.deepcopy( valid_recs)
        for x in invalid_recs:
            x[vc.NIGHTS_FLD] = 0
        filtered_recs = vc.filtered_sorted_xlsaref(invalid_recs, vc.FIRST_NIGHT_FLD, vc.NIGHTS_FLD)
        pp.pprint(filtered_recs)
        self.assertEqual(len(filtered_recs), len(valid_recs))
        
if __name__ == "__main__":
    unittest.main()
