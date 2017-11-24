import xlwings as xw

## Variables ##
# XL Wing Variables #
workbook = xw.Book(r'C:\path\to\file.xlsx')

worksheet_turbos = workbook.sheets['Turbo Bonds']
worksheet_serials = workbook.sheets['Serial Bonds']
worksheet_dsrf = workbook.sheets['DSRF']
worksheet_rev = worksheet.sheets['Pledged Revenue']

# Major Variables #
turbo_bonds = []
serial_bonds = []
pledged_revs = {}

dsrf_max = worksheet_dsrf.range('B3').value
dsrf_current = worksheet_dsrf.range('B4').value
dsrfIsFull = True if (dsrf_max == dsrf_current) else False

current_year = 2017
end_year = 2099

default_has_occurred = False

# Excel-Python Helper Variables #
num_turbos = worksheet_turbos.range('B1').value
num_serials = worksheet_serials.range('B1').value

def build_column_dict():
    '''
    Builds and returns a dictionary
    Keys: Integers between 1 and 701
    Values: String value of excel column corresponding to int (1 is "A", 27 is "AA", 55 is "BC", 701 is "ZZ")
    '''
    column_dictionary = {1: "A", 2: "B", 3: "C", 4: "D", 5: "E", 6: "F", 7: "G", 8: "H", 9: "I", 10: "J",
                        11: "K", 12: "L", 13: "M", 14: "N", 15: "O", 16: "P", 17: "Q", 18: "R", 19: "S", 20: "T",
                        21: "U", 22: "V", 23: "W", 24: "X", 25: "Y", 26: "Z"}

    assistant_to_the_column_dictionary = {27: "A", 53: "B", 79: "C", 105: "D", 131: "E", 157: "F", 183: "G",
                                          209: "H", 235: "I", 261: "J", 287: "K", 313: "L", 339: "M", 365: "N",
                                          391: "O", 417: "P", 443: "Q", 469: "R", 495: "S", 521: "T", 547: "U",
                                          573: "V", 599: "W", 625: "X", 651: "Y", 677: "Z"}
    
    for i in range(27, 702, 26):
        for j in range(0, 26):
            column_dictionary[i+j] = assistant_to_the_column_dictionary[i] + column_dictionary[j+1]
    
    return column_dictionary

col_dict = build_column_dict()        
    
## Initialization ##
# Initializing Turbo Bonds #
for i in range(num_turbos):
    home_col_int = 2 + (i*6)
    home_col = col_dict[home_col_int]
    
    maturity = worksheet_turbos.range(home_col + cstr(4)).value
    coupon = worksheet_turbos.range(home_col + cstr(5)).value
    lien_priority = worksheet_turbos.range(home_col + cstr(6)).value
    amt_outstanding = worksheet_turbos.range(home_col + cstr(7)).value
    
    turbo_bonds.append(Turbo_Bond(maturity, coupon, amt_outstanding, lien_priority, home_col_int))

# Initializing Pledged Revenues #
for i in range(5, 88):
    pledged_revs[current_year + (i - 5)] = worksheet_rev.range('B' + cstr(i)).value
    
# Initializing Serial Bonds #    
for i in range(num_serials):
    home_col_int = 2 + (5*i)
    home_col = col_dict[home_col_int]
    
    maturity = worksheet_serials.range(home_col + cstr(4)).value
    coupon = worksheet_serials.range(home_col + cstr(5)).value
    amt_outstanding = worksheet_serials.range(home_col + cstr(6)).value
    #### YOU NEED TO ADD THIS ROW TO THE EXCEL SHEET! ####
    lien = worksheet_serials.range(home_col + cstr(7)).value
    
    serial_bonds.append(Serial_Bond(maturity, coupon, amt_outstanding, lien))
    
## Main ##   
for year in range(Current_year, end_year):
    
    # Minor Variables #
    december_interest_payments = 0
    available_revs = pledged_revs[year]
    amount_to_turbo = 0
    
    # June Interest Serial Bonds
    available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred)
        
    # June Interest Turbo Bonds #    
    available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred)
      
    # Principal Payment Serial Bonds #
    available_revs, dsrf_current, default_has_occurred = principal_payment(serial_bonds, year, available_revs, dsrf_current, default_has_occurred)
            
    # Principal Payment Turbo Bonds #
    available_revs, dsrf_current, default_has_occurred = principal_payment(turbo_bonds, year, available_revs, dsrf_current, default_has_occurred)
            
    # December Interest Serial Bonds #
    available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred)
    
    # December Interest Turbo Bonds 1st Estimation #
    available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "December Estimate", available_revs, dsrf_current, december_interest_payments, default_has_occurred)
            
    # Turbo Payment Estimation #
    if (available_revs - december_interest_payments) > 0:
        amount_to_turbo = available_revs - december_interest_payments
        
        # Figure out which bonds to turbo - how do I signify priority? #
        while amount_to_turbo > 0:
            
    # December Interest Turbo Bonds #
    available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred)
    
    # Paying the Turbo Bonds #
    
    
    
def format_liens(list_of_bonds):
    '''
    Takes list of bonds as an input
    Returns a dict of liens with bond maturity as key and list of bonds as value
    '''
    unique_liens = []
    return_dict = {}
    
    for bond in list_of_bonds:
        if bond.maturity not in unique_liens:
            unique_liens.append(cstr(bond.maturity))
            
    for lien in unique_liens:
        return_dict[lien] = []
        
    for bond in list_of_bonds:
        if cstr(bond.maturity) in return_dict.keys():
            return_dict[cstr(bond.maturity)].append(bond)
            
    return return_dict        

def has_unique_lien(bond, lien_dict):
    '''
    Returns true if bond is the only one in the lien
    '''
    if len(lien_dict[cstr(bond.maturity)]) = 1:
        return True
    else:
        return False
            
def interest_payment(bond_list, month, revs, dsrf, dec_payments, default):
    for bond in bond_list:
        if bond.is_outstanding():
            if (month == "June") or (month == "December"):
                if revs >= bond.calc_interest_payment():
                    revs -= bond.calc_interest_payment()
                    bond.pay_interest(month)
                elif (revs + dsrf) >= bond.calc_interest_payment():
                    dsrf -= (bond.calc_interest_payment() - revs)
                    revs = 0
                    bond.pay_interest(month)
                else:
                    default = True
            elif month == "December Estimate":
                dec_payments += bond.calc_interest_payment()
    return (revs, dsrf, dec_payments, default)            
        
def principal_payment(bond_list, yr, revs, dsrf, default):
    for bond in bond_list:
        if bond.is_outstanding() and bond.is_maturing(yr):
            if revs >= bond.amount_outstanding:
                revs -= bond.amount_outstanding
                bond.amount_outstanding = 0
                bond.matured = True
            elif (revs + dsrf) >= bond.amount_outstanding:
                dsrf -= (bond.amount_outstanding - revs)
                revs = 0
                bond.amount_outstanding = 0
                bond.matured = True
            else:
                default = True
    return (revs, dsrf, default)
    
    
    
    
    
