import xlwings as xw

# This is the file path to the workbook with the bond structure #
workbook = xw.Book(r'H:\MUNI\RSchuster\Research\Tobacco Model\Tobacco_Model_Beta.xlsm')
# This program will open and overwrite the excel file, but won't save it #
# For best results, keep the file path aimed to an excel template #
# After running, "save-as" the newly filled in excel with a DIFFERENT name #

####################################################################
############################ BOND CLASS ############################
####################################################################

class Bond:
    
    '''
    cusip: string    
    maturity_year: int
    coupon: float
    amount_outstanding: long int
    
    matured: bool initialized to false
    june_coupon_history, dec_coupon_history: dicts w/ year (int) as key and payment (double) as value
    year_paid_in: int
    '''
    
    def __init__(self, cusip, maturity_year, coupon, amount_outstanding, home_column_int, structure):
        # Pulled in from Excel #
        self.cusip = cusip                            
        self.maturity_year = maturity_year
        self.coupon = coupon
        self.amount_outstanding = amount_outstanding
        self.home_column_int = home_column_int
        self.structure = structure
        
        # Histories #
        self.june_coupon_history = {}
        self.dec_coupon_history = {}
        self.turbo_payment_history = {}
        self.year_end_balances = {}
        self.year_paid = 0
        
        # Identifiers #
        self.is_matured = False
        self.is_defaulted = False
        
    def __repr__(self):
        return str(self.cusip)

    def calc_interest_payment(self):            
        return (self.coupon * self.amount_outstanding)
    
    def is_maturing_this_year(self, year):
        if self.maturity_year == year:
            return True
        return False
    
    def mature(self):
        self.amount_outstanding = 0
        self.is_matured = True
        
    def default(self):
        self.is_defaulted = True

    def is_outstanding(self):
        if self.amount_outstanding > 0:
            return True
        else:
            return False

    def update_turbo_history(self, year, payment):
        self.turbo_history[year] = payment
        
        
####################################################################
######################## UNIVERSE FUNCTIONS ########################
####################################################################
        
        
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


def interest_payment(bond_list, bond_type, month, revs, dsrf, dec_payments, default, yr):
    for bond in bond_list:
        if (bond.is_outstanding()) and (not default):
            if (month == "June"):
                if revs >= bond.calc_interest_payment():
                    if bond_type == "serial":
                        revs -= bond.calc_interest_payment()
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                    elif bond_type == "turbo":
                        revs -= bond.calc_interest_payment()
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                elif (revs + dsrf) >= bond.calc_interest_payment():
                    if bond_type == "serial":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).color = (177, 160, 199)
                    elif bond_type == "turbo":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).color = (177, 160, 199)
                else:
                    if bond_type == "serial":
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).value = "DEFAULT"
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).color = (240, 64, 64)
                    elif bond_type == "turbo":
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).value = "DEFAULT"
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 9)).color = (240, 64, 64)
                        
            elif (month == "December"):
                if revs >= bond.calc_interest_payment():
                    if bond_type == "serial":
                        revs -= bond.calc_interest_payment()
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                    elif bond_type == "turbo":
                        revs -= bond.calc_interest_payment()
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                elif (revs + dsrf) >= bond.calc_interest_payment():
                    if bond_type == "serial":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).color = (177, 160, 199)
                    elif bond_type == "turbo":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).color = (177, 160, 199)
                else:
                    if bond_type == "serial":
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).value = "DEFAULT"
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).color = (240, 64, 64)
                    elif bond_type == "turbo":
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).value = "DEFAULT"
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 9)).color = (240, 64, 64)
            elif month == "December Estimate":
                dec_payments += bond.calc_interest_payment()
        
    return (revs, dsrf, dec_payments, default)


def enter_default(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf):
    
    if dsrf_paid_out_in_default:
        revenues_available_for_default = dsrf + available_revs
    else:
        revenues_available_for_default = available_revs
    
    if default_scenario == 'Pro Rata':

            series_amount_outstanding = 0
            
            for bond in turbo_bonds:
                if bond.amount_outstanding > 0:
                    series_amount_outstanding += bond.amount_outstanding
            for bond in serial_bonds:
                if bond.amount_outstanding:
                    series_amount_outstanding += bond.amount_outstanding
            
            # We can't pay off all the bonds #
            if series_amount_outstanding > (revenues_available_for_default):
            
                for bond in turbo_bonds:
                    if bond.amount_outstanding > 0:
                        prop_of_revs = round(float(bond.amount_outstanding / series_amount_outstanding), 9)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).value = prop_of_revs * (revenues_available_for_default)
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).color = (217, 105, 105)
                        
                for bond in serial_bonds:
                    if bond.amount_outstanding > 0:
                        prop_of_revs = round(float(bond.amount_outstanding / series_amount_outstanding), 9)
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).value = prop_of_revs * (revenues_available_for_default)
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).color = (217, 105, 105)
    
                revenues_available_for_default = 0
    
                if dsrf_paid_out_in_default:
                    available_revs = 0
                    dsrf = 0
                else:
                    available_revs = 0
                
            else:
                
                total_payment = 0
                
                for bond in turbo_bonds:
                    if bond.amount_outstanding > 0:
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).value = bond.amount_outstanding
                        total_payment += bond.amount_outstanding
                        bond.amount_outstanding = 0
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).color = (217, 105, 105)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(6)).value = year
                        bond.year_paid= year
                for bond in serial_bonds:
                    if bond.amount_outstanding > 0:
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).value = bond.amount_outstanding
                        total_payment += bond.amount_outstanding
                        bond.amount_outstanding = 0
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).color = (217, 105, 105)
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(6)).value = year + 1
                        bond.year_paid = year + 1
        
                revenues_available_for_default -= total_payment
        
                if dsrf_paid_out_in_default:
                    if total_payment < available_revs:
                        available_revs -= total_payment
                    else:
                        payment_from_dsrf = total_payment - available_revs
                        available_revs = 0
                        dsrf -= payment_from_dsrf
                else:
                    available_revs -= total_payment
        
    elif default_scenario == 'By Maturity':
        
        defaulted_bonds = []
        
        for bond in turbo_bonds:
            if bond.amount_outstanding > 0:
                defaulted_bonds.append(bond)
                
        for bond in serial_bonds:
            if bond.amount_outstanding > 0:
                defaulted_bonds.append(bond)
                
        defaulted_bonds_formatted_by_maturity = format_turbos_by_maturity(defaulted_bonds)
        
        for mat_yr in defaulted_bonds_formatted_by_maturity:
            
            total_mat_yr_outstanding = 0
            
            for bond in mat_yr:
                total_mat_yr_outstanding += bond.amount_outstanding
            
            # Can't pay everything off - exhaust revs and dsrf #
            if total_mat_yr_outstanding > revenues_available_for_default:
                
                for bond in mat_yr:
                    
                    prop_of_revs = round(float(bond.amount_outstanding / total_mat_yr_outstanding), 4)
                    
                    if bond.structure == 'turbo':
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).value = prop_of_revs * (revenues_available_for_default)
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).color = (217, 105, 105)
                    elif bond.structure == 'serial':
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).value = prop_of_revs * (revenues_available_for_default)
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).color = (217, 105, 105)
                
                revenues_available_for_default = 0
                
                if dsrf_paid_out_in_default:
                    available_revs = 0
                    dsrf = 0
                else:
                    available_revs = 0
                    
            # Can pay everything off #
            else:
                
                total_payment = 0
                
                for bond in mat_yr:
                    
                    if revenues_available_for_default != 0:
                    
                        if bond.structure == 'turbo':
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).value = bond.amount_outstanding
                            total_payment += bond.amount_outstanding
                            bond.amount_outstanding = 0
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).color = (217, 105, 105)
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(6)).value = year
                            bond.year_paid = year
                        elif bond.structure == 'serial':
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).value = bond.amount_outstanding
                            total_payment += bond.amount_outstanding
                            bond.amount_outstanding = 0
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 9)).color = (217, 105, 105)
                            worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(6)).value = year
                            bond.year_paid = year
        
                revenues_available_for_default -= total_payment
        
                if dsrf_paid_out_in_default:
                    if total_payment < available_revs:
                        available_revs -= total_payment
                    else:
                        payment_from_dsrf = total_payment - available_revs
                        available_revs = 0
                        dsrf -= payment_from_dsrf
                else:
                    available_revs -= total_payment
                        
    return available_revs, dsrf

def principal_payment(bond_list, bond_type, yr, revs, dsrf, default):
    for bond in bond_list:
        if (bond.is_outstanding()) and (bond.is_maturing_this_year(yr)) and (not default):
            if revs >= bond.amount_outstanding:
                revs -= bond.amount_outstanding
                if bond_type == "serial":
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).value = bond.amount_outstanding
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(6)).value = yr
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = "No"
                elif bond_type == "turbo":
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 9)).value = bond.amount_outstanding
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(6)).value = yr
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = "No"
                bond.mature()
                bond.year_paid = yr
            elif (revs + dsrf) >= bond.amount_outstanding:
                dsrf -= (bond.amount_outstanding - revs)
                revs = 0
                if bond_type == "serial":
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).value = bond.amount_outstanding
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).color = (177 ,160, 199)
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(6)).value = yr
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = "No"
                elif bond_type == "turbo":
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 9)).value = bond.amount_outstanding
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 9)).color = (177, 160, 199)
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(6)).value = yr
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = "No"
                bond.mature()
                bond.year_paid = yr
            else:
                default = True
                if bond_type == "serial":
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).value = "DEFAULT"
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).color = (240, 64, 64)
                elif bond_type == "turbo":
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 9)).value = "DEFAULT"
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 9)).color = (240, 64, 64)
                    
    return (revs, dsrf, default)

def there_is_enough_to_pay_this_year(serial_bonds, turbo_bonds, year, dsrf, revs):
    total_year_payments = 0
    
    for bond in serial_bonds:
        if bond.is_outstanding():
            total_year_payments += bond.calc_interest_payment()
        if (bond.is_outstanding()) and (not bond.is_maturing_this_year(year)):
            total_year_payments += bond.calc_interest_payment()
        if (bond.is_outstanding()) and (bond.is_maturing_this_year(year)):
            total_year_payments += bond.amount_outstanding
            
    for bond in turbo_bonds:
        if bond.is_outstanding():
            total_year_payments += bond.calc_interest_payment()
        if (bond.is_outstanding()) and (not bond.is_maturing_this_year(year)):
            total_year_payments += bond.calc_interest_payment()
        if (bond.is_outstanding()) and (bond.is_maturing_this_year(year)):
            total_year_payments += bond.amount_outstanding
            
    if (revs + dsrf) >= total_year_payments:
        return True
    else:
        return False

def format_turbos_by_maturity(list_of_bonds):
    
    '''
    Takes list of bonds as an input
    Returns an ORDERED list of lists of bonds ASCENDING by maturity
    not even going to lie that last line of code is genius
    ONLY RETURNS BONDS THAT HAVEN'T DEFAULTED OR MATURED!
    '''
    
    unique_years = []
    return_dict = {}
    return_list = []
    
    for bond in list_of_bonds:
        if (bond.amount_outstanding > 0) and (bond.maturity_year not in unique_years):
            unique_years.append(bond.maturity_year)
            
    for year in unique_years:
        return_dict[year] = []
        
    for bond in list_of_bonds:
        if (bond.amount_outstanding > 0) and (bond.maturity_year in return_dict.keys()):
            return_dict[bond.maturity_year].append(bond)
   
    for mat_yr in range(0, len(unique_years)):
        return_list.append(return_dict.pop(min(return_dict.keys())))
 
    return return_list


def turbo_payment(list_of_bonds, yr, turbo_amt, revs):
    
    '''
    Makes a call of format_turbos_by_maturity
    Returns the remaining revenues after the turbo
    '''
    
    formatted_list = format_turbos_by_maturity(list_of_bonds)
    
    for mat_yr in formatted_list:
        
        if turbo_amt > 0:
        
            total_outstanding_mat_yr = 0
            
            for bond in mat_yr:
                total_outstanding_mat_yr += bond.amount_outstanding
           
            # we can pay off everything in the maturity year
            if turbo_amt >= total_outstanding_mat_yr:
            
                turbo_amt -= total_outstanding_mat_yr
                revs -= total_outstanding_mat_yr
            
                for bond in mat_yr:
                    worksheet_turbos.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).value = bond.amount_outstanding
                    bond.turbo_payment_history[yr] = bond.amount_outstanding
                    bond.mature()
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(6)).value = yr
                    bond.year_paid = yr
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = "No"
            
            # we can't pay off everything in the maturity year
            else:
                for bond in mat_yr:
                    prop_of_revs = round(float(bond.amount_outstanding / total_outstanding_mat_yr), 4)
                    prop_of_turbo = (turbo_amt*prop_of_revs)
                    
                    worksheet_turbos.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 9)).value = prop_of_turbo
                    bond.turbo_payment_history[yr] = prop_of_turbo
                    bond.amount_outstanding -= prop_of_turbo
                    
                revs -= turbo_amt
                turbo_amt = 0
                
    return revs


####################################################################
############################ VARIABLES #############################
####################################################################

###### XL Wings Variables ######

worksheet_turbos = workbook.sheets['Turbo Bonds']
worksheet_serials = workbook.sheets['Serial Bonds']
worksheet_dsrf = workbook.sheets['DSRF']
worksheet_rev = workbook.sheets['Pledged Revenue']

###### Major Variables ######

turbo_bonds = []
serial_bonds = []
pledged_revs = {}

dsrf_max = int(worksheet_dsrf.range('B3').value)
dsrf_current = int(worksheet_dsrf.range('B4').value)
dsrf_balances = {}
dsrf_reserves = int(worksheet_dsrf.range('E3').value)

start_year = int(worksheet_rev.range('F8').value)
end_year = int(worksheet_rev.range('F9').value)
turbo_start_year = int(worksheet_rev.range('F10').value)
year_of_default = 0

default_has_occurred = False

###### Toggle Options ######

    # Does DSRF take priority over turbo payments? #
    # This is checked before assigning a value to amonut_to_turbo #

if str(worksheet_rev.range('F11').value) == 'Yes':
    DSRF_takes_priority_over_turbo = True
else:
    DSRF_takes_priority_over_turbo = False

    # What default scenario are we entering? #
    # enter_default method checks this and runs accordingly #
    
if str(worksheet_rev.range("F13").value) == "Pro Rata":
    default_scenario = "Pro Rata"
elif str(worksheet_rev.range("F13").value) == "By Maturity":
    default_scenario = "By Maturity"
    
    # Does default start payment in the current year? #
    # Default condition within payment segment checks for this #
    
if str(worksheet_rev.range("F12").value) == "Yes":
    default_accelerates_in_current_year = True
else:
    default_accelerates_in_current_year = False
    
    # Is the DSRF paid out immediately upon default? #
    # Used in enter_default - determines revenues_available_for_default #

if str(worksheet_rev.range("F14").value) == "Yes":
    dsrf_paid_out_in_default = True
else:
    dsrf_paid_out_in_default = False

# Excel-Python Helper Variables #

num_turbos = int(worksheet_turbos.range('B1').value)
num_serials = int(worksheet_serials.range('B1').value)

col_dict = build_column_dict()        
    

####################################################################
########################## INITIALIZATOIN ##########################
####################################################################

# Turbo Bonds #

for i in range(num_turbos):
    
    home_col_int = 2 + (i*7)
    home_col = col_dict[home_col_int]
    
    cusip = str(worksheet_turbos.range(home_col + str(3)).value)
    maturity = int(worksheet_turbos.range(home_col + str(4)).value)
    coupon = float(worksheet_turbos.range(home_col + str(5)).value)
    amt_outstanding = int(worksheet_turbos.range(home_col + str(6)).value)
    
    turbo_bonds.append(Bond(cusip, maturity, coupon, amt_outstanding, home_col_int, "turbo"))

# Pledged Revenues #
    
for i in range(5, 87):
    
    pledged_revs[start_year + (i - 5)] = int(worksheet_rev.range('B' + str(i)).value)
    
# Serial Bonds #    
    
for i in range(num_serials):
    
    home_col_int = 2 + (6*i)
    home_col = col_dict[home_col_int]
    
    cusip = str(worksheet_serials.range(home_col + str(3)).value)
    maturity = int(worksheet_serials.range(home_col + str(4)).value)
    coupon = float(worksheet_serials.range(home_col + str(5)).value)
    amt_outstanding = int(worksheet_serials.range(home_col + str(6)).value)
    
    serial_bonds.append(Bond(cusip, maturity, coupon, amt_outstanding, home_col_int, "serial"))


####################################################################
############################### MAIN ###############################
####################################################################

def run_model(start_year, end_year, turbo_bonds, serial_bonds, pledged_revs, dsrf_current, worksheet_dsrf, worksheet_rev, worksheet_serials, worksheet_turbos, default_has_occurred):
    
    for year in range(start_year, end_year + 1):
    
        # Note that bonds_outstanding lets us know if we have anything to do #
        # in the current year - if false none of the preceding code will run #
    
        bonds_outstanding = False
    
        for bond in turbo_bonds:
            if bond.amount_outstanding > 0:
                bonds_outstanding = True
        for bond in serial_bonds:
            if bond.amount_outstanding > 0:
                bonds_outstanding = True
        
        if bonds_outstanding:
            available_revs = pledged_revs[year]
            december_interest_payments = 0
            amount_to_turbo = 0
    
        ####################
        # Default Scenario #
        ####################
    
        if bonds_outstanding and default_has_occurred:
            available_revs, dsrf_current = enter_default(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf_current)
     
        ##################################
        #### Turbo hasn't started yet ####
        ##################################    
        
        elif bonds_outstanding and (year < turbo_start_year):
            
            if there_is_enough_to_pay_this_year(serial_bonds, turbo_bonds, year, dsrf_current, available_revs):
            
                # June Interest Serial Bonds
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "serial", "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
        
                # June Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "turbo", "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
      
                # Principal Payment Serial Bonds #
                available_revs, dsrf_current, default_has_occurred = principal_payment(serial_bonds, "serial", year, available_revs, dsrf_current, default_has_occurred)
                
                # Principal Payment Turbo Bonds #
                available_revs, dsrf_current, default_has_occurred = principal_payment(turbo_bonds, "turbo", year, available_revs, dsrf_current, default_has_occurred)
                
                # December Interest Payment Serial Bonds #    
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "serial", "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
        
                # December Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "turbo", "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
              
            #############################
            ##### Default Condition #####    
            #############################    
            
            else:
            
                default_has_occurred = True
            
                for bond in serial_bonds:
                    if bond.is_outstanding():
                        worksheet_serials.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 9) + ":" + col_dict[bond.home_column_int +3] + str(year - start_year + 9)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = "Yes"
                for bond in turbo_bonds:
                    if bond.is_outstanding():
                        worksheet_turbos.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 9) + ":" + col_dict[bond.home_column_int +4] + str(year - start_year + 9)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = "Yes"
                        
                worksheet_rev.range("A" + str(year - start_year + 5) + ":B" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_rev.range("H" + str(year - start_year + 5) + ":I" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_dsrf.range("A" + str(year - start_year + 7)).color = (217, 105, 105)
            
                if default_accelerates_in_current_year:
                    available_revs, dsrf_current = enter_default(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf_current)

                year_of_default = year

        ################################
        ######## Turbo Scenario ########
        ################################
    
        elif bonds_outstanding:
        
            if there_is_enough_to_pay_this_year(serial_bonds, turbo_bonds, year, dsrf_current, available_revs):
            
                # June Interest Serial Bonds
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "serial", "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
        
                # June Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "turbo", "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
      
                # Principal Payment Serial Bonds #
                available_revs, dsrf_current, default_has_occurred = principal_payment(serial_bonds, "serial", year, available_revs, dsrf_current, default_has_occurred)
                
                # Principal Payment Turbo Bonds #
                available_revs, dsrf_current, default_has_occurred = principal_payment(turbo_bonds, "turbo", year, available_revs, dsrf_current, default_has_occurred)
            
                # December Interest Payment Serial Bonds #    
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "serial", "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)

                # December Interest First Estimate Turbo Bonds #
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "turbo", "December Estimate", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
                
                #########################
                ##### Turbo Payment #####
                #########################
                
                if (not default_has_occurred) and (available_revs > december_interest_payments):
                
                    if DSRF_takes_priority_over_turbo:
                    
                        if dsrf_current < dsrf_max:
                            if (available_revs - december_interest_payments) >= (dsrf_max - dsrf_current):
                                available_revs -= (dsrf_max - dsrf_current)
                                dsrf_current = dsrf_max
                                amount_to_turbo = available_revs - december_interest_payments
                            else:
                                dsrf_current += (available_revs - december_interest_payments)
                                available_revs = december_interest_payments
                                amount_to_turbo = 0
                        else:
                            amount_to_turbo = available_revs - december_interest_payments
                            
                    else:
                        amount_to_turbo = available_revs - december_interest_payments
            
                if amount_to_turbo > 0:
                    available_revs = turbo_payment(turbo_bonds, year, amount_to_turbo, available_revs)
                
                # December Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "turbo", "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
          
            #############################
            ##### Default Condition #####    
            #############################
            
            else:
            
                default_has_occurred = True
                
                for bond in serial_bonds:
                    if bond.is_outstanding():
                        worksheet_serials.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 9) + ":" + col_dict[bond.home_column_int +3] + str(year - start_year + 9)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = "Yes"
                for bond in turbo_bonds:
                    if bond.is_outstanding():
                        worksheet_turbos.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 9) + ":" + col_dict[bond.home_column_int +4] + str(year - start_year + 9)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = "Yes"
                        
                worksheet_rev.range("A" + str(year - start_year + 5) + ":B" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_rev.range("H" + str(year - start_year + 5) + ":I" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_dsrf.range("A" + str(year - start_year + 7)).color = (217, 105, 105)
            
                if default_accelerates_in_current_year:
                    available_revs, dsrf_current = enter_default(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf_current)
                
                year_of_default = year        
            
        ######################################################   
        #################### Housekeeping ####################
        ######################################################
    
        if bonds_outstanding:
        
            # Serial Bond Year-End Balances #
    
            for bond in serial_bonds:
                bond.year_end_balances[year] = bond.amount_outstanding
                
                if year == start_year:
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).value = bond.amount_outstanding
                elif bond.year_end_balances[year - 1] != 0:
                    if bond.is_defaulted:
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).color = (217, 105, 105)
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 9)).value = bond.amount_outstanding
    
            # Turbo Bonds Year-End Balances #    
        
            for bond in turbo_bonds:
                bond.year_end_balances[year] = bond.amount_outstanding
            
                if year == start_year:
                    worksheet_turbos.range(col_dict[bond.home_column_int + 4] + str(year - start_year + 9)).value = bond.amount_outstanding
                elif bond.year_end_balances[year - 1] != 0:
                    if bond.is_defaulted:
                        worksheet_turbos.range(col_dict[bond.home_column_int + 4] + str(year - start_year + 9)).color = (217, 105, 105)
                    worksheet_turbos.range(col_dict[bond.home_column_int + 4] + str(year - start_year + 9)).value = bond.amount_outstanding
        
            # DSRF Year-End Balance #
    
            worksheet_dsrf.range("B" + str(year - start_year + 7)).value = dsrf_current + dsrf_reserves
            dsrf_balances[year] = dsrf_current
        
            if year != start_year:
                if (dsrf_balances[year - 1] < dsrf_current) and (dsrf_current != dsrf_max):
                    worksheet_dsrf.range("B" + str(year - start_year + 7)).color = (239, 210, 209)
                elif (dsrf_balances[year - 1] != dsrf_max) and (dsrf_current > dsrf_balances[year - 1]):
                    worksheet_dsrf.range("B" + str(year - start_year + 7)).color = (216, 228, 188)
                elif dsrf_current != dsrf_max:
                    if dsrf_current == 0:
                        worksheet_dsrf.range('B' + str(year - start_year + 7)).color = (205, 115, 113)
                    else:
                        worksheet_dsrf.range("B" + str(year - start_year + 7)).color = (230, 184, 183)
            
            elif dsrf_current != dsrf_max:
                if dsrf_current == 0:
                    worksheet_dsrf.range('B' + str(year - start_year + 7)).color = (205, 115, 113)
                else:
                    worksheet_dsrf.range("B" + str(year - start_year + 7)).color = (230, 184, 183)

            # Excess Revenues #
        
            worksheet_rev.range("I" + str(year - start_year + 5)).value = available_revs
            
    return year_of_default

def_yr = run_model(start_year, end_year, turbo_bonds, serial_bonds, pledged_revs, dsrf_current, worksheet_dsrf, worksheet_rev, worksheet_serials, worksheet_turbos, default_has_occurred)