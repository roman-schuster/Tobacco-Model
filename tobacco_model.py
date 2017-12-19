import xlwings as xw

# Bond Class #
class Bond:
    '''
    cusip: string    
    maturity_year: int
    coupon: float
    amount_outstanding: long int
    proportion_of_revenue: float (should be 100.00 if unique maturity)
    lien_priority: I have to be honest I'm not sure how this is going to be used
    
    matured: bool initialized to false
    june_coupon_history, dec_coupon_history: dicts w/ year (int) as key and payment (double) as value
    year_paid_in: int
    '''
    
    def __init__(self,cusip, maturity_year, coupon, amount_outstanding, lien_priority, home_column_int):
        self.cusip = cusip                            
        self.maturity_year = maturity_year
        self.coupon = coupon
        self.amount_outstanding = amount_outstanding
        self.lien_priority = lien_priority
        self.home_column_int = home_column_int
        
        self.matured = False
        self.june_coupon_history = {}
        self.dec_coupon_history = {}
        self.turbo_payment_history = {}
        self.year_end_balances = {}
        self.year_paid_in = 3000

    def calc_interest_payment(self):            
        return (self.coupon * self.amount_outstanding)
    
    def is_maturing(self, year):
        if self.maturity_year == year:
            return True
        return False
    
    def mature(self, year):
        self.amount_outstanding = 0
        self.matured = True
        self.year_paid_in = year

    def is_outstanding(self, yr):
        if self.maturity_year < yr:
            return False
        else:
            return (not self.matured)
    
    def pay_interest(self, month, year):
        '''
        Calculates current interest and adds to appropriate payment history
        '''
        interest_payment = self.calc_interest_payment()
        
        if month == 'December':
            self.dec_coupon_history[year] = interest_payment
        elif month == "June":
            self.june_coupon_history[year] = interest_payment

    def update_turbo_history(self, year, payment):
        '''
        this is it
        '''
        self.turbo_history[year] = payment
        

    def update_year_end_balance(self, year):
        self.year_end_balances[year] = self.amount_outstanding

    
    
# Universe Functions #
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
    

def format_turbos_by_maturity(list_of_bonds):
    '''
    Takes list of bonds as an input
    Returns a dict with bond maturity as key and list of bonds as value
    ONLY RETURNS BONDS THAT HAVEN'T DEFAULTED OR MATURED!
    '''
    unique_years = []
    return_dict = {}
    
    for bond in list_of_bonds:
        if (bond.is_outstanding()) and (bond.maturity not in unique_years):
            unique_years.append(bond.maturity)
            
    for year in unique_years:
        return_dict[year] = []
        
    for bond in list_of_bonds:
        if (bond.is_outstanding()) and (bond.maturity in return_dict.keys()):
            return_dict[bond.maturity].append(bond)
            
    return return_dict        

def has_unique_lien(bond, lien_dict):
    '''
    Returns true if bond is the only one in the lien
    '''
    if len(lien_dict[bond.maturity]) == 1:
        return True
    else:
        return False
            
def interest_payment(bond_list, month, revs, dsrf, dec_payments, default, yr):
    for bond in bond_list:
        if bond.is_outstanding(yr):
            if (month == "June") or (month == "December"):
                if revs >= bond.calc_interest_payment():
                    revs -= bond.calc_interest_payment()
                    bond.pay_interest(month, yr)
                elif (revs + dsrf) >= bond.calc_interest_payment():
                    dsrf -= (bond.calc_interest_payment() - revs)
                    revs = 0
                    bond.pay_interest(month, yr)
                else:
                    default = True
            elif month == "December Estimate":
                dec_payments += bond.calc_interest_payment()
        
    return (revs, dsrf, dec_payments, default)            
        
def principal_payment(bond_list, yr, revs, dsrf, default):
    for bond in bond_list:
        if bond.is_outstanding(yr) and bond.is_maturing(yr):
            if revs >= bond.amount_outstanding:
                revs -= bond.amount_outstanding
                print('paying ' + bond.cusip)
                bond.mature(yr)
            elif (revs + dsrf) >= bond.amount_outstanding:
                dsrf -= (bond.amount_outstanding - revs)
                revs = 0
                bond.mature(yr)
                print('paying ' + bond.cusip)
            else:
                default = True
        
    return (revs, dsrf, default)    

    
    
    
    
    
    
    
    
    
    
## Variables ##
# XL Wing Variables #
workbook = xw.Book(r'C:\Users\Roman\Desktop/tobacco.xls')

worksheet_turbos = workbook.sheets['Turbo Bonds']
worksheet_serials = workbook.sheets['Serial Bonds']
worksheet_dsrf = workbook.sheets['DSRF']
worksheet_rev = workbook.sheets['Pledged Revenue']

# Major Variables #
turbo_bonds = []
serial_bonds = []
pledged_revs = {}

dsrf_max = int(worksheet_dsrf.range('B3').value)
dsrf_current = int(worksheet_dsrf.range('B4').value)
dsrfIsFull = True if (dsrf_max == dsrf_current) else False
dsrf_balances = {}

current_year = 2017
end_year = 2050
year_of_default = 0

default_has_occurred = False

# Excel-Python Helper Variables #
num_turbos = int(worksheet_turbos.range('B1').value)
num_serials = int(worksheet_serials.range('B1').value)

col_dict = build_column_dict()        
    
## Initialization ##
# Initializing Turbo Bonds #
for i in range(num_turbos):
    home_col_int = 2 + (i*6)
    home_col = col_dict[home_col_int]
    
    cusip = str(worksheet_turbos.range(home_col + str(3)).value)
    maturity = int(worksheet_turbos.range(home_col + str(4)).value)
    coupon = float(worksheet_turbos.range(home_col + str(5)).value)
    lien_priority = str(worksheet_turbos.range(home_col + str(6)).value)
    amt_outstanding = int(worksheet_turbos.range(home_col + str(7)).value)
    
    turbo_bonds.append(Bond(cusip, maturity, coupon, amt_outstanding, lien_priority, home_col_int))

# Initializing Pledged Revenues #
for i in range(5, 88):
    pledged_revs[current_year + (i - 5)] = int(worksheet_rev.range('B' + str(i)).value)
    
# Initializing Serial Bonds #    
for i in range(num_serials):
    home_col_int = 2 + (5*i)
    home_col = col_dict[home_col_int]
    
    cusip = str(worksheet_serials.range(home_col + str(3)).value)
    maturity = int(worksheet_serials.range(home_col + str(4)).value)
    coupon = float(worksheet_serials.range(home_col + str(5)).value)
    lien_priority = str(worksheet_serials.range(home_col + str(6)).value)
    amt_outstanding = int(worksheet_serials.range(home_col + str(7)).value)
    
    serial_bonds.append(Bond(cusip, maturity, coupon, amt_outstanding, lien_priority, home_col_int))
    
    
    

# Main ##   
for year in range(current_year, end_year):
    
    available_revs = pledged_revs[year]
    december_interest_payments = 0
    amount_to_turbo = 0
    
    if default_has_occurred:
        # What happens when we default? #
        default_has_occurred = True
    else:
        # June Interest Serial Bonds
        available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
        
        # June Interest Turbo Bonds #    
        if not default_has_occurred:
            available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "June", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
      
        # Principal Payment Serial Bonds #
        if not default_has_occurred:
            available_revs, dsrf_current, default_has_occurred = principal_payment(serial_bonds, year,
                                                                                   available_revs, dsrf_current,
                                                                                   default_has_occurred)
            
        # Principal Payment Turbo Bonds #
        if not default_has_occurred:
            available_revs, dsrf_current, default_has_occurred = principal_payment(turbo_bonds, year,
                                                                                   available_revs, dsrf_current,
                                                                                   default_has_occurred)
            
        # December Interest Serial Bonds #
        if not default_has_occurred:
            available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(serial_bonds, "December", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
    
        # December Interest Turbo Bonds 1st Estimation #
        if not default_has_occurred:
                available_revs, dsrf_current, december_interest_payments, default_has_occurred = interest_payment(turbo_bonds, "December Estimate", available_revs, dsrf_current, december_interest_payments, default_has_occurred, year)
            
        # Turbo Payment Estimation #
        if (available_revs - december_interest_payments) and (not default_has_occurred) > 0:
            amount_to_turbo = available_revs - december_interest_payments
        
        # Turbo Payment #
        if amount_to_turbo > 0:
            turbo_formatted_bonds = format_turbo_by_maturity(turbo_bonds)
            
            for maturity_year in turbo_formatted_bonds.keys():
                if len(turbo_formatted_bonds[maturity_year]) = 1:
                    if amount_to_turbo >= turbo_formatted_bonds[maturity_year][0].amount_outstanding:
                        turbo_formatted_bonds[maturity_year][0].update_turbo_history(year, turbo_formatted_bonds[maturity_year][0].amount_outstanding)
                        amount_to_turbo -= turbo_formatted_bonds[maturity_year][0].amount_outstanding
                        available_revs -= turbo_formatted_bonds[maturity_year][0].amount_outstanding
                        turbo_formatted_bonds[maturity_year][0].amount_outstanding = 0
                    else:
                        turbo_formatted_bonds[maturity_year][0].update_turbo_history(year, amount_to_turbo)
                        turbo_formatted_bonds[maturity_year][0].amount_outstanding -= amount_to_turbo
                        available_revs -= amount_to_turbo
                        amount_to_turbo = 0

                else:
                    maturity_year_amt_outstanding = 0
                    rev_proportions = []
                    
                    for bond in turbo_formatted_bonds[maturity_year]:
                        maturity_year_amt_outstanding += bond.amount_outstanding
                    
                    if amount_to_turbo >= maturity_year_amt_outstanding:
                        # we can pay off multiple turbos and need to move to a new maturity year #
                    else:
                        # we 
        
        # December Interest on Turbo Bonds #
        # First Instance of Default #
        if default_has_occurred:
            

    for bond in serial_bonds:
        bond.update_year_end_balance(year)

        
    for bond in turbo_bonds:
        bond.update_year_end_balance(year)
        
    dsrf_balances[year] = dsrf_current   
    
# Rewriting #

for year in range(current_year, end_year):
    
    for bond in turbo_bonds:
        if year in bond.june_coupon_history.keys():
            worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - 2007)).value = bond.june_coupon_history[year]
        if year in bond.dec_coupon_history.keys():
            worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - 2007)).value = bond.dec_coupon_history[year]
        if year in bond.year_end_balances.keys():
            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - 2007)).value = bond.year_end_balances[year]
                                   
    for bond in serial_bonds:
        if year in bond.june_coupon_history.keys():
            worksheet_serials.range(col_dict[bond.home_column_int] + str(year - 2007)).value = bond.june_coupon_history[year]
        if year in bond.dec_coupon_history.keys():
            worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - 2007)).value = bond.dec_coupon_history[year]
        if year in bond.year_end_balances.keys():
            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - 2007)).value = bond.year_end_balances[year]
