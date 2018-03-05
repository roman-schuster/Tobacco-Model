# -*- coding: utf-8 -*-
"""
Created on Fri Feb 23 14:52:25 2018

@author: RSCHUSTER
"""

import xlwings as xw

# This is the file path to the workbook with the bond structure #
home_testing_path = r'C:\Users\Roman\Desktop\Golden_Tobacco_Model.xlsm'
work_testing_path = r'H:\MUNI\RSchuster\Research\Tobacco Model\Golden_Tobacco_Model.xlsm'
workbook = xw.Book(home_testing_path)
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
    structure: string (turbo or serial)
    
    june_coupon_history, dec_coupon_history: dicts w/ year (int) as key and payment (double) as value
    year_paid_in: int
    '''
    def __init__(self, cusip, maturity_year, coupon, amount_outstanding, home_column_int, structure, lien, price):

        self.cusip = cusip                            
        self.maturity_year = maturity_year
        self.coupon = coupon
        self.amount_outstanding = amount_outstanding
        self.initial_amount_outstanding = amount_outstanding
        self.lien = lien
        self.price = price
        
        self.home_column_int = home_column_int
        self.structure = structure
        
        self.june_coupon_history = {}
        self.dec_coupon_history = {}
        self.turbo_payment_history = {}
        self.year_end_balances = {}
        self.year_paid = 0
        
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



def calc_irr_with_excel_updates(serial_bonds, turbo_bonds):
    '''
    '''
    for bond in serial_bonds:
        for i in range(10, 92):
            if int(worksheet_serials.range(col_dict[bond.home_column_int - 1] + str(i)).value) <= bond.year_paid:
                total = 0
                if worksheet_serials.range(col_dict[bond.home_column_int] + str(i)).value is not None:
                    total += worksheet_serials.range(col_dict[bond.home_column_int] + str(i)).value
                if worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(i)).value is not None:
                    total += worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(i)).value
                if worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(i)).value is not None:
                    total += worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(i)).value
                
                worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(i)).value = total
        
    for bond in serial_bonds:
        
        last_row = 10
        for i in range(11, 92):
            if worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(i)).value is not None:
                last_row = i
            else:
                break
        
        cash_flows = []
        
        for i in range(10, last_row + 1):
            cash_flows.append(float(worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(i)).value))
        
        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(10)).value = bond.initial_amount_outstanding * -1

        for i in range(11, last_row + 2):
            worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(i)).value = cash_flows[i - 11]
                                    
        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(7)).value = '=irr(' + col_dict[bond.home_column_int + 3] + str(10) + ':' + col_dict[bond.home_column_int + 3] + str(last_row + 1)
        
        irr = float(worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(7)).value)
        
        for i in range(10, last_row + 2):
            if (i == (last_row + 1)) or (i == (last_row + 2)):
                worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(i)).value = ''
            else:
                worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(i)).value = cash_flows[i - 10]
                                        
        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(7)).value = irr                    
                
                                           
                                
    
    for bond in turbo_bonds:
        for i in range(10, 92, 1):
            if int(worksheet_turbos.range(col_dict[bond.home_column_int - 1] + str(i)).value) <= bond.year_paid:
                total = 0
                if worksheet_turbos.range(col_dict[bond.home_column_int] + str(i)).value is not None:
                    total += worksheet_turbos.range(col_dict[bond.home_column_int] + str(i)).value
                if worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(i)).value is not None:
                    total += worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(i)).value
                if worksheet_turbos.range(col_dict[bond.home_column_int + 2] + str(i)).value is not None:
                    total += worksheet_turbos.range(col_dict[bond.home_column_int + 2] + str(i)).value
                if worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(i)).value is not None:
                    total += worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(i)).value
                    
                worksheet_turbos.range(col_dict[bond.home_column_int + 4] + str(i)).value = total
    

def format_turbos_by_maturity(list_of_bonds):
    '''
    Takes list of bonds as an input
    Returns an ORDERED list of lists of bonds ASCENDING by maturity
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



def there_is_enough_to_pay_this_year(serial_bonds, turbo_bonds, year, dsrf, revs):
    '''
    Lets us know if we have enough to pay interest and principal (before turbo)
    for current year
    Takes list of serials, list of turbos, year (int), DSRF balance (long int) and
    available revs (long int)
    Returns False if default would occur, True if there is enough money to pay
    '''
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
    
    

def turbo_payment_with_excel_updates(list_of_bonds, yr, turbo_amt, revs):
    '''
    Makes a call of format_turbos_by_maturity
    Turbo is paid BY MATURITY - NOT PRO RATA
    Returns the remaining revenues after the turbo
    '''
    formatted_list = format_turbos_by_maturity(list_of_bonds)
    
    for mat_yr in formatted_list:
        
        if turbo_amt > 0:

            total_outstanding_mat_yr = 0
            
            for bond in mat_yr:
                total_outstanding_mat_yr += bond.amount_outstanding
           
            # Can pay everything in maturity year #
            if turbo_amt >= total_outstanding_mat_yr:
            
                turbo_amt -= total_outstanding_mat_yr
                revs -= total_outstanding_mat_yr
            
                for bond in mat_yr:
                    worksheet_turbos.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 10)).value = bond.amount_outstanding
                    bond.turbo_payment_history[yr] = bond.amount_outstanding
                    bond.mature()
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = yr
                    bond.year_paid = yr
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(4)).value = "No"
            
            # Can't pay everything in maturity year #
            else:
                
                for bond in mat_yr:
                    prop_of_revs = round(float(bond.amount_outstanding / total_outstanding_mat_yr), 4)
                    prop_of_turbo = (turbo_amt*prop_of_revs)
                    
                    worksheet_turbos.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 10)).value = prop_of_turbo
                    bond.turbo_payment_history[yr] = prop_of_turbo
                    bond.amount_outstanding -= prop_of_turbo
                    
                revs -= turbo_amt
                turbo_amt = 0
                
    return revs    



def principal_payment_with_excel_updates(bond_list, yr, revs, dsrf):
    '''
    Returns:
        revs - revenue left over after principal payments
        dsrf - current dsrf balance after principal payments
    If no bond is maturing this year function does nothing
        and revs, dsrf will not be altered
    Updates bond amount outstanding via bond.mature()
    Note that a default scenario will never occur through this function
        there_is_enough..() would have caught the insufficiency before this
        function was called
    '''
    for bond in bond_list:
        
        if bond.is_outstanding() and bond.is_maturing_this_year(yr):
            
            # Can pay principal w/o DSRF #
            if revs >= bond.amount_outstanding:
                revs -= bond.amount_outstanding
                if bond.structure == "serial":
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 10)).value = bond.amount_outstanding
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = yr
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(4)).value = "No"
                    
                elif bond.structure == "turbo":
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 10)).value = bond.amount_outstanding
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = yr
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(4)).value = "No"
                    
                bond.mature()
                bond.year_paid = yr
                
            # Can pay principal but need to use DSRF #    
            elif (revs + dsrf) >= bond.amount_outstanding:
                
                dsrf -= (bond.amount_outstanding - revs)
                revs = 0
                
                if bond.structure == "serial":
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 10)).value = bond.amount_outstanding
                    worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(yr - start_year + 10)).color = (177 ,160, 199)
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = yr
                    worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(4)).value = "No"
                    
                elif bond.structure == "turbo":
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 10)).value = bond.amount_outstanding
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(yr - start_year + 10)).color = (177, 160, 199)
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = yr
                    worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(4)).value = "No"
                    
                bond.mature()
                bond.year_paid = yr
                    
    return (revs, dsrf)



def interest_payment_with_excel_updates(bond_list, month, revs, dsrf, dec_payments, yr):
    '''
    month is a string - either "June", "December" or "December Estimate"
    "December Estimate" should ONLY be entered with turbo bonds
        Serial bonds can't be turboed, so we don't need december estimates for them
    Returns:
        revs - udpated available revenues after interest payments
        dsrf - updated dsrf balance after interest payments
        dec_payments - int; total value of december payments for turbo estimation
    Note that default will never occur through this function
        there_is_enough..() would have caught the insufficiency before this
        function was called
    '''
    for bond in bond_list:
        if (bond.is_outstanding()):
            
            if (month == "June"):
                
                # Can pay interest w/o DSRF #
                if revs >= bond.calc_interest_payment():
                    if bond.structure == "serial":
                        revs -= bond.calc_interest_payment()
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                    elif bond.structure == "turbo":
                        revs -= bond.calc_interest_payment()
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                        
                # Can pay interest but need to use DSRF #        
                elif (revs + dsrf) >= bond.calc_interest_payment():
                    if bond.structure == "serial":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(yr - start_year + 10)).color = (177, 160, 199)
                    elif bond.structure == "turbo":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.june_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(yr - start_year + 10)).color = (177, 160, 199)
                        
            elif (month == "December"):
                
                # Can pay interest w/o DSRF #
                if revs >= bond.calc_interest_payment():
                    if bond.structure == "serial":
                        revs -= bond.calc_interest_payment()
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                    elif bond.structure == "turbo":
                        revs -= bond.calc_interest_payment()
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                        
                # Can pay interest but need to use DSRF #        
                elif (revs + dsrf) >= bond.calc_interest_payment():
                    if bond.structure == "serial":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 10)).color = (177, 160, 199)
                    elif bond.structure == "turbo":
                        dsrf -= (bond.calc_interest_payment() - revs)
                        revs = 0
                        bond.dec_coupon_history[yr] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(yr - start_year + 10)).color = (177, 160, 199)

            elif (month == "December Estimate") and (bond.structure == "turbo"):
                dec_payments += bond.calc_interest_payment()
        
    return (revs, dsrf, dec_payments)



def defaulted_interest_payment_with_excel_updates(bond_list, month, default_revs, dsrf, year):
    '''
    bond_list is a list of DEFAULTED bonds
    month is a string - either "June" or "December"
    Returns:
        default_revs - updated revenues available after interest payments on defaulted bonds
        dsrf - updated dsrf balance after interest payments on defaulted bonds
    Same as interest_payment function, except with a third conditional
        scenario (default within a default) where all available revs are paid
        pro rata
    '''
    total_interest_amt = 0
    total_interest_paid = 0
    
    for bond in bond_list:
        total_interest_amt += bond.calc_interest_payment()
    
    for bond in bond_list:
        if (month == "June"):
            
            # Can pay interest w/o DSRF #
            if default_revs >= total_interest_amt:
                
                if bond.structure == "serial":
                    default_revs -= bond.calc_interest_payment()
                    total_interest_paid += bond.calc_interest_payment()
                    bond.june_coupon_history[year] = bond.calc_interest_payment()
                    worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                    worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
                    
                elif bond.structure == "turbo":
                    default_revs -= bond.calc_interest_payment()
                    total_interest_paid += bond.calc_interest_payment()
                    bond.june_coupon_history[year] = bond.calc_interest_payment()
                    worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                    worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
            
            # Can pay interest but have to use DSRF #
            elif (default_revs + dsrf) >= total_interest_amt:
                
                if bond.structure == "serial":
                    if (total_interest_paid + bond.calc_interest_payment()) < (default_revs):
                        default_revs -= bond.calc_interest_payment()
                        total_interest_paid += bond.calc_interest_payment()
                        bond.june_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
                    else:
                        dsrf -= (bond.calc_interest_payment() - default_revs)
                        default_revs = 0
                        total_interest_paid += bond.calc_itnerest_payment()
                        bond.june_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)

                elif bond.structure == "turbo":
                    if (total_interest_paid + bond.calc_interest_payment()) < (default_revs):
                        default_revs -= bond.calc_interest_payment()
                        total_interest_paid += bond.calc_interest_payment()
                        bond.june_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
                    else:
                        dsrf -= (bond.calc_interest_payment() - default_revs)
                        default_revs = 0
                        total_interest_paid += bond.calc_itnerest_payment()
                        bond.june_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
            
            # Can't pay interest - deplete DSRF and pay remaining revs out pro rata #
            else:
                
                prop_of_revs = bond.calc_interest_payment() / total_interest_amt
                pro_rata_revs = dsrf + default_revs
                
                if bond.structure == "serial":
                    bond.june_coupon_history[year] = (pro_rata_revs * prop_of_revs)
                    total_interest_paid += (pro_rata_revs * prop_of_revs)
                    worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = (pro_rata_revs * prop_of_revs)
                    worksheet_serials.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
                    default_revs -= (default_revs * prop_of_revs)
                    dsrf -= (dsrf * prop_of_revs)
                    
                elif bond.structure == "turbo":
                    bond.june_coupon_history[year] = (pro_rata_revs * prop_of_revs)
                    total_interest_paid += (pro_rata_revs * prop_of_revs)
                    worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).value = (pro_rata_revs * prop_of_revs)
                    worksheet_turbos.range(col_dict[bond.home_column_int] + str(year - start_year + 10)).color = (217, 105, 105)
                    default_revs -= (default_revs * prop_of_revs)
                    dsrf -= (dsrf * prop_of_revs)
                    
        elif (month == "December"):
            
            # Can pay interest w/o DSRF #
            if default_revs >= total_interest_amt:
                
                if bond.structure == "serial":
                    default_revs -= bond.calc_interest_payment()
                    total_interest_paid += bond.calc_interest_payment()
                    bond.dec_coupon_history[year] = bond.calc_interest_payment()
                    worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                    worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
                    
                elif bond.structure == "turbo":
                    default_revs -= bond.calc_interest_payment()
                    total_interest_paid += bond.calc_interest_payment()
                    bond.dec_coupon_history[year] = bond.calc_interest_payment()
                    worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                    worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
            
            # Can pay interest but have to use DSRF #
            elif (default_revs + dsrf) >= total_interest_amt:
                
                if bond.structure == "serial":
                    if (total_interest_paid + bond.calc_interest_payment()) < (default_revs):
                        default_revs -= bond.calc_interest_payment()
                        total_interest_paid += bond.calc_interest_payment()
                        bond.dec_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
                    else:
                        dsrf -= (bond.calc_interest_payment() - default_revs)
                        default_revs = 0
                        total_interest_paid += bond.calc_itnerest_payment()
                        bond.dec_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)

                elif bond.structure == "turbo":
                    if (total_interest_paid + bond.calc_interest_payment()) < (default_revs):
                        default_revs -= bond.calc_interest_payment()
                        total_interest_paid += bond.calc_interest_payment()
                        bond.dec_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
                    else:
                        dsrf -= (bond.calc_interest_payment() - default_revs)
                        default_revs = 0
                        total_interest_paid += bond.calc_itnerest_payment()
                        bond.dec_coupon_history[year] = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = bond.calc_interest_payment()
                        worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
                 
            # Can't pay interest - pay remaining DSRF and revs out pro rata #        
            else:
                prop_of_revs = bond.calc_interest_payment() / total_interest_amt
                pro_rata_revs = dsrf + default_revs
                
                if bond.structure == "serial":
                    bond.dec_coupon_history[year] = (pro_rata_revs * prop_of_revs)
                    total_interest_paid += (pro_rata_revs * prop_of_revs)
                    worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = (pro_rata_revs * prop_of_revs)
                    worksheet_serials.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
                    default_revs -= (default_revs * prop_of_revs)
                    dsrf -= (dsrf * prop_of_revs)
                    
                elif bond.structure == "turbo":
                    bond.dec_coupon_history[year] = (pro_rata_revs * prop_of_revs)
                    total_interest_paid += (pro_rata_revs * prop_of_revs)
                    worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).value = (pro_rata_revs * prop_of_revs)
                    worksheet_turbos.range(col_dict[bond.home_column_int + 1] + str(year - start_year + 10)).color = (217, 105, 105)
                    default_revs -= (default_revs * prop_of_revs)
                    dsrf -= (dsrf * prop_of_revs)
                    
    return (default_revs, dsrf)



def enter_default_with_excel_updates(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf, year_of_default):
    '''
    
    '''
    
    defaulted_bonds = []
    
    for bond in turbo_bonds:
        if bond.amount_outstanding != 0:
            defaulted_bonds.append(bond)
            
    for bond in serial_bonds:
        if bond.amount_outstanding != 0:
            defaulted_bonds.append(bond)
    
    # Interest Payments #
    revenues_available_for_default = available_revs
    revenues_available_for_default, dsrf = defaulted_interest_payment_with_excel_updates(defaulted_bonds, "June", revenues_available_for_default, dsrf, year)
    revenues_available_for_default, dsrf = defaulted_interest_payment_with_excel_updates(defaulted_bonds, "December", revenues_available_for_default, dsrf, year)
    
    ##################################
    # Accelerated Payments from DSRF #
    ##################################
    
    if ((year == year_of_default) and (dsrf_paid_out_in_default == "year of default")) or ((year == year_of_default + 1) and (dsrf_paid_out_in_default == 'year after default')):
        
        # Pro Rata Scenario #
        if dsrf_scenario == "Pro Rata":
            
            series_amt_outstanding = 0
        
            for bond in defaulted_bonds:
                series_amt_outstanding += bond.amount_outstanding
                
            # Pro Rata Scenario - Can't Pay Everything #    
            if series_amt_outstanding > dsrf:
                
                for bond in defaulted_bonds:
                        prop_of_revs = round(float(bond.amount_outstanding / series_amt_outstanding), 9)
                        bond.amount_outstanding -= prop_of_revs * (dsrf)
                        
                        if bond.structure == "turbo":
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = prop_of_revs * (dsrf)
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                        elif bond.structure == "serial":
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = prop_of_revs * (dsrf)
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
            
                dsrf = 0
            
            # Pro Rata Scenario - Can Pay Everything #
            else:

                for bond in defaulted_bonds:
                    dsrf -= bond.amount_outstanding
                    bond.amount_outstanding = 0
                    bond.year_paid= year
                
                    if bond.structure == "turbo":
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = bond.amount_outstanding
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = year

                    elif bond.structure == "serial":
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = bond.amount_outstanding
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
                
        # By Maturity Scenario #       
        elif dsrf_scenario == "By Maturity":
                
            defaulted_bonds_formatted_by_maturity = format_turbos_by_maturity(defaulted_bonds)
        
            for mat_yr in defaulted_bonds_formatted_by_maturity:
            
                total_mat_yr_outstanding = 0
            
                for bond in mat_yr:
                    total_mat_yr_outstanding += bond.amount_outstanding
                
                # By Maturity Scenario - Can't Pay Everything #
                if (total_mat_yr_outstanding > dsrf) and (dsrf != 0):
                
                    for bond in mat_yr:
                    
                        prop_of_revs = round(float(bond.amount_outstanding / total_mat_yr_outstanding), 4)
                        bond.amount_outstanding -= (prop_of_revs * dsrf)
                        
                        if bond.structure == 'turbo':
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value =( prop_of_revs * dsrf)
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                        elif bond.structure == 'serial':
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = (prop_of_revs * dsrf)
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
                
                    dsrf = 0
                    
                # By Maturity Scenario - Can Pay Everything #
                elif dsrf != 0:
                    
                    total_payment = 0
                
                    for bond in mat_yr:
                        
                        total_payment += bond.amount_outstanding
                        bond.amount_outstanding = 0
                        bond.year_paid = year
                        
                        if bond.structure == 'turbo':
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = bond.amount_outstanding
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
                            
                        elif bond.structure == 'serial':
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = bond.amount_outstanding
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
                            worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
        
                    dsrf -= total_payment
    
    ####################################
    # Payments from Available Revenues #
    ####################################
    
    if default_scenario == 'Pro Rata':

            series_amount_outstanding = 0
            
            for bond in turbo_bonds:
                series_amount_outstanding += bond.amount_outstanding
            for bond in serial_bonds:
                series_amount_outstanding += bond.amount_outstanding
            
            # We can't pay off all the bonds #
            if series_amount_outstanding > (revenues_available_for_default):
            
                for bond in turbo_bonds:
                    if bond.amount_outstanding > 0:
                        prop_of_revs = round(float(bond.amount_outstanding / series_amount_outstanding), 9)
                        prev_amt = worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value
                        if prev_amt is not None:
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default)) + float(prev_amt)
                        else:
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default))
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                        
                for bond in serial_bonds:
                    if bond.amount_outstanding > 0:
                        prop_of_revs = round(float(bond.amount_outstanding / series_amount_outstanding), 9)
                        prev_amt = worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value
                        if prev_amt is not None:
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default)) + float(prev_amt)
                        else:
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default))
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
    
                revenues_available_for_default = 0
                available_revs = 0
                
            else:
                
                total_payment = 0
                
                for bond in turbo_bonds:
                    if bond.amount_outstanding > 0:
                        prev_amt = worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value
                        if prev_amt is not None:
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = bond.amount_outstanding + float(prev_amt)
                        else:
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = bond.amount_outstanding
                        total_payment += bond.amount_outstanding
                        bond.amount_outstanding = 0
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
                        bond.year_paid= year
                for bond in serial_bonds:
                    if bond.amount_outstanding > 0:
                        prev_amt = worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value
                        if prev_amt is not None:
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = bond.amount_outstanding + float(prev_amt)
                        else:
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = bond.amount_outstanding
                        total_payment += bond.amount_outstanding
                        bond.amount_outstanding = 0
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
                        bond.year_paid = year
        
                revenues_available_for_default -= total_payment
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
            if (revenues_available_for_default != 0) and (total_mat_yr_outstanding > revenues_available_for_default):
                
                for bond in mat_yr:
                    
                    prop_of_revs = round(float(bond.amount_outstanding / total_mat_yr_outstanding), 4)
                    
                    if bond.structure == 'turbo':
                        prev_amt = worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value
                        if prev_amt is not None:
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default)) + float(prev_amt)
                        else:
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default))
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                    elif bond.structure == 'serial':
                        prev_amt = worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value
                        if prev_amt is not None:
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default)) + float(prev_amt)
                        else:
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = (prop_of_revs * (revenues_available_for_default))
                        bond.amount_outstanding -= prop_of_revs * (revenues_available_for_default)
                        worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
                
                revenues_available_for_default = 0
                available_revs = 0
                    
            # Can pay everything off #
            else:
                
                total_payment = 0
                
                for bond in mat_yr:
                    
                    if revenues_available_for_default != 0:
                    
                        if bond.structure == 'turbo':
                            prev_amt = worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value
                            if prev_amt is not None:
                                worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = bond.amount_outstanding + float(prev_amt)
                            else:
                                worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).value = bond.amount_outstanding
                            total_payment += bond.amount_outstanding
                            bond.amount_outstanding = 0
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(year - start_year + 10)).color = (217, 105, 105)
                            worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
                            bond.year_paid = year
                        elif bond.structure == 'serial':
                            prev_amt = worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value
                            if prev_amt is not None:
                                worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = bond.amount_outstanding + float(prev_amt)
                            else:
                                worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).value = bond.amount_outstanding
                            total_payment += bond.amount_outstanding
                            bond.amount_outstanding = 0
                            worksheet_serials.range(col_dict[bond.home_column_int + 2] + str(year - start_year + 10)).color = (217, 105, 105)
                            worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(5)).value = year
                            bond.year_paid = year
        
                revenues_available_for_default -= total_payment
                available_revs -= total_payment
                        
    return available_revs, dsrf


####################################################################
############################ VARIABLES #############################
####################################################################

###### XL Wings Variables ######

worksheet_turbos = workbook.sheets['Turbo Bonds']
worksheet_serials = workbook.sheets['Serial Bonds']
worksheet_dsrf = workbook.sheets['DSRF']
worksheet_rev = workbook.sheets['Pledged Revenue']

###### Universe Variables ######

turbo_bonds = []
serial_bonds = []
pledged_revs = {}
excess_revs = {}

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
if str(worksheet_rev.range('F11').value) == 'Yes':
    DSRF_takes_priority_over_turbo = True
else:
    DSRF_takes_priority_over_turbo = False

    # What default scenario are we entering? #
if str(worksheet_rev.range("F12").value) == "Pro Rata":
    default_scenario = "Pro Rata"
elif str(worksheet_rev.range("F12").value) == "By Maturity":
    default_scenario = "By Maturity"

    
    # Is the DSRF paid out immediately upon default? #
if str(worksheet_rev.range("F13").value) == "Yes - In Year of Default":
    dsrf_paid_out_in_default = "year of default"
elif str(worksheet_rev.range("F13").value) == 'Yes - In Year After Default':
    dsrf_paid_out_in_default = "year after default"
else:
    dsrf_paid_out_in_default = "not paid"
    
    # How is the accelerated DSRF payment distributed? #
if str(worksheet_rev.range("F14").value) == "Pro Rata":
    dsrf_scenario = "Pro Rata"
elif str(worksheet_rev.range("F14").value) == "By Maturity":
    dsrf_scenario = "By Maturity"

# Excel-Python Helper Variables #

num_turbos = int(worksheet_turbos.range('B1').value)
num_serials = int(worksheet_serials.range('B1').value)

col_dict = build_column_dict()        
    

####################################################################
########################## INITIALIZATOIN ##########################
####################################################################

# Turbo Bonds #

for i in range(num_turbos):
    
    home_col_int = 2 + (i*8)
    home_col = col_dict[home_col_int]
    
    cusip = str(worksheet_turbos.range(home_col + str(3)).value)
    maturity = int(worksheet_turbos.range(home_col + str(4)).value)
    coupon = round((1/2) * float(worksheet_turbos.range(home_col + str(5)).value), 3)
    amt_outstanding = int(worksheet_turbos.range(home_col + str(6)).value)
    lien = str(worksheet_turbos.range(home_col + str(7)).value)
    price = float(worksheet_turbos.range(col_dict[home_col_int + 3] + str(6)).value)
    
    turbo_bonds.append(Bond(cusip, maturity, coupon, amt_outstanding, home_col_int, "turbo", lien, price))

# Pledged Revenues #
    
for i in range(5, 87):
    
    pledged_revs[start_year + (i - 5)] = int(worksheet_rev.range('B' + str(i)).value)
    
# Serial Bonds #    
    
for i in range(num_serials):
    
    home_col_int = 2 + (7*i)
    home_col = col_dict[home_col_int]
    
    cusip = str(worksheet_serials.range(home_col + str(3)).value)
    maturity = int(worksheet_serials.range(home_col + str(4)).value)
    coupon = round((1/2) * float(worksheet_serials.range(home_col + str(5)).value), 3)
    amt_outstanding = int(worksheet_serials.range(home_col + str(6)).value)
    lien = str(worksheet_serials.range(home_col + str(7)).value)
    price = float(worksheet_serials.range(col_dict[home_col_int + 3] + str(6)).value)
    
    serial_bonds.append(Bond(cusip, maturity, coupon, amt_outstanding, home_col_int, "serial", lien, price))
    
####################################################################
############################### MAIN ###############################
####################################################################

def run_model_with_excel_updates(start_year, end_year, turbo_bonds, serial_bonds, pledged_revs, dsrf_current, worksheet_dsrf, worksheet_rev, worksheet_serials, worksheet_turbos, default_has_occurred):
    
    year_of_default = 0
    
    for year in range(start_year, end_year + 1):
    
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
            available_revs, dsrf_current = enter_default_with_excel_updates(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf_current, year_of_default)
     
        ##################################
        #### Turbo hasn't started yet ####
        ##################################    
        
        elif bonds_outstanding and (year < turbo_start_year):
            
            if there_is_enough_to_pay_this_year(serial_bonds, turbo_bonds, year, dsrf_current, available_revs):
            
                # June Interest Serial Bonds
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(serial_bonds, "June", available_revs, dsrf_current, december_interest_payments, year)
        
                # June Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(turbo_bonds, "June", available_revs, dsrf_current, december_interest_payments, year)
      
                # Principal Payment Serial Bonds #
                available_revs, dsrf_current = principal_payment_with_excel_updates(serial_bonds, year, available_revs, dsrf_current)
                
                # Principal Payment Turbo Bonds #
                available_revs, dsrf_current = principal_payment_with_excel_updates(turbo_bonds, year, available_revs, dsrf_current)
                
                # December Interest Payment Serial Bonds #    
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(serial_bonds, "December", available_revs, dsrf_current, december_interest_payments, year)
        
                # December Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(turbo_bonds, "December", available_revs, dsrf_current, december_interest_payments, year)
              
            #############################
            ##### Default Condition #####    
            #############################    
            
            else:
            
                default_has_occurred = True
                year_of_default = year
            
                for bond in serial_bonds:
                    if bond.is_outstanding():
                        worksheet_serials.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 10) + ":" + col_dict[bond.home_column_int +4] + str(year - start_year + 10)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(4)).value = "Yes"
                for bond in turbo_bonds:
                    if bond.is_outstanding():
                        worksheet_turbos.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 10) + ":" + col_dict[bond.home_column_int +5] + str(year - start_year + 10)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(4)).value = "Yes"
                        
                worksheet_rev.range("A" + str(year - start_year + 5) + ":B" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_rev.range("H" + str(year - start_year + 5) + ":I" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_dsrf.range("A" + str(year - start_year + 7)).color = (217, 105, 105)
            
                available_revs, dsrf_current = enter_default_with_excel_updates(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf_current, year_of_default)

        ################################
        ######## Turbo Scenario ########
        ################################
    
        elif bonds_outstanding:
        
            if there_is_enough_to_pay_this_year(serial_bonds, turbo_bonds, year, dsrf_current, available_revs):
            
                # June Interest Serial Bonds
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(serial_bonds, "June", available_revs, dsrf_current, december_interest_payments, year)
        
                # June Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(turbo_bonds, "June", available_revs, dsrf_current, december_interest_payments, year)
      
                # Principal Payment Serial Bonds #
                available_revs, dsrf_current = principal_payment_with_excel_updates(serial_bonds, year, available_revs, dsrf_current)
                
                # Principal Payment Turbo Bonds #
                available_revs, dsrf_current = principal_payment_with_excel_updates(turbo_bonds, year, available_revs, dsrf_current)
            
                # December Interest Payment Serial Bonds #    
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(serial_bonds, "December", available_revs, dsrf_current, december_interest_payments, year)

                # December Interest First Estimate Turbo Bonds #
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(turbo_bonds, "December Estimate", available_revs, dsrf_current, december_interest_payments, year)
                
                #########################
                ##### Turbo Payment #####
                #########################
                
                if available_revs > december_interest_payments:
                
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
                    available_revs = turbo_payment_with_excel_updates(turbo_bonds, year, amount_to_turbo, available_revs)
                
                # December Interest Payment Turbo Bonds #    
                available_revs, dsrf_current, december_interest_payments = interest_payment_with_excel_updates(turbo_bonds, "December", available_revs, dsrf_current, december_interest_payments, year)
          
            #############################
            ##### Default Condition #####    
            #############################
            
            else:
            
                default_has_occurred = True
                year_of_default = year 
                
                for bond in serial_bonds:
                    if bond.is_outstanding():
                        worksheet_serials.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 10) + ":" + col_dict[bond.home_column_int +4] + str(year - start_year + 10)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_serials.range(col_dict[bond.home_column_int + 3] + str(4)).value = "Yes"
                for bond in turbo_bonds:
                    if bond.is_outstanding():
                        worksheet_turbos.range(col_dict[bond.home_column_int - 1] + str(year - start_year + 10) + ":" + col_dict[bond.home_column_int +5] + str(year - start_year + 10)).color = (217, 105, 105)
                        bond.is_defaulted = True
                        worksheet_turbos.range(col_dict[bond.home_column_int + 3] + str(4)).value = "Yes"
                        
                worksheet_rev.range("A" + str(year - start_year + 5) + ":B" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_rev.range("H" + str(year - start_year + 5) + ":I" + str(year - start_year + 5)).color = (217, 105, 105)
                worksheet_dsrf.range("A" + str(year - start_year + 7)).color = (217, 105, 105)
            
                available_revs, dsrf_current = enter_default_with_excel_updates(default_scenario, turbo_bonds, serial_bonds, year, available_revs, dsrf_current, year_of_default)
            
        ######################################################   
        #################### Housekeeping ####################
        ######################################################
    
        if bonds_outstanding:
        
            # Serial Bond Year-End Balances #
    
            for bond in serial_bonds:
                bond.year_end_balances[year] = bond.amount_outstanding
                
                if year == start_year:
                    worksheet_serials.range(col_dict[bond.home_column_int + 4] + str(year - start_year + 10)).value = bond.amount_outstanding
                elif bond.year_end_balances[year - 1] != 0:
                    if bond.is_defaulted:
                        worksheet_serials.range(col_dict[bond.home_column_int + 4] + str(year - start_year + 10)).color = (217, 105, 105)
                    worksheet_serials.range(col_dict[bond.home_column_int + 4] + str(year - start_year + 10)).value = bond.amount_outstanding
    
            # Turbo Bonds Year-End Balances #    
        
            for bond in turbo_bonds:
                bond.year_end_balances[year] = bond.amount_outstanding
            
                if year == start_year:
                    worksheet_turbos.range(col_dict[bond.home_column_int + 5] + str(year - start_year + 10)).value = bond.amount_outstanding
                elif bond.year_end_balances[year - 1] != 0:
                    if bond.is_defaulted:
                        worksheet_turbos.range(col_dict[bond.home_column_int + 5] + str(year - start_year + 10)).color = (217, 105, 105)
                    worksheet_turbos.range(col_dict[bond.home_column_int + 5] + str(year - start_year + 10)).value = bond.amount_outstanding
        
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
            excess_revs[year] = available_revs
    
    calc_irr_with_excel_updates(serial_bonds, turbo_bonds)
    worksheet_rev.range("E16").value = "Done"
    
    return turbo_bonds, serial_bonds, pledged_revs, dsrf_current, default_has_occurred, year_of_default


# Running w/ Excel Updates #
turbo_bonds, serial_bonds, pledged_revs, dsrf_current, default_has_occurred, year_of_default = run_model_with_excel_updates(start_year, end_year, turbo_bonds, serial_bonds, pledged_revs, dsrf_current, worksheet_dsrf, worksheet_rev, worksheet_serials, worksheet_turbos, default_has_occurred)