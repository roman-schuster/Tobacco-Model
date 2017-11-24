class Bond:
	'''
	Muni bond class - please only use cusips as variable names for these bonds
	
	maturity_year: int
	coupon: float
	amount_outstanding: long int
	proportion_of_revenue: float (should be 100.00 if unique maturity)
	lien_priority: int (must be lower than global max_lien_priority)
	
	matured: bool
	june_coupon_history and dec_coupon_history are lists used in data visualization
	'''
	
	matured = False
	defaulted = False
	june_coupon_history = []
	dec_coupon_history = []
	
	def __init__(self, maturity_year, coupon, amount_outstanding, lien_priority, home_column_int):							
		self.maturity_year = maturity_year
		self.coupon = coupon
		self.amount_outstanding = amount_outstanding
		self.lien_priority = lien_priority
		self.home_column_int = home_column_int

	def calc_interest_payment():			
		return (self.coupon * self.amount_outstanding)
    
    def is_maturing(year):
        if self.year == year:
            return True
        return False
	
	def is_matured():
		return self.matured
	
    def is_defaulted():
	    return self.defaulted

    def is_outstanding():
		if (not self.matured) and (not self.defaulted) and (self.amount_outstanding > 0):
			return True
		return False
	
	def pay_interest(month):
	    '''
        Calculates current interest and adds to appropriate payment history
        '''
		interest_payment = self.calc_interest_payment()
		
		if month = 'December':
			dec_coupon_history.append(interest_payment)
		elif month = "June":
			june_coupon_history.append(interest_payment) 				
	
class TurboBond(Bond):
	'''
	Inherits from class Bond
	'''
	
	turbo_payment_history = []
	
	
class SerialBond(Bond):
	
def calc_turbo(turbo_bonds):
	total_december_payments = 0
	for bond in turbo_bonds:
		total_december_payments += bond.calc_interest_payment()
			
