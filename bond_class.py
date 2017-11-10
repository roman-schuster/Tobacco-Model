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
	june_coupon_history = []
	dec_coupon_history = []
	
	def __init__(self, maturity_year, coupon, amount_outstanding, proportion_of_revenue, lien_priority):							
		self.maturity_year = maturity_year
		self.coupon = coupon
		self.amount_outstanding = amount_outstanding
		self.proportion_of_revenue = proportion_of_revenue
		self.lien_priority = lien_priority

	def calc_interest_payment():			
		return (self.coupon * self.amount_outstanding)
		
	def pay_interest(month):
		interest_payment = self.calc_interest_payment()
		
		if month = 'December':
			dec_coupon_history.append(interest_payment)
		else:
			june_coupon_history.append(interest_payment)
			
		self.amount_outstanding = self.amount_outstanding - interest_payment
		
	def compile_coupon_history():
			
	
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
			
