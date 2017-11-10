class Bond:
	'''
	Muni bond class - please only use cusips to name these bonds. This avoids duplicates
	maturity_year: int
	coupon: float
	amount_outstanding: long int
	matured: bool
	proportion_of_revenue: optional bool
	'''
	 
	maturity_year =
	coupon = 
	amount_outstanding = 
	matured = False
	proportion_of_revenue = 100.00
	lien_priority = 1
	lien_max_priority = 
	
	def __init__(self, maturity_year):							
		self.maturity_year = maturity_year

	def calc_interest_payment():			
		return (self.coupon * self.amount_outstanding)
		
	def pay_interest():
		interest_payment = self.calc_interest_payment()
		
		

class TurboBond(Bond):

class SerialBond(Bond):
	
