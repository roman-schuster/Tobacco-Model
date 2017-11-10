import xlwings as xw

workbook = xw.Book(r'C:\path\to\file.xlsx')

worksheet_turbos = workbook.sheets['Turbo Bonds']
worksheet_serials = workbook.sheets['Serial Bonds']
worksheet_dsrf = workbook.sheets['DSRF']
worksheet_rev = worksheet.sheets['Pledged Revenue']

num_turbos = worksheet_turbos.range('B1').value
num_serials = worksheet_serials.range('B1').value

turbo_bonds = []
serial_bonds = []
pledged_revs = {}
dsrf_max = worksheet_dsrf.range('B3').value
dsrf_current = worksheet_dsrf.range('B4').value
current_year = 2017


turbo_ints_to_cols = {0:'B', 1:'H', 2:'N', 3:'T', 4:'Z', 5:'AF',
    6:'AL', 7:'AR', 8:'AX', 9:'BD', 10:'BJ',
    11:'BP', 12:'BV', 13:'CB', 14:'CH', 15:'CN',
    16:'CT', 17:'CZ', 18:'DF', 19:'DL', 20:'DR'}

serial_ints_to_cols = {0:'B', 1:'G', 2:'L', 3:'Q', 4:'V', 5:'AA',
    6:'AF', 7:'AK', 8:'AP', 9:'AU', 10:'AZ',
    11:'BE', 12:'BJ', 13:'BO', 14:'BT', 15:'BY',
    16:'CD', 17:'CI', 18:'CN', 19:'CS', 20:'CX'}

for i in range(num_turbos):
    cusip = worksheet_turbos.range(turbo_ints_to_cols[i] + cstr(3)).value
    maturity = worksheet_turbos.range(turbo_ints_to_cols[i] + cstr(4)).value
    coupon = worksheet_turbos.range(turbo_ints_to_cols[i] + cstr(5)).value
    prop_of_revs = worksheet_turbos.range(turbo_ints_to_cols[i] + cstr(6)).value
    amt_outstanding = worksheet_turbos.range(turbo_ints_to_cols[i] + cstr(7)).value
    
    turbo_bonds.append(Turbo_Bond(cusip, maturity, coupon, prop_of_revs, amt_outstanding))

for i in range(5, 88):
    pledged_revs[current_year + (i - 5)] = worksheet_rev.range('B' + cstr(i)).value
    
for i in range(num_serials):
    cusip = worksheet_serials.range(serial_ints_to_cols[i] + cstr(3)).value
    maturity = worksheet_serials.range(serial_ints_to_cols[i] + cstr(4)).value
    coupon = worksheet_serials.range(serial_ints_to_cols[i] + cstr(5)).value
    amt_outstanding = worksheet_serials.range(serial_ints_to_cols[i] + cstr(6)).value
    
    serial_bonds.append(Serial_Bond(cusip, maturity, coupon, amt_outstanding))
    
  
