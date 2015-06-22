import xlrd
import csv
import sys
import argparse



parser = argparse.ArgumentParser(description="Convert from HCR excel files to Kaprica csv files")
parser.add_argument('infile', metavar='In', type=str, help="the location of the input excel file")
parser.add_argument('outfile', metavar='Out', type=str, help="the location of the output csv template")
parser.add_argument('--users', dest='csv_type', const='user', default='device', action='store_const')

csv_type = 'device'
constants = {'device_model':'SM-N910V', 'default_config_name':'HCR', 'HCREnrollMobileIron/server':'m.mobileiron.net:10261', 'email/device model/device serial':'/SM-N910V/TEMP'}

def main():
	args = parser.parse_args()
	conversions = {}
	if args.csv_type == 'user':
		csv_type = 'user'
		conversions['email'] = {'sources':['Email']}
		conversions['name'] = {'sources':['First Name', 'Last Name'], 'delimeter':' '}
	else:
		csv_type = 'device'
		conversions['HCREnrollMobileIron/username'] = {'sources':['AD Account']}
		conversions['HCREnrollMobileIron/password'] = {'sources':['AD Password']}
		conversions['HCREnrollMobileIron/devicepw'] = {'sources':['Default PIN']}
		conversions['DoEnc/devicepw'] = {'sources':['Default PIN']}
		conversions['GooglePlay/username'] = {'sources':['Gplay Account']}
		conversions['GooglePlay/password'] = {'sources':['Gplay Password']}
		conversions['email/device model/device serial'] = {'sources':['Email']}


	excel_to_unaltered_csv(args.infile, 'temp.csv')
	reformat_csv('temp.csv', args.outfile, conversions)


def excel_to_unaltered_csv(excel_in, csv_out):
	workbook = xlrd.open_workbook(excel_in)
	worksheet = workbook.sheet_by_index(0)
	csvfile = open(csv_out, 'wb')
	wr = csv.writer(csvfile, quoting=csv.QUOTE_ALL)
	for rownum in xrange(worksheet.nrows):
		wr.writerow(list(x.encode('utf-8') if type(x) == type(u'') else x for x in worksheet.row_values(rownum)))
	csvfile.close()

def reformat_csv(csv_in, csv_out, conversions):
    out_headers = []
    with open(csv_out, 'rU') as csvoutfile:
    	reader = csv.reader(csvoutfile)
    	out_headers = reader.next()
    	for header in out_headers:
    		print header
    with open(csv_in, 'rU') as csvfile, open(csv_out, 'wb') as csvoutfile:
    	inreader = csv.reader(csvfile)
    	outwriter = csv.writer(csvoutfile)
    	outwriter.writerow(out_headers)
    	headers = inreader.next()
    	for row in inreader:
    		output = [None]*len(out_headers)
    		for out_header in out_headers:
    			try:
    				output_conversion = conversions[out_header]
    				if len(output_conversion['sources']) == 1:
    					output[out_headers.index(out_header)] = row[headers.index(output_conversion['sources'][0])]
    				else:
    					output[out_headers.index(out_header)] = ""
    					for source in output_conversion['sources']:
    						output[out_headers.index(out_header)] += (row[headers.index(source)] + output_conversion['delimeter'])
    					output[out_headers.index(out_header)] = output[out_headers.index(out_header)][:-1*len(output_conversion['delimeter'])]
    			except KeyError:
    				print "No source given to fill column " + out_header
    		for constant in constants:
  				try:
  					index = out_headers.index(constant)
  					if index >= 0 and output[index] != None:
  						output[index] += constants[constant]
  					elif index >= 0:
  						output[index] = constants[constant]
  				except ValueError:
  					pass
    		outwriter.writerow(output)

main()

#Conversion format {"excel_header_name": ["csv_header_name_1", "csv_header_name_2"]}
#New conversion fromat {"csv_header_name": {"sources":["excel_header_1", "excel_header_2"], "delimeter":<delim>}}