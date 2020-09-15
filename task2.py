from Google import Create_Service  #this module from https://learndataanalysis.org/
import pandas as pd
from bs4 import BeautifulSoup


CLIENT_SECRET_FILE = 'credentials.json'
API_SERVICE_NAME = 'sheets'
API_VERSION = 'v4'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

service = Create_Service(CLIENT_SECRET_FILE, API_SERVICE_NAME, API_VERSION, SCOPES)

spreadsheet_id = '1bcEBETPaKJSciHS7PPLtL-rDTna8gYXQKbp93NjSoqs'


files = ['Lace-Ups for Men _ Dolce&Gabbana', 'Men\'s Loafers and Moccasins _ Dolce&Gabbana', \
'Men\'s Sneakers _ Dolce&Gabbana', 'Men\'s Boots _ Dolce&Gabbana', 'Men\'s Sandals and Slides _ Dolce&Gabbana']

lst = ['Men - Lace-Ups', 'Men - Loafers and Moccasins', 'Men - Sneakers', 'Men - Boots', 'Men - Sandals and Slides']

for file_html, z in zip(files, lst):
	with open(file_html+'.htm', "r") as f:
	    contents = f.read()
	    soup = BeautifulSoup(contents, 'lxml')

	    '''data-brand'''
	    product_name=[]
	    item_id=[]
	    price=[]
	    variant=[]
	    brand=[]
	    category=[]
	    id=[]
	    dimension9=[]
	    dimension10=[]
	    metric=[]
	    for v in soup.find_all('div', {'class': 'b-product_tile js-product_tile'}):
	    	product_name.append(v['data-product-name'])
	    	item_id.append(v['data-itemid'])
	    	price.append(v['data-price'][:-3])
	    	variant.append(v['data-variant'])
	    	brand.append(v['data-brand'])
	    	category.append(v['data-category'])
	    	id.append(v['id'])
	    	dimension9.append(v['data-dimension9'])
	    	dimension10.append(v['data-dimension10'])
	    	metric.append(v['data-metric2'][:-3])


	    '''delete repeated links'''
	    for v in soup.find_all('div', {'class': 'b-product_tile-body js-product_tile-body'}):
	    	v.decompose()

	    url=[]
	    alt_image=[]
	    for v in soup.find_all('a', {'class': 'js-producttile_link b-product_image-wrapper'}):
	    	url.append(v['href'])
	    	alt_image.append(v['data-altimage'])


	    '''src of image 2 classes'''
	    src=[]
	    for v in soup.find_all('img', {'class': ['js-producttile_image b-product_image lazy-loaded', 'js-producttile_image b-product_image lazy-hidden']}):
		    src.append(v['src'])



	    '''sizes(in the example (appendix) 39.5 is not available for first product) and one_link'''
	    sizes=[]
	    one_link=[]
	    for v in item_id:
	    	lis = []
	    	link=''
	    	div = soup.find('div', {'data-itemid': v})
	    	for d in div.find_all('a', {'class': ['js-togglerhover b-swatches_size-link', 'js-swatchanchor b-swatches_size-link']}):
	    	# for d in div.find_all('a', {'class': 'js-togglerhover b-swatches_size-link'}): # solution for only available sizes
	    		size=d.text.rstrip()[1:]
	    		lis.append(size)
	    		if not link and d.get('href'):
	    			link=d['href']
	    	sizes.append(','.join(lis))
	    	one_link.append(link)


	    dict = {
	    	'data-product-name': product_name,
	    	'data-itemid': item_id,
	    	'data-price': price,
	    	'data-variant': variant,
	    	'data-brand': brand,
	    	'URL': url,
	    	'data-altimage': alt_image,
	    	'src': src,
	    	'data-category': category,
	    	'id': id,
	    	'data-dimension9': dimension9,
	    	'data-dimension10': dimension10,
	    	'data-metric2': metric,
	    	'sizes': sizes,
	    	'one-link': one_link,
	    }

	    df = pd.DataFrame.from_dict(dict)

	    '''write to google sheets'''
	    worksheet_name = z+'!'
	    cell_range_insert = 'A2'
	    values = df.values.tolist()  #dataframes to list
	    value_range_body = {
	    	'majorDimension': 'ROWS',
			'values': values
		}
	    service.spreadsheets().values().update(
	    	spreadsheetId=spreadsheet_id, 
	    	valueInputOption='USER_ENTERED', 
	    	range=worksheet_name + cell_range_insert, 
	    	body=value_range_body
	    ).execute()
