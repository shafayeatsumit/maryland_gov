from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import xlsxwriter

base_url = "http://sdat.dat.maryland.gov/RealProperty/Pages/default.aspx"

driver = webdriver.Firefox()
driver.get(base_url)

#page one select county and search method from dropdown
user_selected_county = "DORCHESTER COUNTY"
user_search_method = "PROPERTY ACCOUNT IDENTIFIER"
county_css_id = "MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucSearchType_ddlCounty"
search_method = "MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucSearchType_ddlSearchType"

select = Select(driver.find_element_by_id(county_css_id))
select.select_by_visible_text(user_selected_county)
time.sleep(2)

select = Select(driver.find_element_by_id(search_method))
select.select_by_visible_text(user_search_method)
time.sleep(2)

#click to continue go to second page
click_continue_id = "MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_StartNavigationTemplateContainerID_btnContinue"
driver.find_element_by_id(click_continue_id).click()
time.sleep(2)

#page two select District and account id from deopdown

user_input_district = '01'
user_input_account_id = '000713'
district_css_id = "MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucEnterData_txtDistrict"
account_id_css_id = "MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucEnterData_txtAccountIdentifier"

#giving input
time.sleep(1)
element_id = driver.find_element_by_id("MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucEnterData_txtDistrict")
print element_id
time.sleep(1)
element_id.send_keys(user_input_district)

element_id = driver.find_element_by_id("MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucEnterData_txtAccountIdentifier")
time.sleep(1)
element_id.send_keys(user_input_account_id)

#click to continue got to detail page
click_continue_id = "MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_StepNavigationTemplateContainerID_btnStepNextButton"
driver.find_element_by_id(click_continue_id).click()
time.sleep(2)
##Detail page
all_keys = driver.find_elements_by_css_selector('td a')
all_values = driver.find_elements_by_css_selector('td span')
# import xlsxwriter

row = 0
col = 0


workbook = xlsxwriter.Workbook('sumit.xls')
worksheet = workbook.add_worksheet()

format = workbook.add_format()
format.set_bold()

for i in all_keys[4:]:
	txt = i.text
	print txt
	worksheet.write_string(row, col,txt, format)
	col = col+ 1
	


#list of xpath
key_list = [i.text for i in all_keys[4:17]]


owner_name = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblOwnerName_0"]'
use = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblUse_0"]'
principal_residence = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblPrinResidence_0"]'
mailing_address ='//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblMailingAddress_0"]'
deed_reference = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblDedRef_0"]'
premises_address = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblPremisesAddress_0"]'
legal_description = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_lblLegalDescription_0"]'
map_xpath = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label5_0"]'
grid = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label6_0"]'
parcel = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label7_0"]'
sub_district = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label8_0"]'
subdivision = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label9_0"]'
block = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label11_0"]'
lot = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label12_0"]'
assessment_year = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label13_0"]'
plat_no = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label1_0"]'
plat_ref = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label14_0"]'
speacial_tax_areas = ''
town = ''
ad_valorem = ''
tax_class = ''
primary_structure_built = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label18_0"]'
above_grade_living_area = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label19_0"]'
finished_basement_area = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label27_0"]'
property_land_area = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label20_0"]'
county_use = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label21_0"]'
stories = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label22_0"]'
basement = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label23_0"]'
type_xpath = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label24_0"]'
exterior = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label25_0"]'
full_half_bath = '<span id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label34_0" class="text">1 full</span>'
garage = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label35_0"]'
last_major_rivision = '//*[@id="MainContent_MainContent_cphMainContentArea_ucSearchType_wzrdRealPropertySearch_ucDetailsSearch_dlstDetaisSearch_Label36_0"]'

all_values_xpath =[
	owner_name,
	use,
	principal_residence,
	mailing_address,
	deed_reference,
	premises_address,
	legal_description,
	map_xpath,
	grid,
	parcel,
	sub_district,
	subdivision,
	block,
	lot,
	assessment_year,
	plat_no,
	plat_ref,
	speacial_tax_areas,
	town,
	ad_valorem,
	tax_class,
	primary_structure_built,
	above_grade_living_area,
	finished_basement_area,
	property_land_area,
	county_use,
	stories,
	basement,
	type_xpath,
	exterior,
	full_half_bath,
	garage,
	last_major_rivision
]
row = 1
col = 0
for i in all_values_xpath:
	xpath = str(i)
	try:
		value = driver.find_element_by_xpath("%s"%xpath).text
		worksheet.write(row, col,value)
	except Exception as e:
		print e
	col = col + 1
workbook.close()
driver.quit()