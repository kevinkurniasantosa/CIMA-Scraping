import requests
import csv
import copy
import datetime
import time
import random
from threading import Thread
import smtplib
import mimetypes
import string
import os
import os.path
import traceback
import pprint
from bs4 import BeautifulSoup
import math
import calendar
import time
import logging
import pandas as pd
from itertools import islice
from pathlib import Path
from datetime import datetime
import json
import urllib
from bs4 import NavigableString as nav
import re
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import math

print('import success')
print('--------------------------')

result = {
    "Company Name": [],
    "Surname": [],
    "First Name": [],
    "Mobile Phone": [],
    "Work Phone": [],
    "Email Address": [],
    "Homepage": [],
    "Address Line 1": [],
    "Address Line 2": [],
    "Address Line 3": [],
    "Postal Code": [],
    "City": [],
    "County": [],
    "Country": [],
    "Size of Practice": [],
    "Functional Specialisms 1": [],
    "Functional Specialisms 2": [],
    "Functional Specialisms 3": [],
    "Functional Specialisms 4": [],
    "Functional Specialisms 5": [],
    "Functional Specialisms 6": [],
    "Functional Specialisms 7": [],
    "Functional Specialisms 8": [],
    "Functional Specialisms 9": [],
    "Functional Specialisms 10": [],
    "Functional Specialisms 11": [],
    "Functional Specialisms 12": [],
    "Functional Specialisms 13": [],
    "Functional Specialisms 14": [],
    "Functional Specialisms 15": []
}

excel_filename = 'CIMA scraping result.xlsx'
excel_sheetname = 'Sheet1'

def clean_string(x):
    try:
        x = x.replace('\"','\'\'').replace('\r',' ').replace('\n',' ')
        x = unicodedata.normalize('NFKD', x).encode('ascii', 'ignore')
        x = x.decode('ascii')
    except:
        x = '?'
    return x
    
def cima_scraping():
    main_url = 'https://www.cimaglobal.com/About-us/Find-a-CIMA-Accountant/'
    res = requests.get(main_url, headers={'User-Agent': 'Mozilla/5.0'})
    soup = BeautifulSoup(res.text, 'html.parser')
    
    num_of_data  = soup.find('div', class_='eventSearchBlock-results-stats').find_all('span')[1].text
    print('Number of data: ' + str(num_of_data))

    num_of_pages = int(num_of_data)/10
    num_of_pages = math.ceil(num_of_pages)
    print('Number of pages: ' + str(num_of_pages))

    print('--------------------------')

    # Loop through all pages
    for x in range(1, num_of_pages+1): # 1 - 172
        url_per_page = 'https://www.cimaglobal.com/About-us/Find-a-CIMA-Accountant/?p={}#Results'.format(x)
        print('Current URL Page: ' + str(url_per_page))

        # Request page for each URL page
        res_page = requests.get(url_per_page, headers={'User-Agent': 'Mozilla/5.0'})
        soup_page = BeautifulSoup(res_page.text, 'html.parser')
        # print(soup_page.prettify())

        accountant_content = soup_page.find('ul', class_='accountantListing')
        accountant_list = accountant_content.find_all('li', class_='accountantListing-item')

        # Loop through list of accountant
        for accountant_l in accountant_list:
            accountant_link = accountant_l.find('p', class_='accountantListing-button').find('a', href=True)['href']
            print('--------------------------')
            print(accountant_link)

            res_page_data = requests.get(accountant_link, headers={'User-Agent': 'Mozilla/5.0'})
            soup_page_data = BeautifulSoup(res_page_data.text, 'html.parser')
            
            # List data (table)
            data_wrapper =  soup_page_data.find('dl', class_='searchResultDetails-list')
            list_data = data_wrapper.find_all('dd')
            # List data (functional specialisms)
            data_functional_specialisms = soup_page_data.find('div', class_='searchResultDetails-details')
            list_functional_specialisms = data_functional_specialisms.find('ul').find_all('li')

            ### TABLE DATA
            # Company Name
            try:
                company_name = list_data[0].text.strip()
            except:
                company_name = '-'
            result['Company Name'].append(company_name)
            print('Company name: ' + str(company_name))

            # Surname
            try:
                surname = list_data[1].text.strip()
            except:
                surname = '-'
            result['Surname'].append(surname)
            print('Surname: ' + str(surname))

            # First Name
            try:
                first_name = list_data[2].text.strip()
            except:
                first_name = '-'
            result['First Name'].append(first_name)
            print('First name: ' + str(first_name))

            # Mobile Phone
            try:
                mobile_phone = list_data[3].text.strip()
            except:
                mobile_phone = '-'
            result['Mobile Phone'].append(mobile_phone)
            print('Mobile phone: ' + str(mobile_phone))
            
            # Work Phone
            try:
                work_phone = list_data[4].text.strip()
            except:
                work_phone = '-'
            result['Work Phone'].append(work_phone)
            print('Work phone: ' + str(work_phone))

            # Email Address
            try:    
                email_address = list_data[5].text.strip()
            except: 
                email_address = '-'
            result['Email Address'].append(email_address)
            print('Email address: ' + str(email_address))

            # Homepage
            try:
                homepage = list_data[6].text.strip()
            except:
                homepage = '-'
            result['Homepage'].append(homepage)
            print('Homepage: ' + str(homepage))

            # Address Line 1
            try:
                address_line_1 = list_data[7].text.strip()
            except:
                address_line_1 = '-'
            result['Address Line 1'].append(address_line_1)
            print('Address line 1: ' + str(address_line_1))

            # Address Line 2
            try:
                address_line_2 = list_data[8].text.strip()
            except:
                address_line_2 = '-'
            result['Address Line 2'].append(address_line_2)
            print('Address line 2: ' + str(address_line_2))

            # Address Line 3
            try:
                address_line_3 = list_data[9].text.strip()
            except:
                address_line_3 = '-'
            result['Address Line 3'].append(address_line_3)
            print('Address line 3: ' + str(address_line_3))

            # Postal code
            try:
                postal_code = list_data[10].text.strip()
            except:
                postal_code = '-'
            result['Postal Code'].append(postal_code)
            print('Postal code: ' + str(postal_code))

            # City
            try:
                city = list_data[11].text.strip()
            except:
                city = '-'
            result['City'].append(city)
            print('City: ' + str(city))

            # County
            try:
                county = list_data[12].text.strip()
            except:
                county = '-'
            result['County'].append(county)
            print('County: ' + str(county))
            
            # Country
            try:
                country = list_data[13].text.strip()
            except:
                country = '-'
            result['Country'].append(country)
            print('Country: ' + str(country))

            # Size of Practices
            try:
                size_of_practices = list_data[14].text.strip()
            except:
                size_of_practices = '-'
            result['Size of Practice'].append(size_of_practices)
            print('Size of practices: ' + str(size_of_practices))  

            ### FUNCTIONAL SPECIALISMS
            # FC 1
            try:
                fc_1 = list_functional_specialisms[0].text.strip()
            except:
                fc_1 = '-'
            result['Functional Specialisms 1'].append(fc_1)
            print('FC 1: ' + str(fc_1))

            # FC 2
            try:
                fc_2 = list_functional_specialisms[1].text.strip()
            except:
                fc_2 = '-'
            result['Functional Specialisms 2'].append(fc_2)
            print('FC 2: ' + str(fc_2))

            # FC 3
            try:
                fc_3 = list_functional_specialisms[2].text.strip()
            except:
                fc_3 = '-'
            result['Functional Specialisms 3'].append(fc_3)
            print('FC 3: ' + str(fc_3))

            # FC 4
            try:
                fc_4 = list_functional_specialisms[3].text.strip()
            except:
                fc_4 = '-'
            result['Functional Specialisms 4'].append(fc_4)
            print('FC 4: ' + str(fc_4))

            # FC 5
            try:
                fc_5 = list_functional_specialisms[4].text.strip()
            except:
                fc_5 = '-'
            result['Functional Specialisms 5'].append(fc_5)
            print('FC 5: ' + str(fc_5))

            # FC 6
            try:
                fc_6 = list_functional_specialisms[5].text.strip()
            except:
                fc_6 = '-'
            result['Functional Specialisms 6'].append(fc_6)
            print('FC 6: ' + str(fc_6))

            # FC 7
            try:
                fc_7 = list_functional_specialisms[6].text.strip()
            except:
                fc_7 = '-'
            result['Functional Specialisms 7'].append(fc_7)
            print('FC 7: ' + str(fc_7))

            # FC 8
            try:
                fc_8 = list_functional_specialisms[7].text.strip()
            except:
                fc_8 = '-'
            result['Functional Specialisms 8'].append(fc_8)
            print('FC 8: ' + str(fc_8))

            # FC 9
            try:
                fc_9 = list_functional_specialisms[8].text.strip()
            except:
                fc_9 = '-'
            result['Functional Specialisms 9'].append(fc_9)
            print('FC 9: ' + str(fc_9))

            # FC 10
            try:
                fc_10 = list_functional_specialisms[9].text.strip()
            except:
                fc_10 = '-'
            result['Functional Specialisms 10'].append(fc_10)
            print('FC 10: ' + str(fc_10))

            # FC 11
            try:
                fc_11 = list_functional_specialisms[10].text.strip()
            except:
                fc_11 = '-'
            result['Functional Specialisms 11'].append(fc_11)
            print('FC 11: ' + str(fc_11))

            # FC 12
            try:
                fc_12 = list_functional_specialisms[11].text.strip()
            except:
                fc_12 = '-'
            result['Functional Specialisms 12'].append(fc_12)
            print('FC 12: ' + str(fc_12))

            # FC 13
            try:
                fc_13 = list_functional_specialisms[12].text.strip()
            except:
                fc_13 = '-'
            result['Functional Specialisms 13'].append(fc_13)
            print('FC 13: ' + str(fc_13))

            # FC 14
            try:
                fc_14 = list_functional_specialisms[13].text.strip()
            except:
                fc_14 = '-'
            result['Functional Specialisms 14'].append(fc_14)
            print('FC 14: ' + str(fc_14))

            # FC 15
            try:
                fc_15 = list_functional_specialisms[14].text.strip()
            except:
                fc_15 = '-'
            result['Functional Specialisms 15'].append(fc_15)
            print('FC 15: ' + str(fc_15))

            ## STORE IN DICTIONARY
            dict_result = {
                "Company Name": company_name,
                "Surname": surname,
                "First Name": first_name,
                "Mobile Phone": mobile_phone,
                "Work Phone": work_phone,
                "Email Address": email_address,
                "Homepage": homepage,
                "Address Line 1": address_line_1,
                "Address Line 2": address_line_2,
                "Address Line 3": address_line_3,
                "Postal Code": postal_code,
                "City": city,
                "County": county,
                "Country": country,
                "Size of Practice": size_of_practices,
                "Functional Specialisms 1": fc_1,
                "Functional Specialisms 2": fc_2,
                "Functional Specialisms 3": fc_3,
                "Functional Specialisms 4": fc_4,
                "Functional Specialisms 5": fc_5,
                "Functional Specialisms 6": fc_6,
                "Functional Specialisms 7": fc_7,
                "Functional Specialisms 8": fc_8,
                "Functional Specialisms 9": fc_9,
                "Functional Specialisms 10": fc_10,
                "Functional Specialisms 11": fc_11,
                "Functional Specialisms 12": fc_12,
                "Functional Specialisms 13": fc_13,
                "Functional Specialisms 14": fc_14,
                "Functional Specialisms 15": fc_15
            }

    # Saved it as DataFrame
    df = pd.DataFrame(result)

    # Convert DataFrame to Excel file
    writer = pd.ExcelWriter(excel_filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name=excel_sheetname, index=False)
    writer.save()

    # data.append(dict_result) 
    # print('updated')

def main():
    cima_scraping()
    print('cima data successfully scraped!')

if __name__ == '__main__':
    main()
    

