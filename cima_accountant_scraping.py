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
            
            data_wrapper =  soup_page_data.find('dl', class_='searchResultDetails-list')
            # List data (table)
            list_data = data_wrapper.find_all('dd')
            # List data (functional specialisms)
            list_functional_specialisms = data_wrapper.find('ul').find_all('li')

            # Company Name
            try:
                company_name = list_data[0].text.strip()
            except:
                company_name = '-'
            print('Company name: ' + str(company_name))

            # Surname
            try:
                surname = list_data[1].text.strip()
            except:
                surname = '-'
            print('Surname: ' + str(surname))

            # First Name
            try:
                first_name = list_data[2].text.strip()
            except:
                first_name = '-'
            print('First name: ' + str(first_name))

            # Mobile Phone
            try:
                mobile_phone = list_data[3].text.strip()
            except:
                mobile_phone = '-'
            print('Mobile phone: ' + str(mobile_phone))
            
            # Work Phone
            try:
                work_phone = list_data[4].text.strip()
            except:
                work_phone = '-'
            print('Work phone: ' + str(work_phone))

            # Email Address
            try:    
                email_address = list_data[5].text.strip()
            except: 
                email_address = '-'
            print('Email address: ' + str(email_address))

            # Homepage
            try:
                homepage = list_data[6].text.strip()
            except:
                homepage = '-'
            print('Homepage: ' + str(homepage))

            # Address Line 1
            try:
                address_line_1 = list_data[7].text.strip()
            except:
                address_line_1 = '-'
            print('Address line 1: ' + str(address_line_1))

            # Address Line 2
            try:
                address_line_2 = list_data[8].text.strip()
            except:
                address_line_2 = '-'
            print('Address line 2: ' + str(address_line_2))

            # Address Line 3
            try:
                address_line_3 = list_data[9].text.strip()
            except:
                address_line_3 = '-'
            print('Address line 3: ' + str(address_line_3))

            # Postal code
            try:
                postal_code = list_data[10].text.strip()
            except:
                postal_code = '-'
            print('Postal code: ' + str(postal_code))

            # City
            try:
                city = list_data[11].text.strip()
            except:
                city = '-'
            print('City: ' + str(city))

            # County
            try:
                county = list_data[12].text.strip()
            except:
                county = '-'
            print('County: ' + str(county))
            
            # Country
            try:
                country = list_data[13].text.strip()
            except:
                country = '-'
            print('Country: ' + str(country))

            # Size of Practices
            try:
                size_of_practices = list_data[14].text.strip()
            except:
                size_of_practices = '-'
            print('Size of practices: ' + str(size_of_practices))           

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
                "Size of Practice": size_of_practices
            }

    # data.append(dict_result) 
    # print('updated')

def main():
    cima_scraping()

if __name__ == '__main__':
    main()
    

