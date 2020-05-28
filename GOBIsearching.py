""""
Katharine Frazier
NC State University Libraries
05/21/20
PURPOSE: Scrapes GOBI, an academic print/ebook vendor, for pricing information based on lists of in-demand items. 
"""

import pandas as pd
from pandas import ExcelWriter
import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
import requests
from fuzzywuzzy import fuzz, process
import re
from itertools import groupby

#create lists for title, author, and year from original file
titlelist = []
authorlist = []
datelist = []
#read in original file of report output
report = pd.read_excel(r'PATH_TO_REPORT')
#append column values from original file to lists
titlelist = report['title'].values
authorlist = report['author'].values
datelist = report['year'].values
#create dictionary to match titles with authors
zipObj = zip(titlelist, authorlist)
titlesandauthors = dict(zipObj)

#establish webdriver (ex: ChromeDriver)
browser = webdriver.Chrome(r'PATH_TO_CHROMEDRIVER')
#tell browser to fetch website
browser.get('http://www.gobi3.com')

#find username field, send username
userElem = browser.find_element_by_id('guser')
userElem.send_keys('USERNAME_HERE')

#find password field, send password
passwordElem = browser.find_element_by_id('gpword')
passwordElem.send_keys('PASSWORD_HERE')
passwordElem.submit()

#create lists for name, binding, year, and price extracted from GOBI
namelist = []
bindinglist = []
yearlist = []
pricelist = []

#create dictionary that will pair data with data type ("Year": 2020)
choices = {}

#Begin iterating through dictionary of title/author in GOBI
for k, v in titlesandauthors.items():
    #find search dropdown, click it
    searchElem = browser.find_element_by_id('menu_li2')
    searchElem.click()
    #find "standard" option: account for varying Xpaths through try/except
    try:
        standardElem = browser.find_element_by_xpath('/html[1]/body[1]/div[1]/div[2]/div[1]/div[12]/a[1]')
        standardElem.click()
    except:
        standardElem = browser.find_element_by_xpath("//a[contains(text(),'Standard')]")
    #find title search bar on the standard search page
    try:
        titleElem = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txttitle"]')))
    except:
        continue
    #send title from dictionary
    try:
        titleElem.send_keys(k)
    except TypeError:
        pass
    #find author search bar on the standard search page
    authorElem = browser.find_element_by_xpath('//*[@id="author"]')
    #send author from dictionary
    try:
        authorElem.send_keys(v)
    except TypeError:
        pass
    #find "search" button on standard search page; click it
    searchElem2 = browser.find_element_by_xpath('/html/body/table[2]/tbody/tr[2]/td[2]/div/span[5]/a/span[2]')
    try:
        searchElem2.click()
    #account for and try again after page timeout due to long search
    except selenium.common.exceptions.TimeoutException:
        pass
    #find all results on results page
    itemElem = browser.find_elements_by_xpath('//div[@id="containeritems"]/div')
    #begin iterating through the results
    for item in itemElem:
        #transform web element to text element
        individualitem = item.text
        #create and apply regex to find price
        priceRegex = re.compile(r'(\d+(?:\.\d+)\s)')
        rawprice = priceRegex.search(str(individualitem))
        #further refine result of price regex
        try:
            moneyRegex = re.compile(r'\d+(?:\.\d+)')
            cleanmoney = moneyRegex.search(str(rawprice.group()))
            price = cleanmoney.group()
        except AttributeError:
            pass
        #create and apply title regex
        nameRegex = re.compile(r'\b(?!Title:\b)+(\w+)+(\S)+\s+(\w+)+\s?(\w+)')
        name = nameRegex.search(str(individualitem))
        title = name.group()
        print(title)
        #create and apply date regex
        dateRegex = re.compile(r'Year:(\d{4})')
        rawdate = dateRegex.search(str(individualitem))
        #create and apply regex to further refine date regex results (remove 'Year')
        try:
            yearRegex = re.compile('\d{4}')
            cleandate = yearRegex.search(str(rawdate.group()))
            #ensure results are integers
            date = int(cleandate.group())
            #sort results
            date.sort()
        except AttributeError:
            pass
        #create and apply binding regex, account for potential errors (lack of binding info)
        try:
            bindingRegex = re.compile(r'Binding:(\w+)')
            ebookorprint = bindingRegex.search(str(individualitem))
            binding = ebookorprint.group()
        except AttributeError:
            pass
        #apply fuzzy string matching to find accurate matches (not 100 to account for punctuation)
        ratio = fuzz.partial_token_set_ratio(k,title)
        #send matches to lists; match set at 70% for best results
        if ratio > 70:
            namelist.append(title)
            bindinglist.append(binding)
            yearlist.append(date)
            pricelist.append(price)
    #add lists to results dictionary
    choices.update({'title':namelist})
    choices.update({'binding':bindinglist})
    choices.update({'date':yearlist})
    choices.update({'price':pricelist})
    
#create dataframe of results dictionary
choices.update({'date':yearlist})
gobidf = pd.DataFrame(choices)
#group dataframe by title
gobidf.groupby('title')
#create sub-dataframe for ebook results
ebookoption = gobidf[gobidf['binding'] == 'Binding:eBook']
#create sub-dataframe for print results
bestprint = gobidf[(gobidf['binding'] == 'Binding:Cloth') | (gobidf['binding'] == 'Binding:Paper')]

#create lists of titles from original file, print file, and ebook file
sirsititles = report['title'].values
printtitles = bestprint['title'].values
etitles = ebookoption['title'].values

#create list to contain confirmed correct print titles
correcttitles=[]
#compare each paper title to original title, keep only best match
for x in printtitles:
    correct = process.extractOne(x,sirsititles)
    correcttitles.append(correct[0])
#swap out list of print titles for list of confirmed correct print titles
bestprint['title'] = pd.Series(correcttitles)

#rename other column titles
bestprint.columns=['title','Format of new copy in GOBI', 'Publication year in GOBI (Print)', 'Price in GOBI (Print)']

#sort dataframe by date of publication, with newest of each title at top
bestprint.sort_values(by=['title','Publication year in GOBI (Print)'], ascending=False, inplace=True)

#drop duplicates (keep only top, or newest, for each title)
bestprint.drop_duplicates(subset='title', keep='first', inplace=True)

#merge print dataframe with original dataframe on title value
dfmerge = original.merge(bestprint,on='title',sort=False,left_index=False,right_index=True,copy=False, how= 'outer')

#create list to contain confirmed correct ebook titles
correctetitles = []
#compare each ebook title to original title, keep only best match
for x in etitles:
    correctE = process.extractOne(x,sirsititles)
    correctetitles.append(correctE[0])
#swap out list of ebook titles for confirmed correct ebook titles
ebookoption['title'] = pd.Series(correctetitles)

#rename other column titles
ebookoption.columns = ['title','Format of new copy in GOBI', 'Publication year in GOBI (Ebook)', 'Price in GOBI (Ebook)']
#sort dataframe by date of pubication, with newest of each title at top
ebookoption.sort_values(by=['Publication year in GOBI (Ebook)'], ascending=False, inplace=True)

#drop duplicates (keep only top, or newest, for each title)
ebookoption = ebookoption.drop_duplicates(subset='title', keep='first', inplace=False)

#merge ebook dataframe with other merged dataframe (print & original combined)
newmerge = dfmerge.merge(ebookoption, on='title', sort=False, left_index=False, right_index=True, copy=False, how='outer')

#send to excel file
newmerge.to_excel(r'PATH-TO-OUTPUT')
