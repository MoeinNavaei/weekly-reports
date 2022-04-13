
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from bs2json import bs2json
import re
import requests
from urllib.request import urlopen as uReq
import time
import sys

############# start crawl #############
page1 = 'https://www.filimo.com/tag/thumbnailspecial'
page2 = 'https://www.filimo.com/tag/FilimoNewShows'
page3 = 'https://www.filimo.com/tag/10052392'

page4 = 'https://www.filimo.com/movies/FilimoNewShows'
page5 = 'https://www.filimo.com/series/FilimoNewShows'

page6 = 'https://www.filimo.com/tag/historic'
page7 = 'https://www.filimo.com/tag/family'
page8 = 'https://www.filimo.com/tag/comedy'
page9 = 'https://www.filimo.com/tag/plays'
page10 = 'https://www.filimo.com/tag/horror'
page11 = 'https://www.filimo.com/tag/short'
page12 = 'https://www.filimo.com/tag/concert'
page13 = 'https://www.filimo.com/tag/action'
page14 = 'https://www.filimo.com/tag/romance'
page15 = 'https://www.filimo.com/tag/documentary'
page16 = 'https://www.filimo.com/tag/animated'
page17 = 'https://www.filimo.com/tag/sci-fi'
page18 = 'https://www.filimo.com/tag/talkshow'
page19 = 'https://www.filimo.com/tag/turkey'

pages = pd.DataFrame()
pages.loc[0, 'page'] = page1
pages.loc[1, 'page'] = page2
pages.loc[2, 'page'] = page3
pages.loc[3, 'page'] = page4
pages.loc[4, 'page'] = page5
pages.loc[5, 'page'] = page6
pages.loc[6, 'page'] = page7
pages.loc[7, 'page'] = page8
pages.loc[8, 'page'] = page9
pages.loc[9, 'page'] = page10
pages.loc[10, 'page'] = page11
pages.loc[11, 'page'] = page12
pages.loc[12, 'page'] = page13
pages.loc[13, 'page'] = page14
pages.loc[14, 'page'] = page15
pages.loc[15, 'page'] = page16
pages.loc[16, 'page'] = page17
pages.loc[17, 'page'] = page18
pages.loc[18, 'page'] = page19

list_title = pd.DataFrame()
i = 0
for k in range(0, len(pages)):
    per_link = pages.loc[k, 'page']
    link_all = requests.get(per_link)
    soup_title = BeautifulSoup(link_all.text, 'html.parser')
    for link_item in soup_title.findAll('div', {'class': "item"}):
        for link_link in link_item.findAll('a'):
            list_title.loc[i, 'LinkAddress'] = link_link.get('href')
            i = i + 1
            print("i: ", i)

list_title = list_title.fillna('AAAA')
list_title = list_title[~list_title.LinkAddress.str.contains('AAAA')]
list_title = list_title [list_title.LinkAddress.str.contains('/m/')]
list_title.drop_duplicates(subset =['LinkAddress'], keep = 'last', inplace = True)
list_title = list_title.reset_index()
del list_title['index']

filimo_output = pd.DataFrame()

for i in range(0, len(list_title)):    #   len(list_title)
    print(i)
    LinkAddress = list_title.loc[i, 'LinkAddress']
#    print(LinkAddress)
    site = requests.get(LinkAddress)
    if site:
        soup_second = BeautifulSoup(site.text, 'html.parser')
        try:
            title = soup_second.find('div', {'class': "fa-title ui-fw-semibold"})
            filimo_output.loc[i, 'Title'] = title.text
        except: pass
        try:
            rate_filimo1 = soup_second.find('span', {'class': "rate_cnt"})
            filimo_output.loc[i, 'Like'] = rate_filimo1.text
        except: pass
        try:
            genre1=soup_second.findAll('li', {'class': "ui-ml-2x"})            
            genre2=""
            for genre_per in genre1:
                genre2=genre2+','+genre_per.text
            filimo_output.loc[i, 'Genres']=genre2
            filimo_output['Genres'] = filimo_output['Genres'].str.strip()
        except: pass
        try:
            MetaData1=soup_second.find('tr', {'class': "details_poster-description-more ui-mb-6x d-flex"})
            MetaData=MetaData1.text.split('-')
            Year=re.findall('\d{4}', MetaData1.text )
            filimo_output.loc[i, 'Year']=Year
            MetaData = pd.DataFrame({'col': MetaData})
            for word in range(0, len(MetaData)):
                meta = MetaData.loc[word, 'col']
                if "محصول" in meta:
                    filimo_output.loc[i, 'Country']=meta
        except: pass

filimo_output['Like'] = filimo_output['Like'].str.strip()
filimo_output['Title'] = filimo_output['Title'].str.strip()
filimo_output['Country'] = filimo_output['Country'].str.strip()
filimo_output['Like'] = filimo_output['Like'].astype(str)
filimo_output = filimo_output[~filimo_output.Like.str.contains("nan")]
filimo_output['Like'] = filimo_output['Like'].astype(int)

filimo_output1 = filimo_output.copy()
filimo_output = filimo_output1.copy()

filimo_output['Year'] = filimo_output['Year'].astype(str)
filimo_output['Year'] = filimo_output['Year'].str.replace('[', '')
filimo_output['Year'] = filimo_output['Year'].str.replace(']', '')
filimo_output['Year'] = filimo_output['Year'].str.replace("'", '')
filimo_output['Year'] = filimo_output['Year'].str.strip()
#filimo_output.to_excel('E:python codesweekly reportsvod_weekly\filimo_output.xlsx', index = False)
#filimo_output.to_excel('filimo_output.xlsx', index=False)

############# comparision with previous data #############
filimo_previous = pd.read_excel(r'E:\python codes\weekly reports\vod_weekly\filimo_output.xlsx')
filimo_previous['Year'] = filimo_previous['Year'].astype(str)
############# Visit #############
del filimo_previous['Genres']
filimo_previous = filimo_previous.rename(columns={"Like":"Like_previous"})
filimo_merge = pd.merge(filimo_output, filimo_previous, on = ['Title', 'Year', 'Country'])
filimo_merge['Visit'] = filimo_merge['Like'] - filimo_merge['Like_previous']

del filimo_merge['Like']
del filimo_merge['Like_previous']
Visit = filimo_merge.copy()
Visit.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
Visit = Visit.reset_index()
del Visit['index']
Visit = Visit.loc[0:9]
Visit.to_excel('محتواهای پربازدید VOD.xlsx', index=False)
#Visit.to_excel('E:\python codes\weekly reports\vod_weekly\Visit.xlsx', index = False)
############# Genres #############
Genres_primary = pd.DataFrame()
Genres_primary['Genres'] = filimo_merge['Genres']
Genres_primary['Visit'] = filimo_merge['Visit']
Genres_primary['Genres'] = Genres_primary['Genres'].str.strip()
Genres_primary['Genres'].replace('', 'nan', inplace=True)
Genres_primary = Genres_primary.fillna('nan')
Genres_primary = Genres_primary[~Genres_primary.Genres.str.contains('nan')]
Genres_primary['Genres'] = Genres_primary['Genres'].str.replace(',', '،')
Genres_primary = Genres_primary.reset_index()
del Genres_primary['index']
Genres = pd.DataFrame()
for i in range(0, len(Genres_primary)):
    print(i)
    Genres_primary1 = Genres_primary.loc[i, 'Genres']
    Genres_primary1 = Genres_primary1.split('،')
    Genres_primary_df = pd.DataFrame({'Genres': Genres_primary1})
    Genres_primary_df.insert(1, 'Visit', '')
    for j in range(0, len(Genres_primary_df)):
        Genres_primary_df.loc[j, 'Visit'] = Genres_primary.loc[i, 'Visit']
    Genres = Genres.append(Genres_primary_df)

Genres['Genres'] = Genres['Genres'].str.strip()
Genres = Genres.reset_index()
del Genres['index']
Genres['Genres'] = Genres['Genres'].str.strip()
Genres['Genres'].replace('', 'nan', inplace=True)
Genres = Genres.fillna('nan')
Genres = Genres[~Genres.Genres.str.contains('nan')]
Genres = Genres.reset_index()
del Genres['index']

Genres = Genres.groupby(['Genres']).sum().reset_index()
Genres.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
Genres = Genres.reset_index()
del Genres['index']
Genres = Genres.loc[0:9]
Genres.to_excel('ژانرهای پربازدید VOD.xlsx', index=False)
#Genres.to_excel('E:\python codes\weekly reports\vod_weekly\Genres.xlsx', index = False)
############# End #############
#filimo_output.to_excel('E:python codesweekly reportsvod_weekly\filimo_output.xlsx', index = False)
filimo_output.to_excel('filimo_output.xlsx', index=False)
















