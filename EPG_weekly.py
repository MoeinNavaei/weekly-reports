import pandas as pd
import numpy as np
import time
import xlsxwriter  
###### Get Data and Classification ######
EPG_main = pd.read_excel(r'E:\python mo\weekly reports\total_EPG.xlsx')

EPG_Repeat = pd.DataFrame()
EPG_Repeat['channel'] = EPG_main['نام شبکه']
EPG_Repeat['program'] = EPG_main['نام برنامه']
EPG_Repeat['Duration'] = EPG_main['مدت بازدید']
EPG_Repeat['Visit'] = EPG_main['تعداد بازدید']
EPG_Repeat['Type'] = EPG_main['جنس']
EPG_Repeat['Service'] = EPG_main['نوع']
EPG_Repeat['Date'] = EPG_main['تاریخ']
EPG_Repeat['Date'] = EPG_Repeat['Date'].astype(str).replace('\.0', '', regex=True)
EPG_Repeat['Date'] = EPG_Repeat['Date'].str[0:8]
EPG_Repeat = EPG_Repeat.groupby(['channel', 'program', 'Date', 'Type', 'Service', 'Visit']).sum().reset_index()

EPG_sima = EPG_Repeat.query('Service == "سراسری"')
EPG_sima_new = EPG_sima.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_radio = EPG_Repeat.query('Service == "رادیویی"')
EPG_radio = EPG_radio.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_ostani = EPG_Repeat.query('Service == "استانی"')
EPG_ostani = EPG_ostani.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_ekhtesasi = EPG_Repeat.query('Service == "اختصاصی"')
EPG_ekhtesasi = EPG_ekhtesasi.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_boronmarzi = EPG_Repeat.query('Service == "برون مرزی"')
EPG_boronmarzi = EPG_boronmarzi.groupby(['channel', 'program', 'Date']).sum().reset_index()

EPG_movie = EPG_sima.query('Type == "فیلم سینمایی"')
EPG_news = EPG_sima.query('Type == "اخبار"')
EPG_news_new = EPG_news.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_series = EPG_sima.query('Type == "مجموعه تلویزیونی"')
EPG_series_new = EPG_series.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_sport = EPG_sima.query('Type == "ورزشی"')
EPG_sport_new = EPG_sport.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_kid = EPG_sima.query('Type == "کودک"')
EPG_kid_new = EPG_kid.groupby(['channel', 'program', 'Date']).sum().reset_index()
EPG_documentary = EPG_sima.query('Type == "مستند"')
EPG_documentary_new = EPG_documentary.groupby(['channel', 'program', 'Date']).sum().reset_index()

######################################################
################## Types of Services #################
######################################################

###### Sima ######
## Visit ##
EPG_sima_new.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_sima_new = EPG_sima_new.reset_index()
del EPG_sima_new['index']
EPG_sima_visit = EPG_sima_new.loc[0:4, :]
#del EPG_sima_visit['Service']
#del EPG_sima_visit['Type']
del EPG_sima_visit['Duration']
## Duration ##
EPG_sima_new.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_sima_new = EPG_sima_new.reset_index()
del EPG_sima_new['index']
EPG_sima_duration= EPG_sima_new.loc[0:4, :]
del EPG_sima_duration['Visit']
###### radio ######
## Visit ##
EPG_radio.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_radio = EPG_radio.reset_index()
del EPG_radio['index']
EPG_radio_visit = EPG_radio.loc[0:4, :]
del EPG_radio_visit['Duration']
## Duration ##
EPG_radio.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_radio = EPG_radio.reset_index()
del EPG_radio['index']
EPG_radio_duration= EPG_radio.loc[0:4, :]
del EPG_radio_duration['Visit']
###### ostani ######
## Visit ##
EPG_ostani.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_ostani = EPG_ostani.reset_index()
del EPG_ostani['index']
EPG_ostani_visit = EPG_ostani.loc[0:4, :]
del EPG_ostani_visit['Duration']
## Duration ##
EPG_ostani.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_ostani = EPG_ostani.reset_index()
del EPG_ostani['index']
EPG_ostani_duration= EPG_ostani.loc[0:4, :]
del EPG_ostani_duration['Visit']
###### ekhtesasi ######
## Visit ##
EPG_ekhtesasi.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_ekhtesasi = EPG_ekhtesasi.reset_index()
del EPG_ekhtesasi['index']
EPG_ekhtesasi_visit = EPG_ekhtesasi.loc[0:4, :]
del EPG_ekhtesasi_visit['Duration']
## Duration ##
EPG_ekhtesasi.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_ekhtesasi = EPG_ekhtesasi.reset_index()
del EPG_ekhtesasi['index']
EPG_ekhtesasi_duration= EPG_ekhtesasi.loc[0:4, :]
del EPG_ekhtesasi_duration['Visit']
###### boronmarzi ######
## Visit ##
EPG_boronmarzi.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_boronmarzi = EPG_boronmarzi.reset_index()
del EPG_boronmarzi['index']
EPG_boronmarzi_visit = EPG_boronmarzi.loc[0:4, :]
del EPG_boronmarzi_visit['Duration']
## Duration ##
EPG_boronmarzi.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_boronmarzi = EPG_boronmarzi.reset_index()
del EPG_boronmarzi['index']
EPG_boronmarzi_duration= EPG_boronmarzi.loc[0:4, :]
del EPG_boronmarzi_duration['Visit']
##################

######################################################
################## Types of Contents #################
######################################################
###### movie ######
## Visit ##
EPG_movie.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_movie = EPG_movie.reset_index()
del EPG_movie['index']
EPG_movie_visit = EPG_movie.loc[0:4, :]
del EPG_movie_visit['Service']
del EPG_movie_visit['Type']
del EPG_movie_visit['Duration']
## Duration ##
EPG_movie.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_movie = EPG_movie.reset_index()
del EPG_movie['index']
EPG_movie_duration= EPG_movie.loc[0:4, :]
del EPG_movie_duration['Service']
del EPG_movie_duration['Type']
del EPG_movie_duration['Visit']
###### news ######
## Visit ##
EPG_news.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_news = EPG_news.reset_index()
del EPG_news['index']
EPG_news_visit = EPG_news.loc[0:4, :]
del EPG_news_visit['Service']
del EPG_news_visit['Type']
del EPG_news_visit['Duration']
## Duration ##
EPG_news.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_news = EPG_news.reset_index()
del EPG_news['index']
EPG_news_duration= EPG_news.loc[0:4, :]
del EPG_news_duration['Service']
del EPG_news_duration['Type']
del EPG_news_duration['Visit']
###### series ######
## Visit ##
EPG_series.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_series = EPG_series.reset_index()
del EPG_series['index']
EPG_series_visit = EPG_series.loc[0:4, :]
del EPG_series_visit['Service']
del EPG_series_visit['Type']
del EPG_series_visit['Duration']
## Duration ##
EPG_series.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_series = EPG_series.reset_index()
del EPG_series['index']
EPG_series_duration= EPG_series.loc[0:4, :]
del EPG_series_duration['Service']
del EPG_series_duration['Type']
del EPG_series_duration['Visit']
###### sport ######
## Visit ##
EPG_sport.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_sport = EPG_sport.reset_index()
del EPG_sport['index']
EPG_sport_visit = EPG_sport.loc[0:4, :]
del EPG_sport_visit['Service']
del EPG_sport_visit['Type']
del EPG_sport_visit['Duration']
## Duration ##
EPG_sport.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_sport = EPG_sport.reset_index()
del EPG_sport['index']
EPG_sport_duration= EPG_sport.loc[0:4, :]
del EPG_sport_duration['Service']
del EPG_sport_duration['Type']
del EPG_sport_duration['Visit']
###### kid ######
## Visit ##
EPG_kid.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_kid = EPG_kid.reset_index()
del EPG_kid['index']
EPG_kid_visit = EPG_kid.loc[0:4, :]
del EPG_kid_visit['Service']
del EPG_kid_visit['Type']
del EPG_kid_visit['Duration']
## Duration ##
EPG_kid.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_kid = EPG_kid.reset_index()
del EPG_kid['index']
EPG_kid_duration= EPG_kid.loc[0:4, :]
del EPG_kid_duration['Service']
del EPG_kid_duration['Type']
del EPG_kid_duration['Visit']
###### documentary ######
## Visit ##
EPG_documentary.sort_values('Visit', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_documentary = EPG_documentary.reset_index()
del EPG_documentary['index']
EPG_documentary_visit = EPG_documentary.loc[0:4, :]
del EPG_documentary_visit['Service']
del EPG_documentary_visit['Type']
del EPG_documentary_visit['Duration']
## Duration ##
EPG_documentary.sort_values('Duration', axis = 0, ascending = False, inplace = True, na_position ='last')
EPG_documentary = EPG_documentary.reset_index()
del EPG_documentary['index']
EPG_documentary_duration= EPG_documentary.loc[0:4, :]
del EPG_documentary_duration['Service']
del EPG_documentary_duration['Type']
del EPG_documentary_duration['Visit']

######################################################
################# TimeLine of Services ###############
######################################################
###### sima ######
EPG_sima_timeline = pd.DataFrame()
EPG_main_sima = EPG_main.query('نوع == "سراسری"')
EPG_sima_timeline['Visit'] = EPG_main_sima['تعداد بازدید']
EPG_sima_timeline['Duration'] = EPG_main_sima['مدت بازدید']
EPG_sima_timeline['Time'] = EPG_main_sima['ساعت']
EPG_sima_timeline['Time'] = EPG_sima_timeline['Time'].astype(str).replace('\.0', '', regex=True)
EPG_sima_timeline['Time'] = EPG_sima_timeline['Time'].apply(lambda x: x.zfill(2))
EPG_sima_timeline = EPG_sima_timeline.groupby(['Time']).sum().reset_index()
EPG_sima_timeline.sort_values('Time', axis = 0, ascending = True, inplace = True, na_position ='last')
###### radio ######
EPG_radio_timeline = pd.DataFrame()
EPG_main_radio = EPG_main.query('نوع == "رادیویی"')
EPG_radio_timeline['Visit'] = EPG_main_radio['تعداد بازدید']
EPG_radio_timeline['Duration'] = EPG_main_radio['مدت بازدید']
EPG_radio_timeline['Time'] = EPG_main_radio['ساعت']
EPG_radio_timeline['Time'] = EPG_radio_timeline['Time'].astype(str).replace('.0', '', regex=True)
EPG_radio_timeline['Time'] = EPG_radio_timeline['Time'].apply(lambda x: x.zfill(2))
EPG_radio_timeline = EPG_radio_timeline.groupby(['Time']).sum().reset_index()
EPG_radio_timeline.sort_values('Time', axis = 0, ascending = True, inplace = True, na_position ='last')
###### ekhtesasi ######
EPG_ekhtesasi_timeline = pd.DataFrame()
EPG_main_ekhtesasi = EPG_main.query('نوع == "اختصاصی"')
EPG_ekhtesasi_timeline['Visit'] = EPG_main_ekhtesasi['تعداد بازدید']
EPG_ekhtesasi_timeline['Duration'] = EPG_main_ekhtesasi['مدت بازدید']
EPG_ekhtesasi_timeline['Time'] = EPG_main_ekhtesasi['ساعت']
EPG_ekhtesasi_timeline['Time'] = EPG_ekhtesasi_timeline['Time'].astype(str).replace('.0', '', regex=True)
EPG_ekhtesasi_timeline['Time'] = EPG_ekhtesasi_timeline['Time'].apply(lambda x: x.zfill(2))
EPG_ekhtesasi_timeline = EPG_ekhtesasi_timeline.groupby(['Time']).sum().reset_index()
EPG_ekhtesasi_timeline.sort_values('Time', axis = 0, ascending = True, inplace = True, na_position ='last')
###### ostani ######
EPG_ostani_timeline = pd.DataFrame()
EPG_main_ostani = EPG_main.query('نوع == "استانی"')
EPG_ostani_timeline['Visit'] = EPG_main_ostani['تعداد بازدید']
EPG_ostani_timeline['Duration'] = EPG_main_ostani['مدت بازدید']
EPG_ostani_timeline['Time'] = EPG_main_ostani['ساعت']
EPG_ostani_timeline['Time'] = EPG_ostani_timeline['Time'].astype(str).replace('.0', '', regex=True)
EPG_ostani_timeline['Time'] = EPG_ostani_timeline['Time'].apply(lambda x: x.zfill(2))
EPG_ostani_timeline = EPG_ostani_timeline.groupby(['Time']).sum().reset_index()
EPG_ostani_timeline.sort_values('Time', axis = 0, ascending = True, inplace = True, na_position ='last')
###### boronmarzi ######
EPG_boronmarzi_timeline = pd.DataFrame()
EPG_main_boronmarzi = EPG_main.query('نوع == "برون مرزی"')
EPG_boronmarzi_timeline['Visit'] = EPG_main_boronmarzi['تعداد بازدید']
EPG_boronmarzi_timeline['Duration'] = EPG_main_boronmarzi['مدت بازدید']
EPG_boronmarzi_timeline['Time'] = EPG_main_boronmarzi['ساعت']
EPG_boronmarzi_timeline['Time'] = EPG_boronmarzi_timeline['Time'].astype(str).replace('.0', '', regex=True)
EPG_boronmarzi_timeline['Time'] = EPG_boronmarzi_timeline['Time'].apply(lambda x: x.zfill(2))
EPG_boronmarzi_timeline = EPG_boronmarzi_timeline.groupby(['Time']).sum().reset_index()
EPG_boronmarzi_timeline.sort_values('Time', axis = 0, ascending = True, inplace = True, na_position ='last')

######################################################
################# Production of Excel ################
######################################################
from Production_Series import *

[EPG_S_sima_visit, EPG_S_sima_duration, EPG_S_radio_visit, EPG_S_radio_duration, 
                                 EPG_S_ostani_visit, EPG_S_ostani_duration, EPG_S_ekhtesasi_visit, EPG_S_ekhtesasi_duration, 
                                 EPG_S_boronmarzi_visit, EPG_S_boronmarzi_duration, EPG_S_news_visit, EPG_S_news_duration, \
                                 EPG_S_series_visit, EPG_S_series_duration, EPG_S_sport_visit, EPG_S_sport_duration, \
                                 EPG_S_kid_visit, EPG_S_kid_duration, EPG_S_documentary_visit, EPG_S_documentary_duration] = Production_Series(EPG_sima_visit, EPG_sima_duration, EPG_radio_visit, EPG_radio_duration, \
                                 EPG_ostani_visit, EPG_ostani_duration, EPG_ekhtesasi_visit, EPG_ekhtesasi_duration, \
                                 EPG_boronmarzi_visit, EPG_boronmarzi_duration, \
                                 EPG_movie_visit, EPG_movie_duration, EPG_news_visit, EPG_news_duration, \
                                 EPG_series_visit, EPG_series_duration, EPG_sport_visit, EPG_sport_duration, \
                                 EPG_kid_visit, EPG_kid_duration, EPG_documentary_visit, EPG_documentary_duration, \
                                 EPG_sima_new, EPG_radio, EPG_ostani, EPG_ekhtesasi, EPG_boronmarzi, \
                                 EPG_news_new, EPG_series_new, EPG_sport_new, \
                                 EPG_kid_new, EPG_documentary_new)
######################################################
################ Production of Series ################
######################################################
writer = pd.ExcelWriter('آمار سرویس های پخش-EPG.xlsx', engine='xlsxwriter')
EPG_sima_visit.to_excel(writer, 'پربازدید سیما-تعداد بازدید', index = False)
EPG_sima_duration.to_excel(writer, 'پربازدید سیما-زمان بازدید', index = False)
EPG_radio_visit.to_excel(writer, 'پربازدید رادیو-تعداد بازدید', index = False)
EPG_radio_duration.to_excel(writer, 'پربازدید رادیو-زمان بازدید', index = False)
EPG_ostani_visit.to_excel(writer, 'پربازدید استانی-تعداد بازدید', index = False)
EPG_ostani_duration.to_excel(writer, 'پربازدید استانی-زمان بازدید', index = False)
EPG_ekhtesasi_visit.to_excel(writer, 'پربازدید اختصاصی-تعداد بازدید', index = False)
EPG_ekhtesasi_duration.to_excel(writer, 'پربازدید اختصاصی-زمان بازدید', index = False)
EPG_boronmarzi_visit.to_excel(writer, 'پربازدید برون مرزی-تعداد بازدید', index = False)
EPG_boronmarzi_duration.to_excel(writer, 'پربازدید برون مرزی-زمان بازدید', index = False)
EPG_sima_timeline.to_excel(writer, 'تایم لاین بازدید سیما', index = False)
EPG_radio_timeline.to_excel(writer, 'تایم لاین بازدید رادیو', index = False)
EPG_ekhtesasi_timeline.to_excel(writer, 'تایم لاین بازدید اختصاصی', index = False)
EPG_ostani_timeline.to_excel(writer, 'تایم لاین بازدید استانی', index = False)
EPG_boronmarzi_timeline.to_excel(writer, 'تایم لاین بازدید برون مرزی', index = False)
writer.save()

writer = pd.ExcelWriter('آمار انواع محتوا-EPG.xlsx', engine='xlsxwriter')
EPG_movie_visit.to_excel(writer, 'پربازدید فیلم-تعداد بازدید', index = False)
EPG_movie_duration.to_excel(writer, 'پربازدید فیلم-زمان بازدید', index = False)
EPG_news_visit.to_excel(writer, 'پربازدید خبر-تعداد بازدید', index = False)
EPG_news_duration.to_excel(writer, 'پربازدید خبر-زمان بازدید', index = False)
EPG_series_visit.to_excel(writer, 'پربازدید سریال-تعداد بازدید', index = False)
EPG_series_duration.to_excel(writer, 'پربازدید سریال-زمان بازدید', index = False)
EPG_sport_visit.to_excel(writer, 'پربازدید ورزشی-تعداد بازدید', index = False)
EPG_sport_duration.to_excel(writer, 'پربازدید ورزشی-زمان بازدید', index = False)
EPG_kid_visit.to_excel(writer, 'پربازدید کودک-تعداد بازدید', index = False)
EPG_kid_duration.to_excel(writer, 'پربازدید کودک-زمان بازدید', index = False)
EPG_documentary_visit.to_excel(writer, 'پربازدید مستند-تعداد بازدید', index = False)
EPG_documentary_duration.to_excel(writer, 'پربازدید مستند-زمان بازدید', index = False)
writer.save()














