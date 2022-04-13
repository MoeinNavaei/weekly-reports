
def Production_Series(EPG_sima_visit, EPG_sima_duration, EPG_radio_visit, EPG_radio_duration, \
                                 EPG_ostani_visit, EPG_ostani_duration, EPG_ekhtesasi_visit, EPG_ekhtesasi_duration, \
                                 EPG_boronmarzi_visit, EPG_boronmarzi_duration, \
                                 EPG_movie_visit, EPG_movie_duration, EPG_news_visit, EPG_news_duration, \
                                 EPG_series_visit, EPG_series_duration, EPG_sport_visit, EPG_sport_duration, \
                                 EPG_kid_visit, EPG_kid_duration, EPG_documentary_visit, EPG_documentary_duration, \
                                 EPG_sima_1, EPG_radio, EPG_ostani, EPG_ekhtesasi, EPG_boronmarzi, \
                                 EPG_news_2, EPG_series_2, EPG_sport_2, EPG_kid_2, EPG_documentary_new2):

    import pandas as pd
    import xlsxwriter 
    
    ###### sima --- visit ######
    list_dup = EPG_sima_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_sima_visit = pd.DataFrame()
    EPG_sima_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_sima_title = EPG_sima_1.query("channel == @var_channel")
        EPG_sima_title = EPG_sima_title.query("program == @var_program")
        if len(EPG_sima_title) > 1:
            EPG_S_sima_visit1 = pd.DataFrame()
            EPG_S_sima_visit1['channel'] = EPG_sima_title['channel']
            EPG_S_sima_visit1['program'] = EPG_sima_title['program']
            EPG_S_sima_visit1['Visit'] = EPG_sima_title['Visit']
            EPG_S_sima_visit1['Date'] = EPG_sima_title['Date']
            EPG_S_sima_visit1 = EPG_S_sima_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_sima_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_sima_visit = pd.concat([EPG_S_sima_visit, EPG_S_sima_visit1])
    ###### sima --- duration ######
    list_dup = EPG_sima_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_sima_duration = pd.DataFrame()
    EPG_sima_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_sima_title = EPG_sima_1.query("channel == @var_channel")
        EPG_sima_title = EPG_sima_title.query("program == @var_program")
        if len(EPG_sima_title) > 1:
            EPG_S_sima_duration1 = pd.DataFrame()
            EPG_S_sima_duration1['channel'] = EPG_sima_title['channel']
            EPG_S_sima_duration1['program'] = EPG_sima_title['program']
            EPG_S_sima_duration1['Duration'] = EPG_sima_title['Duration']
            EPG_S_sima_duration1['Date'] = EPG_sima_title['Date']
            EPG_S_sima_duration1 = EPG_S_sima_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_sima_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_sima_duration = pd.concat([EPG_S_sima_duration, EPG_S_sima_duration1])
    ###### radio --- visit ######
    list_dup = EPG_radio_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_radio_visit = pd.DataFrame()
    EPG_radio_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_radio_title = EPG_radio.query("channel == @var_channel")
        EPG_radio_title = EPG_radio_title.query("program == @var_program")
        if len(EPG_radio_title) > 1:
            EPG_S_radio_visit1 = pd.DataFrame()
            EPG_S_radio_visit1['channel'] = EPG_radio_title['channel']
            EPG_S_radio_visit1['program'] = EPG_radio_title['program']
            EPG_S_radio_visit1['Visit'] = EPG_radio_title['Visit']
            EPG_S_radio_visit1['Date'] = EPG_radio_title['Date']
            EPG_S_radio_visit1 = EPG_S_radio_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_radio_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_radio_visit = pd.concat([EPG_S_radio_visit, EPG_S_radio_visit1])
    ###### radio --- duration ######
    list_dup = EPG_radio_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_radio_duration = pd.DataFrame()
    EPG_radio_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_radio_title = EPG_radio.query("channel == @var_channel")
        EPG_radio_title = EPG_radio_title.query("program == @var_program")
        if len(EPG_radio_title) > 1:
            EPG_S_radio_duration1 = pd.DataFrame()
            EPG_S_radio_duration1['channel'] = EPG_radio_title['channel']
            EPG_S_radio_duration1['program'] = EPG_radio_title['program']
            EPG_S_radio_duration1['Duration'] = EPG_radio_title['Duration']
            EPG_S_radio_duration1['Date'] = EPG_radio_title['Date']
            EPG_S_radio_duration1 = EPG_S_radio_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_radio_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_radio_duration = pd.concat([EPG_S_radio_duration, EPG_S_radio_duration1])  
    ###### ostani --- visit ######
    list_dup = EPG_ostani_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_ostani_visit = pd.DataFrame()
    EPG_ostani_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_ostani_title = EPG_ostani.query("channel == @var_channel")
        EPG_ostani_title = EPG_ostani_title.query("program == @var_program")
        if len(EPG_ostani_title) > 1:
            EPG_S_ostani_visit1 = pd.DataFrame()
            EPG_S_ostani_visit1['channel'] = EPG_ostani_title['channel']
            EPG_S_ostani_visit1['program'] = EPG_ostani_title['program']
            EPG_S_ostani_visit1['Visit'] = EPG_ostani_title['Visit']
            EPG_S_ostani_visit1['Date'] = EPG_ostani_title['Date']
            EPG_S_ostani_visit1 = EPG_S_ostani_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_ostani_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_ostani_visit = pd.concat([EPG_S_ostani_visit, EPG_S_ostani_visit1])
    ###### ostani --- duration ######
    list_dup = EPG_ostani_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_ostani_duration = pd.DataFrame()
    EPG_ostani_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_ostani_title = EPG_ostani.query("channel == @var_channel")
        EPG_ostani_title = EPG_ostani_title.query("program == @var_program")
        if len(EPG_ostani_title) > 1:
            EPG_S_ostani_duration1 = pd.DataFrame()
            EPG_S_ostani_duration1['channel'] = EPG_ostani_title['channel']
            EPG_S_ostani_duration1['program'] = EPG_ostani_title['program']
            EPG_S_ostani_duration1['Duration'] = EPG_ostani_title['Duration']
            EPG_S_ostani_duration1['Date'] = EPG_ostani_title['Date']
            EPG_S_ostani_duration1 = EPG_S_ostani_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_ostani_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_ostani_duration = pd.concat([EPG_S_ostani_duration, EPG_S_ostani_duration1])
    ###### ekhtesasi --- visit ######
    list_dup = EPG_ekhtesasi_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_ekhtesasi_visit = pd.DataFrame()
    EPG_ekhtesasi_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_ekhtesasi_title = EPG_ekhtesasi.query("channel == @var_channel")
        EPG_ekhtesasi_title = EPG_ekhtesasi_title.query("program == @var_program")
        if len(EPG_ekhtesasi_title) > 1:
            EPG_S_ekhtesasi_visit1 = pd.DataFrame()
            EPG_S_ekhtesasi_visit1['channel'] = EPG_ekhtesasi_title['channel']
            EPG_S_ekhtesasi_visit1['program'] = EPG_ekhtesasi_title['program']
            EPG_S_ekhtesasi_visit1['Visit'] = EPG_ekhtesasi_title['Visit']
            EPG_S_ekhtesasi_visit1['Date'] = EPG_ekhtesasi_title['Date']
            EPG_S_ekhtesasi_visit1 = EPG_S_ekhtesasi_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_ekhtesasi_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_ekhtesasi_visit = pd.concat([EPG_S_ekhtesasi_visit, EPG_S_ekhtesasi_visit1])
    ###### ekhtesasi --- duration ######
    list_dup = EPG_ekhtesasi_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_ekhtesasi_duration = pd.DataFrame()
    EPG_ekhtesasi_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_ekhtesasi_title = EPG_ekhtesasi.query("channel == @var_channel")
        EPG_ekhtesasi_title = EPG_ekhtesasi_title.query("program == @var_program")
        if len(EPG_ekhtesasi_title) > 1:
            EPG_S_ekhtesasi_duration1 = pd.DataFrame()
            EPG_S_ekhtesasi_duration1['channel'] = EPG_ekhtesasi_title['channel']
            EPG_S_ekhtesasi_duration1['program'] = EPG_ekhtesasi_title['program']
            EPG_S_ekhtesasi_duration1['Duration'] = EPG_ekhtesasi_title['Duration']
            EPG_S_ekhtesasi_duration1['Date'] = EPG_ekhtesasi_title['Date']
            EPG_S_ekhtesasi_duration1 = EPG_S_ekhtesasi_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_ekhtesasi_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_ekhtesasi_duration = pd.concat([EPG_S_ekhtesasi_duration, EPG_S_ekhtesasi_duration1])
    ###### boronmarzi --- visit ######
    list_dup = EPG_boronmarzi_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_boronmarzi_visit = pd.DataFrame()
    EPG_boronmarzi_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_boronmarzi_title = EPG_boronmarzi.query("channel == @var_channel")
        EPG_boronmarzi_title = EPG_boronmarzi_title.query("program == @var_program")
        if len(EPG_boronmarzi_title) > 1:
            EPG_S_boronmarzi_visit1 = pd.DataFrame()
            EPG_S_boronmarzi_visit1['channel'] = EPG_boronmarzi_title['channel']
            EPG_S_boronmarzi_visit1['program'] = EPG_boronmarzi_title['program']
            EPG_S_boronmarzi_visit1['Visit'] = EPG_boronmarzi_title['Visit']
            EPG_S_boronmarzi_visit1['Date'] = EPG_boronmarzi_title['Date']
            EPG_S_boronmarzi_visit1 = EPG_S_ekhtesasi_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_boronmarzi_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_boronmarzi_visit = pd.concat([EPG_S_boronmarzi_visit, EPG_S_boronmarzi_visit1])
    ###### boronmarzi --- duration ######
    list_dup = EPG_boronmarzi_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_boronmarzi_duration = pd.DataFrame()
    EPG_boronmarzi_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_boronmarzi_title = EPG_boronmarzi.query("channel == @var_channel")
        EPG_boronmarzi_title = EPG_boronmarzi_title.query("program == @var_program")
        if len(EPG_boronmarzi_title) > 1:
            EPG_S_boronmarzi_duration1 = pd.DataFrame()
            EPG_S_boronmarzi_duration1['channel'] = EPG_boronmarzi_title['channel']
            EPG_S_boronmarzi_duration1['program'] = EPG_boronmarzi_title['program']
            EPG_S_boronmarzi_duration1['Duration'] = EPG_boronmarzi_title['Duration']
            EPG_S_boronmarzi_duration1['Date'] = EPG_boronmarzi_title['Date']
            EPG_S_boronmarzi_duration1 = EPG_S_ekhtesasi_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_boronmarzi_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_boronmarzi_duration = pd.concat([EPG_S_boronmarzi_duration, EPG_S_boronmarzi_duration1])
    ###### news --- visit ######
    list_dup = EPG_news_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_new2s_visit = pd.DataFrame()
    EPG_new2s_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_new2s_title = EPG_news_2.query("channel == @var_channel")
        EPG_new2s_title = EPG_new2s_title.query("program == @var_program")
        if len(EPG_new2s_title) > 1:
            EPG_S_new2s_visit1 = pd.DataFrame()
            EPG_S_new2s_visit1['channel'] = EPG_new2s_title['channel']
            EPG_S_new2s_visit1['program'] = EPG_new2s_title['program']
            EPG_S_new2s_visit1['Visit'] = EPG_new2s_title['Visit']
            EPG_S_new2s_visit1['Date'] = EPG_new2s_title['Date']
            EPG_S_new2s_visit1 = EPG_S_new2s_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_new2s_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_new2s_visit = pd.concat([EPG_S_new2s_visit, EPG_S_new2s_visit1])
    ###### news --- duration ######
    list_dup = EPG_news_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_new2s_duration = pd.DataFrame()
    EPG_new2s_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_new2s_title = EPG_news_2.query("channel == @var_channel")
        EPG_new2s_title = EPG_new2s_title.query("program == @var_program")
        if len(EPG_new2s_title) > 1:
            EPG_S_new2s_duration1 = pd.DataFrame()
            EPG_S_new2s_duration1['channel'] = EPG_new2s_title['channel']
            EPG_S_new2s_duration1['program'] = EPG_new2s_title['program']
            EPG_S_new2s_duration1['Duration'] = EPG_new2s_title['Duration']
            EPG_S_new2s_duration1['Date'] = EPG_new2s_title['Date']
            EPG_S_new2s_duration1 = EPG_S_ekhtesasi_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_new2s_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_new2s_duration = pd.concat([EPG_S_new2s_duration, EPG_S_new2s_duration1])
    ###### series --- visit ######
    list_dup = EPG_series_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_series_visit = pd.DataFrame()
    EPG_series_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_series_title = EPG_series_2.query("channel == @var_channel")
        EPG_series_title = EPG_series_title.query("program == @var_program")
        if len(EPG_series_title) > 1:
            EPG_S_series_visit1 = pd.DataFrame()
            EPG_S_series_visit1['channel'] = EPG_series_title['channel']
            EPG_S_series_visit1['program'] = EPG_series_title['program']
            EPG_S_series_visit1['Visit'] = EPG_series_title['Visit']
            EPG_S_series_visit1['Date'] = EPG_series_title['Date']
            EPG_S_series_visit1 = EPG_S_series_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_series_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_series_visit = pd.concat([EPG_S_series_visit, EPG_S_series_visit1])
    ###### series --- duration ######
    list_dup = EPG_series_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_series_duration = pd.DataFrame()
    EPG_series_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_series_title = EPG_series_2.query("channel == @var_channel")
        EPG_series_title = EPG_series_title.query("program == @var_program")
        if len(EPG_series_title) > 1:
            EPG_S_series_duration1 = pd.DataFrame()
            EPG_S_series_duration1['channel'] = EPG_series_title['channel']
            EPG_S_series_duration1['program'] = EPG_series_title['program']
            EPG_S_series_duration1['Duration'] = EPG_series_title['Duration']
            EPG_S_series_duration1['Date'] = EPG_series_title['Date']
            EPG_S_series_duration1 = EPG_S_series_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_series_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_series_duration = pd.concat([EPG_S_series_duration, EPG_S_series_duration1])
    ###### sport --- visit ######
    list_dup = EPG_sport_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_sport_visit = pd.DataFrame()
    EPG_sport_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_sport_title = EPG_sport_2.query("channel == @var_channel")
        EPG_sport_title = EPG_sport_title.query("program == @var_program")
        if len(EPG_sport_title) > 1:
            EPG_S_sport_visit1 = pd.DataFrame()
            EPG_S_sport_visit1['channel'] = EPG_sport_title['channel']
            EPG_S_sport_visit1['program'] = EPG_sport_title['program']
            EPG_S_sport_visit1['Visit'] = EPG_sport_title['Visit']
            EPG_S_sport_visit1['Date'] = EPG_sport_title['Date']
            EPG_S_sport_visit1 = EPG_S_sport_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_sport_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_sport_visit = pd.concat([EPG_S_sport_visit, EPG_S_sport_visit1])
    ###### sport --- duration ######
    list_dup = EPG_sport_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_sport_duration = pd.DataFrame()
    EPG_sport_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_sport_title = EPG_sport_2.query("channel == @var_channel")
        EPG_sport_title = EPG_sport_title.query("program == @var_program")
        if len(EPG_sport_title) > 1:
            EPG_S_sport_duration1 = pd.DataFrame()
            EPG_S_sport_duration1['channel'] = EPG_sport_title['channel']
            EPG_S_sport_duration1['program'] = EPG_sport_title['program']
            EPG_S_sport_duration1['Duration'] = EPG_sport_title['Duration']
            EPG_S_sport_duration1['Date'] = EPG_sport_title['Date']
            EPG_S_sport_duration1 = EPG_S_sport_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_sport_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_sport_duration = pd.concat([EPG_S_sport_duration, EPG_S_sport_duration1])
    ###### kid --- visit ######
    list_dup = EPG_kid_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_kid_visit = pd.DataFrame()
    EPG_kid_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_kid_title = EPG_kid_2.query("channel == @var_channel")
        EPG_kid_title = EPG_kid_title.query("program == @var_program")
        if len(EPG_kid_title) > 1:
            EPG_S_kid_visit1 = pd.DataFrame()
            EPG_S_kid_visit1['channel'] = EPG_kid_title['channel']
            EPG_S_kid_visit1['program'] = EPG_kid_title['program']
            EPG_S_kid_visit1['Visit'] = EPG_kid_title['Visit']
            EPG_S_kid_visit1['Date'] = EPG_kid_title['Date']
            EPG_S_kid_visit1 = EPG_S_kid_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_kid_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_kid_visit = pd.concat([EPG_S_kid_visit, EPG_S_kid_visit1])
    ###### kid --- duration ######
    list_dup = EPG_kid_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_kid_duration = pd.DataFrame()
    EPG_kid_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_kid_title = EPG_kid_2.query("channel == @var_channel")
        EPG_kid_title = EPG_kid_title.query("program == @var_program")
        if len(EPG_kid_title) > 1:
            EPG_S_kid_duration1 = pd.DataFrame()
            EPG_S_kid_duration1['channel'] = EPG_kid_title['channel']
            EPG_S_kid_duration1['program'] = EPG_kid_title['program']
            EPG_S_kid_duration1['Duration'] = EPG_kid_title['Duration']
            EPG_S_kid_duration1['Date'] = EPG_kid_title['Date']
            EPG_S_kid_duration1 = EPG_S_kid_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_kid_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_kid_duration = pd.concat([EPG_S_kid_duration, EPG_S_kid_duration1])
    ###### documentary --- visit ######
    list_dup = EPG_documentary_visit.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_documentary_visit = pd.DataFrame()
    EPG_documentary_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_documentary_title = EPG_documentary_new2.query("channel == @var_channel")
        EPG_documentary_title = EPG_documentary_title.query("program == @var_program")
        if len(EPG_documentary_title) > 1:
            EPG_S_documentary_visit1 = pd.DataFrame()
            EPG_S_documentary_visit1['channel'] = EPG_documentary_title['channel']
            EPG_S_documentary_visit1['program'] = EPG_documentary_title['program']
            EPG_S_documentary_visit1['Visit'] = EPG_documentary_title['Visit']
            EPG_S_documentary_visit1['Date'] = EPG_documentary_title['Date']
            EPG_S_documentary_visit1 = EPG_S_documentary_visit1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_documentary_visit1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_documentary_visit = pd.concat([EPG_S_documentary_visit, EPG_S_documentary_visit1])
    ###### documentary --- duration ######
    list_dup = EPG_documentary_duration.copy()
    list_dup.drop_duplicates(subset =['channel', 'program'], keep = 'first', inplace = True)
    list_dup = list_dup.reset_index()
    del list_dup['index']
    EPG_S_documentary_duration = pd.DataFrame()
    EPG_documentary_title = pd.DataFrame()
    for i in range(len(list_dup)):    # len(list_dup)
        var_channel = list_dup.loc[i, 'channel']
        var_program = list_dup.loc[i, 'program']
        EPG_documentary_title = EPG_documentary_new2.query("channel == @var_channel")
        EPG_documentary_title = EPG_documentary_title.query("program == @var_program")
        if len(EPG_documentary_title) > 1:
            EPG_S_documentary_duration1 = pd.DataFrame()
            EPG_S_documentary_duration1['channel'] = EPG_documentary_title['channel']
            EPG_S_documentary_duration1['program'] = EPG_documentary_title['program']
            EPG_S_documentary_duration1['Duration'] = EPG_documentary_title['Duration']
            EPG_S_documentary_duration1['Date'] = EPG_documentary_title['Date']
            EPG_S_documentary_duration1 = EPG_S_documentary_duration1.groupby(['channel', 'program', 'Date']).sum().reset_index()
            EPG_S_documentary_duration1.sort_values('Date', axis = 0, ascending = True, inplace = True, na_position ='last')
            EPG_S_documentary_duration = pd.concat([EPG_S_documentary_duration, EPG_S_documentary_duration1])
    
    ######################################################
    ################ Production of Series ################
    ######################################################
    writer = pd.ExcelWriter('قسمت های محتواهای پربازدید-EPG.xlsx', engine='xlsxwriter')
    EPG_S_sima_visit.to_excel(writer, 'سیما-تعداد بازدید', index = False)
    EPG_S_sima_duration.to_excel(writer, 'سیما-زمان بازدید', index = False)
    EPG_S_radio_visit.to_excel(writer, 'رادیو-تعداد بازدید', index = False)
    EPG_S_radio_duration.to_excel(writer, 'رادیو-زمان بازدید', index = False)
    EPG_S_ostani_visit.to_excel(writer, 'استانی-تعداد بازدید', index = False)
    EPG_S_ostani_duration.to_excel(writer, 'استانی-زمان بازدید', index = False)
    EPG_S_ekhtesasi_visit.to_excel(writer, 'اختصاصی-تعداد بازدید', index = False)
    EPG_S_ekhtesasi_duration.to_excel(writer, 'اختصاصی-زمان بازدید', index = False)
    EPG_S_boronmarzi_visit.to_excel(writer, 'برون مرزی-تعداد بازدید', index = False)
    EPG_S_boronmarzi_duration.to_excel(writer, 'برون مرزی-زمان بازدید', index = False)
    EPG_S_new2s_visit.to_excel(writer, 'اخبار-تعداد بازدید', index = False)
    EPG_S_new2s_duration.to_excel(writer, 'اخبار-زمان بازدید', index = False)
    EPG_S_series_visit.to_excel(writer, 'سریال-تعداد بازدید', index = False)
    EPG_S_series_duration.to_excel(writer, 'سریال-زمان بازدید', index = False)
    EPG_S_sport_visit.to_excel(writer, 'ورزشی-تعداد بازدید', index = False)
    EPG_S_sport_duration.to_excel(writer, 'ورزشی-زمان بازدید', index = False)
    EPG_S_kid_visit.to_excel(writer, 'کودک-تعداد بازدید', index = False)
    EPG_S_kid_duration.to_excel(writer, 'کودک-زمان بازدید', index = False)
    EPG_S_documentary_visit.to_excel(writer, 'مستند-تعداد بازدید', index = False)
    EPG_S_documentary_duration.to_excel(writer, 'مستند-زمان بازدید', index = False)
    writer.save()
    

    return EPG_S_sima_visit, EPG_S_sima_duration, EPG_S_radio_visit, EPG_S_radio_duration, \
                                 EPG_S_ostani_visit, EPG_S_ostani_duration, EPG_S_ekhtesasi_visit, EPG_S_ekhtesasi_duration, \
                                 EPG_S_boronmarzi_visit, EPG_S_boronmarzi_duration, \
                                 EPG_S_new2s_visit, EPG_S_new2s_duration, EPG_S_series_visit, EPG_S_series_duration, EPG_S_sport_visit, EPG_S_sport_duration, \
                                 EPG_S_kid_visit, EPG_S_kid_duration, EPG_S_documentary_visit, EPG_S_documentary_duration







