#---Load modules
import tkinter as tk
from tkinter import ttk
import pandas as pd
import regex
import os
import ctypes
import matplotlib.pyplot as plt
import datetime as DT

#---Close console/shell window
kernel32=ctypes.WinDLL('kernel32')
user32=ctypes.WinDLL('user32')
SW_HIDE=0
hWnd=kernel32.GetConsoleWindow()
user32.ShowWindow(hWnd, SW_HIDE)
pd.set_option('mode.chained_assignment',None)

#--Functions
def logcheck():
    global filename
    global contu
    os.chdir(foldir.get())
    for filename in os.listdir(foldir.get()):
        if filename.endswith('DataLog.csv'):
            try:
                txtbx.insert(tk.END,f'Reading {filename}\n','info')
                sensor=env.get()
                contu=1
                contin()
                if contu==0:
                    df=pd.read_csv(filename,encoding='cp1252',skip_blank_lines=True)
                    if len(df.columns)!=33:
                        txtbx.insert(tk.END,f'{filename} not 33 column format. Must be all variables currently in Maestro MQB Logger.\n','fail')
                    df.columns=['Time','RPM','Airmass','PUT Setpoint','PUT Actual','Wastegate','Wastegate Setpoint','Ambient Pressure','IAT','Lambda','Lambda Setpoint','STFT','DI Pressure BAR','PI Rail Pressure','LPFP Duty','Gear','Injector pulsewidth','PI Pulsewidth','Advance Cyl 1','KR Cyl 1','KR Cyl 2','KR Cyl 3','KR Cyl 4','Pedal','ECU Pedal','Reported Torque','Torque Limit','Octane Slider','Max Octane Lookup','Min Octane Lookup','Boost Limit','DV Position','DV Setpoint']
                    df['Effective PSI']=(df['PUT Actual'].astype(float) - df['Ambient Pressure'].astype(float)) *0.014538
                    df['OverallTiming']=(df['Advance Cyl 1'] - df['KR Cyl 1'])
                    df['WG Mean']=df['Wastegate'].rolling(window=5).mean().round(0)
                    df['Boost Delta']=((df['PUT Actual'] / df['PUT Setpoint']) * 100 - 100).round(2)
                    df=df[['Time','RPM','Airmass','Effective PSI','PUT Setpoint','PUT Actual','Boost Delta','Wastegate','Wastegate Setpoint','WG Mean','Ambient Pressure','IAT','Lambda','Lambda Setpoint','STFT','DI Pressure BAR','PI Rail Pressure','LPFP Duty','Gear','Injector pulsewidth','PI Pulsewidth','OverallTiming','KR Cyl 1','KR Cyl 2','KR Cyl 3','KR Cyl 4','Pedal','ECU Pedal','Reported Torque','Torque Limit','Octane Slider','Max Octane Lookup','Min Octane Lookup','Boost Limit','DV Position','DV Setpoint']]
                    roundcols=['Wastegate','Airmass','Wastegate Setpoint','PUT Setpoint','PUT Actual','Ambient Pressure','DI Pressure BAR','PI Rail Pressure','LPFP Duty','Pedal','ECU Pedal','Reported Torque','Torque Limit','Boost Limit']
                    df[roundcols]=df[roundcols].round(0)
                    df['DI Pressure BAR']=(df['DI Pressure BAR'] / 10000).round(0)
                    df['Effective PSI']=df['Effective PSI'].round(1)
                    df[['Lambda','Lambda Setpoint','STFT']]=df[['Lambda','Lambda Setpoint','STFT']].round(2)
                    df['Delta']=pd.to_datetime(df['Time'], format='%H:%M:%S:%f')
                    df['Delta']=(df['Delta'] - df['Delta'].min()).dt.seconds
                    df=df.loc[df['Pedal'].rolling(window=3).mean() > 55]
                    df.loc[(df['Delta'] - df['Delta'].shift(1))> 1, 'NewLog']=1
                    df.loc[(df['NewLog'].isnull()), 'NewLog']=0
                    wgcols=['Wastegate','Wastegate Setpoint','WG Mean']
                    df[wgcols]=100 - df[wgcols]
                    if (df.iloc[0,16])>10000.0:
                        df=df.drop(['PI Rail Pressure','PI Pulsewidth'], axis=1)
                        piset=0
                    else:
                        piset=1
                    df.loc[((df['NewLog'].rolling(window=5).mean())==0) & (df['RPM'] > 5000) & (df['WG Mean'] > 80) & (df['PUT Setpoint'] < 2700), 'Boost Suggestion']='REVIEW - Lower boost demand/check for boost leak if BT. Wastegate is above 80% duty at high RPM and not meeting target.'
                    df.loc[((df['NewLog'].rolling(window=5).mean())==0) & ((df['Boost Delta'].rolling(window=5).mean().round(0))<-10) & (df['RPM'] > 5000) & (df['WG Mean'] > 90) & (df['PUT Setpoint'] < 3000), 'Boost Suggestion']='URGENT - Lower boost demand/check for boost leak if BT. Wastegate is above 90% duty at high RPM and not meeting target.'
                    df.loc[((df['NewLog'].rolling(window=5).mean())==0) & ((df['Boost Delta'].rolling(window=5).mean().round(0))>8), 'Boost Suggestion']='REVIEW - Overshooting boost due to wastegate setpoints and/or load. Review WG targets.'
                    df.loc[((df['NewLog'].rolling(window=5).mean())==0) & ((df['Boost Delta'].rolling(window=5).mean().round(0))>10), 'Boost Suggestion']='URGENT - Overshooting boost by over 10% due to wastegate setpoints and/or load. Review WG targets.'
                    df.loc[((df['Boost Limit']) - (df['PUT Setpoint'].max())) > 80, 'Boost Limiter']='REVIEW - Boost slider is above boost control limits in file.'
                    df.loc[(((df['NewLog'].rolling(window=5).mean())==0) & (((df['PUT Actual']).shift(1) - (df['PUT Actual'])) > 80) & (df['Pedal'] > 75)), 'Boost Control']='REVIEW - Large oscillation in boost. If not on initial spool or gear change, check sensor data, WG data/setpoints, or for leaks.'
                    df.loc[(((df['IAT'].shift(2).rolling(window=2).mean()) - (df['IAT'].rolling(window=2).mean())) <= -1) & ((df['NewLog'].rolling(window=5).mean())==0), 'IAT Check']='REVIEW - IAT increasing quickly. Turbo out of efficiency or IC poor.'
                    df.loc[(((df['Lambda'].rolling(window=3).mean()) - (df['Lambda Setpoint'].rolling(window=3).mean())) > 0.03) & (df['Pedal'] > 75) & (((df['Gear'].rolling(window=5).mean()) - (df['Gear'].shift(1).rolling(window=5).mean())) == 0) , 'Lambda Check']='REVIEW - Lambda is over request. If on gear change or initial spool, ignore.'
                    df.loc[(((df['Lambda'].rolling(window=3).mean()) - (df['Lambda Setpoint'].rolling(window=3).mean())) > 0.05) & (df['Pedal'] > 75) & (((df['Gear'].rolling(window=5).mean()) - (df['Gear'].shift(1).rolling(window=5).mean())) == 0), 'Lambda Check']='URGENT - Lambda is considerably over request. If on gear change, ignore.'
                    df.loc[((df['STFT']) < -15) & ((df['PUT Actual']) > 2000) , 'Trim Check']='REVIEW - Rich trims. Adjust E content slider if blending and/or injector constant if using MPI.'
                    df.loc[((df['STFT']) > 15) & ((df['PUT Actual']) > 2000) , 'Trim Check']='REVIEW - Lean trims. Adjust E content slider if blending and/or injector constant if using MPI.'
                    df.loc[((df['DI Pressure BAR']) <16) & ((df['PUT Actual']) > 2000) & ((df['Injector pulsewidth']) > 7) , 'HPFP Check']='URGENT - Fuel pressure VERY LOW on DI rail. Too much airflow or E content.'
                    df.loc[((df['DI Pressure BAR']) <19) & ((df['PUT Actual']) > 2000) & ((df['Injector pulsewidth']) > 7) & (df['RPM'].shift(1) < df['RPM']) , 'HPFP Check']='REVIEW - Fuel pressure dropping on DI rail. Too much airflow or E content.'
                    df.loc[((df['KR Cyl 1']) + (df['KR Cyl 2']) + (df['KR Cyl 3']) + (df['KR Cyl 4'])) < -4.99  , 'Knock Check']='REVIEW - Noticeable knock. If repeated on back to back logs, pull timing out equivalent to half of knock event in area or turn down octane slider.'
                    df.loc[((df['KR Cyl 1']) + (df['KR Cyl 2']) + (df['KR Cyl 3']) + (df['KR Cyl 4'])) < -9.99  , 'Knock Check']='URGENT - Considerable knock. Turn down octane slider 3. Log again. If knock remains, pull timing out equivalent to half of knock event in area/check for false knock.'
                    df.loc[((df['KR Cyl 1']) + (df['KR Cyl 2']) + (df['KR Cyl 3']) + (df['KR Cyl 4'])) < -24  , 'Knock Check']='URGENT - Serious knock. Turn down octane slider to minimum, turn down boost. Log again. If knock remains and overall timing below 0 above 5000 rpm, check for false knock.'
                    df.loc[((df['OverallTiming'].rolling(window=3).mean() <-2) & ((df['PUT Actual']) > 2200) & ((df['RPM']) > 4000)) , 'Timing Check']='REVIEW - Overall timing is low after spool. Efficiency reduced. Can create high EGTs at high boost.'
                    df.loc[(((df['Reported Torque']) - (df['Torque Limit'])) >0) , 'Torque Check']='REVIEW - Over torque limit. Check Max Clutch Torque values. Might need Reference Torque/Airflow targets adjusted.'
                    df.loc[(((df['Torque Limit']) < 400) & (df['Gear']) > 2), 'Torque Check']='URGENT - Safety has been tripped or major error in Max Clutch Torque tables. Review all nearby row variables.'
                    if piset==1:
                        df.loc[((df['PI Pulsewidth']) >0) & ((df['PI Pulsewidth'].shift(1))==0) & ((df['Lambda'].shift(-1) - (df['Lambda'].shift(1)))>0.07), 'Injector Constant']='URGENT - Increase MPI injector constant or confirm MPI working. Lambda considerably increasing after MPI activates.'
                    writer=pd.ExcelWriter(f"Corrected.{filename}"+'.xlsx', engine='xlsxwriter')
                    df.to_excel(writer, sheet_name='Log', index=None)
                    workbook=writer.book
                    wg=pd.read_csv(f'{installdir.get()}\suggestbtwg.csv', header=0,index_col=0)
                    wg.columns=[0,500,1000,1500,2000,2200,2600,2800,3000,3200]
                    wg=wg.astype(int)
                    wg.to_excel(writer, sheet_name='Wastegate',startrow=1,startcol=0)
                    worksheet3=writer.sheets['Wastegate']
                    worksheet3.write_string(0,0, 'Map for boost pressure actuator setpoint suggestion',workbook.add_format({'bold': True}))
                    if len(df.loc[df['Gear'] == 3].index) > 10:
                        boostg3=df.loc[df['Gear'] == 3]
                        boostg3['RPM']=(boostg3['RPM']/500).round(0)*500
                        boostg3['SampleSize']=boostg3.groupby('RPM')['RPM'].transform(len)
                        boostg3=boostg3.groupby('RPM').mean().round(2)
                        boostg3=boostg3.loc[boostg3['SampleSize'] > 4]
                        boostg3=boostg3.drop(['Delta','DV Position','DV Setpoint','Boost Limit','Octane Slider','NewLog'], axis=1)
                        boostg3.to_excel(writer, sheet_name='Analysis',startrow=1,startcol=0)
                        wgauto3=boostg3[['PUT Setpoint','Boost Delta','WG Mean']].round(0)
                        wgauto3=wgauto3.loc[(wgauto3['Boost Delta']>-5) & (wgauto3['Boost Delta']<5)]
                        wgauto3['PUT Setpoint']=wgauto3['PUT Setpoint'].round(-2)
                        wgauto3['Setpoint']=wgauto3['WG Mean'] - (wgauto3['Boost Delta'] * 2)
                        wgauto3=wgauto3[['PUT Setpoint','Setpoint']]
                        wgauto3.to_excel(writer, sheet_name='Wastegate',startrow=20,startcol=0)
                        tmauto3=boostg3[['Airmass','PUT Actual','OverallTiming','KR Cyl 1','KR Cyl 2','KR Cyl 3','KR Cyl 4']].round(1)
                        tmauto3['PUT Actual']=tmauto3['PUT Actual'].round(-2)
                        tmauto3.loc[(((tmauto3['KR Cyl 1']) + (tmauto3['KR Cyl 2']) + (tmauto3['KR Cyl 3']) + (tmauto3['KR Cyl 4'])) > -4), 'Outcome']=3
                        tmauto3.loc[(((tmauto3['KR Cyl 1']) + (tmauto3['KR Cyl 2']) + (tmauto3['KR Cyl 3']) + (tmauto3['KR Cyl 4'])) >= -1), 'Outcome']=1
                        tmauto3.loc[(((tmauto3['KR Cyl 1']) + (tmauto3['KR Cyl 2']) + (tmauto3['KR Cyl 3']) + (tmauto3['KR Cyl 4'])) < -3.9), 'Outcome']=4
                        if (tmauto3['Outcome'].mean()) ==1:
                            suggestion='Clear. If Gear 4 is available, confirm it is clear otherwise deal with those issues first. Review documentation on ideal timing curve. If satisfactory, you may turn up the octane slider 2 notches and relog.'
                        if ((tmauto3['Outcome'].mean()) <2) & ((tmauto3['Outcome'].mean()) >1):
                            suggestion='Review documentation on ideal timing curve. For minor knock, address the area with the timing map using half of the knock value on the log.'
                        if ((tmauto3['Outcome'].mean()) >=2):
                            suggestion='Too many problem areas to properly address with the timing table. It is best to drop the octane slider 3 positions and relog. Review for problem areas at that point and reevaluate.'
                        if ((tmauto3['Outcome'].mean()) >3.5):
                            suggestion='Big issues. Scale octane slider back down and start over if on pump gas. If on E30+, start at the slider level where you are seeing no more than 5* advance which is roughly 91/92 on an almost stock curve.'
                        if len(tmauto3.index) < 5:
                            suggestion='RPM sample size is too small. No suggestions possible for this gear. If no large knock (>3*) on this log, log again from 3000-6500rpm in 4th if possible, otherwise 3rd.'
                        tmauto3=tmauto3.drop(['Outcome'], axis=1)
                        tmauto3.to_excel(writer, sheet_name='Timing',startrow=2,startcol=0)
                        worksheet2=writer.sheets['Analysis']
                        worksheet2.write_string(0,0, 'Gear 3 Mean Analysis',workbook.add_format({'bold': True}))
                        worksheet3.write_string(19,0, 'Gear 3',workbook.add_format({'bold': True}))
                        worksheet4=writer.sheets['Timing']
                        worksheet4.write_string(0,0, 'Suggested Timing Changes',workbook.add_format({'bold': True}))
                        worksheet4.write_string(1,0, f'Gear 3 - {suggestion}',workbook.add_format({'bold': True})) 
                    if len(df.loc[df['Gear'] == 4].index) > 10:
                        boostg4=df.loc[df['Gear'] == 4]
                        boostg4['RPM']=(boostg4['RPM']/500).round(0)*500
                        boostg4['SampleSize']=boostg4.groupby('RPM')['RPM'].transform(len)
                        boostg4=boostg4.groupby('RPM').mean().round(2)
                        boostg4=boostg4.loc[boostg4['SampleSize'] > 8]
                        boostg4=boostg4.drop(['Delta','DV Position','DV Setpoint','Boost Limit','Octane Slider','NewLog'], axis=1)
                        wgauto4=boostg4[['PUT Actual','Boost Delta','WG Mean']].round(0)
                        wgauto4=wgauto4.loc[(wgauto4['Boost Delta']>-5) & (wgauto4['Boost Delta']<5)]
                        wgauto4['PUT Actual']=wgauto4['PUT Actual'].round(-2)
                        wgauto4['Setpoint']=wgauto4['WG Mean'] - (wgauto4['Boost Delta'] * 2)
                        wgauto4=wgauto4[['PUT Actual','Setpoint']]
                        tmauto4=boostg4[['Airmass','PUT Actual','OverallTiming','KR Cyl 1','KR Cyl 2','KR Cyl 3','KR Cyl 4']].round(1)
                        tmauto4['PUT Actual']=tmauto4['PUT Actual'].round(-2)
                        tmauto4.loc[(((tmauto4['KR Cyl 1']) + (tmauto4['KR Cyl 2']) + (tmauto4['KR Cyl 3']) + (tmauto4['KR Cyl 4'])) > -4), 'Outcome']=3
                        tmauto4.loc[(((tmauto4['KR Cyl 1']) + (tmauto4['KR Cyl 2']) + (tmauto4['KR Cyl 3']) + (tmauto4['KR Cyl 4'])) >= -1), 'Outcome']=1
                        tmauto4.loc[(((tmauto4['KR Cyl 1']) + (tmauto4['KR Cyl 2']) + (tmauto4['KR Cyl 3']) + (tmauto4['KR Cyl 4'])) < -3.9), 'Outcome']=4
                        if (tmauto4['Outcome'].mean()) ==1:
                            suggestion='Clear. If Gear 3 is available, confirm it is clear otherwise deal with those issues first. Review documentation on ideal timing curve. If satisfactory, you may turn up the octane slider 2 notches and relog.'
                        if ((tmauto4['Outcome'].mean()) <2) & ((tmauto4['Outcome'].mean()) >1):
                            suggestion='Review documentation on ideal timing curve. Address the small problem area with the timing map using half of the knock value on the log. For a 3* KR event, you would pull 1.5* from the table in that area.'
                        if ((tmauto4['Outcome'].mean()) >=2):
                            suggestion='Too many problem areas to properly address with the timing table. It is best to drop the octane slider 3 positions and relog. Review for problem areas at that point and reevaluate.'
                        if ((tmauto4['Outcome'].mean()) >3.5):
                            suggestion='Big issues. Scale octane slider back down and start over if on pump gas. If on E30+, start at the slider level where you are seeing no more than 5* advance which is roughly 91/92 on an almost stock curve.'
                        if len(tmauto4.index) < 5:
                            suggestion='RPM sample size is too small. No suggestions possible for this gear. If no large knock (>3*) on this log, log again from 3000-6500rpm in 4th if possible, otherwise 3rd.'
                        if len(df.loc[df['Gear'] == 3].index) > 10:
                            boostg4.to_excel(writer, sheet_name='Analysis',startrow=(len(boostg3.index)+4),startcol=0)
                            wgauto4.to_excel(writer, sheet_name='Wastegate',startrow=(len(wgauto3.index)+23),startcol=0)
                            tmauto4=tmauto4.drop(['Outcome'], axis=1)
                            tmauto4.to_excel(writer, sheet_name='Timing',startrow=(len(tmauto3.index)+5),startcol=0)
                            worksheet2.write_string(len(boostg3.index)+3,0, 'Gear 4 Mean Analysis',workbook.add_format({'bold': True}))
                            worksheet3.write_string((len(wgauto3.index)+22),0, 'Gear 4',workbook.add_format({'bold': True}))
                            worksheet4.write_string((len(tmauto3.index)+4),0, f'Gear 4 - {suggestion}',workbook.add_format({'bold': True}))
                        else:
                            boostg4.to_excel(writer, sheet_name='Analysis',startrow=1,startcol=0)
                            wgauto4.to_excel(writer, sheet_name='Wastegate',startrow=20,startcol=0)
                            tmauto4=tmauto4.drop(['Outcome'], axis=1)
                            tmauto4.to_excel(writer, sheet_name='Timing',startrow=2,startcol=0)
                            worksheet2=writer.sheets['Analysis']
                            worksheet4=writer.sheets['Timing']
                            worksheet2.write_string(0,0, 'Gear 4 Mean Analysis',workbook.add_format({'bold': True}))
                            worksheet3.write_string(19,0, 'Gear 4',workbook.add_format({'bold': True}))
                            worksheet4.write_string(0,0, 'Suggested Timing Changes',workbook.add_format({'bold': True}))
                            worksheet4.write_string(1,0, f'Gear 4 - {suggestion}',workbook.add_format({'bold': True}))
                    if sensor==3:
                        df.loc[(df['PUT Actual'].rolling(window=2).mean().round(0)>2960), 'Boost Sensor']='URGENT - Over 3 bar manifold sensor capability. Scale back boost.'
                    worksheet=writer.sheets['Log']
                    wrap=writer.book.add_format({'text_wrap': True})
                    sfmt1=writer.book.add_format({'bg_color': '#F5FB9D'})
                    sfmt2=writer.book.add_format({'bg_color': '#F3FD5D'})
                    sfmt3=writer.book.add_format({'bg_color': '#ECFB05'})
                    hfmt1=writer.book.add_format({'bg_color': '#FA9B00'})
                    hfmt2=writer.book.add_format({'bg_color': '#FA6A00'})
                    hfmt3=writer.book.add_format({'bg_color': '#FA0000'})
                    nfmt=writer.book.add_format({'bg_color': '#BDBDBD'})
                    wfmt=writer.book.add_format({'bg_color': '#FCF4B3'})
                    afmt=writer.book.add_format({'bg_color': '#D7BDE2'})
                    gfmt=writer.book.add_format({'bg_color': '#00FA1A'}) 
                    if piset==0:
                        worksheet.conditional_format('U2:X1000', {'type': 'formula', 'criteria': '=($U2+$V2+$W2+$X2)<-9.99', 'format': hfmt3})
                        worksheet.conditional_format('U2:X1000', {'type': 'cell', 'criteria': '<', 'value': -5.99, 'format': hfmt3})
                        worksheet.conditional_format('U2:X1000', {'type': 'cell', 'criteria': '<', 'value': -4.99, 'format': hfmt2})
                        worksheet.conditional_format('U2:X1000', {'type': 'cell', 'criteria': '<', 'value': -3.99, 'format': hfmt1})
                        worksheet.conditional_format('U2:X1000', {'type': 'cell', 'criteria': '<', 'value': -2.99, 'format': sfmt2})
                        worksheet.conditional_format('U2:X1000', {'type': 'cell', 'criteria': '<', 'value': -1.99, 'format': sfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -26, 'format': hfmt3})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 26, 'format': hfmt3})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -20, 'format': hfmt2})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 20, 'format': hfmt2})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -16, 'format': hfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 16, 'format': hfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -12, 'format': sfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 12, 'format': sfmt1})
                        worksheet.conditional_format('AA2:AB1000', {'type': 'formula', 'criteria': '=$AA2>$AB2', 'format': hfmt3})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.12)<$F2,$Y2>60)', 'format': hfmt3})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.1)<$F2,$Y2>60)', 'format': hfmt2})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.06)<$F2,$Y2>60)', 'format': hfmt1})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.04)<$F2,$Y2>60)', 'format': sfmt1})
                        worksheet.conditional_format('P2:P1000', {'type': 'formula', 'criteria': '=and($P2<18,$F2>2000,$S2>7)', 'format': hfmt3})
                        worksheet.conditional_format('P2:P1000', {'type': 'formula', 'criteria': '=and($P2<20,$F2>2000,$S2>7)', 'format': sfmt2})
                        worksheet.conditional_format('Y2:Z1000', {'type': 'formula', 'criteria': '=and($Y2>95,$Z2<79,$F2>2000)', 'format': hfmt3})
                        worksheet.conditional_format('G2:G1000', {'type': 'formula', 'criteria': '=and(or($G2>10,$G2<-10),or(and($H2<80,$B2<5000),$B2>5000),$Y2>90)', 'format': hfmt3})
                        worksheet.conditional_format('G2:G1000', {'type': 'formula', 'criteria': '=and(or($G2>5,$G2<-5),or(and($H2<80,$B2<5000),$B2>5000),$Y2>90)', 'format': hfmt1})
                        worksheet.conditional_format('G2:G1000', {'type': 'formula', 'criteria': '=and(or($G2>2,$G2<-2),or(and($H2<80,$B2<5000),$B2>5000),$Y2>90)', 'format': sfmt2})
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and($B10>5000,$J10>90,$R10>2,($AI10-$AI9)<2)', 'format': hfmt3})
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and(($J5-$J3)>3,($J7-$I5)>3,($J10-$J8)>3,$B10>5000,$J10>80,$R10>2,($AI10-$AI9)<2)', 'format': hfmt3}) 
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and(($J5-$J3)>2,($J7-$I5)>2,($J10-$J8)>2,$B10>5000,$J10>80,$R10>2,($AI10-$AI9)<2)', 'format': sfmt2})
                        worksheet.conditional_format('M10:N1000', {'type': 'formula', 'criteria': '=and(($M10-$N10)>.04,($M9-$N9)>.04,($M8-$N8)>.04,($AI10-$AI9)<2)', 'format': hfmt3})
                        worksheet.conditional_format('M10:N1000', {'type': 'formula', 'criteria': '=and(($M10-$N10)>.02,($M9-$N9)>.02,($M8-$N8)>.02,($AI10-$AI9)<2)', 'format': sfmt2}) 
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and((($J10+$J9+$J8)/3)>80,$B10>5000,$J10>80,$R10>2,($AI10-$AI9)<2)', 'format': wfmt})
                        worksheet.conditional_format('L4:L1000', {'type': 'formula', 'criteria': '=and(($L3-$L1)>1,($L4-$L2)>1,$F4>2000)', 'format': sfmt2})
                        worksheet.conditional_format('A3:AI1000', {'type': 'formula', 'criteria': '=($AI3-$AI2)>1', 'format': nfmt})
                        worksheet.conditional_format('A2:AI1000', {'type': 'formula', 'criteria': '=or($AG2>0,$AH2>0)', 'format': afmt})
                        worksheet2.conditional_format('T2:W1000', {'type': 'formula', 'criteria': '=and(($T2+$U2+$V2+$W2)>-.05,IF($T2="",FALSE,TRUE),IF($U2="",FALSE,TRUE),IF($V2="",FALSE,TRUE),IF($W2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and($F2<3,$F2>-3,IF($F2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('L2:M1000', {'type': 'formula', 'criteria': '=and(($L2-$M2)>-0.03,($L2-$M2)<0.03,$N2<0,IF($L2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('O2:O1000', {'type': 'formula', 'criteria': '=and($O2>18,$O2<26,IF($O2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('I2:I1000', {'type': 'formula', 'criteria': '=and($I2<85,$I3<90,IF($I2="",FALSE,TRUE))', 'format': gfmt})
                    if piset==1:
                        worksheet.conditional_format('W2:Z1000', {'type': 'formula', 'criteria': '=($W2+$X2+$Y2+$Z2)<-9.99', 'format': hfmt3})
                        worksheet.conditional_format('W2:Z1000', {'type': 'cell', 'criteria': '<', 'value': -5.99, 'format': hfmt3})
                        worksheet.conditional_format('W2:Z1000', {'type': 'cell', 'criteria': '<', 'value': -4.99, 'format': hfmt2})
                        worksheet.conditional_format('W2:Z1000', {'type': 'cell', 'criteria': '<', 'value': -3.99, 'format': hfmt1})
                        worksheet.conditional_format('W2:Z1000', {'type': 'cell', 'criteria': '<', 'value': -2.99, 'format': sfmt2})
                        worksheet.conditional_format('W2:Z1000', {'type': 'cell', 'criteria': '<', 'value': -1.99, 'format': sfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -26, 'format': hfmt3})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 26, 'format': hfmt3})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -20, 'format': hfmt2})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 20, 'format': hfmt2})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -16, 'format': hfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 16, 'format': hfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '<', 'value': -12, 'format': sfmt1})
                        worksheet.conditional_format('O2:O1000', {'type': 'cell', 'criteria': '>', 'value': 12, 'format': sfmt1})
                        worksheet.conditional_format('AC2:AD1000', {'type': 'formula', 'criteria': '=$AC2>$AD2', 'format': hfmt3})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.12)<$F2,$AA2>60)', 'format': hfmt3})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.1)<$F2,$AA2>60)', 'format': hfmt2})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.06)<$F2,$AA2>60)', 'format': hfmt1})
                        worksheet.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and(($E2*1.04)<$F2,$AA2>60)', 'format': sfmt1})
                        worksheet.conditional_format('P2:P1000', {'type': 'formula', 'criteria': '=and($P2<18,$F2>2000)', 'format': hfmt3})
                        worksheet.conditional_format('P2:P1000', {'type': 'formula', 'criteria': '=and($P2<20,$F2>2000)', 'format': sfmt2})
                        worksheet.conditional_format('AA2:AB1000', {'type': 'formula', 'criteria': '=and($AA2>95,$AB2<79,$F2>2000)', 'format': hfmt3})
                        worksheet.conditional_format('G2:G1000', {'type': 'formula', 'criteria': '=and(or($G2>10,$G2<-10),or(and($H2<80,$B2<5000),$B2>5000),$AA2>90)', 'format': hfmt3})
                        worksheet.conditional_format('G2:G1000', {'type': 'formula', 'criteria': '=and(or($G2>5,$G2<-5),or(and($H2<80,$B2<5000),$B2>5000),$AA2>90)', 'format': hfmt1})
                        worksheet.conditional_format('G2:G1000', {'type': 'formula', 'criteria': '=and(or($G2>2,$G2<-2),or(and($H2<80,$B2<5000),$B2>5000),$AA2>90)', 'format': sfmt2})
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and($B10>5000,$J10>90,$R10>2,($AI10-$AI9)<2)', 'format': hfmt3})
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and(($J5-$J3)>3,($J7-$I5)>3,($J10-$J8)>3,$B10>5000,$J10>80,$S10>2,($AI10-$AI9)<2)', 'format': hfmt3}) 
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and(($J5-$J3)>2,($J7-$I5)>2,($J10-$J8)>2,$B10>5000,$J10>80,$S10>2,($AI10-$AI9)<2)', 'format': sfmt2})
                        worksheet.conditional_format('M10:N1000', {'type': 'formula', 'criteria': '=and(($M10-$N10)>.04,($M9-$N9)>.04,($M8-$N8)>.04,($AK10-$AK9)<2)', 'format': hfmt3})
                        worksheet.conditional_format('M10:N1000', {'type': 'formula', 'criteria': '=and(($M10-$N10)>.02,($M9-$N9)>.02,($M8-$N8)>.02,($AK10-$AK9)<2)', 'format': sfmt2}) 
                        worksheet.conditional_format('J10:J1000', {'type': 'formula', 'criteria': '=and((($J10+$J9+$J8)/3)>80,$B10>5000,$J10>80,$R10>2,($AK10-$AK9)<2)', 'format': wfmt})
                        worksheet.conditional_format('L4:L1000', {'type': 'formula', 'criteria': '=and(($L3-$L1)>1,($L4-$L2)>1,$F4>2000)', 'format': sfmt2})
                        worksheet.conditional_format('A3:AL1000', {'type': 'formula', 'criteria': '=($AK3-$AK2)>1', 'format': nfmt})
                        worksheet.conditional_format('A2:AL1000', {'type': 'formula', 'criteria': '=or($AI2>0,$AJ2>0)', 'format': afmt})
                        worksheet2.conditional_format('W2:Z1000', {'type': 'formula', 'criteria': '=and(($V2+$W2+$X2+$Y2)>-.05,IF($V2="",FALSE,TRUE),IF($W2="",FALSE,TRUE),IF($X2="",FALSE,TRUE),IF($Y2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('F2:F1000', {'type': 'formula', 'criteria': '=and($F2<3,$F2>-3,IF($F2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('L2:M1000', {'type': 'formula', 'criteria': '=and(($L2-$M2)>-0.03,($L2-$M2)<0.03,$N2<0,IF($L2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('O2:O1000', {'type': 'formula', 'criteria': '=and($O2>18,$O2<26,IF($O2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet2.conditional_format('I2:I1000', {'type': 'formula', 'criteria': '=and($I2<85,$I3<90,IF($I2="",FALSE,TRUE))', 'format': gfmt})
                        worksheet.conditional_format('T3:U1000', {'type': 'formula', 'criteria': '=and($U3>0,$U2=0,($M4-$M2)>0.05)', 'format': hfmt3})
                    writer.save()
                    txtbx.insert(tk.END,f'Written {filename}\n','pass')
                    krcols=['KR Cyl 1','KR Cyl 2','KR Cyl 3','KR Cyl 4']
                    df[krcols]=df[krcols] *(-1)
                    if plotmake.get()==1:
                        plt.figure(filename)
                        ax=plt.gca()
                        ax1=ax.twinx()
                        if plottype.get()==0:
                            df.plot(kind='line',x='Time',y='Effective PSI',ax=ax)
                            df.plot(kind='line',x='Time',y='OverallTiming',color='red',ax=ax)
                            df.plot(kind='line',x='Time',y='KR Cyl 1',color='darkorchid',ax=ax)
                            df.plot(kind='line',x='Time',y='KR Cyl 2',color='magenta',ax=ax)
                            df.plot(kind='line',x='Time',y='KR Cyl 3',color='crimson',ax=ax)
                            df.plot(kind='line',x='Time',y='KR Cyl 4',color='mediumslateblue',ax=ax)
                            df.plot(kind='line',x='Time',y='RPM',color='green',ax=ax1)
                        if plottype.get()==1:
                            df.plot(kind='line',x='Time',y='WG Mean',color='red',ax=ax)
                            df.plot(kind='line',x='Time',y='Wastegate Setpoint',color='darkorchid',ax=ax)
                            df.plot(kind='line',x='Time',y='Effective PSI',ax=ax1)
                        if plottype.get()==2:
                            df.plot(kind='line',x='Time',y='Lambda',color='red',ax=ax)
                            df.plot(kind='line',x='Time',y='Lambda Setpoint',color='darkorchid',ax=ax)
                            df.plot(kind='line',x='Time',y='Effective PSI',ax=ax1)
                        ax.legend(loc='upper left', ncol=3)
                        ax1.legend(loc='upper right')
                        ax.axes.get_xaxis().set_visible(False)
                        ax1.axes.get_xaxis().set_visible(False)
                        plt.tight_layout()
                        plt.show(block=False)
            except Exception as e:
                txtbx.insert(tk.END,f'Failed - {e}\n','fail')
                pass
def excelrun():
    os.system(f'start excel.exe "{installdir.get()}\\documentation.xlsx"')
def settingtop():
    topl=tk.Toplevel()
    s1frame=tk.Frame(topl)
    s1frame.grid(columnspan=2,row=1,sticky='nsew')
    tk.Label(s1frame, text='Settings', font='Helvetica 14 bold').pack(fill='x',pady=(5,10),padx=(5,5))
    ttk.Separator(s1frame,orient='horizontal').pack(fill='x',pady=(5,0))
    tk.Label(s1frame, text='Graphs', font='Arial 9 bold').pack(anchor='w', padx=(5,5), pady=(3,0))
    tk.Radiobutton(s1frame, text='On', variable=plotmake,value=1).pack(anchor='w',padx=(5,5),pady=(0,0))
    tk.Radiobutton(s1frame, text='Off', variable=plotmake,value=0).pack(anchor='w',padx=(5,5))
    ttk.Separator(s1frame,orient='horizontal').pack(fill='x',pady=(5,0))
    tk.Label(s1frame, text='Graph Type', font='Helvetica 9 bold').pack(anchor='w',padx=(5,5),pady=(5,5))
    tk.Radiobutton(s1frame, text='Boost/Timing/RPM', variable=plottype,value=0).pack(anchor='w',padx=(5,5),pady=(0,0))
    tk.Radiobutton(s1frame, text='Boost/WG', variable=plottype,value=1).pack(anchor='w',padx=(5,5))
    tk.Radiobutton(s1frame, text='Boost/Lambda', variable=plottype,value=2).pack(anchor='w',padx=(5,5))
    window.wait_window(topl)
def contin():
    contw=tk.Toplevel(window)
    def choiceyn(value):
        global contu
        contu=value
        contw.destroy()
    tk.Label(contw,text=f'Process {filename}?').pack(padx=(10,10), pady=(10,10))
    tk.Button(contw,text='Yes',command=lambda *args: choiceyn(0)).pack(fill='x', pady=(5,5), padx=(5,5))
    tk.Button(contw,text='No',command=lambda *args: choiceyn(1)).pack(fill='x', pady=(5,5), padx=(5,5))
    window.wait_window(contw)
window=tk.Tk()
window.title("Maestro Log Tool")
plotmake=tk.IntVar()
plottype=tk.IntVar()
env=tk.IntVar()
foldir=tk.StringVar()
installdir=tk.StringVar()
env.set(3)
plotmake.set(0)
plottype.set(0)
foldir.set("C:\\")
installdir.set(f'{os.getcwd()}')
tipframe=tk.Frame(window)
tipframe2=tk.Frame(window)
t1frame=tk.Frame(window)
t2frame=tk.Frame(window)
butframe=tk.Frame(window)
textframe=tk.Frame(window)
tipframe.grid(column=0,row=0,columnspan=2,sticky='nsew')
tipframe2.grid(column=0,row=1,columnspan=2,sticky='nsew')
t1frame.grid(column=0,row=2,sticky='nsew')
t2frame.grid(column=1,row=2,sticky='nsew')
butframe.grid(row=3,columnspan=2,sticky='nsew')
textframe.grid(row=5,columnspan=2,sticky='nsew')
tk.Label(tipframe, text='Maestro Log Tool Analyzer',font='Arial 12 bold').pack(fill='x', padx=(2,0), pady=(0,2))
ttk.Separator(tipframe,orient='horizontal').pack(fill='x', pady=(5,7))
tk.Label(tipframe2, text='Log location',font='Arial 9 bold').pack(anchor='w', padx=(2,2), pady=(1,2))
tk.Entry(tipframe2, textvariable=foldir, width=68).pack(fill='x', side='left', padx=(2,5), pady=(3,2))
ttk.Separator(t1frame,orient='horizontal').pack(fill='x', pady=(4,3))
ttk.Separator(t2frame,orient='horizontal').pack(fill='x', pady=(4,3))
tk.Label(t1frame, text='PUT Sensor', font='Arial 9 bold').pack(anchor='w', padx=(5,5), pady=(3,0))
tk.Radiobutton(t1frame, text='3Bar', variable=env,value=3).pack(anchor='w',padx=(5,5),pady=(0,0))
tk.Radiobutton(t2frame, text='4Bar', variable=env,value=4).pack(anchor='w',padx=(5,5),pady=(24,0))
ttk.Separator(t1frame,orient='horizontal').pack(fill='x', pady=(5,0))
ttk.Separator(t2frame,orient='horizontal').pack(fill='x', pady=(5,0))
tk.Label(t1frame, text='Actions', font='Arial 9 bold').pack(anchor='w', padx=(5,5), pady=(3,3))
tk.Label(t2frame, text='', font='Arial 9 bold').pack(anchor='w', padx=(5,5), pady=(3,3))
tk.Button(t1frame, text='Launch documentation', command = excelrun).pack(pady=(5,5),padx=(5,5),fill='x')
tk.Button(t2frame, text='Settings', command = settingtop).pack(pady=(5,),padx=(5,5),fill='x')
tk.Button(butframe, text='Analyze logs', command = logcheck).pack(pady=(5,),padx=(5,5),fill='x')
ttk.Separator(textframe,orient='horizontal').pack(fill='x', pady=(5,0))
tk.Label(textframe, text='Status window',font='Arial 9 bold').pack(anchor='w', padx=(2,0), pady=(0,2))
txtbx=tk.Text(textframe, width=68, height=10, font='Arial 8 bold')
txtbx.pack()
txtbx.tag_config('fail',background='white', foreground='red')
txtbx.tag_config('pass',background='white', foreground='green')
txtbx.tag_config('info',background='white', foreground='blue')
window.resizable(0,0)
window.mainloop()
