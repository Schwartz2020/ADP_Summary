import streamlit as st
import pandas as pd
import numpy as np
import base64
from io import BytesIO
# import os
# import openpyxl as op
# from openpyxl import load_workbook
# from openpyxl.utils.dataframe import dataframe_to_rows
# from openpyxl.styles import Alignment
# from openpyxl.styles import Font


def loadRawADPExcel(rawExcelFile):
    df = pd.read_excel(rawExcelFile, index_col=None, usecols="A,C,D,F,G,H,J,K,L", engine="openpyxl", names=[
                       'Work_code', 'Day', 'Date', 'Work_Time_Frame', 'First_Name', 'Lunch', 'Hours', 'Sup_Info', 'Identify'])
    df = df.where(df.notnull(), None)
    dfAsListofLists = df.values.tolist()
    currentEmployeeName = ""
    dfList = []
    for row in dfAsListofLists:
        excelAColumnValue = row[0]
        excelDColumnValue = row[2]
        excelGColumnValue = row[4]
        excelLColumnValue = row[8]
        if (excelAColumnValue is not None and excelAColumnValue != "Last Name") and (excelGColumnValue is not None and excelGColumnValue != "First Name") and (excelLColumnValue is not None and excelLColumnValue != "Position ID"):
            currentEmployeeName = str(
                excelGColumnValue) + " " + str(excelAColumnValue)
        elif excelDColumnValue is not None:
            dfList.append(
                (currentEmployeeName, row[0], row[1], row[2], row[3], row[5], row[6], row[7]))
        else:
            pass

    cleanRawDataDf = pd.DataFrame(dfList, columns=[
        'Full_Name', 'Work_code', 'Day', 'Date', 'Work_Time_Frame', 'Lunch', 'Hours', 'Sup_Info'])
    return cleanRawDataDf


def getTeamWorkWeekStats(rawDataDf):
    rawDataNoPTODf = rawDataDf.query('Sup_Info!="PTO"')
    rawDataPTODf = rawDataDf.query('Sup_Info=="PTO"')
    x = (rawDataNoPTODf.groupby(['Full_Name', 'Date']).sum().groupby(
        level=0).cumsum().rename(columns={'Hours': 'CumSum'})).reset_index()
    y = (rawDataNoPTODf.groupby(['Full_Name', 'Date']).sum()).reset_index()
    z = x.join(y, lsuffix='L_')
    z['Reg_Hours'] = np.where(z['CumSum'] <= 40, z['Hours'], np.where(
        z['CumSum']-z['Hours'] <= 40, z['Hours']-(z['CumSum']-40), 0))
    z['O_Hours'] = np.where(z['CumSum'] <= 40, 0, np.where(
        z['CumSum']-z['Hours'] <= 40, z['CumSum']-40, z['Hours']))
    # 2
    # teamWorkWeekStatsDf = rawDataDf.assign(
    #     ptoResult=np.where(
    #         rawDataDf['Sup_Info'] == 'PTO', rawDataDf.Hours, 0)
    # ).merge(z, how='left', left_on=['Full_Name', 'Date'], right_on=['Full_NameL_', 'DateL_'], suffixes=(None, "_y")).groupby(['Date', 'Day']).agg({'Reg_Hours': 'sum', 'O_Hours': 'sum', 'ptoResult': 'sum', 'Hours': 'sum'}).rename(columns={'Reg_Hours': 'Work_Hours', 'O_Hours': 'Overtime_Hours', 'ptoResult': 'PTO_Hours', 'Hours': 'Total_Hours'}).reset_index()
    appendedDf = z.append(rawDataPTODf, ignore_index=True)
    teamWorkWeekStatsDf = appendedDf.assign(
        ptoResult=np.where(
            appendedDf['Sup_Info'] == 'PTO', appendedDf.Hours, 0)
    ).groupby(['Date']).agg({'Reg_Hours': 'sum', 'O_Hours': 'sum', 'ptoResult': 'sum', 'Hours': 'sum'}).rename(columns={'Reg_Hours': 'Work_Hours', 'O_Hours': 'Overtime_Hours', 'ptoResult': 'PTO_Hours', 'Hours': 'Total_Hours'}).reset_index()
    return teamWorkWeekStatsDf


def sumGreaterThanZero(pandasSeries):
    if pandasSeries.sum() == 1:
        return 'O'  # Worked and did not take lunch
    elif pandasSeries.sum() >= 1:
        return 'X'  # Worked and took lunch
    else:
        return None


def getDaysOfWeek(ungroupedData):
    ungroupedDataDaysOfWeek = ungroupedData.assign(
        Sunday=np.where(ungroupedData['Day'] == 'Sun', 1, 0),
        Monday=np.where(ungroupedData['Day'] == 'Mon', 1, 0),
        Tuesday=np.where(ungroupedData['Day'] == 'Tue', 1, 0),
        Wednesday=np.where(ungroupedData['Day'] == 'Wed', 1, 0),
        Thursday=np.where(ungroupedData['Day'] == 'Thu', 1, 0),
        Friday=np.where(ungroupedData['Day'] == 'Fri', 1, 0),
        Saturday=np.where(ungroupedData['Day'] == 'Sat', 1, 0)
    )
    groupedDataDaysOfWeek = ungroupedDataDaysOfWeek.groupby(['Full_Name']).agg(
        {'Sunday': sumGreaterThanZero, 'Monday': sumGreaterThanZero, 'Tuesday': sumGreaterThanZero, 'Wednesday': sumGreaterThanZero, 'Thursday': sumGreaterThanZero, 'Friday': sumGreaterThanZero, 'Saturday': sumGreaterThanZero})
    return groupedDataDaysOfWeek


def getNonOvertimeHours(pandasSeries):
    if pandasSeries.sum() > 40:
        return 40.0
    else:
        return pandasSeries.sum()


def getOvertimeHours(pandasSeries):
    if pandasSeries.sum() > 40:
        return pandasSeries.sum()-40.0
    else:
        return 0.0


def getDriverWorkDayStats(rawDataDf):
    rawDataNoPTODf = rawDataDf.query('Sup_Info!="PTO"')
    groupedDataDaysOfWeek = getDaysOfWeek(rawDataNoPTODf)
    rdnpGroupByDf = rawDataNoPTODf.groupby(['Full_Name'])
    daysWorked = rdnpGroupByDf['Date'].nunique().to_frame(
        name='Days_Worked')
    driverWorkDayStatsDf = daysWorked.join(
        rdnpGroupByDf.agg({'Hours': getNonOvertimeHours}).rename(columns={'Hours': 'Work_Hours'})).join(rdnpGroupByDf.agg({'Hours': getOvertimeHours}).rename(columns={'Hours': 'Overtime_Hours'})).join(rdnpGroupByDf.agg({'Hours': 'sum'}).rename(columns={'Hours': 'Total_Work_Hours'})).join(groupedDataDaysOfWeek).reset_index()
    driverWorkDayStatsDf['Hours_Owed'] = np.where(
        (8*driverWorkDayStatsDf['Days_Worked']-driverWorkDayStatsDf['Work_Hours']) > 0, 8*driverWorkDayStatsDf['Days_Worked']-driverWorkDayStatsDf['Work_Hours'], 0)
    return driverWorkDayStatsDf


def getPTOStats(rawDataDf):
    rawDataPTODf = rawDataDf.query('Sup_Info=="PTO"')
    groupedDataDaysOfWeek = getDaysOfWeek(rawDataPTODf)
    rdpGroupByDf = rawDataPTODf.groupby(['Full_Name'])
    daysPTO = rdpGroupByDf['Date'].nunique().to_frame(name='Days_PTO')
    ptoStatsDf = daysPTO.join(rdpGroupByDf.agg({'Sup_Info': 'count'}).join(rdpGroupByDf.agg({'Hours': 'sum'})).join(groupedDataDaysOfWeek).rename(columns={
        'Sup_Info': 'PTO_Instances', 'Hours': 'PTO_Hours'})).reset_index()
    return ptoStatsDf


def getMissingLunchClockouts(rawDataDf):
    rawDataNoPTODf = rawDataDf.query('Sup_Info!="PTO"')
    missingLunchClockoutsDf = rawDataNoPTODf.groupby(['Full_Name', 'Date']).filter(
        lambda df: (df.Lunch != "LP").all())
    return missingLunchClockoutsDf


def getMissingLunchInstances(rawDataDf):
    rawDataNoPTODf = rawDataDf.query('Sup_Info!="PTO"')
    missingLunchInstancesDf = rawDataNoPTODf.groupby(['Full_Name', 'Date']).filter(
        lambda df: len(df) < 2)
    return missingLunchInstancesDf


def getNotClockedOutInstances(rawDataDf):
    notClockedOutInstancesDf = rawDataDf.query(
        'Work_Time_Frame.str.count(":") < 2')
    return notClockedOutInstancesDf


def as_text(value):
    if value is None:
        return ""
    return str(value)


def sendDataToExcelFile(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf,
                        ptoStatsDf,  missingLunchClockoutsDf, missingLunchInstancesDf, notClockedOutInstancesDf):

    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    cell_format = workbook.add_format()
    cell_format.set_align('center')
    #
    rawDataDf.to_excel(writer, index=False, startrow=1,
                       header=False, sheet_name='Raw Data')
    columns1 = [{'header': column} for column in rawDataDf.columns]
    (max_row1, max_col1) = rawDataDf.shape
    (writer.sheets['Raw Data']).add_table(
        0, 0, max_row1, max_col1-1, {'columns': columns1})
    (writer.sheets['Raw Data']).set_column('A:M', 18, cell_format)
    #
    teamWorkWeekStatsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Team Hours')
    columns2 = [{'header': column} for column in teamWorkWeekStatsDf.columns]
    (max_row2, max_col2) = teamWorkWeekStatsDf.shape
    (writer.sheets['Team Hours']).add_table(
        0, 0, max_row2, max_col2-1, {'columns': columns2})
    (writer.sheets['Team Hours']).set_column('A:M', 18, cell_format)
    #
    driverWorkDayStatsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Driver Hours')
    columns3 = [{'header': column} for column in driverWorkDayStatsDf.columns]
    (max_row3, max_col3) = driverWorkDayStatsDf.shape
    (writer.sheets['Driver Hours']).add_table(
        0, 0, max_row3, max_col3-1, {'columns': columns3})
    (writer.sheets['Driver Hours']).set_column('A:M', 14, cell_format)
    #
    ptoStatsDf.to_excel(writer, index=False, startrow=1,
                        header=False, sheet_name='PTO Hours')
    columns4 = [{'header': column} for column in ptoStatsDf.columns]
    (max_row4, max_col4) = ptoStatsDf.shape
    (writer.sheets['PTO Hours']).add_table(
        0, 0, max_row4, max_col4-1, {'columns': columns4})
    (writer.sheets['PTO Hours']).set_column('A:M', 18, cell_format)
    #
    missingLunchClockoutsDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Missing Lunch Clockouts')
    columns5 = [{'header': column}
                for column in missingLunchClockoutsDf.columns]
    (max_row5, max_col5) = missingLunchClockoutsDf.shape
    (writer.sheets['Missing Lunch Clockouts']).add_table(
        0, 0, max_row5, max_col5-1, {'columns': columns5})
    (writer.sheets['Missing Lunch Clockouts']
     ).set_column('A:M', 18, cell_format)
    #
    missingLunchInstancesDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Missing Lunches')
    columns6 = [{'header': column}
                for column in missingLunchInstancesDf.columns]
    (max_row6, max_col6) = missingLunchInstancesDf.shape
    (writer.sheets['Missing Lunches']).add_table(
        0, 0, max_row6, max_col6-1, {'columns': columns6})
    (writer.sheets['Missing Lunches']).set_column('A:M', 18, cell_format)
    #
    notClockedOutInstancesDf.to_excel(
        writer, index=False, startrow=1, header=False, sheet_name='Missing Clockouts')
    columns7 = [{'header': column}
                for column in notClockedOutInstancesDf.columns]
    (max_row7, max_col7) = notClockedOutInstancesDf.shape
    (writer.sheets['Missing Clockouts']).add_table(
        0, 0, max_row7, max_col7-1, {'columns': columns7})
    (writer.sheets['Missing Clockouts']).set_column('A:M', 18, cell_format)
    #
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def get_table_download_link(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf,
                            ptoStatsDf,  missingLunchClockoutsDf, missingLunchInstancesDf, notClockedOutInstancesDf):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = sendDataToExcelFile(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf, ptoStatsDf,
                              missingLunchClockoutsDf, missingLunchInstancesDf, notClockedOutInstancesDf)
    b64 = base64.b64encode(val)  # val looks like b'...'
    # decode b'abc' => abc
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="ADP_Summary.xlsx">Download ADP Summary Excel File</a>'


st.set_page_config(page_title='ADP Summary Generator')
st.title("ADP Summary Generator")
st.subheader("Please upload your ADP Excel File:")

uploaded_file = st.file_uploader(
    'Select the ADP Time Sheet .xlsx file', type='xlsx')
if uploaded_file:
    st.write('File Uploaded.')
    rawDataDf = loadRawADPExcel(uploaded_file)
    teamWorkWeekStatsDf = getTeamWorkWeekStats(rawDataDf)
    driverWorkDayStatsDf = getDriverWorkDayStats(rawDataDf)
    ptoStatsDf = getPTOStats(rawDataDf)
    missingLunchClockoutsDf = getMissingLunchClockouts(rawDataDf)
    missingLunchInstancesDf = getMissingLunchInstances(rawDataDf)
    notClockedOutInstancesDf = getNotClockedOutInstances(rawDataDf)
    # Send out data
    st.markdown(get_table_download_link(rawDataDf, teamWorkWeekStatsDf, driverWorkDayStatsDf, ptoStatsDf,
                                        missingLunchClockoutsDf, missingLunchInstancesDf, notClockedOutInstancesDf), unsafe_allow_html=True)
    st.write("File has been processed. Click link to download file.")
