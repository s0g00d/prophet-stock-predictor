#This is good for understanding the output: https://www.mikulskibartosz.name/prophet-plot-explained/
import pandas as pd
from fbprophet import Prophet
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from os import path
import numpy as np
import holidays
from datetime import date
import time
from googleapiclient import discovery


DATA_DIR = 'C:/Users/xbsqu/Desktop/Python Learning/Projects/Machine Learning to Predict Stock Prices'

#######Connecting to G Sheet######
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(path.join(DATA_DIR, 'client_secret.json'), scope)
client = gspread.authorize(creds)
sheet = client.open('ML Stock Price Predictor')
historical_data = sheet.get_worksheet(1) #Change sheet here, zero indexed
dashboard = sheet.get_worksheet(0) #Need this to get the prediction period from dashboard C7
prediction_periods = int(dashboard.acell('C7').value) #User Value: GSheet Dashboard C7


###Now that I have data need to do some cleaning###
list_of_lists = historical_data.get_all_values() #Make list of lists to put in df
full_dataset = pd.DataFrame.from_records(list_of_lists)
new_header = full_dataset.iloc[0] #Need to fix the header row
full_dataset = full_dataset[1:]
full_dataset.columns = new_header
full_dataset = full_dataset.reset_index(drop=True) 
full_dataset = full_dataset.replace('#N/A', np.NaN) #Account for any #N/A errors
full_dataset = full_dataset.dropna(axis=0, how='any')

#Need to change the data types
cols = full_dataset.columns
full_dataset[cols[1:]] = full_dataset[cols[1:]].apply(pd.to_numeric, errors='coerce')

#Remove weekend dates
full_dataset['Date'] = pd.to_datetime(full_dataset['Date'], format='%m/%d/%Y')
full_dataset['dow'] = full_dataset['Date'].dt.day_name()
weekdays = ['Monday','Tuesday','Wednesday','Thursday','Friday']
bool_series = full_dataset.dow.isin(weekdays)
full_dataset = full_dataset[bool_series]
del full_dataset['dow']

#Remove US Holidays
current_year = date.today().year
current_year += 2 #add extra years just in case, depending on prediction period & where we're at in the calendar

us_holidays = []
for i in range(2010, current_year):
    for date in holidays.UnitedStates(years=i).items():
        us_holidays.append(str(date[0]))

us_holidays = pd.to_datetime(us_holidays, format='%Y-%m-%d') #Convert to datetime
full_dataset['is_holiday'] = [1 if date in us_holidays else 0 for date in full_dataset['Date']]
full_dataset = full_dataset[full_dataset.is_holiday != 1]
del full_dataset['is_holiday']

#This is start of Prophet tutorial: https://facebook.github.io/prophet/docs/quick_start.html#python-api
df = full_dataset[['Date', 'Open']]
df.columns = ['ds', 'y'] #need to change column names for Prophet
#df Dataframe is the Historic Stock Tab w/o weekends & holidays

m = Prophet()
m.fit(df)

future = m.make_future_dataframe(periods=prediction_periods)

#######Need to remove weekends & holidays from future dates########
future_dates = future.tail(prediction_periods) #This has index from old df
future_dates['dow'] = future_dates['ds'].dt.day_name()
bool_series = future_dates.dow.isin(weekdays)
future_dates = future_dates[bool_series]
del future_dates['dow']

future_dates['is_holiday'] = [1 if date in us_holidays else 0 for date in future_dates['ds']]
future_dates = future_dates[future_dates.is_holiday != 1]


#Need to remove original future dates from Prophet df Future and then append our version w/o weekends
length = len(future['ds'])
future = future.head(length - prediction_periods)
future = future.append(future_dates, ignore_index=True)

#Here's the forecast part
forecast = m.predict(future)

#Plotting the forecast
#fig1 = m.plot(forecast)
#fig2 = m.plot_components(forecast)


###Splitting the data for G Sheet###
num_days_forecasted = len(future_dates)
past_data_from_forecast = forecast.head(length-prediction_periods)
forecasted_data = forecast.tail(num_days_forecasted)

#Getting the important forecasted data up to GSheet
gsheet_future = forecasted_data[['ds', 'yhat', 'yhat_lower', 'yhat_upper']]
gsheet_future['ds'] = gsheet_future['ds'].dt.strftime('%Y-%m-%d') #Turn datetime to str
gsheet_future = gsheet_future.applymap(str) #Turn the prices to str
gsheet_future.insert(1, 'Placeholder', "") #I need to do this for later on & getting the data into the plot tab
futures_tab = sheet.get_worksheet(2)
list_of_df_rows = gsheet_future.to_numpy().tolist()

#Need to format the df so I can write rows to Gsheet
for list in list_of_df_rows:
    ','.join(map(str,list))

#Writing to the G Sheet    
index = 3 #This is set to 3 to account for the header & startpoint, which is most recent day in historic data set
for list in list_of_df_rows:
    futures_tab.insert_row(list, index, 'USER_ENTERED')
    index += 1
    time.sleep(1.5) #Need to slow the for loop so as to not exhaust G Sheet API quota

time.sleep(1.0)

#Need to change datatypes for G Sheet for plotting
plot_info_tab = sheet.get_worksheet(3) #These are zero indexed

###################################################################################
#backtest_tab = sheet.get_worksheet(5) #This is temporary ###############################
###################################################################################

###Let's format the G Sheet for auto-plotting###
#I NEED TO NOT MOVE OVER WEEKENDS OR HOLIDAYS FROM HISTORICAL TO PLOT INFO DUMP####################

#Writing the historic data to the plot tab
top_row = 2
for i in range(1+num_days_forecasted, 1, -1): #The 1's accounts for the header row 
    row_info = historical_data.row_values(i) #Copy from the bottom
    plot_info_tab.insert_row(row_info, top_row, 'USER_ENTERED') #Paste to the top
    top_row += 1 
    time.sleep(1.5) #Now let it breathe
    
time.sleep(1.0)

#Moving over the prediction data
bottom_row = top_row #This is so the future data can append to the end of the historical data that just pasted in 
#num_prediction_lines = len(gsheet_future)
for i in range(2, 3+num_days_forecasted): #Need to add 3 b/c range is non-inclusive & two locked rows at top of Futures tab
    row_info = futures_tab.row_values(i) #Copy from the top
    plot_info_tab.insert_row(row_info, bottom_row, 'USER_ENTERED')  #Paste to the bottom
    bottom_row += 1
    time.sleep(1.5) #Now let it breathe
    
time.sleep(1.0)

#Need to change the date formatting in G Sheet Plot tab for plotting continuity
service = discovery.build('sheets', 'v4', credentials=creds) #Need to connect to Google Sheets API; earlier was Google Drive API
spreadsheetId = '**************************'

reqs = {"requests": [
    {
        "repeatCell": {
        "range": {
          "sheetId": **********,
          "startRowIndex": 0,
          "startColumnIndex": 0,
          "endColumnIndex": 1
        },
        "cell": {
          "userEnteredFormat": {
            "numberFormat": {
              "type": "DATE",
              "pattern": "mm/dd/yyyy"
            }
          }
        },
        "fields": "userEnteredFormat.numberFormat"
      }
    },
    {
      "repeatCell": {
        "range": {
          "sheetId": *********,
          "startRowIndex": 1,
          "startColumnIndex": 1,
          "endColumnIndex": 5
        },
        "cell": {
          "userEnteredFormat": {
            "numberFormat": {
              "type": "NUMBER",
              "pattern": "#.00"
            }
          }
        },
        "fields": "userEnteredFormat.numberFormat"
      }
    }
  ]
}

res = service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, body=reqs).execute()
