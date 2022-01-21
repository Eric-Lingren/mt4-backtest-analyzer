from bs4 import BeautifulSoup as bs
import pandas as pd


class Mt4_Report_Cleaner():
    def __init__(self, input_file, output_path):
        self.input_file = input_file
        self.output_path = output_path
        self.output_filename = ''
        self.soup = None
        self.summary_df = None
        self.trades_df = None
    

    def run_cleaner(self):
        self.open_report()
        summary_list = self.scrape_summary_data()
        self.build_summary_data_output(summary_list)
        trades_list = self.scrape_trade_data()
        self.build_trade_data_output(trades_list)
        self.write_data_to_xls()
        return self.output_filename
    

    def open_report(self):
        with open(self.input_file) as file:
            self.soup = bs(file, 'html.parser')
        output_filename = self.input_file.replace('.htm', '.xlsx') 
        slice_index = output_filename.rfind('/')
        self.output_filename = self.output_path + output_filename[slice_index:]
    

    #* Scrape the trade sumary out of the report webpage
    def scrape_summary_data(self):
        trade_summary = self.soup.findAll('table')[0]
        trade_summary_rows = trade_summary.find_all('tr')
        trade_summary_list = []
        for tr in trade_summary_rows:
            td = tr.find_all('td')
            row = [tr.text for tr in td]
            trade_summary_list.append(row)
        return trade_summary_list
    

    #* Cleans random table row lengths into 1:1 matching key value pairs
    def build_summary_data_output(self, trade_summary_list):
        system_name = self.soup.findAll('b')[1].text
        trade_summary_list_final = [['System Name', system_name]]
        for row in trade_summary_list:
            clean_row = []
            for item in row:
                clean_row = [item for item in row if item != ""]
            if len(clean_row) == 2:
                if(clean_row[0] == 'Period'): # Break chart period and duration into different data points
                    last_open_parenthesis_index = clean_row[1].rfind("(")
                    cleaned_date = clean_row[1][:last_open_parenthesis_index]
                    last_close_parenthesis_index = cleaned_date.rfind(")")
                    duration = cleaned_date[last_close_parenthesis_index+3:].rstrip()
                    period = cleaned_date[:last_close_parenthesis_index+1]
                    trade_summary_list_final.append([clean_row[0], period])
                    trade_summary_list_final.append(['Duration', duration])
                else:
                    trade_summary_list_final.append(clean_row)
            elif len(clean_row) == 6:
                for i in range(0, len(clean_row), 2):
                    trade_summary_list_final.append(clean_row[i:i+2])
            elif len(clean_row) == 5:
                new_string1_key = clean_row[0] + ' ' + clean_row[1]
                new_string1_value = clean_row[2]
                new_string2_key = clean_row[0] + ' ' + clean_row[3]
                trade_summary_list_final.append([new_string1_key, new_string1_value])
                trade_summary_list_final.append([new_string2_key, clean_row[4]])
        self.summary_df = pd.DataFrame(trade_summary_list_final,columns=['Key', 'Value'])


    #* Scrape the trade data out of the report webpage
    def scrape_trade_data(self):
        tradeData = self.soup.findAll('table')[1]
        trade_data_rows = tradeData.find_all('tr')
        trade_data_list = []
        # Buld a dataframe of the trade data:
        for tr in trade_data_rows:
            td = tr.find_all('td')
            row = [tr.text for tr in td]
            trade_data_list.append(row)
        df = pd.DataFrame(trade_data_list[1:],columns=trade_data_list[0])
        return df


    def build_trade_data_output(self, trades_list):
        df = trades_list
        # Modify DF to get unique orders and their duration
        df['Time'] = pd.to_datetime(df['Time'])
        final_trade_data = pd.DataFrame()
        order_numbers = df['Order'].unique()
        for order_number in order_numbers:
            order_number = int(order_number)
            order_number = order_number - 1
            order_pair = df.loc[df['Order'] == order_numbers[order_number]]
            open_time = order_pair.iloc[0]['Time']
            close_time = order_pair.iloc[1]['Time']
            duration = close_time - open_time
            days, seconds = duration.days, duration.seconds
            hours = days * 24 + seconds // 3600
            minutes = (seconds % 3600) // 60
            seconds = seconds % 60
            duration_in_min = (hours*60)+minutes
            duration_in_hrs = round( ((minutes/60)+hours), 2)
            order_pair.insert(2, 'Duration (hrs)', [duration_in_hrs, duration_in_hrs])
            final_trade_data = pd.concat([final_trade_data, order_pair])
        self.trades_df = final_trade_data


    #*  Output the contents of the trade data table in excel format
    def write_data_to_xls(self):
        #  Create a Pandas Excel writer using XlsxWriter as the engine.
        with pd.ExcelWriter(self.output_filename, engine='xlsxwriter') as writer:    
            # Write each dataframe to a different worksheet.
            self.summary_df.to_excel(writer, sheet_name='summary_data', index=False)
            self.trades_df.to_excel(writer, sheet_name='trade_data', index=False)
    