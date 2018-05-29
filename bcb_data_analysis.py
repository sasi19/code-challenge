import requests
import os
import calendar
import datetime

# 3rd party packages
import xlrd
import xlwt


FOREIGN_EXCHANGE_TRANSACTION_REPORT = 'foreign_exchange_transaction_report'
FOREIGN_EXCHANGE_POSITION_REPORT = 'foreign_exchange_position_report'


class DownloadFiles:
    """
    This class lets you download files from web resource.
    """

    def __init__(self, url, target_path=None, file_name='bcb_input.xlsx'):
        self.url = url
        self.target_path = target_path if target_path else os.path.dirname(os.path.abspath(__file__))
        self.file_name = file_name

    def download(self):
        """
        Request the file content from the source target and save it to user's target directory.
        """
        response = requests.get(self.url)

        if response.status_code != 200:
            raise Exception(f'Error: Received status code {response.status_code} while downloading the file from url {self.url}.')

        file_path = os.path.join(self.target_path, self.file_name)
        with open(file_path, 'wb') as fp:
            fp.write(response.content)

        return file_path


class GenerateAnalysisReport:

    def __init__(self, file_path=None, report_type=None):
        if report_type not in [FOREIGN_EXCHANGE_TRANSACTION_REPORT, FOREIGN_EXCHANGE_POSITION_REPORT]:
            raise Exception(f'Report has to be one of {FOREIGN_EXCHANGE_TRANSACTION_REPORT}, {FOREIGN_EXCHANGE_POSITION_REPORT}')
        self.file_path = file_path
        self.report_type = report_type

    def read_excel_data(self):
        """
        Reads excel file data and returns the result in list of lists.
        :return:
        """
        workbook = xlrd.open_workbook(self.file_path)
        sheet = workbook.sheet_by_index(0)
        result = [sheet.row_values(rowx) for rowx in range(sheet.nrows)]

        return result

    def analyse_data(self, last_record_date=None):
        """
        Read and process the excel data to get the required output.
        """
        year, month, day = None, None, None
        result = []
        last_processed_record_date = datetime.datetime.strptime(last_record_date, '%m/%d/%Y')
        months_list = [calendar.month_abbr[month_val] for month_val in range(1, 13)]

        excel_data = self.read_excel_data()
        for column in excel_data:
            valid_row = False
            if column[0] and (isinstance(column[0], (int, float))):
                year = int(column[0])
                valid_row = True
            if column[1]:
                if isinstance(column[1], str) and column[1] in months_list:
                    month = column[1]
                    valid_row = True
                elif isinstance(column[1], (int, float)) and int(column[1]) in range(1, 32) and self.report_type == FOREIGN_EXCHANGE_TRANSACTION_REPORT:
                    day = int(column[1])
                    valid_row = True

            if not valid_row:
                continue

            if self.report_type == FOREIGN_EXCHANGE_TRANSACTION_REPORT:
                if not (year and month and day):
                    continue
                row_date = datetime.datetime.strptime(f'{month}/{day}/{year}', '%b/%d/%Y')
            else:
                if not (year and month):
                    continue
                row_date = datetime.datetime.strptime(f'{month}/{year}', '%b/%Y')

            if row_date > last_processed_record_date:
                res = [row_date.strftime('%m/%d/%Y')]
                res.extend(column[2:])
                result.append(res)

        return result

    def generate_report(self, output_file_path=None, last_record_date=None):
        """
        Generates the excel report depending on report type.
        """
        result = self.analyse_data(last_record_date=last_record_date)
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(self.report_type[:30])
        row = 0

        # Add headings by report
        if self.report_type == FOREIGN_EXCHANGE_TRANSACTION_REPORT:
            sheet.write(row, 0, 'Date')
            sheet.write(row, 1, 'BCB_Commercial_Exports_Total')
            sheet.write(row, 2, 'BCB_Commercial_Exports_Advances_on_Contracts')
            sheet.write(row, 3, 'BCB_Commercial_Exports_Payment_Advance')
            sheet.write(row, 4, 'BCB_Commercial_Exports_Others')
            sheet.write(row, 5, 'BCB_Commercial_Imports')
            sheet.write(row, 6, 'BCB_Commercial_Balance')
            sheet.write(row, 7, 'BCB_Financial_Purchases')
            sheet.write(row, 8, 'BCB_Financial_Sales')
            sheet.write(row, 9, 'BCB_Financial_Balance')
            sheet.write(row, 10, 'BCB_Balance')
        else:
            sheet.write(row, 0, 'Date')
            sheet.write(row, 1, 'BCB_FX_Position')

        # Write content
        for columns_set in result:
            row += 1
            for col, column_data in enumerate(columns_set):
                sheet.write(row, col, column_data)

        workbook.save(os.path.join(output_file_path, f'{self.report_type}.xls'))


if __name__ == '__main__':

    output_file_path = os.path.dirname(os.path.abspath(__file__))
    # 1. Generate foreign exchange transaction report
    print("Running foreign exchange transaction report")
    # NOTE: Change the last modified date to get the latest record
    # Move this to persistent storage so that we can run it dynamically.
    last_record_date = '12/22/2017'
    file = DownloadFiles(url='http://www.bcb.gov.br/pec/Indeco/Ingl/ie5-24i.xlsx', file_name='bcb_input_1.xlsx')
    file_path = file.download()
    report = GenerateAnalysisReport(file_path=file_path, report_type=FOREIGN_EXCHANGE_TRANSACTION_REPORT)
    report.generate_report(output_file_path=output_file_path, last_record_date=last_record_date)
    print(f"Output file has been generated at: {os.path.join(output_file_path, FOREIGN_EXCHANGE_POSITION_REPORT)}")

    # 1. Generate foreign exchange position report
    print("Running foreign exchange position report")
    # NOTE: Change the last modified date to get the latest record
    # Move this to persistent storage so that we can run it dynamically.
    last_record_date = '12/02/2017'
    file = DownloadFiles(url='http://www.bcb.gov.br/pec/Indeco/Ingl/ie5-26i.xlsx', file_name='bcb_input_2.xlsx')
    file_path = file.download()
    report = GenerateAnalysisReport(file_path=file_path, report_type=FOREIGN_EXCHANGE_POSITION_REPORT)
    report.generate_report(output_file_path=output_file_path, last_record_date=last_record_date)
    print(f"Output file has been generated at: {os.path.join(output_file_path, FOREIGN_EXCHANGE_POSITION_REPORT)}")


