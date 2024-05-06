import os
import time
import numpy
import smtplib
import mimetypes
import pandas as pd
from seleniumbase import BaseCase
from email.message import EmailMessage
from datetime import timedelta, datetime
BaseCase.main(__name__, __file__)

class NomeTest(BaseCase):
    def login_to_swag_labs(self):
        self.maximize_window()
        self.open('https://tetsnrai.bluekaktus.com/#/login/:tetsnr')
        self.wait_for_ready_state_complete
        self.send_keys('#username', 'vivek')
        self.send_keys('#password', 'admin1')
        self.click('.submitForm')
        self.wait_for_ready_state_complete
        # if self.assert_element_not_present("#company"):
        #     self.click_xpath("//span[text()='Confirm']")
        #     pass
        time.sleep(1)
        self.click('#company')
        self.click('#company > div > ul > li:nth-child(2) > a')
        self.wait_for_ready_state_complete
        self.click('#location')
        self.click('#location > div > ul > li:nth-child(2) > a')
        self.click('.submitForm')
        self.wait_for_ready_state_complete

    def date_Selection(self, Element):
        if datetime.today().strftime("%#A") == 'Monday' or datetime.today().strftime("%#A") =='Monday':
            Ydate = datetime.today() - timedelta(days=2)
            Yday = Ydate.strftime("%#d")
        else:
            Ydate = datetime.today() - timedelta(days=1)
            Yday = Ydate.strftime("%#d")
        for elm in self.find_elements(Element):
            if elm.text == Yday:
                elm.click()
                break

    def date_Selections(self, Element):
        if datetime.today().strftime("%A") == 'Monday' or datetime.today().strftime("%A") =='Monday':
            Ydate = datetime.today() - timedelta(days=2)
            Yday = Ydate.strftime("%d")
        else:
            Ydate = datetime.today() - timedelta(days=1)
            Yday = Ydate.strftime("%d")
        for elm in self.find_elements(Element):
            if elm.text == Yday:
                elm.click()
                break

    def Format_Sheet(self, FileName):
        df = pd.read_excel(FileName, header=[7])
        df.drop(['SUPPLIER NAME'], axis=1, inplace=True)
        writer = pd.ExcelWriter(FileName, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Stock_Report', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Stock_Report']
        worksheet.set_zoom(90)
        currency_format = workbook.add_format({'num_format': '##,##,##0.0'})
        header_format = workbook.add_format(
            {"valign": "vcenter", "align": "center", "bg_color": "#951F06", "bold": True, 'font_color': 'white', 'font_size': 14})
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.set_column('A:Q', 16, currency_format)
        for column in df:
            column_length = max(df[column].astype(
                str).map(len).max(), len(column)+4)
            col_idx = df.columns.get_loc(column)
            writer.sheets['Stock_Report'].set_column(
                col_idx, col_idx, column_length)
        worksheet.set_column('B:B', 70)
        writer.close()

    def Send_selenium_report(self, recipients, Body, directory, subject, filesName):
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['To'] = ', '.join(recipients)
        msg['From'] = 'erpokhla@tetsnrai.com'
        msg.add_alternative('This is a PLAIN TEXT', subtype='plain')
        msg.add_alternative(Body, subtype='html')
        for file in filesName:
            path = os.path.join(directory, file)
            if not os.path.isfile(path):
                continue
            ctype, encoding = mimetypes.guess_type(path)
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'
            maintype, subtype = ctype.split('/', 1)
            with open(path, 'rb') as fp:
                msg.add_attachment(fp.read(),
                                   maintype=maintype,
                                   subtype=subtype,
                                   filename=file)
        server = smtplib.SMTP('mail.tetsnrai.com', 587)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login('erpokhla@tetsnrai.com', 'erpokhla@123')
        server.send_message(msg)
        server.quit()

    def test1_Stock_Report(self):
        self.login_to_swag_labs()
        self.click('#topButton', timeout=50)
        self.wait_for_ready_state_complete
        self.click_link_text('Stock Ledger (New)', timeout=50)
        self.switch_to_window(1)
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="locationID"]', 'FABRIC STORE-O', timeout=50)
        self.wait_for_ready_state_complete
        self.select_option_by_text( '[name="ledgerTypeID"]', 'Store Wise Summary', timeout=50)
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="item_group_id"]', 'Fabric', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="StartDate"]', timeout=50)
        self.click('.btn-clear-wrapper', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="StartDate"]', timeout=50)
        self.date_Selection('div.bs-datepicker-body > table > tbody > tr > td > span[class="ng-star-inserted"]')
        self.click('input[name="EndDate"]', timeout=50)
        self.click('.btn-clear-wrapper', timeout=50)
        self.wait_for_ready_state_complete
        self.click('input[name="EndDate"]', timeout=50)
        self.date_Selection('div.bs-datepicker-body > table > tbody > tr > td > span[class="ng-star-inserted"]')
        self.click('[class="btn btn-info btn-sm"]', timeout=50)
        self.wait_for_element_present('div.ag-pinned-left-cols-container > div:nth-child(3) > div:nth-child(3)',timeout=50)
        self.wait_for_ready_state_complete
        self.click('[class="btn btn-primary btn-sm"]', timeout=50)
        self.wait_for_ready_state_complete
        time.sleep(6)
        old_name = r"E:\ScriptsFiles\downloaded_files\Summary.xlsx"
        new_name = r"E:\ScriptsFiles\downloaded_files\Fabric Stock Report.xlsx"
        os.rename(old_name, new_name)
        self.Format_Sheet(new_name)
        self.refresh_page()
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="locationID"]', 'ACC STORE-O', timeout=50)
        self.wait_for_ready_state_complete
        self.select_option_by_text( '[name="ledgerTypeID"]', 'Store Wise Summary', timeout=50)
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="item_group_id"]', 'Trim', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="StartDate"]', timeout=50)
        self.click('.btn-clear-wrapper', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="StartDate"]', timeout=50)
        self.date_Selection('div.bs-datepicker-body > table > tbody > tr > td > span[class="ng-star-inserted"]')
        self.click('input[name="EndDate"]', timeout=50)
        self.click('.btn-clear-wrapper', timeout=50)
        self.wait_for_ready_state_complete
        self.click('input[name="EndDate"]', timeout=50)
        self.date_Selection('div.bs-datepicker-body > table > tbody > tr > td > span[class="ng-star-inserted"]')
        self.click('[class="btn btn-info btn-sm"]', timeout=50)
        self.wait_for_element_present('div.ag-pinned-left-cols-container > div:nth-child(3) > div:nth-child(3)',timeout=600)
        self.wait_for_ready_state_complete
        self.click('[class="btn btn-primary btn-sm"]', timeout=50)
        self.wait_for_ready_state_complete
        time.sleep(6)
        old_name = r"E:\ScriptsFiles\downloaded_files\Summary.xlsx"
        new_name = r"E:\ScriptsFiles\downloaded_files\Trim Stock Report.xlsx"
        os.rename(old_name, new_name)
        self.Format_Sheet(new_name)
        self.refresh_page()
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="locationID"]', 'ACC STORE-O', timeout=50)
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="ledgerTypeID"]', 'Store Wise Summary', timeout=50)
        self.wait_for_ready_state_complete
        self.select_option_by_text('[name="item_group_id"]', 'Packing', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="StartDate"]', timeout=50)
        self.click('.btn-clear-wrapper', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="StartDate"]', timeout=50)
        self.date_Selection('div.bs-datepicker-body > table > tbody > tr > td > span[class="ng-star-inserted"]')
        self.click('input[name="EndDate"]', timeout=50)
        self.click('.btn-clear-wrapper', timeout=50)
        self.wait_for_ready_state_complete
        self.click('input[name="EndDate"]', timeout=50)
        self.date_Selection('div.bs-datepicker-body > table > tbody > tr > td > span[class="ng-star-inserted"]')
        self.click('[class="btn btn-info btn-sm"]', timeout=50)
        self.wait_for_element_present('div.ag-pinned-left-cols-container > div:nth-child(3) > div:nth-child(3)',timeout=600)
        self.wait_for_ready_state_complete
        self.click('[class="btn btn-primary btn-sm"]', timeout=50)
        self.wait_for_ready_state_complete
        time.sleep(6)
        old_name = r"E:\ScriptsFiles\downloaded_files\Summary.xlsx"
        new_name = r"E:\ScriptsFiles\downloaded_files\Packing Stock Report.xlsx"
        os.rename(old_name, new_name)
        self.Format_Sheet(new_name)
        directory = r'E:/ScriptsFiles/downloaded_files/'
        # recipients = ['erpokhla@tetsnrai.com']
        recipients = [ 'raghav@tetsnrai.com', 'snarang@tetsnrai.com','sandeepm@tetsnrai.com' , 'rajeshsharma@tetsnrai.com', 'erpokhla@tetsnrai.com', 'storeacc@tetsnrai.com', 'fabric@tetsnrai.com', 'mukesh@tetsnrai.com']
        subject = 'Stock Report Okhla ' + datetime.now().strftime('%#d-%b-%Y')
        Body = '''<html>
            <head></head>
            <body>
                <p style="margin: 0;">Dear All</p>
                <br>
                <p style="margin: 0;">Please find attached the Production Detail Report for your review and comments:</p>
                <br>
                <p style="margin: 0;">Thanks & Regards</p>
                <p style="margin: 0;">Raushan Kumar Jha (ERP)</p>
            </body>
            </html>'''
        filesName = ['Packing Stock Report.xlsx', 'Trim Stock Report.xlsx', 'Fabric Stock Report.xlsx']
        self.Send_selenium_report( recipients, Body, directory, subject, filesName)
        os.remove(r"E:\ScriptsFiles\downloaded_files\Packing Stock Report.xlsx")
        os.remove(r"E:\ScriptsFiles\downloaded_files\Trim Stock Report.xlsx")
        os.remove(r"E:\ScriptsFiles\downloaded_files\Fabric Stock Report.xlsx")

    def test2_Production_Report(self):
        self.login_to_swag_labs()
        self.click('#topButton', timeout=50)
        self.wait_for_ready_state_complete
        self.click_link_text('Production Detail Report', timeout=50)
        self.wait_for_ready_state_complete
        self.click( '[admincontrol="main_model.from_date"]>[class="btn btn-primary ng-scope"]', timeout=50)
        self.click('[class="btn btn-sm btn-danger ng-binding"]', timeout=50)
        self.wait_for_ready_state_complete
        self.click( '[admincontrol="main_model.from_date"]>[class="btn btn-primary ng-scope"]', timeout=50)
        self.date_Selections( 'div[ng-switch="datepickerMode"] > table > tbody > tr > td > button > span[class="ng-binding"]')
        self.click('[admincontrol="main_model.to_date"]>[class="btn btn-primary ng-scope"]', timeout=50)
        self.click('[class="btn btn-sm btn-danger ng-binding"]', timeout=50)
        self.wait_for_ready_state_complete
        self.click( '[admincontrol="main_model.to_date"]>[class="btn btn-primary ng-scope"]', timeout=50)
        self.date_Selections('div[ng-switch="datepickerMode"] > table > tbody > tr > td > button > span[class="ng-binding"]')
        time.sleep(6)
        self.click('[id="print1"]', timeout=50)
        self.wait_for_ready_state_complete
        time.sleep(50)
        directory = r'E:\ScriptsFiles\downloaded_files'
        files = os.listdir(directory)
        for index, file in enumerate(files):
            os.rename(os.path.join(directory, file), os.path.join(
                directory, ''.join(['Production Detail Report', '.pdf'])))
        time.sleep(1)
        Ydate = datetime.today() - timedelta(days=1)
        # recipients = ['erpokhla@tetsnrai.com']
        recipients = ['vivek@tetsnrai.com', 'yuvraj@tetsnrai.com', 'sandeepm@tetsnrai.com', 'patomd@tetsnrai.com','raghav@tetsnrai.com', 'snarang@tetsnrai.com', 'rajeshsharma@tetsnrai.com', 'vikram@tetsnrai.com', 'storeacc@tetsnrai.com', 'fabric@tetsnrai.com', 'mukesh@tetsnrai.com', 'shailendra@tetsnrai.com', 'rameshrao@tetsnrai.com', 'tnr@tetsnrai.com', 'erpokhla@tetsnrai.com']
        subject = 'Production Detail Report Okhla ' + Ydate.strftime('%#d-%b-%Y')
        Body = '''<html>
            <head></head>
            <body>
                <p style="margin: 0;">Dear All</p>
                <br>
                <p style="margin: 0;">Please find attached the Production Detail Report for your review and comments:</p>
                <br>
                <p style="margin: 0;">Thanks & Regards</p>
                <p style="margin: 0;">Raushan Kumar Jha (ERP)</p>
            </body>
            </html>'''
        filesName = ['Production Detail Report.pdf']
        time.sleep(1)
        self.Send_selenium_report(recipients, Body, directory, subject, filesName)
        os.remove(r'E:\ScriptsFiles\downloaded_files\Production Detail Report.pdf')

    def test3_Procurement_Report(self):
        self.login_to_swag_labs()
        self.click('#topButton', timeout=50)
        self.wait_for_ready_state_complete
        self.click_link_text('Procurement Status', timeout=50)
        self.wait_for_ready_state_complete
        self.click('[name="item_type"]', timeout=50)
        self.click_link_text("Fabric", timeout=50)
        self.click("button[ng-hide='from_order_date$hide']", timeout=50)
        self.click('[class="btn btn-sm btn-danger ng-binding"]', timeout=50)
        self.click('button[ng-hide="ship_from_date$hide"]', timeout=50)
        self.click('[class="btn btn-sm btn-danger ng-binding"]', timeout=50)
        self.wait_for_ready_state_complete
        self.click("a#SearchList", timeout=50)
        self.wait_for_element_present("[class='ui-grid-row ng-scope']" , timeout=50)
        self.wait_for_ready_state_complete
        self.click("div[ng-if='grid.options.enableSelectAll']", timeout=50)
        self.click("i.ui-grid-icon-menu", timeout=50)
        self.click("li#menuitem-0>button", timeout=50)
        self.wait_for_ready_state_complete
        self.assert_downloaded_file
        old_name = r'E:/ScriptsFiles/downloaded_files/Sheet.xlsx'
        new_name = r'E:/ScriptsFiles/downloaded_files/Procurement Status.xlsx'
        os.rename(old_name, new_name)
        require_cols = ["Style no","Order no","Order Date","Order qty","Order ex qty","Ship Quantity","item","Item qty","Item Average","Total Avg.","Excess (%)","Po no","UOM","Vendor","Po qty","balance to po qty","Receive Qty","Receive (%)","Net Received","Ok Qty.","Qty. Issued to Prod","Misc Issue","Qty. Issued for jobwork","Net Issue","Balance In Hand","Balance in Hand Value"]
        # require_cols = [ 1, 2, 3, 4, 5, 8, 9, 10, 11, 12, 16, 17, 18, 19, 24, 23, 38, 47, 48]
        df = pd.read_excel(new_name, usecols=require_cols)
        writer = pd.ExcelWriter(new_name, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Fabric_Status', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Fabric_Status']
        worksheet.set_zoom(90)
        currency_format = workbook.add_format({'num_format': '##,##,##0.0'})
        currency_format2 = workbook.add_format({'num_format': '##,##,##0'})
        header_format = workbook.add_format( {"valign": "vcenter", "align": "center", "bg_color": "#3C4748", "bold": True, 'font_color': 'white', 'font_size': 13})
        worksheet.autofit()
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        worksheet.set_column('D:E', 15.5, currency_format2)
        worksheet.set_column('G:G', 15.5,currency_format2)
        worksheet.set_column('N:S', 15.5,currency_format2)
        worksheet.set_column('H:I', 15.5,currency_format)
        writer.close()
        directory = r'E:/ScriptsFiles/downloaded_files/'
        # recipients = ['erpokhla@tetsnrai.com']
        recipients = [ 'erpokhla@tetsnrai.com','raghav@tetsnrai.com', 'snarang@tetsnrai.com', 'rajeshsharma@tetsnrai.com', 'vikram@tetsnrai.com', 'sandeepm@tetsnrai.com', 'storeacc@tetsnrai.com', 'fabric@tetsnrai.com', 'mukesh@tetsnrai.com', 'shailendra@tetsnrai.com' ]
        subject = 'Procurement Status Okhla ' + datetime.now().strftime('%#d-%b-%Y')
        Body = '''<html>
            <head></head>
            <body>
                <p style="margin: 0;">Dear All</p>
                <br>
                <p style="margin: 0;">Please find attached the Procurement Status for your review and comments:</p>
                <br>
                <p style="margin: 0;">Thanks & Regards</p>
                <p style="margin: 0;">Raushan Kumar Jha (ERP)</p>
            </body>
            </html>'''
        filesName = ['Procurement Status.xlsx']
        self.Send_selenium_report(recipients, Body, directory, subject, filesName)