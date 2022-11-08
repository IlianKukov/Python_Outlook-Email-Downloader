from functools import reduce
import win32com.client
import pandas as pd
import zipfile
import pyodbc
import time
import sys
import os
import re

##/ - Connection Variables - /##
conn_server_name = 'Example_Server'  # Enter your server name
conn_db_name = 'ExampleDB'  # Enter your DB
conn_db_view = '[dbo].[v_sys_email_download]'  # Enter your DB Email Table
conn_db_log = '[dbo].[sys_email_download_log]'  # Enter your DB Log Table


class EmailDownloader:

    def __init__(self, server_name, db_name, db_view, db_log):
        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.server_name = 'Driver={SQL Server};''Server=' + server_name + ';''Database=' \
                           + db_name + ';''Trusted_Connection=yes;'
        self.get_mails = 'SELECT * FROM [' + db_name + '].' + db_view
        self.command_line = None
        self.db_name = db_name
        self.SSMS_job_name = None
        self.SSMS_job = None
        self.args_arr = None
        self.log_table = db_log
        self.mail_counter = 0
        self.del_items = None
        self.msg = None
        self.time_now = 0
        self.log_id = 0
        self.atm_save_file = None
        self.save_dir = None
        self.atm_name = None
        self.new_name = None
        self.res_unzip = False
        self.res_rename = False
        self.res_csv = False
        self.res_del = False
        self.res_unr = False
        self.file_rework = False
        self.dict_args = {}

    def db_exec(self, command, sp=False):
        conn = pyodbc.connect(self.server_name)
        cursor = conn.cursor()
        cursor.execute(command)
        if not sp:
            return cursor.fetchall()
        else:
            self.time_now = time.strftime('%Y-%m-%d %H:%M:%S')
            return True

    def download_mails(self):
        self.dict_args = self.command_line[4].split(";")
        self.save_dir = self.command_line[2]
        self.SSMS_job_name = self.command_line[3]
        self.SSMS_job = "EXEC msdb.dbo.sp_start_job '" + self.SSMS_job_name + "'"
        self.args_arr = self.command_line[4].split('/')
        self.log_id = self.command_line[5]

        emails = eval(reduce(lambda x, y: x + ".Folders['" + y + "']",
                             self.command_line[0].split('/'), "self.outlook") + '.Items')

        self.mail_counter = 0
        for self.msg in reversed(emails):
            if re.search(self.command_line[1], self.msg.subject) is not None:
                for atm in self.msg.Attachments:
                    self.atm_name = atm.FileName
                    self.atm_save_file = os.path.join(str(self.save_dir), str(self.atm_name))
                    atm.SaveAsFile(self.atm_save_file)
                    self.mail_counter += 1
                    EmailDownloader.actions(self)
        if self.mail_counter > 0:
            return True
        else:
            return False

    def actions(self):
        if self.dict_args != '':
            for i in self.dict_args:
                if 'unzip' in i:
                    self.res_unzip = EmailDownloader.unzip_att(self)
                elif 'rename' in i:
                    self.res_rename = EmailDownloader.rename(self)
                elif 'csv' in i:
                    self.res_csv = EmailDownloader.conv_to_csv(self)
                elif 'unr' in i:
                    self.res_unr = self.msg.Unread = True
                elif 'del' in i:
                    self.res_del = EmailDownloader.msg_delete(self)

            sys.stdout.write("\n %s/ %s Atms/ Subj: %s/ UNZIP: %s/ Rename: %s/ CSV: %s/ DEL: %s/ UnRead: %s"
                             % (time.strftime("%Y-%m-%d %H:%M:%S"), self.mail_counter, self.msg.subject,
                                self.res_unzip, self.res_rename, self.res_csv, self.res_del,
                                self.res_unr))

    def rename(self):
        for i in self.dict_args:
            if i.startswith("rename"):
                self.new_name = i.split("=")[1]
                pre, ext = os.path.splitext(self.atm_save_file)
                filename_final = os.path.join(str(self.save_dir), str(self.new_name + ext))
                if os.path.exists(filename_final):
                    os.remove(filename_final)
                time.sleep(0.5)
                os.rename(self.atm_save_file, filename_final)
                self.atm_save_file = filename_final
                return True
            elif i.startswith("Drename"):
                self.new_name = i.split("=")[1]
                pre, ext = os.path.splitext(self.atm_save_file)
                msg_time = self.msg.ReceivedTime.strftime('%Y%m%d_%H%M%S')
                filename_final = os.path.join(str(self.save_dir), str(self.new_name + msg_time + ext))
                if os.path.exists(filename_final):
                    os.remove(filename_final)
                time.sleep(0.5)
                os.rename(self.atm_save_file, filename_final)
                self.atm_save_file = filename_final
                return True
        else:
            return False

    def unzip_att(self):
        with zipfile.ZipFile(self.atm_save_file, 'r') as zip_ref:
            zip_ref.extractall(self.save_dir)
            old_file = self.atm_save_file
            self.atm_name = zip_ref.namelist()[0]
            self.atm_save_file = os.path.join(str(self.save_dir), str(self.atm_name))
            zip_ref.close()
            time.sleep(1)
            os.remove(old_file)
        return True

    def conv_to_csv(self):
        old_filename = self.atm_save_file
        read_file = pd.read_excel(self.atm_save_file)
        pre, ext = os.path.splitext(self.atm_save_file)
        self.atm_save_file = pre + '.csv'
        if os.path.exists(self.atm_save_file):
            os.remove(self.atm_save_file)
        read_file.to_csv(self.atm_save_file, index=False, header=True)
        time.sleep(1)
        if os.path.exists(old_filename):
            os.remove(old_filename)
        return True

    def msg_delete(self):
        self.del_items = self.outlook.Folders(self.command_line[0].split('/')[0]).Folders('Deleted Items')
        self.msg.Move(self.del_items)
        return True

    def result_logger(self):
        conn = pyodbc.connect(self.server_name)
        cursor = conn.cursor()
        if Job.SSMS_job_name == '0':
            self.time_now = time.strftime('%Y-%m-%d %H:%M:%S')
        query = "INSERT INTO [" + self.db_name + "]." + self.log_table + " VALUES('" + str(self.log_id) + "','" + str(
            self.time_now) + "','" + str(self.mail_counter) + "')"
        # print(query)
        try:
            cursor.execute(query)
            conn.commit()
        except Exception as e:
            (error_code, error_message) = e


##/ - START ORDER - /##
script_iter = 1
while True:
    Job = EmailDownloader(conn_server_name, conn_db_name, conn_db_view, conn_db_log)
    Mails_list = Job.db_exec(Job.get_mails)
    for row in range(len(Mails_list)):
        Job.command_line = Mails_list[row]
        is_downloaded = Job.download_mails()
        if is_downloaded and not Job.SSMS_job_name == '0':
            print('\n SP %s : %s' % (str(Job.SSMS_job_name), str(Job.db_exec(Job.SSMS_job, sp=True))))
            Job.result_logger()
        elif Job.SSMS_job_name.startswith('0'):
            Job.result_logger()
    print("\n\n\n Iteration count: %s - %s \n\n" % (script_iter, time.strftime('%Y-%m-%d %H:%M:%S')))
    script_iter += 1
    time.sleep(15)

##__name__ == '__main__'
__author__ = "IlianKukov"
