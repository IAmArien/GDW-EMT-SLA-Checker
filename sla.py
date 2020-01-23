import time, sys
import pytz as time_zone
import re as regex
import json as javascript_notation
import pdb as debugger
import openpyxl as py_excel
from os import path, listdir, getcwd, chdir
from openpyxl import load_workbook
from datetime import datetime

class Contents(object):
    def __init__(self, email_contents: list):
        super().__init__()
        self.email_contents = email_contents    
        
    def substantiate_email_contents(self, email_columns: list) -> bool:
        valid_columns_count = []
        for each_email_columns in email_columns:
            if each_email_columns in self.email_contents:
                valid_columns_count.append(each_email_columns)        
        return len(valid_columns_count) == len(self.email_contents)

    def yield_excel_columns(self, work_sheet: object, columns: int) -> list:
        excel_columns = []
        for i in range(1, sum([columns, 1])):
            excel_columns.append(work_sheet.cell(1, i).value)
        return excel_columns        

class Auth(object):
    def __init__(self, class_pst, class_email_conf):
        super().__init__()
        self.class_pst = class_pst
        self.class_email_conf = class_email_conf
    
    def configure(self, config_file_dict: dict) -> dict:               
        config_file = self.class_email_conf.get_email_config_file(config_file_dict["Email-Monitoring"]["Config-Directory-Name"], config_file_dict["Email-Monitoring"]["Config-File-Extension"])
        if not config_file == None:
            valid_emails = self.class_email_conf.acquire_valid_emails(config_file)
            current_time = self.class_pst.get_current_time()
            if not current_time == None:
                standard_pt = self.class_pst.time_conversion(current_time, config_file_dict["Email-Monitoring"]["Datetime-Format"])                                
                if standard_pt == None:                    
                    sys.exit(0)
            else:                
                sys.exit(0)
        else:            
            sys.exit(0)
        return dict(emails=valid_emails, standard_pt=standard_pt)

    @staticmethod
    def confirm_sla(current_datetime: datetime, email_contents: dict, valid_emails_list: list, configuration_file: dict) -> tuple:        
        if email_contents["load_end_time"] == None:
            return (None)
        if email_contents["load_end_time"] <= current_datetime:
            is_in_valid_emails = False            
            for each_valid_emails in valid_emails_list:                                              
                if email_contents["email"] in each_valid_emails:                   
                    is_in_valid_emails = True                                        
                else:
                    is_in_valid_emails = False                
            if not is_in_valid_emails:
                return ('Invalid Email')
            # MATCH DT THROWN WITH VALID EMAILS LIST SCHEDULE
            # UPDATE ROW FLAG
            # REPORT STATUS VIA NOTEPAD            
        else:
            return ('Load End Time is greater than Current Date Time')  

class Render_Excel(object):
    def __init__(self):
        super().__init__()
        self.excel_data = object

    # GET LOAD END TIME FROM THE BODY
    @classmethod
    def yield_load_end_time(self, email_body: str, email_address: str, config_file: dict) -> datetime:
        try:
            date = regex.search(r"\d{2}/\d{2}/\d{2}", email_body)
            time = regex.search(r"\d{2}:\d{2}:\d{2}", email_body)
            if date is not None: date = date.group()
            else: date = ""
            if time is not None: time = time.group()
            else: time = ""
            self.excel_data = "%s %s" % (date, time)
        except Exception as e:
            print("Regex Error LET: %s" % e)
            return None
        return datetime.strptime(self.excel_data, config_file["Email-Monitoring"][email_address]["LET-Format"][0])

    # GET EMAIL ADDRESS FROM 'FROM' COLUMN
    @classmethod
    def yield_email_address(self, from_content: str) -> str:
        try:
            email_address = regex.search(r'\S+@\S+', from_content).group()        
            if len(regex.findall(r'(<|>)', email_address)) > 0 :
                email_address = regex.sub(r"(<|>)","", email_address)          
            self.excel_data = "%s" % (email_address)
        except (Exception, AttributeError) as e:
            print("Regex Error EA: %s" % e)
            return None
        return self.excel_data

    # GET COLUMN INDEX OF A SPECIFIC EMAIL CONTENT SUCH AS From, To, Subject etc.
    # in a SPECIFIC ROW in an EXCEL FILE
    @staticmethod
    def get_column_index(column_name: str, excel_columns: list) -> int:
        if column_name in excel_columns:
            return sum([excel_columns.index(column_name), 1])
        return 0

# NOTE SYNTAX
# INITIALIZATION
# GETTING LOAD END TIME OF THE EMAIL
# GETTING THE EMAIL ADDRESS OF THE SENDER
# DECLARE THE FOLLOWING SYNTAX BELOW THE CLASS
# --------------------------------------------------------------------------------
# sla_excel = Excel()                                                               // init
# email_address = sla_excel.yield_email_address(work_sheet.cell(rows, cols).value)  // getting 'From' value of an email from a specific rows and cols and retrieve its respective email address
# load_end_time = sla_excel.yield_load_end_time(work_sheet.cell(rows, cols).value)  // getting 'Subject' valud of an email from a specific rows and cols of an excel file and retrieve its respective load end time
# ---------------------------------------------------------------------------------

class Email_Config(object):
    def __init__(self):
        super().__init__()

    # RETRIEVE ALL VALID EMAILS NEEDED
    @staticmethod
    def acquire_valid_emails(email_config_file: str) -> list:
        valid_emails = []
        try:
            with open(r"%s\%s" % (getcwd(), email_config_file), "r", encoding="utf-8") as config_file:
                emails = config_file.read()          
                config_file.close() 
            emails = javascript_notation.loads(emails)       
            for each_email in emails["Email-Monitoring"]["Emails"]:
                if len(regex.findall("@x*", each_email)) > 0:          
                    valid_emails.append(each_email) 
        except FileNotFoundError as file_error:
            print("Config File Not Found: %s" % file_error)
            return None       
        return valid_emails    

    # GET THE DIRECTORY AND FILENAME OF config FILE CONTAINING THE VALID EMAILS
    @staticmethod
    def get_email_config_file(directory: str, config_extension: str) -> str:        
        if path.isdir(r"%s\%s" % (getcwd(), directory)):
            for each_files in listdir(path=directory):
                spl_filename = path.splitext(each_files)                     
                if spl_filename[1] == config_extension:
                    return r"%s\{filename}{extension}".format(filename=spl_filename[0], extension=spl_filename[1]) % directory
            else:
                return None
        else:
            print("Path '{directory}' not found".format(directory=directory))
            return None

# NOTE SYNTAX
# INITIALIZATION
# GETTING directory of config file containing the list of valid emails
# OBTAINING the list of valid emails from the directory given
# DECLARE THE FOLLOWING CODES BELOW THE CLASS
# --------------------------------------------------------------
# >>> config = Email_Config()                                   // init
# >>> config_file = config.get_email_config_file("emails")      // get directory
# >>> config.acquire_valid_emails(config_file)                  // get valid emails
# --------------------------------------------------------------

class PST(object):
    def __init__(self):
        super().__init__()
        self.local_date_time = object

    # GETTING THE CURRENT TIME
    @staticmethod
    def get_current_time() -> datetime:
        try:
            utc_moment_naive = datetime.utcnow()
            utc_moment = utc_moment_naive.replace(tzinfo=time_zone.utc)
        except Exception as e:
            print("Date Time Error: %s" % e)
            return None
        return utc_moment    
        
    # CONVERTING TIME TO PST AND ITS RESPECTIVE FORMAT mm/dd/YY H:M:S
    @classmethod
    def time_conversion(self, current_time, local_time_format: str) -> datetime:
        try:
            local_format = local_time_format
            timezones = ['Etc/GMT+8']
            for tz in timezones:
                date_time = current_time.astimezone(time_zone.timezone(tz))                          
                self.local_date_time = date_time.strftime(local_format)               
        except Exception as e:
            print("Date Error: %s" % e)
            return None        
        return datetime.strptime(str(self.local_date_time), local_format)     

# NOTE SYNTAX
# INITIALIZATION of PST class, 
# GETTING CURRENT TIME and STANDARD PT
# DECLARE THE FOLLOWING CODES BELOW THE CLASS
# ---------------------------------------------------------------
# >>> pacific_time = PST()                                          // initialization
# >>> current_time = pacific_time.get_current_time()                // getting current time
# >>> standard_pt = pacific_time.time_conversion(current_time)      // getting standard pt
# ---------------------------------------------------------------

# ---------------------------------------------------------------
# START OF THE PROGRAM
# Email Monitoring and SLA Checker Tool using Python 3.8
# Developed By:
# - Cabrera Troy A.
# - Canolas Jevb John
# - Narido Carlo
# - Palisoc Norman
# ---------------------------------------------------------------

#CONFIG FILE
configuration_file_path = r"%s\%s" % (getcwd(), r"config\__config__.json")

if not path.exists(configuration_file_path):
    print("[ x ] Cannot Find __config__ file for configuration structures ...")
    sys.exit("... Script Terminated Successfully ...")

try:
    with open(file=configuration_file_path, mode="r", encoding="utf-8") as json_configuration_file:
        try:        
            config_file = javascript_notation.loads(json_configuration_file.read())    
        except javascript_notation.decoder.JSONDecodeError as json_error:
            print("JSONDecodeError: %s" % json_error)
            sys.exit("... Script Terminated Successfully ...")
except PermissionError as conf_file_permission:
    print("PermissionError: %s" % conf_file_permission)
    sys.exit("... Script Terminated Successfully ...")

auth = Auth(class_pst=PST(), class_email_conf=Email_Config())
content_checker = Contents(email_contents=config_file["Email-Monitoring"]["Email-Contents"])
excel_content = Render_Excel()

if __name__ == "__main__":
    configures = auth.configure(config_file)
    path_to_scan = config_file["Email-Monitoring"]["Path-To-Scan-Excels"]
    if not path.isdir(path_to_scan):
        print("DirectoryNotFoundError: No '%s' directory to scan, Please check __config__ file." % path_to_scan)
        sys.exit("... Script Terminated Successfully ...")
    for each_files in listdir(path=path_to_scan):
        if path.splitext(each_files)[1] in config_file["Email-Monitoring"]["Excel-File-Extension"]:
            chdir(path_to_scan)
            try:
                work_book = load_workbook(each_files)            
            except PermissionError as excel_file_permission:
                print("PermissionError: %s" % excel_file_permission)
                continue
            except py_excel.utils.exceptions.InvalidFileException as excel_ext_error:
                print("InvalidFileException: with (%s) %s" % (each_files, excel_ext_error))
                continue
            for sheets in work_book.sheetnames:
                work_sheet = work_book[sheets]
                excel_columns = content_checker.yield_excel_columns(work_sheet, work_sheet.max_column)
                if content_checker.substantiate_email_contents(excel_columns):                                                    
                    for rows in range(2, sum([work_sheet.max_row, 1])):                                                                                                                       
                        try:
                            valid_columns = config_file["Email-Monitoring"]["Email-Contents"]
                            email_address = excel_content.yield_email_address(work_sheet.cell(rows, excel_content.get_column_index(valid_columns[0], excel_columns)).value)
                            subject = str(work_sheet.cell(rows, excel_content.get_column_index(valid_columns[1], excel_columns)).value)
                            time_received = str(work_sheet.cell(rows, excel_content.get_column_index(valid_columns[2], excel_columns)).value)
                            body = work_sheet.cell(rows, excel_content.get_column_index(valid_columns[3], excel_columns)).value
                            load_end_time = excel_content.yield_load_end_time(body, email_address, config_file)
                        except ValueError as value_error:
                            print("ValueError: %s" % value_error)
                            continue
                        email_contents = dict(email=email_address, subject=subject, time=time_received, load_end_time=load_end_time, body=body)
                        print(email_contents)
                        #print(auth.confirm_sla(configures["standard_pt"], email_contents, configures["emails"], config_file))
                        #break

