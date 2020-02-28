import win32com.client as win32
from datetime import datetime
from csv import writer

class Email(object):
    def __init__(self, recv_email: str, report_csv: str, late_jobs: str, missing_jobs: str, var_conf: dict):
        super().__init__()
        self.recv_email = recv_email
        self.report_csv = report_csv
        self.late_jobs = late_jobs
        self.missing_jobs = missing_jobs
        self.var_conf = var_conf

    def yield_missing_jobs(self, max_row: int, sheet: object) -> dict:
        try:            
            for row_counter in range(2, max_row):                
                sla_str = str(sheet.cell(row=row_counter,column=4).value)
                let_str = str(sheet.cell(row=row_counter,column=5).value)
                if '.' in let_str:
                    let_str = let_str.split('.',1)[0]                    
                if let_str == "None":
                    self.var_conf["total_missing"] = self.var_conf["total_missing"] + 1
                    self.missing_jobs = self.missing_jobs + """
                        <tr>
                            <td>""" + str(self.var_conf["total_missing"]) + """</td>
                            <td>""" + str(sheet.cell(row=row_counter,column=2).value) + """</td>
                            <td>""" + str(sheet.cell(row=row_counter,column=3).value) + """</td>
                            <td>""" + str(sheet.cell(row=row_counter,column=4).value) + """</td>
                            <td>No execution recorded to date.</td>
                        </tr> 
                        """
                else:
                    let_column = (datetime.strptime(str(let_str),'%Y-%m-%d %H:%M:%S')).date()
                    sla_column = (datetime.strptime(str(sla_str),'%H:%M:%S')).time()
                    now = datetime.now().time()
                    if ((let_column == datetime.now().date() and sla_column < now) or (let_column < datetime.now().date())):
                        self.var_conf["total_missing"] = self.var_conf["total_missing"] + 1
                        self.missing_jobs = self.missing_jobs + """
                            <tr>
                                <td>""" + str(self.var_conf["total_missing"]) + """</td>
                                <td>""" + str(sheet.cell(row=row_counter,column=2).value) + """</td>
                                <td>""" + str(sheet.cell(row=row_counter,column=3).value) + """</td>
                                <td>""" + str(sheet.cell(row=row_counter,column=4).value) + """</td>
                                <td>""" + str(sheet.cell(row=row_counter,column=5).value) + """</td>
                            </tr> 
                            """

            return dict(missing_jobs=self.missing_jobs, total_missing=self.var_conf["total_missing"])

        except Exception as mis_job_err:
            print(mis_job_err)
            return False
    
    def yield_late_jobs(self, max_row: int, sheet: object, rows: int) -> dict:
        try:
            while(rows <= max_row):
                if(sheet.cell(row=rows, column=5).value == "1"):
                    excel_data = [
                                    sheet.cell(row=rows, column=1).value,
                                    sheet.cell(row=rows, column=2).value,
                                    str(sheet.cell(row=rows, column=3).value).replace('T',' ').replace('+00:00',''),
                                    sheet.cell(row=rows, column=4).value,
                                    sheet.cell(row=rows, column=5).value,
                                    sheet.cell(row=rows, column=6).value
                                ]
                    if(sheet.cell(row=rows, column=6).value == "False"):
                        self.var_conf["total_fail"] = self.var_conf["total_fail"] + 1
                        print(str(self.var_conf["total_fail"])+" jobs have been deleted from the monitoring list.")
                        self.late_jobs = self.late_jobs + """
                        <tr>
                            <td>""" + str(self.var_conf["total_fail"]) + """</td>
                            <td>""" + excel_data[0] + """</td>
                            <td>""" + excel_data[2] + """</td>
                            <td>""" + excel_data[3] + """</td>
                        </tr> 
                        """
                    else:
                        self.var_conf["total_success"] = self.var_conf["total_success"] + 1
                    with open(r"Reports\%s" % (self.report_csv), 'a+', newline='') as write_obj:
                        csv_writer = writer(write_obj)
                        csv_writer.writerow(excel_data)                        

                    sheet.delete_rows(rows)                    
                else:
                    rows = rows + 1            

            return dict(late_jobs=self.late_jobs, total_fail=self.var_conf["total_fail"])

        except Exception as late_job_err:
            print(late_job_err)
            return False

    def yield_jobs_headers(self, total_fail: int, total_missing: int) -> dict:
        try:
            if total_fail < 1:
                self.late_jobs = "<tr><td colspan='4'>No late jobs scanned.</td></tr>"
            if total_missing < 1:
                self.missing_jobs = "<tr><td colspan='5'>No missing jobs scanned.</td></tr>"
            return dict(late_jobs=self.late_jobs, missing_jobs=self.missing_jobs)
        except Exception as job_header_err:
            print(job_header_err)
            return False

    def construct_email_body(self, total_job_counter: int) -> str:
        try:
            email_body = ('''
                <html>
                    <style>
                    table, th, td{
                        border: 1px solid black;
                        margin: auto;
                        text-align: center;
                        padding: 5px;
                    }
                    table{
                        width: 65%;
                    }
                    h2{ 
                        text-align: center;
                    }
                    </style>
                    <body>
                        <center>
                        <table style='border-collapse: collapse; width: 50%;'>
                            <tr>
                                <td colspan="2" style="background:#5bb84f;"><marquee behavior="alternate"><h2>Email Monitoring Report</h2></marquee></td>
                            </tr>
                            <tr>
                                <td>Date and Time Executed</td>
                                <td>'''+ str((datetime.now()).strftime("%b %d, %Y %H:%M:%S")) +''', Manila Time</td>
                            </tr>
                            <tr>
                                <td>Total Jobs Scanned</td>
                                <td>'''+ str(total_job_counter) +'''</td>
                            </tr>
                            <tr>
                                <td>Jobs Executed on Time</td>
                                <td>'''+ str(self.var_conf["total_success"]) +'''</td>
                            </tr>
                            <tr>
                                <td>Jobs Executed Late</td>
                                <td>'''+ str(self.var_conf["total_fail"]) +'''</td>
                            </tr>
                            <tr>
                                <td>Total Missing Jobs</td>
                                <td>'''+ str(self.var_conf["total_missing"]) +'''</td>
                            </tr>
                        </table>
                        <br/>
                        <table style='border-collapse: collapse;'>
                            <tr>
                                <td colspan="5" style="background:#d92b1e;"><h2>List of Missing Jobs</h2></td>
                            </tr>
                            <tr>
                                <th>#</td>
                                <th>Job Name</th>
                                <th>Frequency</th>
                                <th>SLA</th>
                                <th>Last Run</th>
                            </tr>
                            ''' + self.missing_jobs + '''
                        </table>
                        <br/>
                        <table style='border-collapse: collapse;'>
                            <tr>
                                <td colspan="4" style="background:#f2d75e;"><h2>List of Late Jobs</h2></td>
                            </tr>
                            <tr>
                                <th>#</td>
                                <th>Sender</th>
                                <th>Time Sent</th>
                                <th>Email Body</th>
                            </tr>
                            ''' + self.late_jobs + '''
                        </table>
                        </center>
                    </body>
                </html>''')

            return email_body

        except Exception as em_body_err:
            print(em_body_err)
            return False
    
    def send_email(self, email_body: str, subject: str):
        try:
            outlook = win32.Dispatch('Outlook.Application')
            mail = outlook.CreateItem(0)
            mail.To = self.recv_email
            mail.Subject = subject
            mail.HTMLBody = email_body     
            mail.Send()
        except Exception as em_err:
            print(em_err)
            return False
