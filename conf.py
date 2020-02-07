import sys as system
import re as regex
from datetime import datetime

class Conf(object):
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
                    system.exit(0)
            else:                
                system.exit(0)
        else:            
            system.exit(0)
        return dict(emails=valid_emails, standard_pt=standard_pt)

    @staticmethod
    def collect_valid_email_body(current_datetime: datetime, email_contents: dict, valid_emails_list: list, configuration_file: dict) -> tuple:        
        if email_contents["load_end_time"] == None:
            return ('Error', 'No Load End Time')
        if datetime.now() >= email_contents["load_end_time"]:            
            if email_contents["email"] in valid_emails_list:
                models = configuration_file["Email-Monitoring"]["Accounts"][email_contents["email"]]["Models"]
                __models__keys = {}
                for each_controls in models:
                    for start_end in models[each_controls]:
                        if start_end == "Dependencies":
                            if models[each_controls][start_end] != None:
                                __models__keys.update({"Dependencies": models[each_controls][start_end]})
                        for each_rules in models[each_controls][start_end]:
                            rule_key = regex.search("%s" % each_rules, email_contents["body"])                                                       
                            if rule_key != None:
                                __models__keys.update({"email-address": email_contents["email"]})
                                __models__keys.update({"subject": email_contents["subject"]})
                                __models__keys.update({"load_end_time": email_contents["load_end_time"]})
                                __models__keys.update({each_controls: {start_end: regex.sub(r"(_x000D_|\n)", "", email_contents["body"])}})
                                return ('Complete', __models__keys)
            else:
                return ('Error', 'Invalid Email')

    def validate_job_loads(self, container: list) -> dict:
        if container[0]["email-address"] == "GDW@ingrammicro.com":
            return self.load_gdw_configurations(container[0], container[1], "Dependencies", "Key-Sources")

    @staticmethod
    def load_gdw_configurations(dict_logs: dict, configur_file: dict, depd_key: str, key_src: str) -> dict:
        if depd_key in dict_logs:
            if key_src in dict_logs["Dependencies"]:
                dependency_label = dict_logs["Dependencies"]["Key-Sources"]
                em_add = dict_logs["email-address"]
                let = dict_logs["load_end_time"]
                subj = dict_logs["subject"]
                del dict_logs["Dependencies"]
                del dict_logs["email-address"]
                del dict_logs["subject"]
                del dict_logs["load_end_time"]
                for each_models in dict_logs.keys():
                    for each_values in dict_logs[each_models]:
                        if each_values == depd_key:
                            continue
                        body_start_end = dict_logs[each_models][each_values]
                        start_end = each_values
                        if dependency_label in body_start_end:
                            runs = configur_file["Email-Monitoring"]["Accounts"][em_add]["GDW-Runs"].keys()                            
                            ext_runs = ""
                            for each_runs in runs:
                                if each_runs in body_start_end:
                                    ext_runs = each_runs
                                    break
                            try:                                
                                ks_dict = configur_file["Email-Monitoring"]["Accounts"][em_add]["Key-Sources"]
                                source = ks_dict[dependency_label]["Jobs"][start_end]
                                for each_source in source:
                                    for each_run_nmbrs in source[each_source]:
                                        if source[each_source][each_run_nmbrs] == ext_runs:
                                            job_attrib = source[each_source][each_run_nmbrs]
                                            job_name = each_source
                                gdw_runs = configur_file["Email-Monitoring"]["Accounts"][em_add]["GDW-Runs"][job_attrib][dependency_label][job_name]
                                __email__struct = {}
                                for each_time in gdw_runs:
                                    regex_str = regex.search(r"\d{4}-\d{2}-\d{2}", str(let))                                
                                    combine_str = "{date_re} {sla}".format(date_re=regex_str.group(), sla=gdw_runs[each_time])
                                    convert_datetime = datetime.strptime(combine_str, '%Y-%m-%d %H:%M:%S')
                                    __email__struct.update({"Job": job_name})
                                    __email__struct.update({"Email-Notif": body_start_end})
                                    __email__struct.update({"Subject": subj})
                                    __email__struct.update({"Elements": each_time})
                                    __email__struct.update({"Ave-Start-End": (let <= convert_datetime)})
                                    __email__struct.update({"Ave-Start-End (Time)": gdw_runs[each_time]})
                                    if not let <= convert_datetime:
                                        if not each_time == "Start":
                                            __email__struct.update({"SLA": False})
                                            __email__struct.update({"SLA (Time)": gdw_runs[each_time]})                                        
                                        return __email__struct
                            except KeyError as e:
                                print(e)


