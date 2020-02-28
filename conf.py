import sys as system
import re as regex
import uuid as unique_id
from datetime import datetime
from treelib import Node, Tree

class Conf(object):
    def __init__(self, class_pst, class_email_conf):
        super().__init__()
        self.class_pst = class_pst
        self.class_email_conf = class_email_conf
        self.tree = Tree()        
        
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
        if current_datetime >= email_contents["load_end_time"]:            
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
                                __models__keys.update({"key-search": each_rules})
                                __models__keys.update({each_controls: {start_end: regex.sub(r"(_x000D_|\n)", "", email_contents["body"])}})
                                return ('Complete', __models__keys)
            else:
                return ('Error', 'Invalid Email')

    def validate_job_loads(self, container: list) -> tuple:
        if container[0]["email-address"] == "GDW@ingrammicro.com":
            return self.load_gdw_configurations(
                dict_logs=container[0], 
                configur_file=container[1], 
                depd_key="Dependencies", 
                key_src="Key-Sources",
                rowvalue=container[2],
                work_book=container[3]
            )

    @staticmethod
    def load_gdw_configurations(dict_logs: dict, configur_file: dict, depd_key: str, key_src: str, rowvalue: int, work_book: object) -> tuple:        
        sheet = work_book['Jobs']
        checksheet = work_book['Checklist']
        if depd_key in dict_logs:
            if key_src in dict_logs[depd_key]:
                dependency_label = dict_logs[depd_key][key_src]
                em_add = dict_logs["email-address"]
                let = dict_logs["load_end_time"]
                key_search = dict_logs["key-search"]                
                #subj = dict_logs["subject"]
                process = ""
                hierarchy_structures = []
                for each in dict_logs.keys():
                    key_model = each                
                for each_keys in dict(dict_logs[key_model]).keys():
                    process = each_keys
                hierarchy_structures.append(dependency_label)
                hierarchy_structures.append(process)                
                del dict_logs[depd_key]
                del dict_logs["email-address"]
                del dict_logs["subject"]
                del dict_logs["load_end_time"]
                del dict_logs["key-search"]
                for each_models in dict_logs.keys():
                    for each_values in dict_logs[each_models]:
                        if each_values == depd_key:
                            continue
                        body_start_end = dict_logs[each_models][each_values]
                        start_end = each_values
                        hierarchy_structures.append(body_start_end)
                        if dependency_label in body_start_end:
                            runs = configur_file["Email-Monitoring"]["Accounts"][em_add]["Job-Runs"].keys()                            
                            ext_runs = ""
                            for each_runs in runs:
                                if each_runs in body_start_end:
                                    ext_runs = each_runs
                                    break
                            try:
                                hierarchy_structures.append(ext_runs)
                                ks_dict = configur_file["Email-Monitoring"]["Accounts"][em_add][key_src]
                                source = ks_dict[dependency_label]["Jobs"][start_end]
                                for each_source in source:
                                    for each_run_nmbrs in source[each_source]:
                                        if source[each_source][each_run_nmbrs] == ext_runs:
                                            job_attrib = source[each_source][each_run_nmbrs]
                                            job_name = each_source
                                gdw_runs = configur_file["Email-Monitoring"]["Accounts"][em_add]["Job-Runs"][job_attrib][dependency_label][job_name]
                                __email__struct = {}
                                for each_time in gdw_runs:
                                    regex_str = regex.search(r"\d{4}-\d{2}-\d{2}", str(let))                                
                                    combine_str = "{date_re} {sla}".format(date_re=regex_str.group(), sla=gdw_runs[each_time])
                                    convert_datetime = datetime.strptime(combine_str, '%Y-%m-%d %H:%M:%S')
                                    __email__struct.update({"Job": job_name})
                                    __email__struct.update({"Key Search": key_search})
                                    __email__struct.update({"Datetime Received (PST)": let})
                                    __email__struct.update({"Process": each_time})
                                    __email__struct.update({"Average Start/End (Time)": gdw_runs[each_time]})
                                    __email__struct.update({"Average Start/End (Bool)": (let <= convert_datetime)})

                                    intime = True                              
                                    if let <= convert_datetime: intime = True
                                    else: intime = False
                                    if(sheet.cell(row=rowvalue, column=5).value != "1"):
                                        sheet["E"+str(rowvalue)] = "1"
                                    if(sheet.cell(row=rowvalue, column=6).value is None):
                                        sheet["F"+str(rowvalue)] = str(intime)

                                    for counter in range(2,checksheet.max_row):
                                        checklist_job_name = checksheet.cell(row=counter, column=2).value
                                        if job_name == checklist_job_name:
                                            job_name_last_run = checksheet.cell(row=counter, column=5).value
                                            if job_name_last_run is not None:
                                                temp_date = datetime.strptime(str(checksheet.cell(row=counter, column=5).value).split(".")[0], '%Y-%m-%d %H:%M:%S')
                                                if let >= temp_date:
                                                    checksheet["E"+str(counter)] = let  
                                            else:
                                                checksheet["E"+str(counter)] = let
                                            checksheet["F"+str(counter)] = datetime.now()
                                            break  

                                    if not each_time == "SLA":
                                        if let <= convert_datetime:
                                            if each_time == "Start":
                                                __email__struct.update({"SLA (Time)": "-----"})  
                                                __email__struct.update({"SLA": "-----"})
                                            else:
                                                __email__struct.update({"SLA (Time)": gdw_runs[each_time]})                                                
                                                __email__struct.update({"SLA": True})
                                        else:
                                            if each_time == "Start":
                                                __email__struct.update({"SLA (Time)": "-----"})  
                                                __email__struct.update({"SLA": "-----"})
                                            else:
                                                __email__struct.update({"SLA (Time)": gdw_runs[each_time]})                                  
                                                __email__struct.update({"SLA": False})                                       
                                        return (hierarchy_structures, __email__struct)                                    

                            except KeyError as e:
                                print(e)
    
    def hierarchy_structures(self, hierarchy_list: list, main_job: str) -> Tree:
        hierarchy_entity = {}
        html_hierarchy = ""
        try:
            if not self.tree.contains(main_job):
                self.tree.create_node(main_job, main_job)
                hierarchy_entity.update({main_job: {}})
            try:
                for each_h_list in hierarchy_list:
                    if not self.tree.contains(each_h_list[0]):
                        self.tree.create_node(each_h_list[0], each_h_list[0], parent=main_job)
                        hierarchy_entity[main_job].update({each_h_list[0]: {}})
                    if not self.tree.contains("%s-%s" % (each_h_list[0], each_h_list[3])):
                        self.tree.create_node(each_h_list[3], "%s-%s" % (each_h_list[0], each_h_list[3]), parent=each_h_list[0])                        
                        hierarchy_entity[main_job][each_h_list[0]].update({each_h_list[3]: {}})
                    if not self.tree.contains("%s-%s-%s" % (each_h_list[0], each_h_list[3], each_h_list[1])):
                        self.tree.create_node(each_h_list[1], "%s-%s-%s" % (each_h_list[0], each_h_list[3], each_h_list[1]), parent="%s-%s" % (each_h_list[0], each_h_list[3]))                        
                        hierarchy_entity[main_job][each_h_list[0]][each_h_list[3]].update({each_h_list[1]: []})
                    self.tree.create_node(each_h_list[2], unique_id.uuid4(), parent="%s-%s-%s" % (each_h_list[0], each_h_list[3], each_h_list[1]))
                    hierarchy_entity[main_job][each_h_list[0]][each_h_list[3]][each_h_list[1]].append(each_h_list[2])
            except Exception as e:
                print(e)
        except Exception as e_1:
            print(e_1)  

        for a in hierarchy_entity:
            html_hierarchy = "%s%s" % (html_hierarchy, "<li>%s<ul>" % a)
            for b in hierarchy_entity[a]:
                html_hierarchy = "%s%s" % (html_hierarchy, "<li>%s<ul>" % b)
                for c in hierarchy_entity[a][b]:
                    html_hierarchy = "%s%s" % (html_hierarchy, "<li>%s<ul>" % c)
                    for d in hierarchy_entity[a][b][c]:
                        html_hierarchy = "%s%s" % (html_hierarchy, "<li>%s<ul>" % d)
                        for e in hierarchy_entity[a][b][c][d]:
                            html_hierarchy = "%s%s" % (html_hierarchy, "<li>%s</li>" % e)
                        html_hierarchy = "%s%s" % (html_hierarchy, "</ul>")
                        html_hierarchy = "%s%s" % (html_hierarchy, "</li>")
                    html_hierarchy = "%s%s" % (html_hierarchy, "</ul>")
                    html_hierarchy = "%s%s" % (html_hierarchy, "</li>")
                html_hierarchy = "%s%s" % (html_hierarchy, "</ul>")
                html_hierarchy = "%s%s" % (html_hierarchy, "</li>")
            html_hierarchy = "%s%s" % (html_hierarchy, "</ul>")
            html_hierarchy = "%s%s" % (html_hierarchy, "</li>")
        
        with open(r"..\struct\HTML\index.html", mode="r", encoding="utf-8") as html_file:
            html_prt = str(html_file.read()).split("\n")
            html_file.close()
        
        html_prt.insert(81, html_hierarchy)
        struct_html = "\n".join(html_prt)

        with open(r"..\struct\HTML\index.html", mode="w", encoding="utf-8") as html_file_1:
            html_file_1.write(struct_html)
            html_file_1.close()

        return self.tree
