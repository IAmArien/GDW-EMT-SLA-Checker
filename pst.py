import pytz as time_zone
from datetime import datetime

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

