import win32com.client
import json

# read documentation : https://docs.microsoft.com/en-us/office/vba/api/Outlook.NameSpace
class outlook:
    '''
    This will access the outlook application to read the emails. It contains email_reader() function
    which reads email and return the subject line and the timestamp of the latest email. It also uses config.json and
    reads folder hierarchy as parameter. i.e. base = 'xyz@outlook.com', default_folder = 'Inbox', sub_folder = 'WFM_Monitoring'
    '''

    def __init__(self):

        # reading the outlook email path  parameters from config file

        path = r'config.json'
        with open(path, 'r') as p:
            param = json.load(p)

        self.base = param['base']
        self.default_folder = param['default_folder']
        self.sub_folder = param['sub_folder']
        self.critical = param['critical']
        p.close()

    def email_reader(self):
        '''
        It reads the email from the specified path and return the timestamp and
        subject of the last email in a dict object.
        '''

        outlook = win32com.client.Dispatch("Outlook.Application")
        mapi = outlook.GetNamespace("MAPI")

        alert = mapi.Folders[str(self.base)].Folders[str(self.default_folder)].Folders[str(self.sub_folder)]

        messages = alert.Items
        subject = messages.GetLast().Subject
        time = messages.GetLast().CreationTime

        email_data = {'time_stamp': str(time), 'subject': subject, 'critical':self.critical}

        return email_data