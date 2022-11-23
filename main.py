import sys
import configparser
import os
import webbrowser
import requests
import base64
import shutil
import re

from tqdm import tqdm
import PyPDF2
import pandas as pd

# Microsoft Authentication Library
import msal

# Global Variables
Download_Cache = './response_downloads/'
Output = './output/'

# Microsoft Graph API Class
class microsoftGraph:

    def __init__(self, cfgFile):
        config = configparser.ConfigParser()
        config.read([cfgFile])
        azure_settings = config['azure']

        # client and tenant, id and scopes
        self.client_id = azure_settings['clientId']
        self.tenant_id = azure_settings['tenantId']
        self.scopes = azure_settings['graphUserScopes'].split(' ')
        
        self.Graph_API_Endpoint = 'https://graph.microsoft.com/v1.0'
        
        self.token = self.get_token()    

        self.headers = {
            'Authorization': 'Bearer ' + self.token['access_token']
        }

    # return token for further api request
    def get_token(self):

        if os.path.exists('access_token.json'):
            print('Found access_token.json')

            access_token_cache = msal.SerializableTokenCache()

            with open('access_token.json', 'r') as f:
                access_token_cache.deserialize(f.read())

            client = msal.PublicClientApplication(self.client_id, token_cache=access_token_cache)
            
            
            accounts = client.get_accounts()[0]
            print(accounts)

            token = client.acquire_token_silent(self.scopes, account=accounts)

        else:
            print('No access_token.json found')
            token = self.create_access_token()        
        
        return token
    
    # creating access token JSON file for the account
    def create_access_token(self):
        print("Creating access token")

        access_token_cache = msal.SerializableTokenCache()

        client = msal.PublicClientApplication(self.client_id, token_cache=access_token_cache)
        flow = client.initiate_device_flow(scopes=self.scopes)

        verification_uri = flow['verification_uri']
        user_code = flow['user_code']
        print(flow['message'])
        webbrowser.open(verification_uri)

        # response access token and saving it
        access_token = client.acquire_token_by_device_flow(flow)

        with open('access_token.json', 'w') as f:
            f.write(access_token_cache.serialize())
        
        return access_token


    # Search Mail with the search text
    def searchMail(self, searchText, hasAttachments=False):
        # print('Searching for correct mails for {}'.format(searchText))

        endpoint = self.Graph_API_Endpoint + f'/me/messages?$search="{searchText}"'

        response = requests.get(endpoint, headers=self.headers)
        
        mails = response.json()['value']

        if hasAttachments:
            temp = mails
            mails = []
            for mail in temp:
                if mail['hasAttachments'] == True:
                    mails.append(mail)
        
        return mails

    def getAttachments(self, mailId, download=False, downloadPath='./'):
        # print('Getting attachments for mail {}'.format(mailId))

        endpoint = self.Graph_API_Endpoint + f'/me/messages/{mailId}/attachments'
        response = requests.get(endpoint, headers=self.headers)

        attachments = response.json()['value']

        if download:
            for attachment in attachments:
                with open(downloadPath + attachment['name'], 'wb') as f:
                    f.write(base64.b64decode(attachment['contentBytes']))

        return attachments

    
    # Searching for the appropriate attachment from the mail
    def searchAndFind(self, df):
        
        data = df.to_dict('records')

        for company in tqdm(data, desc='Searching for mails', unit='mails'):
            searchText = format(int(company['searchItem']), ',d')
            mails = self.searchMail(searchText, True)
            # print('Found {} mails for {}'.format(len(mails), searchText))
            # print(len(mails), searchText)

            for mail in mails:
                attachments = self.getAttachments(mail['id'], download=True, downloadPath=Download_Cache)
                
                attachmentsList = os.listdir(Download_Cache)
                attachmentsList.remove('.gitkeep')

                # searching for the keyword inside the file
                for attachment in attachmentsList:
                    if searchFile(Download_Cache + attachment, searchText):
                        # print( f'Found file with keyword({searchText}) inside attachment: {attachment}')
                        shutil.copy(Download_Cache + attachment, Output + attachment)
                
                # deleting the downloaded cache
                for attachment in attachmentsList:
                    os.remove(Download_Cache + attachment)



# Searching for the keyword inside the file
def searchFile(filePath, searchText):
    pdfFile = PyPDF2.PdfFileReader(filePath)
    
    totalPages = pdfFile.getNumPages()
    
    found = False

    for i in range(totalPages):
        page = pdfFile.getPage(i)
        # pageContent = page.extractText().replace('\n', '').replace(',','')
        pageContent = page.extractText().replace('\n', '')
        # print(searchText, pageContent)
        if  re.search(r'{}'.format(searchText), pageContent):
            found = True
    
    return found

    
    
            

if __name__ == '__main__':

    # Get the path to the file
    filePath = sys.argv[1]

    # Read the file
    df = pd.read_excel(filePath)
    # print(df.head())

    mail = microsoftGraph('config.dev.cfg')

    mail.searchAndFind(df)
