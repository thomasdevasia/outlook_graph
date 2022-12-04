import os
import sys
import configparser
import webbrowser
import requests
import base64


# Microsoft Authentication Library
import msal


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

        try:
            verification_uri = flow['verification_uri']
            user_code = flow['user_code']
            print(flow['message'])
            webbrowser.open(verification_uri)

            # response access token and saving it
            access_token = client.acquire_token_by_device_flow(flow)

            with open('access_token.json', 'w') as f:
                f.write(access_token_cache.serialize())
        
        except:
            print('Error in creating access token')
            print(flow)
            sys.exit(1)
        return access_token


    # Send request to the Graph API
    def sendRequest(self, endpoint):

        headers = {
            'Authorization': 'Bearer ' + self.token['access_token']
        }

        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            return response.json()
        else:
            print('Error in sending request')
            # print(response.json())
            sys.exit(1)

    # Search Mail with the search text
    def searchMail(self, searchText, hasAttachments=False):
        # print('Searching for correct mails for {}'.format(searchText))

        endpoint = self.Graph_API_Endpoint + f'/me/messages?$search="{searchText}"'

        response = self.sendRequest(endpoint)
        
        mails = response['value']
        # print(mails)

        if hasAttachments:
            temp = mails
            mails = []
            for mail in temp:
                if mail['hasAttachments'] == True:
                    mails.append(mail)
        
        return mails

    def getAttachments(self, mailId, download=False, downloadPath='./'):
        # print('Getting attachments for mail {}'.format(mailId))

        # check download path exists
        if not os.path.exists(downloadPath):
            os.makedirs(downloadPath)

        endpoint = self.Graph_API_Endpoint + f'/me/messages/{mailId}/attachments'
        response = self.sendRequest(endpoint)

        attachments = response['value']

        if download:
            for attachment in attachments:
                with open(downloadPath + attachment['name'], 'wb') as f:
                    f.write(base64.b64decode(attachment['contentBytes']))

        return attachments