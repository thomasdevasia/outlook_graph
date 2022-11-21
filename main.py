import sys
import configparser
import os
import webbrowser
import requests
import base64
import shutil
import re

import PyPDF2
import pandas as pd

# Microsoft Authentication Library
import msal


# Global Variables
Graph_API_Endpoint = 'https://graph.microsoft.com/v1.0'
Download_Cache = './response_downloads/'
Output = './output/'

# creating access token JSON file for the account
def create_access_token(client_id, scopes):
    print("Creating access token")

    access_token_cache = msal.SerializableTokenCache()

    client = msal.PublicClientApplication(client_id, token_cache=access_token_cache)
    flow = client.initiate_device_flow(scopes=scopes)

    verification_uri = flow['verification_uri']
    user_code = flow['user_code']
    print(flow['message'])
    webbrowser.open(verification_uri)

    # response access token and saving it
    access_token = client.acquire_token_by_device_flow(flow)

    with open('access_token.json', 'w') as f:
        f.write(access_token_cache.serialize())


# return token for further api request
def get_token(client_id, scopes):

    access_token_cache = msal.SerializableTokenCache()

    with open('access_token.json', 'r') as f:
        access_token_cache.deserialize(f.read())

    client = msal.PublicClientApplication(client_id, token_cache=access_token_cache)
    accounts = client.get_accounts()[0]

    print(accounts)

    token = client.acquire_token_silent(scopes, account=accounts)

    return token


# Searching for the keyword inside the file
def searchFile(filePath, searchText):
    pdfFile = PyPDF2.PdfFileReader(filePath)
    
    totalPages = pdfFile.getNumPages()
    
    found = False

    for i in range(totalPages):
        page = pdfFile.getPage(i)
        pageContent = page.extractText()
        if  re.search(r'{}'.format(searchText), pageContent):
            found = True
    
    return found

# Searching inside attachment
def searchInsideAttachment(token, mailId, searchText):
    print('Searching inside attachment')

    headers = {
        'Authorization': 'Bearer ' + token['access_token']
    }

    endpoint = Graph_API_Endpoint + f'/me/messages/{mailId}/attachments'
    response = requests.get(endpoint, headers=headers)

    attachments = response.json()['value']

    for attachment in attachments:
        with open(Download_Cache + attachment['name'], 'wb') as f:
            f.write(base64.b64decode(attachment['contentBytes']))
    
    attachmentsList = os.listdir(Download_Cache)
    attachmentsList.remove('.gitkeep')

    # searching for the keyword inside the file
    for attachment in attachmentsList:
        if searchFile(Download_Cache + attachment, searchText):
            print( f'Found file with keyword inside attachment: {attachment}')
            shutil.copy(Download_Cache + attachment, Output + attachment)
    
    # deleting the downloaded cache
    for attachment in attachmentsList:
        os.remove(Download_Cache + attachment)
    

# Searching for the appropriate attachment from the mail
def searchAndFind(token, searchText):
    print('Searching for correct mails for {}'.format(searchText))

    headers = {
        'Authorization': 'Bearer ' + token['access_token']
    }

    endpoint = Graph_API_Endpoint + f'/me/messages?$search="{searchText}"'

    response = requests.get(endpoint, headers=headers)
    mails = response.json()['value']

    for mail in mails:
        if mail['hasAttachments'] == True:
            print('Found mail with attachment')
            searchInsideAttachment(token, mail['id'], searchText)
    
            

if __name__ == '__main__':

    # Get the path to the file
    filePath = sys.argv[1]

    # Read the file
    df = pd.read_excel(filePath)

    config = configparser.ConfigParser()
    config.read(['config.dev.cfg'])
    azure_settings = config['azure']

    # client and tenant, id and scopes
    client_id = azure_settings['clientId']
    tenant_id = azure_settings['tenantId']
    scopes = azure_settings['graphUserScopes'].split(' ')

    data = df.to_dict('records')

    if os.path.exists('access_token.json'):
        print('Found access_token.json')
        token = get_token(client_id, scopes)
        for company in data:
            searchAndFind(token, company['searchItem'])
    else:
        print('No access_token.json found')
        create_access_token(client_id, scopes)