import os
import sys
import re
import shutil

from tqdm import tqdm

import PyPDF2
import pandas as pd


# Microsoft API Class
from microsoftgraph import microsoftGraph 

# Global Variables
Download_Cache = './response_downloads/'
Output = './output/'


# Searching for the keyword inside the PDF file
def searchInsidePdf(filePath, searchText):
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

# Searching for files with mail
def main(df, graphApi):

    data = df.to_dict('records')

    for company in data:
        searchText = format(int(company['searchItem']), ',.2f')

        mails = graphApi.searchMail(searchText, True)
        
        print('Found {} mails for {}'.format(len(mails), searchText))

        for mail in tqdm(mails, desc='Searching the mails', unit='mails'):

            attachments = graphApi.getAttachments(mail['id'], download=True, downloadPath=Download_Cache)
            
            attachmentsList = os.listdir(Download_Cache)

            # searching for the keyword inside the file
            for attachment in attachmentsList:
                if searchInsidePdf(Download_Cache + attachment, searchText):
                    # print( f'Found file with keyword({searchText}) inside attachment: {attachment}')
                    shutil.copy(Download_Cache + attachment, Output + attachment)
            
            # deleting the downloaded cache
            for attachment in attachmentsList:
                os.remove(Download_Cache + attachment)
    
            

if __name__ == '__main__':

    if len(sys.argv) != 2:
        print('Usage: python3 main.py \{fileName\}')
        sys.exit(1)
    else:
        # path to the excel file
        filePath = sys.argv[1]
        if not os.path.exists(filePath):
            print('File not found')
            sys.exit(1)

    # Read the file
    df = pd.read_excel(filePath)
    # print(df.head())

    graphApi = microsoftGraph('config.dev.cfg')

    if not os.path.exists(Output):
        os.makedirs(Output)
        
    main(df, graphApi)
