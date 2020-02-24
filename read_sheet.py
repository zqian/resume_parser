from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import requests
from pyresparser import ResumeParser
import nltk
import json
import pandas as pd 

from apiclient import discovery
from httplib2 import Http
from oauth2client import client
from oauth2client import file
from oauth2client import tools

#import python-docx
# Set the scopes and discovery info
DOC_SCOPES = 'https://www.googleapis.com/auth/documents.readonly'
DISCOVERY_DOC = ('https://docs.googleapis.com/$discovery/rest?'
                    'version=v1')
# the Google Doc ID with required skills
DOCUMENT_ID = '<id>'


# If modifying these scopes, delete the file token.pickle.
SHEET_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '<id>'
#READ_RANGE_NAME = 'default_title!<I1>:J10'
READ_RANGE_NAME = '<sheet_title>!<range_start>:<range_end>'

def main():

    nltk.download('stopwords')

    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    # Set the scopes and discovery info

    # Initialize credentials and instantiate Docs API service
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret_doc.json', DOC_SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('docs', 'v1', http=creds.authorize(Http()), discoveryServiceUrl=DISCOVERY_DOC)

    # Do a document "get" request and print the results as formatted JSON
    intern_preferred_skill_list = []
    required_skills = service.documents().get(documentId=DOCUMENT_ID).execute()
    print(json.dumps(required_skills, indent=4, sort_keys=True))
    doc_body = required_skills["body"]["content"]
    for doc_body_item in doc_body:
        if "paragraph" in doc_body_item:
            doc_body_item_paragraph = doc_body_item["paragraph"]
            if "elements" in doc_body_item_paragraph:
                doc_body_item_paragraph_elements = doc_body_item_paragraph["elements"]
                for doc_body_item_paragraph_element in doc_body_item_paragraph_elements:
                    if "textRun" in doc_body_item_paragraph_element:
                        doc_body_item_paragraph_element_textRun = doc_body_item_paragraph_element["textRun"]
                        if "content" in doc_body_item_paragraph_element_textRun:
                            # add the skill to list 
                            doc_body_item_paragraph_element_textRun_content = doc_body_item_paragraph_element_textRun["content"].strip().rstrip('\r\n').lower()
                            print(doc_body_item_paragraph_element_textRun_content)
                            intern_preferred_skill_list.append(doc_body_item_paragraph_element_textRun_content)
    
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret_doc.json', SHEET_SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    # Call the Sheets API
    sheet_service = build('sheets', 'v4', credentials=creds)
    sheet = sheet_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                range=READ_RANGE_NAME).execute()
    values = result.get('values', [])

    ## starting from the second row
    ## email address is in column E
    ## skills is in column C

    sheet_row = 0
    if not values:
        print('No data found.')
    else:

        # Create the pandas DataFrame 
        df = pd.DataFrame(columns = ['Skills', 'Mapped Skills', 'Email', 'Resume_Url'])

        row_count=0
        for row in values:
            row_count += 1
            print(row_count)
            print(row)
            if row is None:
                # jump over empty row
                print("empty data row ")
                continue
            elif len(row) < 2:
                # jump over empty row
                print("row has less than two elements: " + str(row))
                continue

            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[1]))

            # download resume file
            resume_file_url = row[1]

            skills = ""
            email = ""

            try:
                myfile = requests.get(resume_file_url, allow_redirects=True)
                my_file_name = resume_file_url.rsplit('/', 1)[-1]
                print(my_file_name)
                resume_file = open(my_file_name, 'wb')
                resume_file.write(myfile.content)

                data = ResumeParser(my_file_name).get_extracted_data()
                output_json = json.loads(json.dumps(data))
                print(output_json)

                email = output_json['email']
                skills = output_json['skills']
                lower_case_skills=map(str.lower, skills)
                mapped_skills = []
                for skill in lower_case_skills:
                    m_skills = [ele for ele in intern_preferred_skill_list if(ele in skill)]
                    if len(m_skills) > 1:
                        print(skill)
                        mapped_skills.append(skill)

            except:
                print("unexpected error ")

            row = [skills, mapped_skills, email, resume_file_url]
            df.loc[len(df)] = row
            # reset
            mapped_skills = []

            sheet_row +=1

            #remove file
            os.remove(my_file_name)
            print("\n\n\n")

        print(df)
        df.to_csv('resume_skills_email.csv')
            

if __name__ == '__main__':
    main()
