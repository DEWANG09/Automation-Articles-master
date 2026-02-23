# import pandas as pd
# import webbrowser
# import requests
# from bs4 import BeautifulSoup
# #<------------------------------------------------------------------------------------------------------------------------------------------------------------>
# def read_excel_file(file_path):
#     try:
#         df = pd.read_excel(file_path)
#         return df
#     except Exception as e:
#         print("An error occurred:", e)

# # Function to extract article text from a given URL
# def extract_article_text(url):
#     try:
#         response = requests.get(url)
#         if response.status_code == 200:
#             soup = BeautifulSoup(response.content, 'html.parser')
#             title = soup.title.get_text()
#             article_paragraphs = soup.find_all('p')
#             article_text = ' '.join([p.get_text() for p in article_paragraphs])
#             return title, article_text
#         else:
#             print("Failed to fetch URL:", url)
#             return None, None
#     except Exception as e:
#         print("An error occurred while extracting article text:", e)
#         return None, None
# #<------------------------------------------------------------------------------------------------------------------------------------------------------------>
# # Provide the path to your Excel file
# directory_path = r"C:\Users\aring\OneDrive\Desktop\20211030 Test Assignment-20240218T042812Z-001\20211030 Test Assignment\Python\ExtractedArticleData"
# file_path = r"C:\Users\aring\OneDrive\Desktop\20211030 Test Assignment-20240218T042812Z-001\20211030 Test Assignment\Input.xlsx"

# # Read the Excel file
# data_frame = read_excel_file(file_path)

# # Extract only the URL_ID and URL columns
# extracted_data = data_frame[['URL_ID', 'URL']]
# first_url = extracted_data.iloc[0]['URL']
# # Open the first URL in the default web browser
# webbrowser.open(first_url)

# # Extract article text from the first URL
# title, article_text = extract_article_text(first_url)

# # Save extracted article text to a text file
# filename = directory_path+"blackassign001.txt"
# with open(filename, 'w', encoding='utf-8') as file:
#     file.write(article_text)

# print(f"Article text extracted and saved in file: {filename}")


# print(article_text)






import openpyxl
import pandas as pd
import requests
from bs4 import BeautifulSoup

def read_excel_file(file_path):
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        return df
    except Exception as e:
        print("An error occurred:", e)

def extract_article_text(url):
    try:
        # Fetch webpage content
        response = requests.get(url)
        if response.status_code == 200:
            # Parse HTML content
            soup = BeautifulSoup(response.content, 'html.parser')
            # Extract article title
            title = soup.title.get_text()
            # Extract article text (assuming the main article text is contained within <p> tags)
            article_paragraphs = soup.find_all('p')
            article_text = ' '.join([p.get_text() for p in article_paragraphs])
            return title, article_text
        else:
            print("Failed to fetch URL:", url)
            return None, None
    except Exception as e:
        print("An error occurred while extracting article text:", e)
        return None, None


# Provide the path to your Excel file
file_path = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\Input.xlsx"
directory_path = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\Python\ExtractedData"
# Read the Excel file
data_frame = read_excel_file(file_path)

# Iterate over each row in the DataFrame
for index, row in data_frame.iterrows():
    # Extract URL_ID and URL
    url_id, url = row['URL_ID'], row['URL']
    # Extract article text
    title, article_text = extract_article_text(url)
    if article_text is not None:
        # Save extracted article text to a text file
        filename = directory_path+'\\'+f"{url_id}.txt"
        with open(filename, 'w', encoding='utf-8') as file:
            file.write(f"Title: {title}\n\n")
            file.write(article_text)
        print(f"Article text extracted and saved for URL_ID: {url_id} in file: {filename}")




