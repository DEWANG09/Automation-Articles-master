import os
import re

# Directory paths
stopwords_directory = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\StopWords"
extracted_data_directory = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\Python\ExtractedData"
analysed_data_directory = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\Python\AnalysedData"

# Function to extract stop words from files
def extract_stopwords(directory_path):
    stop_words = set()
    for filename in os.listdir(directory_path):
        file_path = os.path.join(directory_path, filename)
        if os.path.isfile(file_path):
            with open(file_path, 'r') as file:
                stop_words.update(file.read().split())
    return stop_words

# Function to clean text by removing stop words
def clean_text(text, stop_words):
    words = re.findall(r'\b\w+\b', text)
    cleaned_words = [word for word in words if word.lower() not in stop_words]
    cleaned_text = ' '.join(cleaned_words)
    return cleaned_text

# Function to read files from directory, clean them, and save cleaned files
# Function to read files from directory, clean them, and save cleaned files
def clean_and_save_files(stop_words):
    for filename in os.listdir(extracted_data_directory):
        input_file_path = os.path.join(extracted_data_directory, filename)
        output_file_path = os.path.join(analysed_data_directory, filename)
        if os.path.isfile(input_file_path):
            with open(input_file_path, 'r', encoding='utf-8') as input_file:
                text = input_file.read()
                cleaned_text = clean_text(text, stop_words)
                with open(output_file_path, 'w', encoding='utf-8') as output_file:
                    output_file.write(cleaned_text)
                    print(f"File '{filename}' cleaned and saved successfully.")


# Extract stop words
stop_words = extract_stopwords(stopwords_directory)

# Clean text
clean_and_save_files(stop_words)
