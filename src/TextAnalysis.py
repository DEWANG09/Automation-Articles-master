import os
import re
import nltk
from nltk.corpus import stopwords
import openpyxl

negative_words_file = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\MasterDictionary\negative-words.txt"
positive_words_file = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\MasterDictionary\positive-words.txt"

def read_words_from_file(file_path):
    words_list = []
    encodings = ['utf-8', 'utf-8-sig', 'latin-1']
    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as file:
                for line in file:
                    words_list.append(line.strip())
            return words_list
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError("Unable to decode the file with any of the specified encodings.")
negative_words = read_words_from_file(negative_words_file)
positive_words = read_words_from_file(positive_words_file)
#<-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------->
nltk.download('stopwords')
def syllable_count(word):
    vowels = 'aeiouy'
    count = 0
    for char in word:
        if char.lower() in vowels:
            count += 1
    if word.endswith(('es', 'ed')):
        count -= 1
    if word.endswith('le'):
        count += 1
    if count == 0:
        count = 1
    return count

#<-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------->
analysed_data_directory = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\Python\ExtractedData"
def calculate_scores_and_metrics(text, positive_words, negative_words):
    positive_score = sum(1 for word in text.split() if word in positive_words)
    negative_score = sum(1 for word in text.split() if word in negative_words)
    polarity_score = (positive_score - negative_score) / (positive_score + negative_score + 0.000001)
    subjectivity_score = (positive_score + negative_score) / (len(text.split()) + 0.000001)
    sentences = re.split(r'[.!?]', text)
    num_sentences = len(sentences)
    total_words = len(text.split())
    average_sentence_length = total_words / num_sentences
    complex_words = [word for word in text.split() if len(word) > 2 and syllable_count(word) > 2]
    num_complex_words = len(complex_words)
    percentage_complex_words = num_complex_words / total_words
    fog_index = 0.4 * (average_sentence_length + percentage_complex_words)
    average_words_per_sentence = total_words / num_sentences
    word_count = total_words
    syllable_count_per_word = sum(syllable_count(word) for word in text.split())
    personal_pronouns = re.findall(r'\b(I|we|my|ours|us)\b', text, flags=re.IGNORECASE)
    personal_pronouns_count = len(personal_pronouns)
    total_characters = sum(len(word) for word in text.split())
    average_word_length = total_characters / total_words
    return (positive_score, negative_score, polarity_score, subjectivity_score,
            average_sentence_length, percentage_complex_words, fog_index,
            average_words_per_sentence,num_complex_words,word_count, syllable_count_per_word,
            personal_pronouns_count, average_word_length)
start = 2
count = 1
excel_file_path = r"D:\New folder\Automation-Articles-master\Automation-Articles-master\Output Data Structure.xlsx"
workbook = openpyxl.load_workbook(excel_file_path)
worksheet = workbook.active
for filename in os.listdir(analysed_data_directory):
    file_path = os.path.join(analysed_data_directory, filename)
    if os.path.isfile(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
            (positive_score, negative_score, polarity_score, subjectivity_score,
            average_sentence_length, percentage_complex_words, fog_index,
            average_words_per_sentence,num_complex_word,word_count, syllable_count_per_word,
            personal_pronouns_count, average_word_length) = calculate_scores_and_metrics(text, positive_words, negative_words)
            data_list = [1,1,1,positive_score, negative_score, polarity_score, subjectivity_score,
            average_sentence_length, percentage_complex_words, fog_index,
            average_words_per_sentence,num_complex_word,word_count, syllable_count_per_word,
            personal_pronouns_count, average_word_length,1,1]
            for row in range(start,start+1):  # Rows 2 to 101
                for col_idx, column in enumerate(range(3, 16), start=3):  # Columns C1 to O1
                    worksheet.cell(row=row, column=column, value=data_list[column])
            print(data_list)
            print(count)
    start += 1
    count += 1
#<-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------->
workbook.save(excel_file_path)