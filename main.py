import pandas as pd
from textblob import TextBlob
import newspaper
import nltk

nltk.download('punkt')


# Function to clean and extract article text
def clean_and_extract_article(url):
    try:
        article = newspaper.Article(url)

        # Download and parse the article
        article.download()
        article.parse()

        return article.text
    except Exception as e:
        return None


# Function to perform text analysis
def analyze_text(text):
    blob = TextBlob(text)
    
    # Additional analysis as per your requirements
    positive_score = blob.sentiment.polarity
    negative_score = -blob.sentiment.polarity
    polarity_score = blob.sentiment.polarity
    subjectivity_score = blob.sentiment.subjectivity
    avg_sentence_length = len(blob.words) / len(blob.sentences)
    percentage_of_complex_words = len([word for word in blob.words if len(word) > 2]) / len(blob.words)
    fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words)
    avg_words_per_sentence = len(blob.words) / len(blob.sentences)
    complex_word_count = len([word for word in blob.words if len(word) > 2])
    word_count = len(blob.words)
    syllable_per_word = sum([syllables(word) for word in blob.words]) / len(blob.words)
    personal_pronouns = len([word for word in blob.words if word.lower() in ['i', 'me', 'my', 'mine', 'myself']])
    avg_word_length = sum([len(word) for word in blob.words]) / len(blob.words)
    
    return [positive_score, negative_score, polarity_score, subjectivity_score, avg_sentence_length,
            percentage_of_complex_words, fog_index, avg_words_per_sentence, complex_word_count,
            word_count, syllable_per_word, personal_pronouns, avg_word_length]

# Function to count syllables in a word
def syllables(word):
    word = word.lower()
    count = 0
    vowels = "aeiouy"
    if word[0] in vowels:
        count += 1
    for i in range(1, len(word)):
        if word[i] in vowels and word[i - 1] not in vowels:
            count += 1
    if word.endswith("e"):
        count -= 1
    if count == 0:
        count += 1
    return count

# Read input data
input_df = pd.read_excel("Input.xlsx")

# Create an empty DataFrame to store the analysis results
output_columns = input_df.columns.tolist() + [
    'POSITIVE SCORE', 'NEGATIVE SCORE', 'POLARITY SCORE', 'SUBJECTIVITY SCORE',
    'AVG SENTENCE LENGTH', 'PERCENTAGE OF COMPLEX WORDS', 'FOG INDEX',
    'AVG NUMBER OF WORDS PER SENTENCE', 'COMPLEX WORD COUNT', 'WORD COUNT',
    'SYLLABLE PER WORD', 'PERSONAL PRONOUNS', 'AVG WORD LENGTH'
]
output_df = pd.DataFrame(columns=output_columns)

# Process each row in the input dataframe
for index, row in input_df.iterrows():
    url_id = row['URL_ID']
    url = row['URL']

    # Extract article text
    article_text = clean_and_extract_article(url)

    if article_text:
        # Analyze text
        analysis_result = analyze_text(article_text)

        # Create a new dataframe for the analysis result
        result_df = pd.DataFrame([row.tolist() + analysis_result], columns=output_columns)

        # # Append the result to the output DataFrame
        output_df = pd.concat([output_df, result_df], ignore_index=True)

        # Save the article text to a text file
        # with open(f"{url_id}.txt", 'w', encoding='utf-8') as file:
        #     file.write(article_text)

# Save the result to an output Excel file
output_df.to_excel("Output.xlsx", index=False)
