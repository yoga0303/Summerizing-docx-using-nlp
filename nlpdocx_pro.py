import streamlit as st
from docx import Document
import re
import nltk
import heapq
from collections import defaultdict

# Download necessary NLTK data
nltk.download('punkt')
nltk.download('stopwords')

# Function to extract text from DOCX file
def extract_text_from_docx(docx_file):
    doc = Document(docx_file)
    full_text = []
    for paragraph in doc.paragraphs:
        full_text.append(paragraph.text)
    return '\n'.join(full_text)

# Function to preprocess text
def preprocess_text(text):
    text = re.sub(r'\s+', ' ', text)
    sentences = nltk.sent_tokenize(text)
    return sentences

# Function to compute word frequencies
def compute_word_frequencies(text):
    stop_words = set(nltk.corpus.stopwords.words('english'))
    words = nltk.word_tokenize(text)
    
    word_frequencies = defaultdict(int)
    for word in words:
        if word not in stop_words:
            word_frequencies[word.lower()] += 1
    
    max_frequency = max(word_frequencies.values())
    for word in word_frequencies:
        word_frequencies[word] /= max_frequency
        
    return word_frequencies

# Function to score sentences using word frequencies
def score_sentences(sentences, word_frequencies):
    sentence_scores = defaultdict(int)
    for sentence in sentences:
        words = nltk.word_tokenize(sentence.lower())
        for word in words:
            if word in word_frequencies:
                sentence_scores[sentence] += word_frequencies[word]
    
    return sentence_scores

# Function to summarize text
def summarize_text(text, num_sentences=3):
    sentences = preprocess_text(text)
    word_frequencies = compute_word_frequencies(text)
    sentence_scores = score_sentences(sentences, word_frequencies)
    
    summary_sentences = heapq.nlargest(num_sentences, sentence_scores, key=sentence_scores.get)
    summary = ' '.join(summary_sentences)
    
    return summary

# Streamlit app interface
st.title("Extract and Summarize Text from DOCX File")

# Upload a DOCX file
uploaded_file = st.file_uploader("Choose a DOCX file", type="docx")

if uploaded_file is not None:
    # Extract text from the uploaded DOCX file
    text = extract_text_from_docx(uploaded_file)
    
    # Display the extracted text
    st.subheader("Extracted Text")
    st.write(text)
    
    # Input for the number of sentences in the summary
    num_sentences = st.number_input("Number of sentences in the summary", min_value=1, value=3)
    
    # Summarize button
    if st.button("Summarize"):
        # Summarize the extracted text
        summary = summarize_text(text, num_sentences)
        
        # Display the summary
        st.subheader("Summary")
        st.write(summary)
else:
    st.info("Please upload a DOCX file to extract and summarize text.")