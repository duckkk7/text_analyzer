import spacy
import os
import pandas as pd
from docx import Document

nlp = spacy.load("ru_core_news_md")


def read_text_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read()


def read_doc_file(file_path):
    doc = Document(file_path)
    text = "\n".join([para.text for para in doc.paragraphs])
    return text


def extract_entities(text):
    doc = nlp(text)
    entities = []
    for ent in doc.ents:
        entities.append((ent.text, ent.label_))
    return entities


def process_file(file_path):
    if file_path.endswith('.txt'):
        text = read_text_file(file_path)
    elif file_path.endswith('.doc') or file_path.endswith('.docx'):
        text = read_doc_file(file_path)
    else:
        raise ValueError("Unsupported file type")

    entities = extract_entities(text)
    return entities


def save_entities_to_csv(entities, output_file):
    df = pd.DataFrame(entities, columns=["Entity", "Label"])
    df.to_csv(output_file, index=False, encoding='utf-8')


def analyze_text(file_path, output_file):
    entities = process_file(file_path)
    save_entities_to_csv(entities, output_file)
    print(f"Сущности были сохранены в файл: {output_file}")


file_path = "C:/Users/Vlada/Desktop/о стиве джобсе.docx"
output_file = "entities_steve_jobs.csv"
analyze_text(file_path, output_file)
