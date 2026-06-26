import re
import numpy as np


def clean_text(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^\w\s]", "", text)
    return text


def tokenize(text: str) -> list:
    return text.split()


def remove_stopwords(tokens: list, stopwords: set = None) -> list:
    if stopwords is None:
        stopwords = set()
    return [t for t in tokens if t not in stopwords and t.isalpha()]
