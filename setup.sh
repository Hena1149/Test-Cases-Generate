#!/bin/bash
python -m pip install --upgrade pip
python -m spacy download fr_core_news_sm
python -m nltk.downloader punkt stopwords