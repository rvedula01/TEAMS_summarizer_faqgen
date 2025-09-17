# -*- coding: utf-8 -*-
"""
Created on Fri Jun  6 15:17:32 2025

@author: ShivakrishnaBoora
"""


import re
import nltk
from nltk.tokenize import sent_tokenize
from bert_score import score
import numpy as np
from sentence_transformers import SentenceTransformer

def normalize_text(text: str) -> str:
    """
    Clean text for display purposes - preserves timestamps but removes punctuation and filler words.
    """
    text = text.lower()
    text = re.sub(r"[“”‘’]", "'", text)          # unify smart quotes
    text = re.sub(r"[^a-z0-9\s']", " ", text)    # keep letters, digits, apostrophes, and spaces
    text = re.sub(r"\s+", " ", text).strip()     # collapse whitespace
    
    # Remove filler words
    fillers = ["um", "uh", "you know", "like", "so"]
    for f in fillers:
        text = re.sub(rf"\b{f}\b", "", text)
    return text.strip()

def clean_text_for_similarity(text: str) -> str:
    """
    Clean text specifically for similarity calculations by removing timestamps and other non-content elements.
    This version is used only for similarity calculations, not for display or downloads.
    """
    # Remove timestamps in format M/DD/YYYY MM:SS
    text = re.sub(r'\b\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}\b', '', text)
    # Remove timestamps in format MM/DD/YYYY, HH:MM:SS
    text = re.sub(r'\b\d{1,2}/\d{1,2}/\d{4},\s+\d{1,2}:\d{2}:\d{2}\b', '', text)
    # Clean up any extra whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Lowercase and remove punctuation for similarity
    text = text.lower()
    text = re.sub(r"[“”‘’]", "'", text)          # unify smart quotes
    text = re.sub(r'[\W_]', ' ', text)    # remove all non-word characters
    text = re.sub(r'[^a-z0-9\s\']', ' ', text)    # keep only letters, digits, spaces, and apostrophes
    text = re.sub(r"\s+", " ", text).strip()     # collapse whitespace

    # Remove filler words for similarity
    fillers = ["um", "uh", "you know", "like", "so"]
    for f in fillers:
        text = re.sub(rf"\b{f}\b", "", text)
    return re.sub(r"\s+", " ", text).strip()

# Sentence tokenization
nltk.download("punkt", quiet=True)

def compute_rouge_scores(reference: str, hypothesis: str) -> dict:
    """
    Compute ROUGE-L F-measure score.
    Returns dictionary with rougeL score.
    """
    # Use the global scorer instance
    try:
        scores = _rouge_scorer.score(reference, hypothesis)
        return {
            'rougeL': scores['rougeL'].fmeasure
        }
    except Exception as e:
        print(f"Error computing ROUGE scores: {e}")
        return {'rougeL': 0.0}

def compute_bert_score(reference: str, hypothesis: str) -> tuple:
    """
    Compute BERTScore precision, recall, and F1.
    Returns tuple of (precision, recall, F1) scores.
    """
    P, R, F1 = score([hypothesis], [reference], lang="en")
    return P.item(), R.item(), F1.item()

def split_into_sentences(text: str) -> list:
    """Split text into sentences using NLTK."""
    return sent_tokenize(text)

def align_sentences(reference_sentences: list, hypothesis_sentences: list) -> list:
    """
    Align sentences from reference and hypothesis using sentence-level cosine similarity.
    Returns list of tuples (reference_idx, hypothesis_idx, similarity_score)
    """
    alignments = []
    for ref_idx, ref_sent in enumerate(reference_sentences):
        best_match = None
        best_score = 0.0
        for hyp_idx, hyp_sent in enumerate(hypothesis_sentences):
            # Compute cosine similarity between sentence embeddings
            ref_emb = get_sentence_embedding(ref_sent)
            hyp_emb = get_sentence_embedding(hyp_sent)
            score = cosine_similarity(ref_emb, hyp_emb)
            if score > best_score:
                best_score = score
                best_match = hyp_idx
        if best_match is not None:
            alignments.append((ref_idx, best_match, best_score))
    return alignments

# Add helper functions for sentence embeddings and cosine similarity
def get_sentence_embedding(sentence: str) -> np.ndarray:
    """Get sentence embedding using BERT"""
    # Use BERT model to get embeddings
    model = SentenceTransformer('all-MiniLM-L6-v2')
    return model.encode([sentence])[0]

def cosine_similarity(vec1: np.ndarray, vec2: np.ndarray) -> float:
    """Calculate cosine similarity between two vectors"""
    return np.dot(vec1, vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2))

def compute_overall_metrics(reference: str, hypothesis: str) -> dict:
    """
    Compute overall metrics directly from reference and hypothesis texts.
    Returns dictionary with alignment and semantic accuracy scores.
    """
    # Split into sentences
    ref_sents = split_into_sentences(reference)
    hyp_sents = split_into_sentences(hypothesis)
    
    # Align sentences
    alignments = align_sentences(ref_sents, hyp_sents)
    
    # Initialize scores
    scores = {
        'alignment_score': [],
        'bert_f1': []
    }
    
    # Compute metrics for each aligned pair
    for ref_idx, hyp_idx, alignment_score in alignments:
        ref_sent = ref_sents[ref_idx]
        hyp_sent = hyp_sents[hyp_idx]
        
        bert_scores = compute_bert_score(ref_sent, hyp_sent)
        
        scores['alignment_score'].append(alignment_score)
        scores['bert_f1'].append(bert_scores[2])  # BERTScore returns (P, R, F1) tuple, we want F1 score
    
    return {
        'alignment_score': float(np.mean(scores['alignment_score'])),
        'bert_f1': float(np.mean(scores['bert_f1']))
    }
