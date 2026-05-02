import pandas as pd
import os

FILE = "data/faq_history.csv"

def save(topic, difficulty, faqs):
    if not os.path.exists(FILE) or os.path.getsize(FILE) == 0:
        df = pd.DataFrame(columns=["topic","difficulty","faqs"])
    else:
        df = pd.read_csv(FILE)

    df.loc[len(df)] = [topic, difficulty, faqs]
    df.to_csv(FILE, index=False)

def load():
    if not os.path.exists(FILE) or os.path.getsize(FILE) == 0:
        return pd.DataFrame(columns=["topic","difficulty","faqs"])
    return pd.read_csv(FILE)