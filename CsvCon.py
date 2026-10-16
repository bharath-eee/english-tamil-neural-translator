from datasets import load_dataset
import pandas as pd

# Step 1: Load dataset
dataset = load_dataset("gopi30/english-tamil")

# Step 2: Convert to pandas
df = dataset['train'].to_pandas()

# Step 3: Save to CSV
df.to_csv("english_tamil_dataset.csv", index=False, encoding="utf-8")

print("Saved ✅")