from transformers import AutoTokenizer, AutoModelForSeq2SeqLM

checkpoint = "suriya7/English-to-Tamil"

tokenizer = AutoTokenizer.from_pretrained(checkpoint)
model = AutoModelForSeq2SeqLM.from_pretrained(checkpoint)

print("Model downloaded ✅")