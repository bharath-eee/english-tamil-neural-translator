from transformers import AutoTokenizer, AutoModelForSeq2SeqLM

# Load model
checkpoint = "suriya7/English-to-Tamil"

tokenizer = AutoTokenizer.from_pretrained(checkpoint)
model = AutoModelForSeq2SeqLM.from_pretrained(checkpoint)

# Translation function
def translate(text):
    inputs = tokenizer(text, return_tensors="pt")
    outputs = model.generate(**inputs, max_length=128)
    return tokenizer.decode(outputs[0], skip_special_tokens=True)

# Menu loop
while True:
    print("\n--- Translator ---")
    print("1. English → Tamil")
    print("2. Tamil → English (may not work well)")
    print("3. Exit")

    choice = input("Choose option: ")

    if choice == "1":
        text = input("Enter English text: ")
        print("Tamil:", translate(text))

    elif choice == "2":
        text = input("Enter Tamil text: ")
        print("English:", translate(text))  # model not trained for reverse

    elif choice == "3":
        print("Exiting...")
        break

    else:
        print("Invalid choice ❌")


        Enter English text: what is your name, you look soo beautiful
Tamil: உங்களது பெயர் என்ன? அழகான தோற்றம்தான்.