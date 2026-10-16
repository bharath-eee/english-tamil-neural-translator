import os
from huggingface_hub import login

token = os.getenv("HUGGINGFACE_TOKEN")
if not token:
    token = input("Enter your Hugging Face API token (press Enter to skip): ").strip()

if token:
    login(token)
    print("Successfully authenticated with Hugging Face Hub.")
else:
    print("No token provided. Skipping login.")
