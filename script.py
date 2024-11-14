import openpyxl
import os
import pandas as pd
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Load the bias dataset
dataset_file = "bias_dataset.xlsx"
if not os.path.exists(dataset_file):
    raise FileNotFoundError(f"Dataset file '{dataset_file}' not found in the current directory.")

bias_dataset = pd.read_excel(dataset_file)

# Check for existing labels file
labels_files = [f for f in os.listdir('.') if f.endswith('_labels.xlsx')]
if labels_files:
    labels_file = labels_files[0]
    labels_df = pd.read_excel(labels_file)
else:
    # Ask for user initials to create a new file
    user_initials = input("Enter your initials: ").strip()
    labels_file = f"{user_initials}_labels.xlsx"
    labels_df = pd.DataFrame(columns=[
        "preceding", "target", "following", "score", 
        "discussed", "autistic", "keyword", "language"
    ])

# Helper to ensure valid responses
def get_valid_input(prompt, valid_responses):
    while True:
        response = input(f"\033[96m{prompt}\033[0m").strip().lower()
        if response == "q":
            labels_df.to_excel(labels_file, index=False)
            print(f"Progress saved to {labels_file}.")
            exit()
        if response in valid_responses:
            return response
        else:
            print("\033[91mInvalid character entered. Please try again.\033[0m")

# Alternate logic
first_split = 153
tasks_sequence = ["classification", "info_collection"] * len(bias_dataset)
tasks_sequence[:first_split] = ["classification"] * first_split
tasks_sequence[first_split:] = ["info_collection"]

# Iterate through sentences in the dataset
for idx, row in bias_dataset.iterrows():
    if idx < len(labels_df):  # Skip already labeled entries
        continue

    # Calculate progress
    if idx < first_split:
        round_number = 1
        remaining_sentences = first_split - idx  # Sentences remaining in round 1
    else:
        round_number = 2
        remaining_sentences = len(bias_dataset) - idx  # Sentences remaining in round 2

    # Display progress
    print(f"\033[91mRound: {round_number}")
    print(f"Sentences remaining in this round: {remaining_sentences}\033[0m")

    print(f"\nPreceding: {row['preceding']}")
    print(f"\n\033[92mTarget: {row['target']}\033[0m")
    print(f"\nFollowing: {row['following']}")

    if tasks_sequence[idx] == "classification":
        # Classification first
        score = get_valid_input("Is this sentence ableist? (Y for yes, N for no): ", ["y", "n"])
        score = 1 if score == "y" else 0
        discussed = get_valid_input("Who is being discussed? (S for self, O for other): ", ["s", "o"])
        autistic = get_valid_input("Is the speaker autistic? (Y for yes, N for no): ", ["y", "n"])
        autistic = 1 if autistic == "y" else 0
        keyword = input("\033[96mKeywords or phrases (free-text): \033[0m").strip()
        language = get_valid_input("Does the sentence use IFL, PFL, or both? (I, P, B): ", ["i", "p", "b"])
        language = {"i": 0, "p": 1, "b": -1}[language]
    else:
        # Information collection first
        discussed = get_valid_input("Who is being discussed? (S for self, O for other): ", ["s", "o"])
        autistic = get_valid_input("Is the speaker autistic? (Y for yes, N for no): ", ["y", "n"])
        autistic = 1 if autistic == "y" else 0
        keyword = input("\033[96mKeywords or phrases (free-text): \033[0m").strip()
        language = get_valid_input("Does the sentence use IFL, PFL, or both? (I, P, B): ", ["i", "p", "b"])
        language = {"i": 0, "p": 1, "b": -1}[language]
        score = get_valid_input("Is this sentence ableist? (Y for yes, N for no): ", ["y", "n"])
        score = 1 if score == "y" else 0

    # Append results to DataFrame
    labels_df = pd.concat([labels_df, pd.DataFrame([{
        "preceding": row["preceding"],
        "target": row["target"],
        "following": row["following"],
        "score": score,
        "discussed": discussed,
        "autistic": autistic,
        "keyword": keyword,
        "language": language
    }])], ignore_index=True)

    # Save progress iteratively
    labels_df.to_excel(labels_file, index=False)

print(f"All sentences have been labeled. Final file saved to {labels_file}.")
