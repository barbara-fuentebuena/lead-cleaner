import pandas as pd
import unicodedata
import re
from rapidfuzz import process, fuzz

# --- Normalize company names ---
def normalize(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    text = re.sub(r"[^\w\s]", " ", text)       # replace special chars with space
    text = re.sub(r"\s+", " ", text)           # collapse multiple spaces
    text = text.replace("\u00A0", " ")         # replace non-breaking space
    return text.strip().lower()

# --- CONFIG ---
leads_file = "leads_sn.xlsx"
clients_file = "client_list.xlsx"
leads_column = "companyName"
clients_column = "Company"
output_cleaned = "cleaned_leads.csv"
output_fuzzy = "potential_matches_to_review.csv"
output_exact = "exact_matches_removed.csv"

print("📥 Loading files...")

try:
    leads_df = pd.read_excel(leads_file)
    clients_df = pd.read_excel(clients_file)

    leads_df.columns = leads_df.columns.str.strip()
    clients_df.columns = clients_df.columns.str.strip()

    print("✅ Files loaded successfully.")
except Exception as e:
    print("❌ Error loading files:", e)
    exit()

# --- Normalize names ---
leads_df["normalized_name"] = leads_df[leads_column].apply(normalize)
clients_df["normalized_name"] = clients_df[clients_column].apply(normalize)

# --- Exact match removal ---
exact_normalized_set = set(clients_df["normalized_name"])
leads_df["is_exact_match"] = leads_df["normalized_name"].isin(exact_normalized_set)

# Save exact matches
exact_matches_df = leads_df[leads_df["is_exact_match"]]
exact_matches_df.drop(columns=["normalized_name", "is_exact_match"]).to_csv(output_exact, index=False)

# Remove exact matches from leads
filtered_leads = leads_df[~leads_df["is_exact_match"]].copy()

print(f"🔍 Exact matches removed: {leads_df['is_exact_match'].sum()}")

# --- Fuzzy matching ---
matches = []
client_names_set = set(clients_df["normalized_name"].unique())
filtered_unique_names = filtered_leads["normalized_name"].drop_duplicates()

for lead_name in filtered_unique_names:
    if lead_name in exact_normalized_set:
        continue
    similars = process.extract(
        lead_name,
        client_names_set,
        scorer=fuzz.token_sort_ratio,
        limit=3
    )
    for match_name, score, _ in similars:
        if 85 <= score < 100 and lead_name != match_name:
            matches.append({
                "Name in leads_sn.xlsx": lead_name,
                "Name in client_list.xlsx": match_name,
                "Similarity": score
            })

fuzzy_df = pd.DataFrame(matches)

# Add original name from leads
fuzzy_df = fuzzy_df.merge(
    leads_df[[leads_column, "normalized_name"]],
    left_on="Name in leads_sn.xlsx",
    right_on="normalized_name",
    how="left"
).rename(columns={leads_column: "Original name in leads_sn.xlsx"}).drop(columns=["normalized_name"])

# Final cleaned leads
final_leads = filtered_leads[~filtered_leads["normalized_name"].isin(fuzzy_df["Name in leads_sn.xlsx"])]
final_leads = final_leads.drop(columns=["normalized_name", "is_exact_match"])

# Save outputs
final_leads.to_csv(output_cleaned, index=False)
fuzzy_df.to_csv(output_fuzzy, index=False)

# Summary
print("✅ Cleaning complete.")
print(f"- Total leads: {len(leads_df)}")
print(f"- Exact matches removed: {len(exact_matches_df)}")
print(f"- Fuzzy matches flagged for review: {len(fuzzy_df)}")
print(f"- Final cleaned leads: {len(final_leads)}")
print("\n📄 Files generated:")
print(f"- {output_cleaned}")
print(f"- {output_fuzzy}")
print(f"- {output_exact}")
