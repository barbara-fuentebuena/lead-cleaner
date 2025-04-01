import streamlit as st
import pandas as pd
import unicodedata
import re
from rapidfuzz import process, fuzz
from io import BytesIO

# --- Normalization function ---
def normalize(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    text = re.sub(r"[^\w\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    text = text.replace("\u00A0", " ")
    return text.strip().lower()

# --- Streamlit App ---
st.set_page_config(page_title="Lead Cleaner", layout="centered")
st.title("🧹 Lead Cleaner")
st.markdown("Upload your **leads_sn.xlsx** and **client_list.xlsx** files. We'll clean them for you ✨")

leads_file = st.file_uploader("Upload leads_sn.xlsx", type=["xlsx"], key="leads")
clients_file = st.file_uploader("Upload client_list.xlsx", type=["xlsx"], key="clients")

if leads_file and clients_file:
    with st.spinner("Processing files..."):
        # Load files
        leads_df = pd.read_excel(leads_file)
        clients_df = pd.read_excel(clients_file)

        # Column names
        leads_column = "companyName"
        clients_column = "Company"

        # Normalize
        leads_df["normalized_name"] = leads_df[leads_column].apply(normalize)
        clients_df["normalized_name"] = clients_df[clients_column].apply(normalize)

        # Exact match removal
        exact_normalized_set = set(clients_df["normalized_name"])
        leads_df["is_exact_match"] = leads_df["normalized_name"].isin(exact_normalized_set)
        exact_matches_df = leads_df[leads_df["is_exact_match"]].copy()
        filtered_leads = leads_df[~leads_df["is_exact_match"]].copy()

        # Fuzzy matching
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

        # Save to buffers
        def to_excel_buffer(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.success("✅ Done! Download your results below:")
        st.download_button("📥 Download Cleaned Leads", data=to_excel_buffer(final_leads), file_name="cleaned_leads.xlsx")
        st.download_button("📥 Download Fuzzy Matches to Review", data=to_excel_buffer(fuzzy_df), file_name="potential_matches_to_review.xlsx")
        st.download_button("📥 Download Exact Matches Removed", data=to_excel_buffer(exact_matches_df), file_name="exact_matches_removed.xlsx")
