import pandas as pd
import streamlit as st
from io import BytesIO
from rapidfuzz import process, fuzz
import unicodedata
import re

# --- Normalization function with extra cleaning ---
def normalize_extra(text):
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("utf-8")
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()

# --- Streamlit App ---
st.set_page_config(page_title="Lead Cleaner", layout="centered")
st.title("🧹 Lead Cleaner")
st.markdown("Upload your **leads_sn.xlsx** and **client_list.xlsx** files. We'll clean them for you ✨")

# File uploaders for the two Excel files
leads_file = st.file_uploader("Upload leads_sn.xlsx", type=["xlsx"], key="leads")
clients_file = st.file_uploader("Upload client_list.xlsx", type=["xlsx"], key="clients")

if leads_file and clients_file:
    with st.spinner("Processing files..."):
        # Load files
        leads_df = pd.read_excel(leads_file)
        clients_df = pd.read_excel(clients_file)

        # Columns that need to be used for matching
        leads_column = "companyName"  
        clients_column = "companyName"  

        # Normalize company names in both DataFrames
        leads_df["normalized_name"] = leads_df[leads_column].apply(normalize_extra)
        clients_df["normalized_name"] = clients_df[clients_column].apply(normalize_extra)

        # Exact match removal
        exact_normalized_set = set(clients_df["normalized_name"])
        leads_df["is_exact_match"] = leads_df["normalized_name"].isin(exact_normalized_set)
        exact_matches_df = leads_df[leads_df["is_exact_match"]].copy()
        filtered_leads = leads_df[~leads_df["is_exact_match"]].copy()

        # Fuzzy matching for non-exact matches
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
                if 75 <= score < 100 and lead_name != match_name:
                    matches.append({
                        "Lead Name": lead_name,
                        "Client Name": match_name,
                        "Similarity": score
                    })

        fuzzy_df = pd.DataFrame(matches)

        # Merge fuzzy matches with original names
        if "Lead Name" in fuzzy_df.columns and not fuzzy_df.empty:
            fuzzy_df = fuzzy_df.merge(
                leads_df[[leads_column, "normalized_name"]],
                left_on="Lead Name",
                right_on="normalized_name",
                how="left"
            ).rename(columns={leads_column: "Original name in leads_sn.xlsx"}).drop(columns=["normalized_name"])

        # **Aquí está el fix: ahora NO eliminamos fuzzy matches del archivo final**
        final_leads = filtered_leads.drop(columns=["normalized_name", "is_exact_match"])

        # Row counts
        exact_matches_count = len(exact_matches_df)
        fuzzy_matches_count = len(fuzzy_df)
        final_leads_count = len(final_leads)

        # Function to convert DataFrame to Excel buffer for download
        def to_excel_buffer(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            return output.getvalue()

        st.success("✅ Done! Download your results below:")
        
        # Download buttons with descriptions
        st.write("#### ✅ Cleaned Leads")
        st.write("This file contains the final list of leads after removing all client companies from the original file:")
        st.download_button(f"📥 Download Cleaned Leads ({final_leads_count} rows)", 
                           data=to_excel_buffer(final_leads), 
                           file_name="cleaned_leads.xlsx")

        st.write("#### 🔍 Similar Matches to Review")
        st.write("This file contains company names that are similar but not exact matches to the exclusion list. You might want to review these manually:")
        st.download_button(f"📥 Download Similar Matches to Review ({fuzzy_matches_count} rows)", 
                           data=to_excel_buffer(fuzzy_df), 
                           file_name="potential_matches_to_review.xlsx")

        st.write("#### ❌ Exact Matches Removed")
        st.write("This file lists all the company names that were an exact match with the exclusion list and were removed from the leads. This is only info, YOU DON'T NEED TO DOWNLOAD 😄.")
        st.download_button(f"📥 Download Exact Matches Removed ({exact_matches_count} rows)", 
                           data=to_excel_buffer(exact_matches_df), 
                           file_name="exact_matches_removed.xlsx")

