import streamlit as st
import pandas as pd
from dev import load_verticals, extract_assignments, get_hidden_fields, save_excel_clickable
import requests

LISTING_URL = "https://www.devnetjobsindia.org/rfp_assignments.aspx"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    )
}

st.set_page_config(page_title="DevNetJobs Scraper", layout="wide")

st.title("📊 DevNetJobs Scraper")

if st.button("Scrape RFP Assignments"):
    with st.spinner("Fetching RFP assignments..."):
        verticals = load_verticals("keywords.json")
        session = requests.Session()
        resp = session.get(LISTING_URL, headers=HEADERS, timeout=30)
        resp.raise_for_status()
        hidden = get_hidden_fields(resp.text)

        rows = extract_assignments(session, resp.text, hidden, verticals)

        if not rows:
            st.error("❌ No relevant assignments found with given keywords.")
        else:
            df = pd.DataFrame(rows)
            st.success(f"✅ Found {len(rows)} relevant assignments")

            # Display dataframe
            st.dataframe(df, use_container_width=True)

            # Download Excel
            save_excel_clickable(rows, "devnetjobindiascraper.xlsx")
            with open("devnetjobindiascraper.xlsx", "rb") as f:
                st.download_button(
                    "⬇️ Download Excel",
                    f,
                    file_name="devnetjobindiascraper.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
