import streamlit as st
from ms_auth import get_token
from process_files import fetch_and_process_files
import openai

st.set_page_config(page_title="AI SharePoint Agent", layout="wide")
st.title("🔗 AI Agent for SharePoint Files")

openai.api_key = st.secrets["OPENAI_API_KEY"]

with st.sidebar:
    st.header("🔐 Authentication")
    client_id = st.text_input("Client ID", type="password")
    tenant_id = st.text_input("Tenant ID", type="password")
    site_url = st.text_input("Site Hostname (e.g., tiiuae.sharepoint.com/sites/xyz)")
    drive_id = st.text_input("Drive ID")
    folder_path = st.text_input("Folder Path (e.g., Shared Documents/xyz)")
    if st.button("Authenticate"):
        token = get_token(client_id, tenant_id)
        st.session_state["token"] = token
        st.success("Authenticated")

if "token" in st.session_state:
    st.subheader("📂 Files from SharePoint")
    data = fetch_and_process_files(
        access_token=st.session_state["token"],
        site_url=site_url,
        drive_id=drive_id,
        folder_path=folder_path
    )

    for fname, content in data.items():
        st.write(f"### {fname}")
        st.write(content[:500])
        if st.button(f"Summarize {fname}"):
            res = openai.ChatCompletion.create(
    model="gpt-4",
    messages=[
        {
            "role": "user",
            "content": f"Summarize this:\n{content[:3000]}"
        }
    ]
)

            st.success(res['choices'][0]['message']['content'])
