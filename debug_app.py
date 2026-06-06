# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Debug", layout="wide")
st.title("🔍 Debug - بصمات")

att_file = st.file_uploader("📎 ارفع ملف البصمات", type=["xlsx", "xls"])

if att_file:
    try:
        df = pd.read_excel(att_file)
        st.write("**أعمدة:**", df.columns.tolist())
        st.write("**أول 3 صفوف:**")
        st.write(df.head(3))
        st.write(f"**الصفوف الكلية:** {len(df)}")
        st.write(f"**C/In count:** {len(df[df['State'].isin(['C/In','In'])])}")
    except Exception as e:
        st.error(f"❌ {e}")
