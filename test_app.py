import streamlit as st

# Simple test to check if comparison modes show up
st.title("Test - Excel Comparison Tool")

comparison_mode = st.radio(
    "🔄 Comparison Mode:",
    options=["Two Different Files", "Same File (Different Sheets)", "Multi-Sheet Batch Processing", "Cross-Sheet Data Consolidation", "Historical Comparison Mode"],
    index=0,
    help="Choose comparison type"
)

st.write(f"Selected mode: {comparison_mode}")

if comparison_mode == "Cross-Sheet Data Consolidation":
    st.success("✅ Cross-Sheet Data Consolidation mode detected!")
    
if comparison_mode == "Historical Comparison Mode":
    st.success("✅ Historical Comparison Mode detected!")
    
if comparison_mode == "Multi-Sheet Batch Processing":
    st.success("✅ Multi-Sheet Batch Processing mode detected!")