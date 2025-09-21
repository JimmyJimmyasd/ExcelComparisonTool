# This is a template for the correct BI tab structure
# Each tab should have content indented with exactly 16 spaces (4 levels)

# Executive Summary Tab
with bi_tabs[0]:
    st.subheader("📈 Executive Summary & Strategic Insights")
    st.markdown("**Comprehensive business analysis summary with key findings and recommendations**")
    
    # Generate executive summaries for both datasets
    exec_summary_a = bi_analyzer_a.generate_executive_summary()
    exec_summary_b = bi_analyzer_b.generate_executive_summary()
    
    # Tab content continues...

# Data Overview Tab  
with bi_tabs[1]:
    st.subheader("📊 Business Overview & Key Insights")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write(f"**{display_sheet_a} Overview**")
        # Tab content continues...
    
    # Tab content continues...

# Financial Analysis Tab
with bi_tabs[2]:
    st.subheader("💰 Financial Analysis")
    
    # Tab content continues...

# And so on for each tab...

# The key is that EVERY line inside a tab should be indented with exactly 16 spaces
# Comments between tabs should be at the same level as "with bi_tabs[X]:"