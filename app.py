import streamlit as st
import scripts
import os
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="Order Management System",
    page_icon="📦",
    layout="wide"
)

# Simple password protection
def check_password():
    """Returns True if the user has entered a correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == "popcorn2026":
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input(
            "Password", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        return False
    elif not st.session_state["password_correct"]:
        st.text_input(
            "Password", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("Password incorrect")
        return False
    else:
        return True

# Main app
if check_password():
    
    st.title("📦 Order Management System")
    st.markdown("---")
    
    # Create columns
    col1, col2 = st.columns(2)
    
    with col1:
        st.header("🏫 Order Processing")
st.header("🏫 Order Processing")
        
        # Generate Order Forms with school selector
        st.subheader("📄 Generate Order Forms")
        
        # Get list of schools from spreadsheet
        try:
            import gspread
            creds = scripts.get_credentials()
            gc = gspread.authorize(creds)
            spreadsheet = gc.open('MASTER SPRING 2026')
            all_sheets = spreadsheet.worksheets()
            school_sheets = [sheet.title.replace(' MASTER', '') for sheet in all_sheets if sheet.title.endswith(' MASTER') and sheet.title != 'MASTER']
            
            if school_sheets:
                selected_school = st.selectbox(
                    "Select School:",
                    options=sorted(school_sheets),
                    key="school_selector"
                )
                
                if st.button("🖨️ Generate Order Forms", use_container_width=True, key="generate_forms"):
                    with st.spinner(f"Generating order forms for {selected_school}..."):
                        try:
                            result, error, pdf_file = scripts.export_order_forms(selected_school)
                            
                            if error:
                                st.error(f"Error: {error}")
                            else:
                                st.success(f"Order forms generated for {selected_school}!")
                                
                                if pdf_file and os.path.exists(pdf_file):
                                    with open(pdf_file, 'rb') as f:
                                        st.download_button(
                                            label="📥 Download Order Forms PDF",
                                            data=f.read(),
                                            file_name=pdf_file,
                                            mime='application/pdf',
                                            key="download_forms_pdf"
                                        )
                            
                            st.text_area("Output:", result, height=300)
                            
                        except Exception as e:
                            st.error(f"Error: {str(e)}")
            else:
                st.warning("No school sheets found. Please run 'Update School Sheets' first.")
                
        except Exception as e:
            st.error(f"Error loading schools: {str(e)}")
        
        st.markdown("---")
        
        st.markdown("---")
        
        # Update School Sheets
        if st.button("📊 Update School Sheets", use_container_width=True, key="update_sheets"):
            with st.spinner("Organizing school data..."):
                try:
                    result, error = scripts.organize_schools()
                    
                    if error:
                        st.error(f"Error: {error}")
                    else:
                        st.success("School sheets updated successfully!")
                    
                    st.text_area("Output:", result, height=300)
                    
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        st.markdown("---")
        
        # Generate Production Report
        if st.button("📦 Generate Production Report", use_container_width=True, key="prod_report"):
            with st.spinner("Creating production report..."):
                try:
                    result, error, pdf_file = scripts.create_production_report()
                    
                    if error:
                        st.error(f"Error: {error}")
                    else:
                        st.success("Production report created!")
                        
                        if pdf_file and os.path.exists(pdf_file):
                            with open(pdf_file, 'rb') as f:
                                st.download_button(
                                    label="📥 Download PDF Report",
                                    data=f.read(),
                                    file_name=pdf_file,
                                    mime='application/pdf',
                                    key="download_prod_pdf"
                                )
                    
                    st.text_area("Output:", result, height=300)
                    
                except Exception as e:
                    st.error(f"Error: {str(e)}")
    
    with col2:
        st.header("📈 Reports & Analytics")
        
        st.info("More features coming soon!")
    
    # Footer
    st.markdown("---")
    st.markdown(f"Last refreshed: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
    
    # Sidebar
    with st.sidebar:
        st.header("ℹ️ System Info")
        st.info("""
        **Order Management System**
        
        This dashboard allows you to:
        - Update and organize school data
        - Generate production reports
        
        All files are saved to Google Drive.
        """)
        
        st.markdown("---")

        st.markdown("**Need help?** Contact the administrator")


