import streamlit as st
import subprocess
import os
import glob
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
        if st.session_state["password"] == "popcorn2026":  # Change this password!
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
        
        # Generate Order Forms
        if st.button("🖨️ Generate Order Forms", use_container_width=True, key="order_forms"):
            with st.spinner("Generating order forms..."):
                try:
                    result = subprocess.run(
                        ["python", "export_orders.py"],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    
                    if result.returncode == 0:
                        st.success("Order forms generated successfully!")
                        st.text_area("Output:", result.stdout, height=200)
                        st.info("Files are ready in your Google Drive folders")
                    else:
                        st.error("Error generating order forms")
                        st.text_area("Error details:", result.stderr, height=200)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        st.markdown("---")
        
        # Update School Sheets
        if st.button("📊 Update School Sheets", use_container_width=True, key="update_sheets"):
            with st.spinner("Organizing school data..."):
                try:
                    result = subprocess.run(
                        ["python", "organize_schools.py"],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    
                    if result.returncode == 0:
                        st.success("School sheets updated successfully!")
                        st.text_area("Output:", result.stdout, height=200)
                    else:
                        st.error("Error updating school sheets")
                        st.text_area("Error details:", result.stderr, height=200)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        st.markdown("---")
        
        # Generate Production Report
        if st.button("📦 Generate Production Report", use_container_width=True, key="prod_report"):
            with st.spinner("Creating production report..."):
                try:
                    result = subprocess.run(
                        ["python", "create_production_report.py"],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    
                    if result.returncode == 0:
                        st.success("Production report created!")
                        st.text_area("Output:", result.stdout, height=200)
                        st.info("Check the 'Production' sheet in your spreadsheet")
                        
                        # Find the PDF file
                        pdf_files = glob.glob("Production_Report_*.pdf")
                        if pdf_files:
                            latest_pdf = max(pdf_files, key=os.path.getctime)
                            with open(latest_pdf, 'rb') as f:
                                st.download_button(
                                    label="📥 Download PDF Report",
                                    data=f.read(),
                                    file_name=latest_pdf,
                                    mime='application/pdf',
                                    key="download_prod_pdf"
                                )
                    else:
                        st.error("Error creating production report")
                        st.text_area("Error details:", result.stderr, height=200)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
    
    with col2:
        st.header("📈 Reports & Analytics")
        
        # Create Leaderboards
        if st.button("🏆 Create Leaderboards", use_container_width=True, key="leaderboards"):
            with st.spinner("Generating leaderboards..."):
                try:
                    result = subprocess.run(
                        ["python", "create_all_leaderboards.py"],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    
                    if result.returncode == 0:
                        st.success("Leaderboards created!")
                        st.text_area("Output:", result.stdout, height=200)
                        
                        # List HTML files
                        html_files = [f for f in os.listdir('.') if f.startswith('leaderboard_') and f.endswith('.html')]
                        
                        if html_files:
                            st.info(f"Created {len(html_files)} leaderboard files")
                            for idx, file in enumerate(html_files):
                                with open(file, 'rb') as f:
                                    st.download_button(
                                        label=f"📥 Download {file}",
                                        data=f.read(),
                                        file_name=file,
                                        mime='text/html',
                                        key=f"download_lb_{idx}"
                                    )
                    else:
                        st.error("Error creating leaderboards")
                        st.text_area("Error details:", result.stderr, height=200)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
        
        st.markdown("---")
        
        # Check for Data Errors
        if st.button("⚠️ Check for Data Errors", use_container_width=True, key="error_check"):
            with st.spinner("Checking for errors..."):
                try:
                    result = subprocess.run(
                        ["python", "find_data_errors.py"],
                        capture_output=True,
                        text=True,
                        timeout=300
                    )
                    
                    if result.returncode == 0:
                        st.success("Error check complete!")
                        st.text_area("Output:", result.stdout, height=200)
                        st.info("Check the 'Error Log' sheet in your spreadsheet")
                    else:
                        st.error("Error during error check")
                        st.text_area("Error details:", result.stderr, height=200)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
    
    # Footer
    st.markdown("---")
    st.markdown(f"Last refreshed: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
    
    # Sidebar
    with st.sidebar:
        st.header("ℹ️ System Info")
        st.info("""
        **Order Management System**
        
        This dashboard allows you to:
        - Generate order forms for any school
        - Update and organize school data
        - Create sales leaderboards
        - Generate production reports
        - Check for data errors
        
        All files are saved to Google Drive.
        """)
        
        st.markdown("---")
        st.markdown("**Need help?** Contact the administrator")