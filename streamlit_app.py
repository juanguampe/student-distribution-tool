import streamlit as st
import pandas as pd
import random
import time
import io
import zipfile

# Configure the page
st.set_page_config(
    page_title="Student Distribution Tool",
    page_icon="📚",
    layout="wide"
)

# Initialize session state variables if they don't exist
if 'processed' not in st.session_state:
    st.session_state.processed = False
if 'seed' not in st.session_state:
    st.session_state.seed = int(time.time()) % 10000

st.title("Student Group Assignment Tool")
st.markdown("""
This tool helps distribute students into different seminar groups while ensuring:
- Students are placed in their chosen seminars (CN1/CN2, CS1/CS2, CF1/CF2)
- No scheduling conflicts (students won't be in G1, G1, G1 for example)
- Even distribution of students across groups
""")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file with student data", type=['xlsx'])

if uploaded_file is not None:
    # Show file details
    st.success(f"File uploaded: {uploaded_file.name}")
    
    # Show a preview of the data
    try:
        data = pd.read_excel(uploaded_file)
        st.write("Preview of uploaded data:")
        st.dataframe(data.head())
        
        # Verify required columns exist
        required_columns = ['NOMBRE', 'CN', 'CS', 'CF']
        missing_columns = [col for col in required_columns if col not in data.columns]
        
        if missing_columns:
            st.error(f"Missing required columns: {', '.join(missing_columns)}")
            st.write("Your file should have columns: NOMBRE, CN, CS, CF (and optionally GROUP)")
        else:
            # Add a checkbox to control whether to generate a new distribution
            generate_new = st.checkbox(
                "Generate new random distribution", 
                value=not st.session_state.get('processed', False),
                help="Uncheck to keep the current distribution if you've already processed once"
            )
            
            # Update seed based on checkbox
            if generate_new:
                st.session_state.seed = int(time.time()) % 10000
                st.session_state.processed = False
            
            # Add a process button
            if st.button("Process Student Assignments"):
                # Processing status indicator
                status = st.empty()
                status.info("Processing student assignments... This may take a moment.")
                
                # Use the seed from session state
                seed = st.session_state.seed
                st.write(f"Using random seed: {seed}")
                
                # Calculate dynamic max students per group
                max_students_per_group = {}
                for column in ['CN', 'CS', 'CF']:
                    choices = data[column].unique()
                    for choice in choices:
                        # Count students who chose this option
                        choice_count = (data[column] == choice).sum()
                        # Calculate optimal number per group (divide by 3 and round up)
                        max_per_group = (choice_count + 2) // 3  # Ensures all students fit
                        max_students_per_group[choice] = max_per_group
                        st.write(f"Maximum students per {choice} group: {max_per_group} (for {choice_count} students)")
                
                # Initialize subgroups for all combinations
                subgrupos = {}
                for curso in sorted(list(data['CN'].unique()) + list(data['CS'].unique()) + list(data['CF'].unique())):
                    for g in range(1, 4):
                        subgrupos[f"{curso}-G{g}"] = []
                
                # Set random seed and shuffle data
                random.seed(seed)
                data_shuffled = data.sample(frac=1, random_state=seed).reset_index(drop=True)
                
                # Progress bar
                progress_bar = st.progress(0)
                
                # Define assignment function
                def asignar_subgrupos(row, index, total):
                    try:
                        # Update progress 
                        progress_bar.progress(index / total)
                        
                        # Get student info
                        student_name = row['NOMBRE']
                        student_group = row.get('GROUP', '') # Get GROUP if it exists
                        
                        # Get the student's chosen seminars
                        cn_choice = row['CN']
                        cs_choice = row['CS']
                        cf_choice = row['CF']
                        
                        # Generate all possible group combinations
                        posibles_g = [(x, y, z) for x in range(1, 4) 
                                    for y in range(1, 4) 
                                    for z in range(1, 4) 
                                    if x != y and y != z and x != z]
                        
                        # Shuffle combinations to add randomness
                        random.shuffle(posibles_g)
                        
                        # Sort combinations to prioritize less filled groups
                        posibles_g.sort(key=lambda combo: 
                            sum(len(subgrupos[f"{choice}-G{g}"]) 
                                for choice, g in zip([cn_choice, cs_choice, cf_choice], combo)))
                        
                        # Try to assign the student to a valid combination of groups
                        for g_combo in posibles_g:
                            # Form the group keys correctly
                            grupo_keys = [
                                f"{cn_choice}-G{g_combo[0]}",
                                f"{cs_choice}-G{g_combo[1]}",
                                f"{cf_choice}-G{g_combo[2]}"
                            ]
                            
                            # Check if all groups have space
                            if all(len(subgrupos[gk]) < max_students_per_group.get(gk.split('-')[0], 21) for gk in grupo_keys):
                                # Assign student to these groups
                                for gk in grupo_keys:
                                    # Store both name and original group
                                    subgrupos[gk].append((student_name, student_group))
                                return True
                                
                        return False
                    except Exception as e:
                        st.error(f"Error processing student {row.get('NOMBRE', 'unknown')}: {str(e)}")
                        return False
                
                # Process all students
                results = []
                for i, (_, row) in enumerate(data_shuffled.iterrows()):
                    result = asignar_subgrupos(row, i, len(data_shuffled))
                    results.append(result)
                
                # Complete progress
                progress_bar.progress(1.0)
                
                # Show assignment summary
                successful_assignments = sum(results)
                status.success(f"Assigned {successful_assignments} out of {len(data)} students!")
                
                # Create Excel files in memory for download
                group_excel = io.BytesIO()
                summary_excel = io.BytesIO()
                
                # Create a dictionary of DataFrames, one for each group
                dfs_subgrupos = {}
                for nombre, estudiantes in subgrupos.items():
                    if estudiantes:  # Skip empty groups
                        # Create DataFrame with both name and original group
                        df = pd.DataFrame(estudiantes, columns=['NOMBRE', 'GROUP'])
                        dfs_subgrupos[nombre] = df
                
                # Create Excel writer for groups
                with pd.ExcelWriter(group_excel, engine='xlsxwriter') as writer:
                    # Write each DataFrame to a different sheet
                    for nombre_subgrupo, df in dfs_subgrupos.items():
                        sheet_name = nombre_subgrupo[:31]  # Excel has a 31-character limit for sheet names
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Create summary file showing each student's assignments
                student_assignments = []
                
                # For each student in the original data
                for _, row in data.iterrows():
                    student_name = row['NOMBRE']
                    student_group = row.get('GROUP', '')
                    
                    # Get the student's choices
                    cn_choice = row['CN']
                    cs_choice = row['CS']
                    cf_choice = row['CF']
                    
                    # Initialize with empty group assignments
                    cn_group = ""
                    cs_group = ""
                    cf_group = ""
                    
                    # Check each group for this student
                    for group_name, students in subgrupos.items():
                        # Look for the student's name in the tuple
                        student_names = [s[0] for s in students]
                        if student_name in student_names:
                            # Extract the G number
                            g_number = group_name.split('-G')[1]
                            
                            # Assign to the correct seminar
                            if group_name.startswith(cn_choice):
                                cn_group = g_number
                            elif group_name.startswith(cs_choice):
                                cs_group = g_number
                            elif group_name.startswith(cf_choice):
                                cf_group = g_number
                    
                    # Add to our list of assignments with all details
                    student_assignments.append({
                        'NOMBRE': student_name,
                        'GROUP': student_group,
                        'CN_ELECCION': cn_choice,
                        'CN_GRUPO': cn_group,
                        'CS_ELECCION': cs_choice,
                        'CS_GRUPO': cs_group,
                        'CF_ELECCION': cf_choice,
                        'CF_GRUPO': cf_group
                    })
                
                # Create summary DataFrame
                summary_df = pd.DataFrame(student_assignments)
                summary_df.to_excel(summary_excel, index=False)
                
                # Reset pointer to the beginning of files
                group_excel.seek(0)
                summary_excel.seek(0)
                
                # Group distribution statistics
                st.subheader("Group Size Summary")
                stats_cols = st.columns(3)
                
                with stats_cols[0]:
                    st.write("CN Groups:")
                    cn_stats = {}
                    for group in sorted([g for g in subgrupos.keys() if g.startswith('CN')]):
                        cn_stats[group] = len(subgrupos[group])
                    st.write(pd.Series(cn_stats))
                
                with stats_cols[1]:
                    st.write("CS Groups:")
                    cs_stats = {}
                    for group in sorted([g for g in subgrupos.keys() if g.startswith('CS')]):
                        cs_stats[group] = len(subgrupos[group])
                    st.write(pd.Series(cs_stats))
                
                with stats_cols[2]:
                    st.write("CF Groups:")
                    cf_stats = {}
                    for group in sorted([g for g in subgrupos.keys() if g.startswith('CF')]):
                        cf_stats[group] = len(subgrupos[group])
                    st.write(pd.Series(cf_stats))
                
                # Save the distribution results in session state
                st.session_state.processed = True
                st.session_state.group_excel = group_excel.getvalue()
                st.session_state.summary_excel = summary_excel.getvalue()
                
                # Display a sample of the student assignments
                st.subheader("Sample Student Assignments")
                st.dataframe(summary_df.head(10))
            
            # Download buttons section - only show if processing has been done
            if st.session_state.processed:
                st.subheader("Download Results")
                
                # Create three columns for download buttons
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.download_button(
                        label="📥 Download Group Assignments",
                        data=st.session_state.group_excel,
                        file_name="subgrupos_asignados.xlsx",
                        mime="application/vnd.ms-excel",
                        key="download_group"
                    )
                
                with col2:
                    st.download_button(
                        label="📥 Download Student Summary",
                        data=st.session_state.summary_excel,
                        file_name="asignaciones_por_estudiante.xlsx",
                        mime="application/vnd.ms-excel",
                        key="download_summary"
                    )
                
                # Create a combined download option (both files in a ZIP)
                with col3:
                    # Create a zip file in memory
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zip_file:
                        zip_file.writestr("subgrupos_asignados.xlsx", st.session_state.group_excel)
                        zip_file.writestr("asignaciones_por_estudiante.xlsx", st.session_state.summary_excel)
                    
                    # Reset pointer to beginning
                    zip_buffer.seek(0)
                    
                    st.download_button(
                        label="📥 Download All Files (ZIP)",
                        data=zip_buffer,
                        file_name="student_assignments.zip",
                        mime="application/zip",
                        key="download_zip"
                    )
                
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
        st.write("Please make sure your Excel file has the correct format.")
else:
    # Show sample data format when no file is uploaded
    st.info("Please upload an Excel file with the following structure:")
    sample_data = {
        'NOMBRE': ['STUDENT 1', 'STUDENT 2', 'STUDENT 3'],
        'GROUP': ['11A', '11B', '11C'],
        'CN': ['CN1', 'CN2', 'CN1'],
        'CS': ['CS1', 'CS2', 'CS1'],
        'CF': ['CF1', 'CF2', 'CF1']
    }
    st.dataframe(pd.DataFrame(sample_data))