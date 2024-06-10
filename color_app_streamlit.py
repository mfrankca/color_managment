import streamlit as st
import pandas as pd

EXCEL_FILE = 'colors.xlsx'

# Initialize session state for colors
if "colors" not in st.session_state:
    st.session_state.colors = pd.DataFrame(columns=["Color Name", "Pantone Number", "Hex Code", "RGB Values", "Notes"])

def load_colors(file=None):
    if file:
        df = pd.read_excel(file, sheet_name='Colors')
    else:
        try:
            df = pd.read_excel(EXCEL_FILE, sheet_name='Colors')
        except FileNotFoundError:
            df = pd.DataFrame(columns=["Color Name", "Pantone Number", "Hex Code", "RGB Values", "Notes"])
    st.session_state.colors = df.copy()  # Update session state with loaded colors
    return df

def save_colors(df):
    try:
        df.to_excel(file, sheet_name='Colors', index=False)
        st.sidebar.success("Colors saved to file")
        st.session_state.colors = df.copy()  # Update session state with saved colors
    except Exception as e:
        st.sidebar.error(f"Error saving colors to file: {e}")

def display_colors(df):
    for index, row in df.iterrows():
        with st.expander(f"{row['Color Name']} ({row['Pantone Number']})"):
            st.write(f"**{row['Color Name']}**")
            st.write(f"Pantone Number: {row['Pantone Number']}")
            st.write(f"Hex Code: {row['Hex Code']}")
            st.write(f"RGB Values: {row['RGB Values']}")
            st.write(f"Notes: {row['Notes']}")
            st.markdown(f'<div style="width:100px;height:50px;background-color:{row["Hex Code"]};border:1px solid #000;"></div>', unsafe_allow_html=True)

            if st.button(f"Edit {row['Color Name']}", key=f"edit_{index}"):
                new_color_name = st.text_input("Color Name", value=row['Color Name'], key=f"name_{index}")
                new_pantone_number = st.text_input("Pantone Number", value=row['Pantone Number'], key=f"pantone_{index}")
                new_hex_code = st.text_input("Hex Code", value=row['Hex Code'], key=f"hex_{index}")
                new_rgb_values = st.text_input("RGB Values", value=row['RGB Values'], key=f"rgb_{index}")
                new_notes = st.text_input("Notes", value=row['Notes'], key=f"notes_{index}")
                
                if st.button(f"Save {row['Color Name']}", key=f"save_{index}"):
                    df.at[index, 'Color Name'] = new_color_name
                    df.at[index, 'Pantone Number'] = new_pantone_number
                    df.at[index, 'Hex Code'] = new_hex_code
                    df.at[index, 'RGB Values'] = new_rgb_values
                    df.at[index, 'Notes'] = new_notes
                    save_colors(df)
                    st.success(f"Updated {new_color_name}")

# File uploader to load colors from an input file
st.sidebar.header("Load Colors from File")
uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file:
    st.session_state.colors = pd.DataFrame(columns=["Color Name", "Pantone Number", "Hex Code", "RGB Values", "Notes"])  # Clear session state
    df = load_colors(uploaded_file)
    st.sidebar.success("Colors loaded from file")

# Load colors from default Excel file if no file uploaded or if session state is empty
if not uploaded_file and st.session_state.colors.empty:
    df = load_colors()

# Main section to display colors
st.header("Pantone Color Manager")
display_colors(df)
