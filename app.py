import streamlit as st
import pandas as pd
import os

# Function to load or create the Excel file
def load_or_create_excel(file_path):
    if not os.path.exists(file_path):
        df = pd.DataFrame(columns=['Name', 'Email', 'Phone Number', 'Organization', 'Notes'])
        df.to_excel(file_path, index=False)
    else:
        df = pd.read_excel(file_path)
    return df

# Function to save data to Excel file
def save_to_excel(df, file_path):
    df.to_excel(file_path, index=False)

# Main function to run the Streamlit app
def main():
    st.title("Contact Manager System")

    # Load or create Excel file
    file_path = "contacts.xlsx"
    df = load_or_create_excel(file_path)

    # Add person form
    st.header("Add New Person")
    name = st.text_input("Name:")
    email = st.text_input("Email:")
    phone_number = st.text_input("Phone Number:")
    organization = st.text_input("Organization:")
    notes = st.text_area("Notes:")
    if st.button("Add Person"):
        new_row = pd.DataFrame({'Name': [name], 'Email': [email], 'Phone Number': [phone_number],
                                'Organization': [organization], 'Notes': [notes]})
        df = pd.concat([df, new_row], ignore_index=True)
        save_to_excel(df, file_path)
        st.success("Person added successfully!")

    # Live search functionality
    st.header("Live Search")
    search_term = st.text_input("Enter name or organization to search:", key="search")

    # Filter DataFrame based on search term
    filtered_df = df[(df['Name'].str.contains(search_term, case=False)) | 
                     (df['Organization'].str.contains(search_term, case=False))]

    # Display filtered results in real-time
    @st.cache_data
    def update_search_results(search_term, df):
        return df[(df['Name'].str.contains(search_term, case=False)) | 
                  (df['Organization'].str.contains(search_term, case=False))]

    search_results = update_search_results(search_term, df)

    # Show All button
    if st.button("Show All"):
        search_results = df

    # Display the table
    st.table(search_results)

    # Delete and Edit functionality
    with st.expander("Edit/Delete"):
        for index, row in search_results.iterrows():
            st.write(f"Index: {index}")
            name = st.text_input("Name", row['Name'], key=f"name_{index}")
            email = st.text_input("Email", row['Email'], key=f"email_{index}")
            phone_number = st.text_input("Phone Number", row['Phone Number'], key=f"phone_{index}")
            organization = st.text_input("Organization", row['Organization'], key=f"organization_{index}")
            notes = st.text_area("Notes", row['Notes'], key=f"notes_{index}")
            if st.button(f"Update {index}"):
                df.at[index, 'Name'] = name
                df.at[index, 'Email'] = email
                df.at[index, 'Phone Number'] = phone_number
                df.at[index, 'Organization'] = organization
                df.at[index, 'Notes'] = notes
                save_to_excel(df, file_path)
                st.success("Entry updated successfully!")
            
            if st.button(f"Delete {index}"):
                df = df.drop(index)
                save_to_excel(df, file_path)
                st.success("Entry deleted successfully!")

if __name__ == "__main__":
    main()
