import pandas as pd
import re, os, sys
import chardet

# --- VALIDATOR FUNCTIONS ---
# These functions run pre-flight checks to ensure the script has everything it needs.

def validate_dependencies():
    """Checks if all required third-party libraries are installed."""
    print("Validator: Checking for required libraries...")
    try:
        import pandas
        import chardet
    except ImportError as e:
        library_name = str(e).split("'")[1]
        raise ImportError(f"Missing required library '{library_name}'. Please install it by running: pip install {library_name}")
    print(" -> Dependencies are satisfied.")

def validate_paths_and_files(input_folder, output_folder, required_files):
    """Validates the existence of folders and a list of required files."""
    print("Validator: Checking folder paths and file existence...")
    # 1. Check folders
    if not os.path.isdir(input_folder):
        raise FileNotFoundError(f"Input folder not found: {input_folder}")
    if not os.path.isdir(output_folder):
        raise FileNotFoundError(f"Output folder not found: {output_folder}")

    # 2. Check for write permissions in the output folder
    try:
        test_file = os.path.join(output_folder, 'permission_test.tmp')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
    except (IOError, OSError) as e:
        raise PermissionError(f"Cannot write to the output folder: {output_folder}. Please check permissions. Error: {e}")

    # 3. Check for all required input files
    for file_path in required_files:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Required input file not found: {file_path}")
            
    print(" -> Paths and files are valid.")

def validate_excel_schema(file_path, required_sheet, file_description, required_cols_indices=None):
    """Checks for a specific sheet in an Excel file and optionally validates column indices."""
    print(f"Validator: Checking schema for '{file_description}'...")
    try:
        xls = pd.ExcelFile(file_path)
        if required_sheet not in xls.sheet_names:
            raise ValueError(f"File '{file_description}' is missing the required sheet: '{required_sheet}'. Found sheets: {xls.sheet_names}")
        
        # Optional: Check if column indices are valid
        if required_cols_indices:
            df_cols = pd.read_excel(file_path, sheet_name=required_sheet, nrows=0).columns
            num_cols = len(df_cols)
            for desc, col_index in required_cols_indices.items():
                if col_index >= num_cols:
                    raise IndexError(f"The '{desc}' index ({col_index}) is out of bounds for sheet '{required_sheet}' in '{file_description}'. The sheet only has {num_cols} columns (indices 0 to {num_cols-1}).")

    except FileNotFoundError:
        raise FileNotFoundError(f"Could not find the Excel file to validate: {file_path}")
    except Exception as e:
        raise e # Re-raise other errors
    print(f" -> Schema for '{file_description}' is valid.")

def validate_inc_file_content(file_path, pattern):
    """Checks the .inc file for the presence of the essential variable pattern."""
    print(f"Validator: Checking content integrity of '{os.path.basename(file_path)}'...")
    try:
        encoding = chardet.detect(open(file_path, 'rb').read())['encoding']
        with open(file_path, 'r', encoding=encoding) as f:
            for line in f:
                if re.search(pattern, line):
                    print(f" -> Content integrity for '{os.path.basename(file_path)}' is OK.")
                    return True # Found at least one match
    except (FileNotFoundError, TypeError):
         raise FileNotFoundError(f"Could not read .inc file for validation: {file_path}")

    # If the loop finishes without returning, no matches were found
    raise ValueError(f"Data Integrity Error in '{os.path.basename(file_path)}': No lines containing the required variable pattern were found. The script cannot extract any variables.")

# --- MAIN SCRIPT LOGIC ---

def main_process(LongVarFileName,TabPlanFileName,InputDir,OutputDir,BannerFileInfo,tabPlanNumber,QuesionIndex,LabelIndex,baseTextIndex,SheetNameInput):
    """
    Contains the entire processing logic of the original script.
    """
    # --- CONFIGURATION ---
    # File names
    countFileName = LongVarFileName
    bannerFileName = BannerFileInfo
    NodesFileName = 'Var_name'
    tabplan_file_name = TabPlanFileName

    MatchedVariablesFileName = 'Matched_Variables'
    MatchingFileName = 'Matched_df'
    summeryFileName = 'Unmatched_Summary'

    # Folder Paths
    inputFolderPath = InputDir
    outputFolderPath = OutputDir

    # --- STATIC VALIDATION (Before user input) ---
    incFile = os.path.join(inputFolderPath, f'{countFileName}.inc')
    bannerFile = os.path.join(inputFolderPath, f'{bannerFileName}.xlsx')
    nodesFile = os.path.join(inputFolderPath, f'{NodesFileName}.txt')
    tabplan_file_path = os.path.join(inputFolderPath, f'{tabplan_file_name}.xlsm')
    
    validate_paths_and_files(inputFolderPath, outputFolderPath, [incFile, bannerFile, nodesFile, tabplan_file_path])
    validate_excel_schema(bannerFile, 'Titles', f"{bannerFileName}.xlsx")
    inc_pattern = r'TableDoc\.Coding\.CreateCategorizedVariable\("([^"]+)"'
    validate_inc_file_content(incFile, inc_pattern)
    print("\n--- Initial validation passed. Please provide input. ---\n")

    # --- USER INPUT ---
    print("Select Tab Plan Type:")
    print("1. Autotabplan")
    print("2. Tabplan A")
    print("3. Custom")

    while True:
        try:
            tab_plan_choice = tabPlanNumber#int(input("Enter your choice (1, 2, or 3): "))
            if tab_plan_choice in [1, 2, 3]:
                break
            else:
                print("Please enter 1, 2, or 3")
        except ValueError:
            print("Please enter a valid number (1, 2, or 3)")

    # Set tab plan parameters based on user choice
    if tab_plan_choice == 1:  # Autotabplan
        questionColumn, labelColumn, basetextColumn, sheetName = 2, 3, 7, 'Stub Specs'
    elif tab_plan_choice == 2:  # Tabplan A
        questionColumn, labelColumn, basetextColumn, sheetName = 0, 4, 9, 'STUB SPECS'
    else:  # Custom
        print("Selected: Custom - Please enter the parameters:")
        questionColumn = QuesionIndex #int(input("Enter Question Column index (0-based): "))
        labelColumn = LabelIndex #int(input("Enter Label Column index (0-based): "))
        basetextColumn = baseTextIndex #int(input("Enter Base Text Column index (0-based): "))  
        sheetName = SheetNameInput #input("Enter Sheet Name: ")

    print(f"\nUsing settings: Question Column={questionColumn}, Label Column={labelColumn}, Base Text Column={basetextColumn}, Sheet Name='{sheetName}'\n")

    # --- DYNAMIC VALIDATION (After user input) ---
    col_indices_to_check = {
        "Question Column": questionColumn,
        "Label Column": labelColumn,
        "Base Text Column": basetextColumn
    }
    validate_excel_schema(tabplan_file_path, sheetName, f"{tabplan_file_name}.xlsm", col_indices_to_check)
    print("\n--- All validations passed. Starting main process. ---\n")

    # --- SCRIPT EXECUTION ---
    # Path Creation for output files
    MatchedVariablesFile = os.path.join(outputFolderPath, f'{MatchedVariablesFileName}.xlsx')
    MatchedFile = os.path.join(outputFolderPath, f'{MatchingFileName}.xlsx')
    summeryFile = os.path.join(outputFolderPath, f'{summeryFileName}.txt')

    # Reading the files
    banner_titles = pd.read_excel(bannerFile, sheet_name='Titles')
    count_file_variable = pd.read_csv(nodesFile, header=None)

    # ... (The rest of your script's logic from 'INC TO TEXT' onwards) ...
    # NOTE: I've copied your entire script logic below this comment block.
    
    '''INC TO TEXT'''
    # Function to detect file encoding
    def detect_encoding(file_path):
        with open(file_path, 'rb') as file:
            result = chardet.detect(file.read())
        return result['encoding']

    # Detect the encoding of the .INC file
    encoding = detect_encoding(incFile)
    print(f"Detected encoding: {encoding}")

    # Read from the .INC file and extract variables
    variables = []
    pattern = r'TableDoc\.Coding\.CreateCategorizedVariable\("([^"]+)"'

    with open(incFile, 'r', encoding=encoding) as inc_file:
        for line in inc_file:
            line = line.strip()
            if not line:
                continue
            matches = re.findall(pattern, line)
            if matches:
                variables.extend(matches)

    # Print the list of extracted variable names
    print("Extracted variables:")
    for variable in variables:
        print(variable)
        
    '''Creating Matched File'''
    # Load the data
    count_file_variable.columns = ['Label']

    Definitions = banner_titles[banner_titles.columns[0]].tolist()

    Original_labels = Definitions

    Titles = [str(x).split('.', 2)[0].strip() if (len(str(x).split('.')) == 2) else str(x).split('.', 2)[:2][0].strip()
            for x in banner_titles[banner_titles.columns[0]]]

    Definitions = ["".join(str(x).split('.')[1:]) if (len(str(x).split('.')) >= 2) else str(x).split('.')[1:]
            for x in banner_titles[banner_titles.columns[0]]]

    # Create matching_df
    matching_df = pd.DataFrame(list(zip(Original_labels, Titles, Definitions)), columns=['Original Labels', 'Title', 'Definition'])

    # Initialize the second column with None
    matching_df['Matched_Label'] = None

    # Function to remove a matched label from count_file_variable['Label']
    def remove_label(label):
        nonlocal count_file_variable
        count_file_variable = count_file_variable[count_file_variable['Label'] != label]

    # Iterate through matching_df and match with count_file_variable['Label']
    for i, row in matching_df.iterrows():
        title = row['Title']
        definition = row['Definition']
        
        if 'Summary' in definition:
            continue
        
        title_lower = title.lower()
        count_file_variable_lower = count_file_variable['Label'].str.lower()
        
        if title_lower in count_file_variable_lower.values:
            original_label = count_file_variable[count_file_variable_lower == title_lower].iloc[0]['Label']
            matching_df.at[i, 'Matched_Label'] = original_label
            remove_label(original_label)
        else:
            pattern = re.escape(title_lower) + r'\[\{'  
            matched_label = count_file_variable[count_file_variable_lower.str.contains(pattern, regex=True)]
            if not matched_label.empty:
                first_match = matched_label.iloc[0]['Label']
                matching_df.at[i, 'Matched_Label'] = first_match
                remove_label(first_match)
            else:
                if pd.isna(matching_df.at[i, 'Matched_Label']):
                    for label in count_file_variable['Label']:
                        if '.' in label:
                            post_dot_variable = label.split('.')[1]
                            if post_dot_variable.lower() == title_lower:
                                matching_df.at[i, 'Matched_Label'] = label
                                remove_label(label)
                                break

    matching_df.to_excel(MatchedFile, index=False)


    '''Updated Part'''
    matching_df['Definition'] = matching_df['Definition'].fillna('').astype(str)
    filtered_matching_df = matching_df[~matching_df['Definition'].str.contains(r'\[Top 2 Box - Summary\]|\[Bottom 2 Box - Summary\]|\[Summary - Mean\]', regex=True)]
    total_entries = len(filtered_matching_df)
    filled_entries = filtered_matching_df['Matched_Label'].notnull().sum()
    blank_entries = filtered_matching_df['Matched_Label'].isnull().sum()
    percentage_matched = (filled_entries / total_entries) * 100 if total_entries > 0 else 0

    with open(summeryFile, 'w') as file:
        file.write("Unmatched Titles Summary\n========================\n\n")
        unmatched_titles = filtered_matching_df[filtered_matching_df['Matched_Label'].isnull()]['Title']
        for title in unmatched_titles:
            file.write(f"{title}\n")
        file.write("\nSummary Statistics\n==================\n")
        file.write(f"Total Entries: {total_entries}\nFilled Entries: {filled_entries}\n")
        file.write(f"Blank Entries: {blank_entries}\nPercentage of Tables Matched: {percentage_matched:.2f}%\n")

    print(f"Summary of unmatched titles and statistics has been written to {summeryFile}")

    def append_cat_and_update_mean_table(excel_file_path, variables, output_excel_file_path):
        matching_df = pd.read_excel(excel_file_path)
        if 'Matched_Label' not in matching_df.columns:
            print("Error: 'Matched_Label' column not found in the Excel file.")
            return None
        if 'Mean_Table' not in matching_df.columns:
            matching_df['Mean_Table'] = False
        matching_df['Matched_Label'] = matching_df['Matched_Label'].astype(str)
        for variable in variables:
            for index, row in matching_df.iterrows():
                matched_label = row['Matched_Label']
                if matched_label and matched_label != 'nan':
                    if (matched_label.startswith(variable.split('.')[0]) and 
                        matched_label.endswith(variable.split('.')[-1])):
                        print(f"Appending '_cat' to '{matched_label}' and setting Mean_Table to TRUE")
                        matching_df.at[index, 'Matched_Label'] = matched_label + "_cat"
                        matching_df.at[index, 'Mean_Table'] = True
        matching_df.to_excel(output_excel_file_path, index=False)
        print(f"Updated DataFrame has been saved to {output_excel_file_path}")
        return matching_df

    updated_matching_df = append_cat_and_update_mean_table(MatchedFile, variables, MatchedVariablesFile)
    if updated_matching_df is not None:
        print(updated_matching_df)
        
    '''Tab Plan Titles Fetching'''
    def process_tab_plan_title(title):
        if pd.isna(title): return ''
        title_str = str(title)
        if '.' in title_str[:15]:
            dot_index = title_str.find('.')
            return title_str[dot_index + 1:].strip()
        else:
            return title_str

    stub_specs_df = pd.read_excel(tabplan_file_path, sheet_name=sheetName)
    print(f"Successfully loaded '{sheetName}' sheet from {tabplan_file_name}.xlsm")
    selected_columns = stub_specs_df.iloc[1:, [questionColumn, labelColumn, basetextColumn]]
    selected_columns.columns = ['Question', 'Table Title', 'Base Text']
    print("Applying new rule to Tab Plan titles...")
    selected_columns['Table Title'] = selected_columns['Table Title'].apply(process_tab_plan_title)
    print("Rule applied. Titles have been processed.")
    selected_columns.index = range(1, len(selected_columns) + 1)
    matched_variables_df_file = pd.read_excel(MatchedVariablesFile)
    print(f"Successfully loaded data from {MatchedVariablesFileName}.xlsx")
    if 'Tab Plan Titles' not in matched_variables_df_file.columns:
        matched_variables_df_file['Tab Plan Titles'] = ''
    if 'Base Text' not in matched_variables_df_file.columns:
        matched_variables_df_file['Base Text'] = ''
    question_title_dict = {q.lower() if isinstance(q, str) else q: t for q, t in zip(selected_columns['Question'], selected_columns['Table Title'])}
    question_base_text_dict = {q.lower() if isinstance(q, str) else q: bt for q, bt in zip(selected_columns['Question'], selected_columns['Base Text'])}

    def extract_base_pattern(question):
        if not isinstance(question, str): return question
        match = re.match(r'^([QqA-Za-z]+\d+)', question)
        if match: return match.group(1).lower()
        return question.lower()

    pattern_title_dict, pattern_base_text_dict = {}, {}
    for q, t, bt in zip(selected_columns['Question'], selected_columns['Table Title'], selected_columns['Base Text']):
        base_pattern = extract_base_pattern(q)
        if base_pattern not in pattern_title_dict:
            pattern_title_dict[base_pattern] = t
            pattern_base_text_dict[base_pattern] = bt

    exact_matches, pattern_matches = 0, 0
    for idx, row in matched_variables_df_file.iterrows():
        title = row['Title']
        title_lower = title.lower() if isinstance(title, str) else title
        if title_lower in question_title_dict:
            matched_variables_df_file.at[idx, 'Tab Plan Titles'] = question_title_dict[title_lower]
            matched_variables_df_file.at[idx, 'Base Text'] = question_base_text_dict[title_lower]
            exact_matches += 1
        else:
            base_pattern = extract_base_pattern(title)
            if base_pattern in pattern_title_dict:
                matched_variables_df_file.at[idx, 'Tab Plan Titles'] = pattern_title_dict[base_pattern]
                matched_variables_df_file.at[idx, 'Base Text'] = pattern_base_text_dict[base_pattern]
                pattern_matches += 1
                print(f"Pattern match found: '{title}' -> '{base_pattern}' -> '{pattern_title_dict[base_pattern]}'")

    matched_variables_df_file.to_excel(MatchedVariablesFile, index=False)
    print(f"Successfully updated {MatchedVariablesFileName}.xlsx with Tab Plan Titles and Base Text.")
    matched_count = matched_variables_df_file['Tab Plan Titles'].replace('', pd.NA).count()
    total_count = len(matched_variables_df_file)
    print(f"\nMatching Statistics:\nTotal rows: {total_count}\nExact matches: {exact_matches}\nPattern matches: {pattern_matches}\nTotal matched: {matched_count}\nMatching percentage: {matched_count/total_count*100:.2f}%")
    unmatched_titles_df = matched_variables_df_file[matched_variables_df_file['Tab Plan Titles'] == '']
    matched_questions_lower = {title.lower() if isinstance(title, str) else title for title in matched_variables_df_file[matched_variables_df_file['Tab Plan Titles'] != '']['Title']}
    extra_titles_in_tabplan = {q: t for q, t in zip(selected_columns['Question'], selected_columns['Table Title']) if (q.lower() if isinstance(q, str) else q) not in matched_questions_lower}
    file_exists = os.path.isfile(summeryFile)
    with open(summeryFile, 'a' if file_exists else 'w') as f:
        if not file_exists: f.write("=== UNMATCHED TITLES SUMMARY ===\n\n")
        else: f.write("\n\n=== NEW UNMATCHED TITLES SUMMARY ===\n\n")
        f.write("TITLES IN MATCHED_VARIABLES_DF WITHOUT MATCHING TAB PLAN TITLES:\n" + "=" * 60 + "\n")
        if len(unmatched_titles_df) > 0:
            for idx, row in unmatched_titles_df.iterrows(): f.write(f"{row['Title']}\n")
        else: f.write("All titles in matched_variables_df were matched with Tab Plan titles.\n")
        f.write("\n\nEXTRA TITLES IN TAB PLAN NOT FOUND IN MATCHED_VARIABLES_DF:\n" + "=" * 60 + "\n")
        if extra_titles_in_tabplan:
            for question, title in extra_titles_in_tabplan.items(): f.write(f"Question: {question}\nTitle: {title}\n\n")
        else: f.write("No extra titles found in Tab Plan.\n")
        f.write(f"\n\nMATCHING SUMMARY:\n" + "=" * 60 + "\n")
        f.write(f"Exact matches: {exact_matches}\nPattern matches: {pattern_matches}\nTotal matches: {matched_count}\nMatching percentage: {matched_count/total_count*100:.2f}%\n")
    print(f"\nUnmatched titles summary appended to {summeryFile}")


if __name__ == "__main__":
    try:
        # First, check for essential library dependencies
        validate_dependencies()
        # Then, run the main process which includes all other validations
        main_process()
        print("\nScript finished successfully!")
    except (FileNotFoundError, ValueError, PermissionError, ImportError, IndexError) as e:
        print(f"\nVALIDATION FAILED: A critical error occurred that prevented the script from running.")
        print(f"Error Details: {e}")
        sys.exit(1) # Exit the script with an error code
    except Exception as e:
        print(f"\nAn unexpected error occurred during execution: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1) # Exit the script with an error code