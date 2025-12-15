import pandas as pd
import numpy as np
import re 

def update_and_validate_members(file_2026_path, file_2025_path, output_file_path):
    # --- CONFIRMED COLUMN NAMES ---
    COL_2026_FIRST_NAME = 'First Name'
    COL_2026_LAST_NAME = 'Last Name'
    COL_2026_EMAIL = 'Email'
    COL_2026_STUDENT_ID = 'Student ID'
    COL_2026_USER_TYPE = 'User Type'
    
    COL_2025_FULL_NAME = 'name' 
    COL_2025_UNI_ONE = 'On UniOne?'
    COL_2025_VERIFICATION = 'Selected correct membership type (student/general)?'
    
    # --- STATIC CHECK VALUES ---
    STUDENT_EMAIL_DOMAIN = '@student.monash.edu'
    STUDENT_USER_TYPE = 'Monash Student'

    try:
        # --- 1. Load the Excel Sheets (Omitted for brevity, loading unchanged) ---
        print(f"Loading 2026 data from: {file_2026_path}")
        df_2026 = pd.read_excel(file_2026_path)
        
        print(f"Loading 2025 data from: {file_2025_path}")
        df_2025 = pd.read_excel(file_2025_path, engine='openpyxl')

        # --- PREPARE FOR CHANGE TRACKING (Omitted for brevity, tracking unchanged) ---
        original_uni_one = df_2025[COL_2025_UNI_ONE].astype(str).str.strip().str.lower().copy()
        original_verification = df_2025[COL_2025_VERIFICATION].astype(str).str.strip().str.lower().copy()
        
        # --- 2. Standardize Names for Comparison (Omitted for brevity, keys unchanged) ---
        name_key_2026 = (
            df_2026[COL_2026_FIRST_NAME].astype(str).str.strip() + ' ' + df_2026[COL_2026_LAST_NAME].astype(str).str.strip()
        ).str.replace(r'\s+', ' ', regex=True).str.lower()
        df_2026['Name_Key'] = name_key_2026
        df_2025['Name_Key'] = (
            df_2025[COL_2025_FULL_NAME].astype(str).str.strip().str.lower()
        )
        names_2026_set = set(df_2026['Name_Key'].unique())
        print("Comparison keys created successfully.")
        
        # --- 3. Update 'On UniOne?' Column (Omitted for brevity, UniOne logic unchanged) ---
        df_2025['UniOne_Normalized'] = df_2025[COL_2025_UNI_ONE].astype(str).str.strip().str.lower()
        is_re_registered = df_2025['Name_Key'].isin(names_2026_set)
        is_default_or_empty = df_2025['UniOne_Normalized'].isin(['nan', '', 'no'])
        df_2025.loc[is_re_registered & is_default_or_empty, COL_2025_UNI_ONE] = 'yes'
        is_empty = df_2025['UniOne_Normalized'].isin(['nan', ''])
        df_2025.loc[~is_re_registered & is_empty, COL_2025_UNI_ONE] = 'no'
        df_2025 = df_2025.drop(columns=['UniOne_Normalized'])
        print(f"'{COL_2025_UNI_ONE}' column updated (custom values preserved).")
        
        # --- 4. Perform Student Membership Verification (Priority Fix Applied) ---
        if COL_2025_VERIFICATION not in df_2025.columns:
             df_2025[COL_2025_VERIFICATION] = np.nan 

        df_2025_for_check = df_2025[df_2025[COL_2025_UNI_ONE].astype(str).str.lower() == 'yes']
        
        cols_for_merge = ['Name_Key', COL_2026_EMAIL, COL_2026_STUDENT_ID, COL_2026_USER_TYPE]
        df_2026_lookup = df_2026[cols_for_merge].drop_duplicates(subset=['Name_Key'])

        merged_df = df_2025_for_check.merge(
            df_2026_lookup, 
            on='Name_Key', 
            how='left',
            suffixes=('_2025', '_2026') 
        )
        
        merged_df['Verification_Status'] = np.nan 

        # Dynamically determine the column names
        email_col = f'{COL_2026_EMAIL}_2026' if f'{COL_2026_EMAIL}_2026' in merged_df.columns else COL_2026_EMAIL
        student_id_col = f'{COL_2026_STUDENT_ID}_2026' if f'{COL_2026_STUDENT_ID}_2026' in merged_df.columns else COL_2026_STUDENT_ID
        user_type_col = f'{COL_2026_USER_TYPE}_2026' if f'{COL_2026_USER_TYPE}_2026' in merged_df.columns else COL_2026_USER_TYPE
        
        # --- Define Condition Masks (Operating on merged_df) ---
        is_monash_student = merged_df[user_type_col].astype(str).str.strip() == STUDENT_USER_TYPE
        is_correct_student_email = merged_df[email_col].astype(str).str.lower().str.endswith(STUDENT_EMAIL_DOMAIN)
        # Check if student ID is NOT NaN AND the string length is > 0
        has_student_id = merged_df[student_id_col].notna() & (merged_df[student_id_col].astype(str).str.strip().str.len() > 0)


        # --- Case 2 (Highest Priority for NEW Status): Student ID Missing ---
        # Correct Email AND Missing Student ID
        # Since 'yes' implies a missing ID is impossible, this must be evaluated BEFORE 'yes' is applied.
        is_missing_id = is_correct_student_email & (~has_student_id)
        
        merged_df.loc[is_missing_id, 'Verification_Status'] = 'Missing Student ID'


        # --- Case 3 (Next Priority): Non-Standard Student Email ---
        # Monash Student User Type BUT Incorrect Email domain
        is_needs_validation = (~is_correct_student_email) & is_monash_student
        
        # Apply only if Verification_Status is still empty (not marked 'Missing Student ID')
        merged_df.loc[is_needs_validation & merged_df['Verification_Status'].isna(), 
                      'Verification_Status'] = 'Requires Validation'

        # --- Case 1 (Lowest Priority, Default 'Yes'): Fully Verified Student ---
        # Correct Email AND Student ID present
        is_verified = is_correct_student_email & has_student_id
        
        # Apply 'yes' only if Verification_Status is still empty
        # This is the key change: 'yes' can ONLY be applied if it hasn't failed the previous checks.
        merged_df.loc[is_verified & merged_df['Verification_Status'].isna(), 
                      'Verification_Status'] = 'yes'
        
        # --- Apply the status back to the original 2025 DataFrame ---
        status_map = merged_df.set_index('Name_Key')['Verification_Status'].dropna().to_dict()
        
        for name_key, status in status_map.items():
            rows_to_update = df_2025['Name_Key'] == name_key
            if pd.notna(status):
                df_2025.loc[rows_to_update, COL_2025_VERIFICATION] = status

        print(f"'{COL_2025_VERIFICATION}' column updated with multi-case validation.")
        
        # --- 5. Finalize and Save the Updated 2025 Sheet (Omitted for brevity) ---
        df_2025 = df_2025.drop(columns=['Name_Key'])
        df_2025.to_excel(output_file_path, index=False)

        # --- 6. PRINT CHANGE SUMMARY (Summary logic omitted for brevity, unchanged from last step) ---
        current_uni_one = df_2025[COL_2025_UNI_ONE].astype(str).str.strip().str.lower()
        uni_one_changed_to_yes_mask = (current_uni_one == 'yes') & (original_uni_one != 'yes')
        uni_one_updated_names = df_2025.loc[uni_one_changed_to_yes_mask, COL_2025_FULL_NAME].drop_duplicates().tolist()

        current_verification = df_2025[COL_2025_VERIFICATION].astype(str).str.strip()
        verification_changed_mask = (current_verification.str.lower() != original_verification)
        
        verification_updates = {}
        
        for status in current_verification.loc[verification_changed_mask].unique():
            status_title_case = status.title()
            status_mask = (current_verification == status) & (verification_changed_mask)
            names = df_2025.loc[status_mask, COL_2025_FULL_NAME].drop_duplicates().tolist()
            
            if names:
                if status_title_case in verification_updates:
                     verification_updates[status_title_case].extend(names)
                else:
                     verification_updates[status_title_case] = names

        
        print("\n" + "="*70)
        print("✅ AUDIT SUMMARY OF CHANGES")
        print("="*70)
        
        print(f"\n✨ {len(uni_one_updated_names)} Members Updated to 'yes' in '{COL_2025_UNI_ONE}':")
        if uni_one_updated_names:
            for name in uni_one_updated_names:
                print(f"  -> {name}")
        else:
            print("  (No new members were marked 'yes' in this run.)")
            
        print(f"\n✨ SUMMARY of '{COL_2025_VERIFICATION}' Updates:")
        
        total_verification_updates = sum(len(names) for names in verification_updates.values())
        
        if total_verification_updates > 0:
            for status in sorted(verification_updates.keys()):
                names = verification_updates[status]
                print(f"\n  -> {len(names)} Members marked '{status}':")
                for name in names:
                    print(f"     * {name}")
        else:
            print("  (No members had their verification status changed in this run.)")


        print("="*70)
        print(f"The updated 2025 data has been saved to: {output_file_path}")

    except FileNotFoundError as e:
        print(f"\n❌ Error: The file was not found. Please check the path.")
        print(f"Missing file: {e.filename}")
    except KeyError as e:
        print(f"\n❌ Error: Column not found in one of the spreadsheets.")
        print(f"Missing column: {e}. Please ensure headers match exactly (case-sensitive).")
    except Exception as e:
        print(f"\n❌ An unexpected error occurred: {e}")


# --- CONFIGURATION ---
FILE_2026 = "202512140830_MonashBadmintonClubMembers.xlsx"
FILE_2025 = "MONASHBADDY Member Audit Sheet.xlsx" 
OUTPUT_FILE = "MONASHBADDY_Audit_Sheet_UPDATED.xlsx"

# --- RUN THE SCRIPT ---
update_and_validate_members(FILE_2026, FILE_2025, OUTPUT_FILE)