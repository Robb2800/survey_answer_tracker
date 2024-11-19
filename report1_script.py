# %%
import pandas as pd
import os
from datetime import datetime
import shutil

# %%
def load_and_combine_questionnaires(folder_path, questionnaire_files):
    questionnaires = []
    for survey in questionnaire_files:
        file_path = os.path.join(folder_path, survey['file'])
        df = pd.read_excel(file_path)
        print(f"Loaded {survey['file']} with shape {df.shape}")
        df['survey_type'] = survey['survey_name']
        df['CSC'] = survey['CSC']
        questionnaires.append(df)
    
    combined_df = pd.concat(questionnaires, ignore_index=True)
    print(f"Combined DataFrame shape: {combined_df.shape}")
    return combined_df


# %%
def create_concat_column(df, new_column_name):

    df[new_column_name] = df[['survey_type', 'CSC','Id']].astype(str).agg('_'.join, axis=1)
    return df

# %%
def create_dir_n_move():
    
    
    current = os.getcwd()
    print('Current Folder:', current)

    
    archive_folder = os.path.join(current, '_Archive')

    if not os.path.exists(archive_folder):
        os.mkdir(archive_folder)
        print(f'Created Archive folder: {archive_folder}')
    else:
        print(f'Archive folder already exists: {archive_folder}')

    
    dtm = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    new_folder = os.path.join(archive_folder, dtm)
    os.mkdir(new_folder) 
    print(f'Creating New Folder: {new_folder}')

    
    for file_name in os.listdir(current):
        file_path = os.path.join(current, file_name)
        
        if os.path.isfile(file_path) and file_name.startswith('CS Survey 2024'):  
            shutil.copy(file_path, new_folder)
            print(f'Copied: {file_name} to {new_folder}')

    return new_folder

# %%
def save_combined_df_to_result_folder(folder_path, combined_df):
    
    result_folder = os.path.join(folder_path, 'Result')

    if not os.path.exists(result_folder):
        os.mkdir(result_folder)
        print(f"Created folder: {result_folder}")
    else:
        print(f"Folder already exists: {result_folder}")

    
    dtm = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_filename = f'survey_update_report_{dtm}.xlsx'

    
    output_file_combined = os.path.join(result_folder, output_filename)
    combined_df.to_excel(output_file_combined, index=False)
    print(f"Combined survey data saved to {output_file_combined}")

# %%
def main():
    
    
    folder_path = os.getcwd()
    
    questionnaire_files = [
        {'file': 'CS Survey 2024 - Cloud Services_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Cloud', 'CSC':'0'},
        {'file': 'CS Survey 2024 - Colocation_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Colocation', 'CSC':'0'},
        {'file': 'CS Survey 2024 - DC & Service Management_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'DC + Service Management','CSC':'1'},
        {'file': 'CS Survey 2024 - DC & SM (Non CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'DC + Service Management', 'CSC':'0'},
        {'file': 'CS Survey 2024 - Desktop Support (No CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Desktop Support', 'CSC':'0'},
        {'file': 'CS Survey 2024 - Desktop Support _PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Desktop Support', 'CSC':'1'},
        {'file': 'CS Survey 2024 - IT Security (No CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'IT Security','CSC':'0'},
        {'file': 'CS Survey 2024 - IT Security_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'IT Security','CSC':'1'},
        {'file': 'CS Survey 2024 - Mail Hosting_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Mail Hosting', 'CSC':'0'},
        {'file': 'CS Survey 2024 - Maintenance Supp (No CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Maintenance Support','CSC':'0'},
        {'file': 'CS Survey 2024 - Maintenance Support (CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Maintenance Support','CSC':'1'},
        {'file': 'CS Survey 2024 - Managed Service (No CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Managed Services','CSC':'0'},
        {'file': 'CS Survey 2024 - Managed Service_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Managed Services','CSC':'1'},
        {'file': 'CS Survey 2024 - Resource Based & CSC_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Resource Fulfillment', 'CSC':'1'},
        {'file': 'CS Survey 2024 - Resource Based Operation_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Resource Fulfillment', 'CSC':'0'},
        {'file': 'CS Survey 2024 - Seat Management (No CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Seat Management', 'CSC':'0'},
        {'file': 'CS Survey 2024 - Seat Management Service _PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Seat Management', 'CSC':'1'},
        {'file': 'CS Survey 2024 - Voucher Based (No CSC)_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Voucher Based','CSC':'0'},
        {'file': 'CS Survey 2024 - Voucher Based Operation_PT Astra Graphia Information Technology (AGIT).xlsx', 'survey_name': 'Voucher Based','CSC':'1'}

    ]

    
    selected_columns = ['survey_type','CSC','Id','Nama Responden','Company',
        ]
    
    
    combined_df = load_and_combine_questionnaires(folder_path, questionnaire_files)
    combined_df = combined_df[selected_columns]
    
    
    combined_df = create_concat_column(combined_df,'Concated')
    
    output_file_combined=save_combined_df_to_result_folder(folder_path,combined_df)
    


    print(f"Combined survey data saved to {output_file_combined}")
    create_dir_n_move()

   

if __name__ == "__main__":
    main()


