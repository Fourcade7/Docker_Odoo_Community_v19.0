from . import models

import os
import pandas as pd
from odoo import models, api


def post_init_hook(env):
    import_bank_names(env)


def import_bank_names(env):
    print('Entered The IMPORT FUNCTION')
    module_path = os.path.dirname(os.path.abspath(__file__))  # Get the current module's path
    excel_folder_path = os.path.join(module_path, 'data', 'excel_files')
    excel_files = [
        'Nbfc_Companies_1-232.xlsx',
        'Nbfc_Companies_233-464.xlsx',
        'Nbfc_Companies_465-end.xlsx'
    ]
    print('FOUND FILES')

    for ex_file in excel_files:
        print('STARTED ROLLING')
        try:
            excel_file_path = os.path.join(excel_folder_path, ex_file)
            bank_names = read_bank_names_from_excel(excel_file_path)
            insert_bank_names_to_res_bank(env, bank_names)
            print('ROLLING OUT --------')
        except Exception as e:
            print(f"Error processing file {ex_file}: {e}")


def read_bank_names_from_excel(excel_file):
    try:
        df = pd.read_excel(excel_file)
        return df.iloc[:, 2].tolist()  # Assuming bank names are in the third column (index 2)
    except Exception as e:
        print(f"Error reading Excel file {excel_file}: {e}")
        return []


def insert_bank_names_to_res_bank(env, bank_names):
    try:
        for bank_name in bank_names:
            env['res.bank'].create({'name': bank_name})
            print(f"Inserted bank name: {bank_name}")
    except Exception as e:
        print(f"Error inserting bank name {bank_name}: {e}")