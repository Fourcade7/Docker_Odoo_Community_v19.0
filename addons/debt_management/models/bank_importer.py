import os
from odoo import models, api
import pandas as pd
import xmlrpc.client

class BankImport(models.Model):
    _name = 'bank.import'
    _description = 'Bank Import'

    @api.model
    def import_bank_names(self):
        # Path to the folder where your Excel files are stored
        module_path = os.path.dirname(os.path.abspath(__file__))  # Get the current module's path
        excel_folder_path = os.path.join(module_path, 'data', 'excel_files')

        # List of Excel files
        excel_files = [
            'Nbfc_Companies_1-232.xlsx',
            'Nbfc_Companies_233-464.xlsx',
            'Nbfc_Companies_465-end.xlsx'
        ]

        # Automatically use the current Odoo instance configuration
        odoo_url = self.env['ir.config_parameter'].sudo().get_param('web.base.url')
        db_name = self.env.cr.dbname  # Get the current database name
        username = self.env.user.login  # Get the current logged-in user's username
        password = self.env.user._password  # Get the user's password

        for ex_file in excel_files:
            excel_file_path = os.path.join(excel_folder_path, ex_file)
            bank_names = self.read_bank_names_from_excel(excel_file_path)

            # Connect to Odoo using the automatically detected parameters
            models, uid = self.connect_to_odoo(odoo_url, db_name, username, password)

            # Insert bank names into res.bank model
            self.insert_bank_names_to_res_bank(models, uid, db_name, bank_names, password)

    def read_bank_names_from_excel(self, excel_file):
        # Read the Excel file
        df = pd.read_excel(excel_file)
        # Assuming the bank name is in the third column (index 2)
        return df.iloc[:, 2].tolist()

    def connect_to_odoo(self, odoo_url, db_name, username, password):
        # Set up the URL for the XML-RPC connection
        url = odoo_url
        common = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/common')

        # Authenticate the user
        uid = common.authenticate(db_name, username, password, {})

        # Set up the object proxy to interact with the models
        models = xmlrpc.client.ServerProxy(f'{url}/xmlrpc/2/object')

        return models, uid

    def insert_bank_names_to_res_bank(self, models, uid, db_name, bank_names, password):
        for bank_name in bank_names:
            # Insert the bank name as the bank name in res.bank
            model_name = 'res.bank'
            bank_data = {
                'name': bank_name,  # Insert the bank name as the bank name or another field
            }

            # Create a new record in the res.bank model
            models.execute_kw(
                db_name, uid, password,
                model_name, 'create',
                [bank_data]
            )
            print(f"Inserted bank name: {bank_name}")
