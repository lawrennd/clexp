import os
from datetime import date
import pandas as pd
import yaml


# Map from expensify categories to Cambridge Uni Catagories
category_map = {'Car, Van and Travel Expenses: Hotel Rooms': 'Hotel',
    'Car, Van and Travel Expenses: Taxi': 'Taxi',
       'Car, Van and Travel Expenses: Train': 'Train', 
       'Car, Van and Travel Expenses: Meal (Overnight Business Trip)': 'Meal',
       'Car, Van and Travel Expenses: Air': 'Flight'}

template_file = 'template-expense-claims-partii.xlsx'


with open('_config.yaml') as file:
    config = yaml.load(file, Loader=yaml.FullLoader)


def prompt_stdin(prompt):
    """Ask user for agreement to overwrite."""
    yes = set(['yes', 'y'])
    no = set(['no','n'])

    try:
        print(prompt)
        choice = input().lower()
    # TODO would like to test for which exceptions here
    except:
        print('Stdin is not implemented.')
        print('You need to set')
        print('overide_manual_authorize=True')
        print('to proceed with the download. Please set that variable and continue.')
        raise


    if choice in yes:
        return True
    elif choice in no:
        return False
    else:
        print("Your response was a " + choice)
        print("Please respond with 'yes', 'y' or 'no', 'n'")



def excel_style(df):
    """Convert column headings and row numbers of sheet to excel style"""
    letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
    col_names = []
    for i in range(len(df.columns)):
        col = i+1
        result = ''
        while col:
            col, rem = divmod(col-1, 26)
            result = letters[rem] + result
        col_names.append(result)
    df.columns=col_names
    df.index==range(1, len(df.index)+1)

def write_claim(data_df, template_df):
    """Write a claim spreadsheet using given data and template"""
    
    if len(data_df)>12:
        raise RuntimeError("Error, can only place up to twelve expense items in one sheet.")

    report_title = pd.unique(data_df.report_title)[0]
    report_id = pd.unique(data_df.report_id)[0]
    report_date = pd.unique(data_df.report_enddate)[0]
    print('Writing Report ID:', report_id, 'titled:', report_title, 'for', report_date)
    
    store_df = template_df.copy()
    store_df['D'][1] = config['title'] + ' ' + config['full_name']
    store_df['D'][3] = 'Report {report_id}. {report_title}'.format(report_id=report_id, report_title=report_title)
    
    start_row = 8
    for i, idx in enumerate(data_df.index):
        category = data_df['category'][idx]
        description=data_df['description'][idx]
        if category in category_map:
            reason = category_map[category]
        else:
            reason = category
        store_df['A'][start_row+i] = data_df['date'][idx]
        if reason in ['Taxi', 'Train', 'Flight']:
            from_val = ''
            to_val = ''
            m = re.match(r"From: ([^;]+); To: ([^;]+)", description)
            m2 = re.match(r"From: ([^;]+); To: ([^;]+); (.*)", description)
            if m2:
                from_val = m2.group(1) # From
                to_val = m2.group(2) # To
                description = m2.group(3) # description
        
            elif m:
                from_val = m.group(1) # From
                to_val = m.group(2) # To
                description = ''
            store_df['B'][start_row+i] = from_val
            store_df['C'][start_row+i] = to_val
        if str(description) == 'nan':
            description = ''
        
        store_df['D'][start_row+i] = '{reason}: Vendor: {merchant}.{description}'.format(
                                                    merchant=data_df['merchant'][idx],
                                                    reason=reason,
                                                    description=description)
 
                

        original_currency = data_df['original_currency'][idx]
        if original_currency != 'GBP':
            store_df['E'][start_row+i] = original_currency
            store_df['F'][start_row+i] = data_df['original_amount'][idx]
            store_df['G'][start_row+i] = data_df['exchange_rate_used'][idx]
        store_df['H'][start_row+i] = data_df['amount'][idx]
        
    store_df['D'][21] = 'RG ' + config['grant_code']
    store_df['H'][20] = data_df['amount'].sum()
    store_df['D'][29] = config['name']
    store_df['F'][29] = date.today().strftime("%Y-%m-%d")
    store_df['D'][30] = config['name']
    store_df['F'][30] = date.today().strftime("%Y-%m-%d")
    store_df['F'][27] = config['four_digits']
    store_df['C'][27] = config['payroll']
    date.today().strftime("%Y-%m-%d")
    report_filename = '{report_date}-claim-{report_id}-{report_title}.xlsx'.format(report_date=report_date,
                                                                                   report_id = report_id,
                                                                                   report_title=report_title.lower().replace(' ', '-'))
    filename = os.path.join(expense_directory, report_filename)
    from shutil import copyfile
    write_file = True
    if os.path.exists(filename):
        write_file = prompt_stdin('The file {filename} exists, overwrite (y/n)?'.format(filename=filename))
    if write_file:
        copyfile(os.path.join(expense_directory, template_file), filename)
        writer = pd.ExcelWriter(filename,
                               engine='xlsxwriter')
        store_df.to_excel(writer, sheet_name='Expenses claim', index=False,header=False)
        writer.save()
    else:
        print("Skipping {filename} as it already exists.".format(filename=filename))
    return store_df
