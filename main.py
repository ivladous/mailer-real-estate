import json
from datetime import datetime

from selenium import webdriver
from selenium.common import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string
import time

START_ROW = 2
END_ROW = 70

FULL_NAME_COLUMN = 'G'
PHONE_NUMBER_COLUMNS = ['H', 'J', 'K']
NUMBER_BEDROOMS_COLUMN = 'P'
MESSAGES_LOGS_COLUMN = 'Q'
# BUILDING_NAME_COLUMN = 'D'

receiver_col_index = column_index_from_string(FULL_NAME_COLUMN)
messages_logs_col_ind = column_index_from_string(MESSAGES_LOGS_COLUMN)
# building_name_col = column_index_from_string(BUILDING_NAME_COLUMN)
number_bedroom_col = column_index_from_string(NUMBER_BEDROOMS_COLUMN)

success_list: list = []
fail_list: list = []


def name_convertor(full_name: str) -> str:
    first_name = full_name.split()[0]
    return first_name.lower().capitalize()


# Load the Excel sheet containing the receivers' phone numbers
wb = load_workbook('boulevard_p_corr.xlsx')
ws = wb.active

# Set up the Chrome driver and navigate to WhatsApp Web
driver = webdriver.Chrome()
driver.get('https://web.whatsapp.com/')

# Wait for the user to log in to WhatsApp Web
input('Please log in to WhatsApp Web, then press Enter to continue...')

# Loop through each receiver's phone number and send a message
for row_index, row in enumerate(ws.iter_rows(min_row=START_ROW, max_row=END_ROW, values_only=True), start=2):
    print(row_index, row)
    logs_cell = ws.cell(row=row_index, column=messages_logs_col_ind)
    for phone_col_letter in PHONE_NUMBER_COLUMNS:

        with open('success_list_bp0_70.json', 'w') as file:
            # Write the list to the file in JSON format
            json.dump(success_list, file)

        with open('failed_list_bp0_70.json', 'w') as file:
            # Write the list to the file in JSON format
            json.dump(fail_list, file)

        wb.save('bp0_70_result.xlsx')

        phone_col_index = column_index_from_string(phone_col_letter)

        # Get the receiver's phone number and message from the Excel sheet
        phone_number = str(row[phone_col_index-1])

        if phone_number == '0':
            continue
        phone_number = '+' + phone_number
        if phone_number in success_list:
            continue
        if phone_number in fail_list:
            continue

        receiver = name_convertor(str(row[receiver_col_index-1]))
        number_bedroom = row[number_bedroom_col-1].split()[0]

        # Navigate to the chat with the receiver
        driver.get(f'https://web.whatsapp.com/send?phone={phone_number}')

        # Wait for the chat to load
        time.sleep(10)

        message = (
            f'''Dear {receiver},
            This is Vlad, I am a downtown specialist. Just want to check if you are still the owner of {number_bedroom} bedroom apartment in Boulevard Point and I was wondering if you might be interested to sell or lease it. I am dealing with many investors that are interested in your property. Feel free to contact me for any inquiries.
            Thank you & best regards,
            Vlad'''
        )
        # Type the message and send it

        try:
            input_box = driver.find_element(By.XPATH, '//div[@class="_3Uu1_"]')
            input_box.send_keys(message)
            input_box.send_keys(Keys.RETURN)
        except NoSuchElementException:
            print(f'Failed send to {phone_number} on Whatsapp;\n')
            fail_list.append(phone_number)
            if logs_cell.value is None:
                logs_cell.value = f'Failed send to {phone_number} on Whatsapp;\n'
            else:
                logs_cell.value += f'Failed send to {phone_number} on Whatsapp;\n'
            continue

        # Wait for the message to be sent
        time.sleep(10)
        success_list.append(phone_number)
        print(f'Message to {phone_number} sent {str(datetime.now())}!\n')
        if logs_cell.value is None:
            logs_cell.value = f'Message to {phone_number} sent {str(datetime.now())}!;\n'
        else:
            logs_cell.value += f'Message to {phone_number} sent {str(datetime.now())}!;\n'


# Close the Chrome driver
driver.quit()

print('All messages sent!')
