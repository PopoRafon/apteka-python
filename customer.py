"""
NAME
    customer

DESCRIPTION
    This module allows customers to be created, updated, deleted from database
    and lets them buy drugs from pharmacy. Before any function can be used it must be
    intialize with "init_database" function.

    This script requires pandas and openpyxl to be installed within the Python
    environment you are running this script in.

FUNCTIONS
    This module contains the following functions:
    * init_database()

    * add_customer(name, email, phone, street, city, country) - returns None if adding customer goes successfully otherwise string with error message

    * remove_customer(id, name) - returns None if removing customer goes successfully otherwise string with error message

    * customer_buy_drug(customer_id, amount, drug_id, drug_name, prescription) - returns None if buying drug goes successfully otherwise string with error message

    * get_customer(id) - returns dictionary of customer data if customer is found else returns None

    * update_customer(id, name, email, phone, street, city, country) - returns None if updating customer goes successfully otherwise string with error message

    * multi_purchase(func)

EXAMPLES
    add_customer(name="Kamil Kowal", email="Youdeen87@gmail.com", phone="606787997", street="Braci Jeziorowskich", city="Plock", country="Poland")

    remove_customer(name="Kamil Kowal")

    customer_buy_drug(5, [{'amount': 2, 'drug_id': 2}, {'amount': 5, 'drug_name': 'amoTax', 'prescription': '123'}])

    get_customer(id=4)

    update_customer(id=7, email='test@gmail.com', city='Bialystok', country='Poland')
"""

import csv, random, os
import pandas as pd
from datetime import date
from const import *

def init_database() -> None:
    """
    Creates database directory, customer.csv, address.csv and drug.xlsx files if they don't exist.

    Returns:
        None
    """
    if not os.path.isdir(DATABASE_DIR_PATH):
        os.makedirs(DATABASE_DIR_PATH)

    if not os.path.isfile(CUSTOMER_FILE_PATH):
        with open(CUSTOMER_FILE_PATH, mode='w') as csv_file:
            writer = csv.DictWriter(csv_file, CUSTOMER_FIELDNAMES, lineterminator='\r')
            writer.writeheader()

    if not os.path.isfile(ADDRESS_FILE_PATH):
        with open(ADDRESS_FILE_PATH, mode='w') as csv_file:
            writer = csv.DictWriter(csv_file, ADDRESS_FIELDNAMES, lineterminator='\r')
            writer.writeheader()

    if not os.path.isfile(DRUGS_FILE_PATH):
        df = pd.DataFrame({fieldname: [] for fieldname in DRUGS_FIELDNAMES})
        df.to_excel(DRUGS_FILE_PATH, index=False)

def add_customer(name: str, email: str, phone: str, street: str, city: str, country: str) -> str | None:
    """
    Creates random id for new customer. Adds his data to customer.csv and address.csv files. Creates customer e-book ".txt" file with his id.

    Args:
        name (str): new customer name and surname
        email (str): new customer email
        phone (str): new customer phone number
        street (str): new customer street address
        city (str): new customer city
        country (str): new customer country

    Returns:
        message (str | None): None if operation goes successfully otherwise string with error message
    """
    ids_range = [1000, 9999]
    id = random.randint(ids_range[0], ids_range[1])
    created = date.today()
    updated = date.today()

    with open(CUSTOMER_FILE_PATH, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        file_ids = set()

        for row in csv_reader:
            file_ids.add(int(row['ID']))

        if len(file_ids) >= (ids_range[1] - ids_range[0]):
            return 'Maksymalna ilość id została przekroczona.'

        while id in file_ids:
            id = random.randint(ids_range[0], ids_range[1])

    with open(CUSTOMER_FILE_PATH, mode='a+', encoding='utf-8') as csv_file:
        csv_writer = csv.DictWriter(csv_file, CUSTOMER_FIELDNAMES, lineterminator='\r')
        csv_writer.writerow({'ID': id, 'NAME': name, 'E-MAIL': email, 'PHONE': phone, 'CREATED': created, 'UPDATED': updated})

    with open(ADDRESS_FILE_PATH, mode='a+', encoding='utf-8') as csv_file:
        csv_writer = csv.DictWriter(csv_file, ADDRESS_FIELDNAMES, lineterminator='\r')
        csv_writer.writerow({'ID': id, 'STREET': street, 'CITY': city, 'COUNTRY': country})

    open(os.path.join(DATABASE_DIR_PATH, f'{id}.txt'), mode='w').close()

    return None

def remove_customer(id: int | None = None, name: str | None = None) -> str | None:
    """
    Removes customer data from customer.csv and address.csv files and removes his e-book ".txt" file.

    Args:
        id (int | None): id of customer to be removed
        name (str | None): name of customer to be removed

    Returns:
        message (str | None): None if operation goes successfully otherwise string with error message
    """
    is_customer_found = False

    with open(CUSTOMER_FILE_PATH, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)

        for row in csv_reader:
            if int(row['ID']) == id or row['NAME'] == name:
                id = int(row['ID'])
                is_customer_found = True
                break

    if not is_customer_found:
        return 'Pacjent z tym id nie istnieje.'

    with open(CUSTOMER_FILE_PATH, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        rows = [row for row in csv_reader if int(row['ID']) != id]

    with open(CUSTOMER_FILE_PATH, mode='w') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=CUSTOMER_FIELDNAMES, lineterminator='\r')
        csv_writer.writeheader()
        csv_writer.writerows(rows)

    with open(ADDRESS_FILE_PATH, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        rows = [row for row in csv_reader if int(row['ID']) != id]

    with open(ADDRESS_FILE_PATH, mode='w') as csv_file:
        csv_writer = csv.DictWriter(csv_file, fieldnames=ADDRESS_FIELDNAMES, lineterminator='\r')
        csv_writer.writeheader()
        csv_writer.writerows(rows)

    customer_database_file_path = os.path.join(DATABASE_DIR_PATH, f'{id}.txt')

    if os.path.isfile(customer_database_file_path):
        os.remove(customer_database_file_path)

    return None

def update_customer(
        id: int,
        name: str | None = None,
        email: str | None = None,
        phone: str | None = None,
        street: str | None = None,
        city: str | None = None ,
        country: str | None = None
    ) -> str | None:
    """
    Updates customer data to customer.csv and address.csv files based on given id.

    Args:
        id (int): id of customer to be updated
        name (str | None): customer new name and surname
        email (str | None): customer new email
        phone (str | None): customer new phone number
        street (str | None): customer new street address
        city (str | None): customer new city
        country (str | None): customer new country

    Returns:
        message (str | None): None if operation goes successfully otherwise string with error message
    """
    updated = date.today()
    is_customer_found = False

    with open(CUSTOMER_FILE_PATH, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        rows = []

        for row in csv_reader:
            if int(row['ID']) == id:
                row['NAME'] = row['NAME'] if name is None else name
                row['E-MAIL'] = row['E-MAIL'] if email is None else email
                row['PHONE'] = row['PHONE'] if phone is None else phone
                row['UPDATED'] = updated
                is_customer_found = True

            rows.append(row)

    if not is_customer_found:
        return 'Pacjent z tym id nie istnieje.'

    with open(CUSTOMER_FILE_PATH, mode='w') as csv_file:
        csv_writer = csv.DictWriter(csv_file, CUSTOMER_FIELDNAMES, lineterminator='\r')
        csv_writer.writeheader()
        csv_writer.writerows(rows)

    with open(ADDRESS_FILE_PATH, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        rows = []

        for row in csv_reader:
            if int(row['ID']) == id:
                row['STREET'] = row['STREET'] if street is None else street
                row['CITY'] = row['CITY'] if city is None else city
                row['COUNTRY'] = row['COUNTRY'] if country is None else country

            rows.append(row)

    with open(ADDRESS_FILE_PATH, mode='w') as csv_file:
        csv_writer = csv.DictWriter(csv_file, ADDRESS_FIELDNAMES, lineterminator='\r')
        csv_writer.writeheader()
        csv_writer.writerows(rows)

    return None

def get_customer(id: int) -> dict[str, str] | None:
    """
    Gets and returns customer data from customer.csv and address.csv files based on given id.

    Args:
        id (int): id of customer to be searched

    Returns:
        customer_data (dict[str, str] | None): dictionary containing all customer data, None if no customer is found
    """
    customer_data = {}

    with open(CUSTOMER_FILE_PATH, 'r') as csv_file:
        csv_reader = csv.DictReader(csv_file)

        for row in csv_reader:
            if int(row['ID']) == id:
                customer_data.update(**row)
                break

    with open(ADDRESS_FILE_PATH, 'r') as csv_file:
        csv_reader = csv.DictReader(csv_file)

        for row in csv_reader:
            if int(row['ID']) == id:
                customer_data.update(**row)
                break

    if len(customer_data.keys()) == 0:
        return None

    return customer_data

def multi_purchase(func):
    """
    Decorator for customer_buy_drug function which allows customer to buy multiple drugs at once.
    """
    def inner(customer_id: int, purchases: list[dict[str, str | int]]) -> list[str] | None:
        messages = []

        for purchase in purchases:
            message = func(
                customer_id,
                amount=purchase['amount'],
                drug_id=purchase.get('drug_id'),
                drug_name=purchase.get('drug_name'),
                prescription=purchase.get('prescription'),
            )

            if not message is None:
                messages.append(message)
        
        return messages if len(messages) > 0 else None
    return inner

@multi_purchase
def customer_buy_drug(
        customer_id: int,
        amount: int,
        drug_id: int | None = None,
        drug_name: str | None = None,
        prescription: str | None = None
    ) -> str | None:
    """
    Adds drug to customer e-book ".txt" file and removes amount customer bought from "drugs.xlsx" file.

    Args:
        customer_id (int): id of customer who buys drug
        amount (int): amount of drugs to be bought
        drug_id (int | None): id of drug to be bought
        drug_name (str | None): name of drug to be bought
        prescription (str | None): optional prescription number for drug

    Returns:
        message (str | None): None if operation goes successfully otherwise string with error message
    """
    if not os.path.isfile(os.path.join(DATABASE_DIR_PATH, f'{customer_id}.txt')):
        return 'Pacjent z tym id nie istnieje.'

    df = pd.read_excel(DRUGS_FILE_PATH)
    drug = None

    if not drug_id is None:
        drug = df['ID'] == drug_id
    elif not drug_name is None:
        drug = df['DRUG'].str.upper() == drug_name.upper()

    if drug is None or not drug.any():
        return 'Lek z tym id nie istnieje.'

    index = df[drug].index[0]
    available_drug_amount = df.at[index, 'NO_PACKAGES_AVAILABLE']
 
    if available_drug_amount < amount:
        return 'Niewystarczająca ilość leku w sklepie.'

    if df.at[index, 'ON_RECEPT'] == 'YES' and not prescription == 'YES':
        return 'Ten lek wymaga recepty.'

    df.at[index, 'NO_PACKAGES_AVAILABLE'] = available_drug_amount - amount
    df.to_excel(DRUGS_FILE_PATH, index=False)

    with open(os.path.join(DATABASE_DIR_PATH, f'{customer_id}.txt'), mode='a') as customer_text_file:
        drug_name = df.at[index, 'DRUG'].upper()
        customer_text_file.write(f'drug: {drug_name}, amount: {amount}, prescription: {prescription}\n')

    return None
