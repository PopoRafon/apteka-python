"""
NAME
    customer

DESCRIPTION
    This module allows customers to be created, updated, deleted from database
    and lets them buy drugs from pharmacy.

    This script requires pandas and openpyxl to be installed within the Python
    environment you are running this script in.

FUNCTIONS
    This module contains the following functions:
    * init_customer_database()

    * add_customer(name, email, phone, street, city, country)

    * remove_customer(id, name)

    * customer_buy_drug(customer_id, drug_id, drug_name, prescription)

EXAMPLES
    add_customer(name="Kamil Kowal", email="Youdeen87@gmail.com", phone="606787997", street="Braci Jeziorowskich", city="Plock", country="Poland")
    remove_customer(name="Kamil Kowal")
    customer_buy_drug(customer_id=5, drug_id=2)
"""

import csv, random, os
import pandas as pd
from datetime import date

DATABASE_DIR_PATH = os.path.join('database')
CUSTOMER_FILE_PATH = os.path.join(DATABASE_DIR_PATH, 'customer.csv')
ADDRESS_FILE_PATH = os.path.join(DATABASE_DIR_PATH, 'address.csv')
DRUG_FILE_PATH = os.path.join(DATABASE_DIR_PATH, 'drugs.xlsx')

CUSTOMER_FIELDNAMES = ['ID', 'NAME', 'E-MAIL', 'PHONE', 'CREATED', 'UPDATED']
ADDRESS_FIELDNAMES = ['ID', 'STREET', 'CITY', 'COUNTRY']

def init_customer_database() -> None:
    """
    Creates database directory, customer.csv and address.csv files if they don't exist.

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

def add_customer(name: str, email: str = '', phone: str = '', street: str = '', city: str = '', country: str = '') -> None:
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
        None
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
            print("Maximum number of ids exceeded.")
            return

        while id in file_ids:
            id = random.randint(ids_range[0], ids_range[1])

    with open(CUSTOMER_FILE_PATH, mode='a+') as csv_file:
        csv_writer = csv.DictWriter(csv_file, CUSTOMER_FIELDNAMES, lineterminator='\r')
        csv_writer.writerow({'ID': id, 'NAME': name, 'E-MAIL': email, 'PHONE': phone, 'CREATED': created, 'UPDATED': updated})

    with open(ADDRESS_FILE_PATH, mode='a+') as csv_file:
        csv_writer = csv.DictWriter(csv_file, ADDRESS_FIELDNAMES, lineterminator='\r')
        csv_writer.writerow({'ID': id, 'STREET': street, 'CITY': city, 'COUNTRY': country})

    open(os.path.join(DATABASE_DIR_PATH, f'{id}.txt'), mode='w').close()

def remove_customer(id: int = 0, name: str = '') -> None:
    """
    Removes customer data from customer.csv and address.csv files and removes his e-book ".txt" file.

    Args:
        id (int): id of customer to be removed
        name (str): name of customer to be removed

    Returns:
        None
    """
    if name and not id:
        with open(CUSTOMER_FILE_PATH, mode='r') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            for row in csv_reader:
                if row['NAME'] == name:
                    id = int(row['ID'])
                    break

    if id:
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
    else:
        print('Customer not found.')

def customer_buy_drug(customer_id: int, drug_id: int = None, drug_name: str = '', prescription: str = '') -> None:
    """
    Adds drug to customer e-book ".txt" file and removes amount customer bought from "drugs.xlsx" file.

    Args:
        customer_id (int): id of customer who buys drug
        drug_id (int): id of drug to be bought
        drug_name (str): name of drug to be bought
        prescription (str): optional prescription number for drug

    Returns:
        None
    """
    df = pd.read_excel(DRUG_FILE_PATH)
    drug = df['ID'] == drug_id if drug_id else df['DRUG'].str.upper() == drug_name.upper()

    if not drug.any():
        print('Drug not found.')
        return

    index = df[drug].index[0]
    drug_amount = df.at[index, 'NO_PACKAGES_AVAILABLE']
 
    if drug_amount <= 0:
        print('No drugs available.')
        return

    if df.at[index, 'ON_RECEPT'] == 'YES' and not prescription:
        print('This drug needs prescription.')
        return

    df.at[index, 'NO_PACKAGES_AVAILABLE'] = drug_amount - 1
    df.to_excel(DRUG_FILE_PATH, index=False)

    with open(os.path.join(DATABASE_DIR_PATH, f'{customer_id}.txt'), mode='a') as customer_text_file:
        customer_text_file.write(f'drug: {df.at[index, 'DRUG'].upper()}, amount: {1}, prescription: {prescription}\n')

init_customer_database()
