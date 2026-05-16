import csv, random, os
from datetime import date

database_dir_path = 'database'
customer_file_path = os.path.join(database_dir_path, 'customer.csv')
address_file_path = os.path.join(database_dir_path, 'address.csv')

customer_fieldnames = ['ID', 'NAME', 'E-MAIL', 'PHONE', 'CREATED', 'UPDATED']
address_fieldnames = ['ID', 'STREET', 'CITY', 'COUNTRY']

def init_customer_database():
    """
    Creates database directory, customer.csv and address.csv files if they don't exist.
    """
    if not os.path.isdir(database_dir_path):
        os.makedirs(database_dir_path)

    if not os.path.isfile(customer_file_path):
        with open(customer_file_path, mode='w') as csv_file:
            writer = csv.DictWriter(csv_file, customer_fieldnames, lineterminator='\r')
            writer.writeheader()

    if not os.path.isfile(address_file_path):
        with open(address_file_path, mode='w') as csv_file:
            writer = csv.DictWriter(csv_file, address_fieldnames, lineterminator='\r')
            writer.writeheader()

def add_customer():
    """
    Adds new customer data to customer.csv and address.csv files. Creates customer e-book ".txt" file with his id.
    """
    id = random.randint(1000, 9999)
    name = ''
    email = ''
    phone = ''
    street = ''
    city = ''
    country = ''
    created = date.today()
    updated = date.today()

    with open(customer_file_path, mode='r') as csv_file:
        csv_reader = csv.DictReader(csv_file)
        file_ids = set()

        for row in csv_reader:
            file_ids.add(int(row['ID']))

        if len(file_ids) >= (9999 - 1000):
            print("Maximum number of ids exceeded.")
            return

        while id in file_ids:
            id = random.randint(1000, 9999)

    with open(customer_file_path, mode='a+') as csv_file:
        csv_writer = csv.DictWriter(csv_file, customer_fieldnames, lineterminator='\r')
        csv_writer.writerow({'ID': id, 'NAME': name, 'E-MAIL': email, 'PHONE': phone, 'CREATED': created, 'UPDATED': updated})

    with open(address_file_path, mode='a+') as csv_file:
        csv_writer = csv.DictWriter(csv_file, address_fieldnames, lineterminator='\r')
        csv_writer.writerow({'ID': id, 'STREET': street, 'CITY': city, 'COUNTRY': country})

    open(os.path.join(database_dir_path, f'{id}.txt'), mode='w').close()

def remove_customer():
    """
    Removes customer data from customer.csv and address.csv files. Also removes customer e-book ".txt" file.
    """
    id = 0
    name = ''

    if name and not id:
        with open(customer_file_path, mode='r') as csv_file:
            csv_reader = csv.DictReader(csv_file)

            for row in csv_reader:
                if row['NAME'] == name:
                    id = int(row['ID'])

    if id:
        with open(customer_file_path, mode='r') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            rows = [row for row in csv_reader if int(row['ID']) != id]

        with open(customer_file_path, mode='w') as csv_file:
            csv_writer = csv.DictWriter(csv_file, fieldnames=customer_fieldnames, lineterminator='\r')
            csv_writer.writeheader()
            csv_writer.writerows(rows)

        with open(address_file_path, mode='r') as csv_file:
            csv_reader = csv.DictReader(csv_file)
            rows = [row for row in csv_reader if int(row['ID']) != id]

        with open(address_file_path, mode='w') as csv_file:
            csv_writer = csv.DictWriter(csv_file, fieldnames=address_fieldnames, lineterminator='\r')
            csv_writer.writeheader()
            csv_writer.writerows(rows)

        customer_database_file_path = os.path.join(database_dir_path, f'{id}.txt')

        if os.path.isfile(customer_database_file_path):
            os.remove(customer_database_file_path)


init_customer_database()
