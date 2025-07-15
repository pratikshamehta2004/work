import pandas as pd
import random
from faker import Faker

# Initialize Faker and column headers
fake = Faker()
columns = [
    "ISIN No", "Security Name", "Client Name", "PAN", "Bank A/C", "Bank Name", "IFSC",
    "Holding", "Face Value", "Amount", "From", "To", "No of days", "ROI",
    "Interest Payable", "TDS", "Net Interest", "Principal Repayment", "Total"
]

# Define number of rows for each sheet
sheet_info = {
    "Sept_2024": 11, "Oct_2024": 48, "Nov_2024": 55, "Dec_2024_1": 55, "Dec_2024_2": 57, "Dec_2024": 59,
    "Jan_2025_1": 59, "Jan_2025_2": 65, "Jan_2025_3": 66, "Jan_2025": 66, "Feb_2025_1": 66, "Feb_2025": 66,
    "March_2025_1": 67, "March_2025": 65, "April_2025_1": 68, "April_2025": 70, "May_2025_1": 72,
    "May_2025": 73, "June_2025": 82
}

# Shared ISIN and Security Name
shared_isin = f"INE{fake.random_uppercase_letter()}{fake.random_uppercase_letter()}{random.randint(10000,99999)}"
shared_security_name = fake.company()

# Random data sources
bank_names = [fake.company() + " Bank" for _ in range(20)]
ifsc_list = [f"{fake.random_uppercase_letter() * 4}{random.randint(100000,999999)}" for _ in range(50)]
face_values = [100000, 500000, 1000000]
rois = ["12%", "14%", "15%", "16%", "18%", "20%"]

# Generate enough unique client names
unique_clients = list(set([fake.name() for _ in range(1000)]))  # More clients to avoid duplication

# PAN Generator: 5 uppercase letters + 4 digits + 1 uppercase letter
def generate_pan():
    letters = ''.join(random.choices('ABCDEFGHIJKLMNOPQRSTUVWXYZ', k=5))
    digits = ''.join(random.choices('0123456789', k=4))
    last = random.choice('ABCDEFGHIJKLMNOPQRSTUVWXYZ')
    return letters + digits + last

# Function to generate one row of financial data
def generate_random_row(client_name):
    pan = generate_pan()
    bank_ac = fake.bban()
    bank = random.choice(bank_names)
    ifsc = random.choice(ifsc_list)
    holding = random.randint(100, 500)
    face_value = random.choice(face_values)
    amount = holding * face_value
    from_date = fake.date_between(start_date='-1y', end_date='today')
    to_date = fake.date_between(start_date=from_date, end_date='+30d')
    no_of_days = (to_date - from_date).days
    roi = random.choice(rois)
    interest_payable = int(amount * int(roi.strip('%')) * no_of_days / 36500)
    tds = int(interest_payable * 0.10)
    net_interest = interest_payable - tds
    principal_repayment = amount if random.choice([True, False]) else None
    total = net_interest + (principal_repayment if principal_repayment else 0)

    return [
        shared_isin, shared_security_name, client_name, pan, bank_ac, bank, ifsc, holding,
        f"{face_value:,}", amount, from_date.strftime('%d-%b-%Y'), to_date.strftime('%d-%b-%Y'),
        no_of_days, roi, f"{interest_payable:,}", tds, net_interest, principal_repayment, total
    ]

# Generate Excel file
output_filename = "Financial_Report_Final.xlsx"
with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
    for sheet_name, row_count in sheet_info.items():
        # Pick unique clients just for this sheet
        clients_for_sheet = random.sample(unique_clients, row_count)
        data_rows = [generate_random_row(client) for client in clients_for_sheet]
        df = pd.DataFrame(data_rows, columns=columns)
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        worksheet = writer.sheets[sheet_name]
        worksheet.freeze_panes(1, 0)

print(f"✅ Excel file '{output_filename}' created successfully with consistent ISIN and Security Name.")
