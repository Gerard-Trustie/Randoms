import pandas as pd
import random
from datetime import timedelta, datetime

# Define spending categories and some example merchants or details for each
spending_categories = {
    "Bills": ["Electricity Bill", "Water Bill", "Internet Bill"],
    "Business": ["Office Supplies", "Freelance Web Design", "Business Lunch"],
    "Car": ["Fuel", "Car Insurance", "Car Wash"],
    "Charges & Taxes": ["Bank Account Fee", "Late Payment Fee", "Income Tax Payment"],
    "Clothing": ["Department Store", "Online Retailer", "Shoe Store"],
    "Eats & Drinks": ["Coffee Shop", "Restaurant", "Fast Food"],
    "Education": ["Textbooks", "Online Course", "Tuition Fee"],
    "Entertainment": ["Movie Tickets", "Streaming Service Subscription", "Music Concert"],
    "Gifts": ["Birthday Gift", "Wedding Gift", "Charity Donation"],
    "Groceries": ["Supermarket", "Farmers Market", "Online Groceries"],
    "Health & Fitness": ["Gym Membership", "Pharmacy", "Yoga Class"],
    "Home": ["Furniture Store", "Home Improvement Store", "Gardening Supplies"],
    "Income": ["Salary", "Freelance Income", "Gift"],
    "Kids": ["Toys", "School Supplies", "Childcare"],
    "Money Transfers": ["Bank Transfer", "Payment App Transfer", "Wire Transfer"],
    "Pets": ["Pet Food", "Vet Visit", "Pet Toys"],
    "Self-care": ["Spa Day", "Beauty Products", "Hair Salon"],
    "Tech": ["Electronics Store", "Online Tech Retailer", "App Purchase"],
    "Transport": ["Train Ticket", "Bus Fare", "Taxi"],
    "Travel": ["Airline Ticket", "Hotel Booking", "Travel Insurance"]
}

# Helper function to generate a random date within the last year
def generate_random_date():
    start_date = datetime.now() - timedelta(days=365)
    end_date = datetime.now()
    random_date = start_date + timedelta(days=random.randint(0, 365))
    return random_date.strftime("%Y-%m-%d")

# Generate the transactions
transactions = []
for _ in range(100):
    category = random.choice(list(spending_categories.keys()))
    detail = random.choice(spending_categories[category])
    date = generate_random_date()
    amount = round(random.uniform(5, 500), 2)  # Amounts between £5 and £500
    if category == "Income":
        amount *= -1  # Negative amount for income
    transactions.append({"Date": date, "Category": category, "Detail": detail, "Amount (£)": amount})

# Create a DataFrame
df_transactions = pd.DataFrame(transactions)

df_transactions.head()  # Show a preview of the data

# print the transactions to a csv file

df_transactions.to_csv('transactions.csv', index=False)
