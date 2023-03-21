# Import uuid for creating bank number
import uuid
import pandas as pd


# Person Class
class Person:

    def __init__(self, name, surname, account_number, money):
        self.name = name
        self.surname = surname
        self.account_number = account_number
        self.money = money
        self.userlist = []
        self.stocklist = []



    # Show user account details
    def show_user_info(self):
        print(f'Full Name: {self.name} {self.surname}\n'f'Bank Name: {self.account_number}\n'f'Total Money: ${self.money}')



    # Show user money
    def show_money(self):
        print(f'{self.name} has ${self.money}.\n')



    # Add new user to person class
    def add_new_user(self, name, surname, account_number, money):

        for user in self.userlist:
            if user["Account Number"] == account_number:
                print(f'{account_number} already exists!\n')
                return

        new_user = {
            "Name": name,
            "Surname": surname,
            "Account Number": account_number,
            "Money": money
        }

        # Append it to user list
        self.userlist.append(new_user)
        print(f'{name} {surname} has been added to the user list!\n')



    # Delete an existing user
    def delete_user(self, account_number):

        for account in self.userlist:
            if account["Account Number"] == account_number:
                self.userlist.remove(account)
                print(f'{account_number} has been deleted from user list!\n')
                return

        print(f'User not found.\n')
        return



    def users_to_excel(self):

        # Create a DataFrame from userlist
        df = pd.DataFrame(self.userlist)

        # Write the DataFrame to an Excel file
        with pd.ExcelWriter('koray.xlsx', mode='a', engine='openpyxl') as writer:
            # Check if the Sheet1 already exists in the workbook
            sheet_names = writer.book.sheetnames
            if 'Sheet1' in sheet_names:
                # If Sheet1 exists, get its index
                sheet_index = sheet_names.index('Sheet1')
                # Remove Sheet1 from the workbook
                writer.book.remove(writer.book.worksheets[sheet_index])
            # Write the DataFrame to the Excel file
            df.to_excel(writer, sheet_name='Sheet1', index=False)

        print('Users have been added to the Excel file!\n')



# Bank class
class Bank:
    def __init__(self,stock_name, value):
        self.stock_name = stock_name
        self.value = value
        self.stocklist = []



    # Show the value of the stocks
    def show_stock_value(self, stock_name):

        if stock_name == self.stock_name:
            print(f'{self.stock_name} is recently ${self.value}')

        else:
            print(f'No stocks found named {stock_name}!\n')



    # Add new stock to the market
    def add_new_stock(self, stock_name, value):

        # Check if stock already exists
        for item in self.stocklist:
            if item["Stock Name"] == stock_name:
                print(f'{stock_name} already exists in market!\n')
                return

        # If stock doesn't exist, add it to the list
        new_stock = {
            "Stock Name": stock_name,
            "Current Value": value
        }

        self.stocklist.append(new_stock)
        print(f'{stock_name} has been added to the market!\n')



    # Delete a stock from market
    def delete_stock(self, stock_name):

        for item in self.stocklist:
            if item["Stock Name"] == stock_name:
                self.stocklist.remove(item)

            print(f'No stocks found named {stock_name}!\n')



    # Update stock value
    def update_stock_value(self, stock_name, new_value):

        for item in self.stocklist:
            if item["Stock Name"] == stock_name:
                item["Current Value"] = new_value
                print(f'{stock_name} stock value has been updated to {new_value}!\n')
                return

            else:
                print(f'No stocks found named {stock_name}!\n')



user1 = Person("Koray", "Alkan", 100, 200)
user1.add_new_user("Koray", "Alkan", 100, 200)

user1.show_user_info()
user1.users_to_excel()