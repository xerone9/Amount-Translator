import os
from amount_to_crore import *
from amount_to_million import *

print("""
SELECT CONVERSION MODE
=======================

[1] - Translation In Millions
[2] - Translation In Crores
For Exit Type - exit\n=============================\n""")
while True:
    selection = input("Enter Selection: ")

    if selection == "1":
        print("")
        while True:
            value = input("Enter Value: ")
            if value.lower() == "exit":
                os.system('cls')
                print("")
                print("""
SELECT CONVERSION MODE
=======================

[1] - Translation In Millions
[2] - Translation In Crores
For Exit Type - exit\n=============================\n""")
                break
            else:
                try:
                    print(amount_to_million(value) + "\n")
                except ValueError:
                    print("Incorrect Value\n")

    elif selection == "2":
        print("")
        while True:
            value = input("Enter Value: ")
            if value.lower() == "exit":
                os.system('cls')
                print("""
SELECT CONVERSION MODE
=======================

[1] - Translation In Millions
[2] - Translation In Crores
For Exit Type - exit\n=============================\n""")
                break
            else:
                try:
                    print(amount_to_crore(value)+ "\n")
                except ValueError:
                    print("Incorrect Value\n")

    elif selection.lower() == "exit":
        break

    else:
        print("Incorrect Choice Made\n")
