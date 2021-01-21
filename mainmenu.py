class MainMenu:
    def __init__(self):
        self.display_options()

    def display_options(self):
        print('Hello, what is your name?')
        name = input()
        print('Hello ' + name + '!!!') 
        print('Choose an option from the menu below:')

        not_chosen_yet = True

        while not_chosen_yet:
            options = {
                "1":"Press 1 to SEARCH an excel spreadsheet.",
                "2":"Press 2 to COUNT the instances in a column of an excel spreadsheet.",
                "3":"Press 3 to CREATE a certificate for every student in an excel spreadsheet."
            }
            for i in range(1, len(options) + 1):
                print(options[str(i)])
            print("What would you like to do?")
            choice = input()
            if choice in options:
                break
            else:
                print("Sorry, please only enter a number from the menu.")

        self.choice = choice
        print("You have chosen " + choice)

        