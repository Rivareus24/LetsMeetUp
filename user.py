class User:
    'Common base class for all employees'

    def __init__(self, firstName, secondName, email):
        self.firstName = firstName
        self.secondName = secondName
        self.email = email

    def displayEmployee(self):
        print(f"Name : {self.firstName}, Salary: {self.secondName}")
