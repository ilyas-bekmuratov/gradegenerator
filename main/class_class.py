class Class:
    def __init__(self, name, subjects, teachers, hours):
        self.name = name
        self.students = []
        self.subjects = subjects
        self.teachers = teachers
        self.hours = hours
        self.grades = {}

        self.is_kz = False  # by default

        print(f"class {self.name} has been created!")

    def add_students(self, students):
        self.students = students

    def add_grades(self, grades):
        self.grades = grades
