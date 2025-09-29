from typing import List, Dict


class Subject:
    def __init__(self, name: str, teacher: str, hours: int):
        self.name = name
        self.teacher = teacher
        self.hours = hours
        self.grades = ""
        self.homework = []
        self.topics = []

    def __repr__(self):
        return f"{self.name}"

    def __eq__(self, other):
        return self.name == other.name  # let the name contain class name as well


class Class:
    def __init__(self, name: str, students: List[str],):
        self.name = name
        self.subjects: Dict[str, Subject] = {}
        self.students = students

        self.is_kz = False  # by default

        print(f"class {self.name} has been created!")

    def __repr__(self):
        return f"class(name='{self.name}') has {self.students.count()} students, {self.subjects.count()} subjects"
