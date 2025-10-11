from typing import List, Dict


class Subject:
    def __init__(self, name: str, teacher: str):
        self.name = name
        self.teacher = teacher
        self.grades = ""
        self.has_exam = False
        self.homework = []
        self.topics = []
        self.hours_in_days = [0, 0, 0, 0, 0]  # 5 days in a week

    def __repr__(self):
        return f"{self.name}"

    def __eq__(self, other):
        return self.name == other.name  # let the name contain class name as well

    def hours(self):
        return sum(self.hours_in_days)


class Class:
    def __init__(self, name: str, subjects: Dict[str, Subject]):
        self.name = name
        self.subjects: Dict[str, Subject] = subjects
        self.students: List[str] = []

        self.is_kz = False  # by default

        print(f"class {self.name} has been created!")

    def __repr__(self):
        return f"class(name='{self.name}') has {len(self.students)} students"
