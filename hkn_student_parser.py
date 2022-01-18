import re
import math
import datetime
import pandas as pd

# Useful globals (I know ew, but just a few)
current_year = datetime.date.today().year

if datetime.date.today().month > 6:
  current_semester = 'Fall'
else:
  current_semester = 'Spring'

current_term = current_semester + ' ' \
               + str(current_year)


def main():
  excelsheet_name = 'Record Copy - PUID Entry.xlsx'
  workbook = pd.ExcelFile(excelsheet_name)

  seniors = list()
  juniors = list()
  sophomores = list()
  
  for sheetname in workbook.sheet_names:
    sheet = workbook.parse(sheetname)
    student = Student(sheet)
    year = student.year_classification
    if year == 2:
      sophomores.append(student)
    elif year == 3:
      juniors.append(student)
    elif year == 4:
      seniors.append(student)

  seniors = filter_out_students(seniors, .30)
  juniors = filter_out_students(juniors, .25)
  sophomores = filter_out_students(sophomores, .20)

  write_student_list_to_file("seniors.csv", seniors)
  write_student_list_to_file("juniors.csv", juniors)
  write_student_list_to_file("sophomores.csv", sophomores)


def filter_out_students(students, top_percent):
  students = sort_students_by_gpa(students)
  students = take_only_top_percentile(students, top_percent)
  students = remove_less_than_10_ece_credits(students)
  return sort_alphabetically(students)

def sort_alphabetically(students):
  students.sort(key=lambda student: student.name)
  return students

def remove_less_than_10_ece_credits(students):
  students = filter(lambda student: student.num_ece_credits >= 10,
                    students)
  return list(students)

def take_only_top_percentile(students, top_percent):
  index_cutoff = math.ceil(len(students) * top_percent)
  return students[:index_cutoff]

def sort_students_by_gpa(students):
  students.sort(key=lambda student: student.gpa,
                reverse=True)
  return students

def write_student_list_to_file(filename, students):
  with open(filename, "w") as outfile:
    for student in students:
      outline = "%s, %s" % (student.name, student.puid)
      outfile.write(outline)


class Student:
  def __init__(self, sheet):
    self._get_identifying(sheet)
    self._get_gpa(sheet)
    self._get_year_and_credit_number(sheet)

  def _get_identifying(self, sheet):
    name_puid_str = sheet.iloc[4][1] 
    pattern = r"(.*?)\s+(\d+)"
    matches = re.search(pattern, name_puid_str)
    self.name = matches.group(1)
    self.puid = matches.group(2)

  def _get_gpa(self, sheet):
    self.gpa = sheet.iloc[7][9]

  def _get_year_and_credit_number(self, sheet):
    startrow = 10
    maxrow = sheet.index.values[-1]
    num_fallspring_semesters = 0
    num_ece_credits = 0
    for row in range(startrow, maxrow + 1):
      rowval = sheet.iloc[row][0]
      if not isinstance(rowval, str):
        continue
      if 'Fall' in rowval or 'Spring' in rowval:
        num_fallspring_semesters += 1
      if 'ECE' in rowval:
        num_ece_credits += sheet.iloc[row][9]
      if rowval is current_term:
        break
    self.num_ece_credits = num_ece_credits
    self.year_classification = math.ceil(num_fallspring_semesters / 2)


if __name__ == "__main__":
  main()
