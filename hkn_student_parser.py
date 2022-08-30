import re
import os
import math
import tqdm
import datetime
import collections
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
  seniors = list()
  juniors = list()
  sophomores = list()
  
  wb_iter = WorkbookIterator()
  for sheet in tqdm.tqdm(wb_iter, desc="Parsing sheets"):
    student = Student(sheet)
    # print(student)
    nsemesters = student.num_nonsummer_semesters
    if nsemesters >  6:
      seniors.append(student)
    elif nsemesters > 4:
      juniors.append(student)
    elif nsemesters > 2:
      sophomores.append(student)

  print("Number of students parsed:", len(wb_iter))
  totseniors = len(seniors) 
  totjuniors = len(juniors) 
  totsophomores = len(sophomores) 

  seniors = filter_out_students(seniors, .30)
  juniors = filter_out_students(juniors, .25)
  sophomores = filter_out_students(sophomores, .20)

  print_stats("seniors", seniors, totseniors)
  print_stats("juniors", juniors, totjuniors)
  print_stats("sophomores", sophomores, totsophomores)

  write_student_list_to_file("seniors.csv", seniors)
  write_student_list_to_file("juniors.csv", juniors)
  write_student_list_to_file("sophomores.csv", sophomores)



def filter_out_students(students, top_percent):
  students = sort_students_by_gpa(students)
  students = take_only_top_percentile(students, top_percent)
  students = remove_less_than_10_ece_credits(students)
  return students

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

def get_gpa_cutoff(students):
  if len(students) == 0:
    return float("NaN")
  return min(s.gpa for s in students)

def print_stats(class_string, students, totnum):
  fstr = "Number %10s Qualified: %3d  Out of: %3d  GPA cutoff: %.2f"
  nstudents = len(students)
  gpa = get_gpa_cutoff(students)
  print(fstr % (class_string, nstudents, totnum, gpa))

def write_student_list_to_file(filename, students):
  # assumed students list is already filtered by filter_out_students
  with open(filename, "w") as outfile:
    if len(students) > 0:
      outfile.write("gpa cutoff:, %f\n" % students[-1].gpa)
    students = sort_alphabetically(students)
    for student in students:
      outline = "%s, %s\n" % (student.name, student.puid)
      outfile.write(outline)


class WorkbookIterator:
  def __init__(self, path='.'):
    filenames = os.listdir(path)
    xlsxfilter = lambda name: os.path.splitext(name)[1] == '.xlsx'
    self.workbook_filenames = list(filter(xlsxfilter, filenames))
    if len(self.workbook_filenames) == 0:
      raise RuntimeError("No .xlsx files in directory")
    self._open_workbooks()
    self.workbook_index = 0
    self.curr_workbook = pd.ExcelFile(self.workbook_filenames[0])
    self.curr_sheet_index = 0

  def __len__(self):
    return sum(len(w.sheet_names) for w in self.workbooks)

  def __iter__(self):
    return self

  def __next__(self):
    if self.curr_sheet_index == len(self.curr_workbook.sheet_names):
      self.workbook_index += 1
      if self.workbook_index == len(self.workbooks):
        raise StopIteration
      self.curr_workbook = self.workbooks[self.workbook_index]
      self.curr_sheet_index = 0
    sheetname = self.curr_workbook.sheet_names[self.curr_sheet_index]
    self.curr_sheet_index += 1
    return self.curr_workbook.parse(sheetname)

  def _open_workbooks(self):
    self.workbooks = list()
    for filename in self.workbook_filenames:
      workbook = pd.ExcelFile(filename)
      self.workbooks.append(workbook)


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
    self.num_nonsummer_semesters = num_fallspring_semesters
    
  def __repr__(self):
    return "Student({}, sems={}, creds={}, gpa={})".format(
      self.name, self.num_nonsummer_semesters,
      self.num_ece_credits, self.gpa
    )


if __name__ == "__main__":
  main()
