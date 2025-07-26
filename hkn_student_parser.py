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

# Configuration flag for GPA logging
INCLUDE_GPA_IN_CSV = False  # Set to True to include individual GPAs in CSV files (this is for testing purposes)

# Cutoff percentages by grade level
SENIOR_CUTOFF = 0.30   # 20% for seniors
JUNIOR_CUTOFF = 0.25   # 15% for juniors
SOPHOMORE_CUTOFF = 0.20  # 10% for sophomores

# Sample data generation parameters
SAMPLE_GPA_MEAN = 3.5
SAMPLE_GPA_STD = 0.5
SAMPLE_MAJOR_WEIGHTS = [0.46, 0.53, 0.01]  # [EE, CMPE, Other]
SAMPLE_GRADE_WEIGHTS = [0.32, 0.34, 0.34]  # [Sophomore, Junior, Senior]
SAMPLE_ECE_CREDITS_PER_SEMESTER = 3.5 # number of ECE classes per semester
SAMPLE_ECE_CREDITS_VARIATION = 2.0
SAMPLE_DEFAULT_COUNT = 2177

# Error logging
ERROR_LOG = []  # Store errors for logging
OTHER_MAJORS = set()

def log_error(error_type, message, student_name=None, sheet_name=None):
  """Log an error with context information"""
  error_entry = {
    'type': error_type,
    'message': message,
    'student_name': student_name,
    'sheet_name': sheet_name
  }
  ERROR_LOG.append(error_entry)
  print(f"ERROR: {error_type} - {message}")


def main(use_sample_data=False, bymajor=True):
  seniors = list()
  juniors = list()
  sophomores = list()
  
  if use_sample_data:
    # Use sample data for testing
    print("Generating 500 sample students for testing...")
    all_students = generate_sample_students()
    
    # Classify sample students by semester count
    for student in all_students:
      nsemesters = student.num_nonsummer_semesters
      if nsemesters > 6:
        seniors.append(student)
      elif nsemesters > 4:
        juniors.append(student)
      elif nsemesters > 2:
        sophomores.append(student)
  else:
    # Use real Excel data
    wb_iter = WorkbookIterator()
    for sheet in tqdm.tqdm(wb_iter, desc="Parsing sheets"):
      try:
        student = Student(sheet)
        # Validate student data
        if not hasattr(student, 'name') or not student.name:
          log_error("VALIDATION", "Student has no name", sheet_name=getattr(sheet, 'name', 'Unknown'))
          continue
        if not hasattr(student, 'puid') or not student.puid:
          log_error("VALIDATION", "Student has no PUID", student_name=getattr(student, 'name', 'Unknown'), sheet_name=getattr(sheet, 'name', 'Unknown'))
          continue
        if not hasattr(student, 'gpa') or student.gpa is None or pd.isna(student.gpa):
          log_error("VALIDATION", "Student has invalid GPA", student_name=student.name, sheet_name=getattr(sheet, 'name', 'Unknown'))
          continue
          
        # Check for reasonable GPA range
        if student.gpa < 0 or student.gpa > 4.0:
          log_error("VALIDATION", f"Student has unrealistic GPA: {student.gpa}")
        
        # Classify by semester count
        nsemesters = student.num_nonsummer_semesters
        if nsemesters >  6:
          seniors.append(student)
        elif nsemesters > 4:
          juniors.append(student)
        elif nsemesters > 2:
          sophomores.append(student)
        # else:
        #   log_error("CLASSIFICATION", f"Student has too few semesters ({nsemesters}) to classify", student_name=student.name)
          
      except Exception as e:
        log_error("PARSING", f"Failed to parse student sheet: {str(e)}", sheet_name=getattr(sheet, 'name', 'Unknown'))

  if use_sample_data:
    print(f"Generated {len(seniors) + len(juniors) + len(sophomores)} sample students")
  else:
    print("Number of students parsed:", len(wb_iter))
  totseniors = len(seniors) 
  totjuniors = len(juniors) 
  totsophomores = len(sophomores) 

  # Filter students based on bymajor parameter
  if bymajor:
    # Filter by major within each grade level
    filtered_seniors = filter_students_by_major(seniors, SENIOR_CUTOFF)
    filtered_juniors = filter_students_by_major(juniors, JUNIOR_CUTOFF)
    filtered_sophomores = filter_students_by_major(sophomores, SOPHOMORE_CUTOFF)
    
    print_stats_by_major("seniors", filtered_seniors, totseniors)
    print_stats_by_major("juniors", filtered_juniors, totjuniors)
    print_stats_by_major("sophomores", filtered_sophomores, totsophomores)
    
    write_student_list_to_file_by_major("seniors.csv", filtered_seniors)
    write_student_list_to_file_by_major("juniors.csv", filtered_juniors)
    write_student_list_to_file_by_major("sophomores.csv", filtered_sophomores)
  else:
    # Filter without considering major - use original approach
    filtered_seniors = filter_out_students(seniors, SENIOR_CUTOFF)
    filtered_juniors = filter_out_students(juniors, JUNIOR_CUTOFF)
    filtered_sophomores = filter_out_students(sophomores, SOPHOMORE_CUTOFF)
    
    print_stats("seniors", filtered_seniors, totseniors)
    print_stats("juniors", filtered_juniors, totjuniors)
    print_stats("sophomores", filtered_sophomores, totsophomores)
    
    write_student_list_to_file("seniors.csv", filtered_seniors)
    write_student_list_to_file("juniors.csv", filtered_juniors)
    write_student_list_to_file("sophomores.csv", filtered_sophomores)

  # Generate report file
  generate_report("report.txt", filtered_seniors, filtered_juniors, filtered_sophomores, totseniors, totjuniors, totsophomores, use_sample_data, seniors, juniors, sophomores, bymajor)



def filter_out_students(students, top_percent):
  students = sort_students_by_gpa(students)
  students = take_only_top_percentile(students, top_percent)
  students = remove_less_than_10_ece_credits(students)
  return students

def sort_alphabetically(students):
  if isinstance(students, list):
    students_copy = students.copy()
  else:
    students_copy = list(students)
  students_copy.sort(key=lambda student: student.name)
  return students_copy

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


def filter_students_by_major(students, top_percent):
  """Filter students by major, taking top percentage from each major"""
  # Group students by major
  students_by_major = group_students_by_major(students)
  
  filtered_students = []
  for major, major_students in students_by_major.items():
    # Filter each major group
    filtered_major_students = filter_out_students(major_students, top_percent)
    filtered_students.extend(filtered_major_students)
  
  return filtered_students

def group_students_by_major(students):
  """Group students by their major filtering group (EE, CMPE, Other)"""
  students_by_major = collections.defaultdict(list)
  
  for student in students:
    if student.major is None:
      log_error("VALIDATION", "Student has no major assigned", student_name=student.name)
      continue
    # Use the major_group property for filtering
    students_by_major[student.major_group].append(student)
  
  return dict(students_by_major)

def print_stats_by_major(class_string, students, totnum):
  """Print statistics broken down by major"""
  print(f"\n=== {class_string.upper()} STATISTICS ===")
  students_by_major = group_students_by_major(students)
  
  total_selected = len(students)
  print(f"Total {class_string} selected: {total_selected} out of {totnum}")
  
  for major, major_students in students_by_major.items():
    num_students = len(major_students)
    gpa_cutoff = get_gpa_cutoff(major_students)
    print(f"  {major}: {num_students} students, GPA cutoff: {gpa_cutoff:.2f}")

def write_student_list_to_file_by_major(filename, students):
  """Write student list to file, organized by major group but showing actual major"""
  students_by_major_group = group_students_by_major(students)
  
  with open(filename, "w") as outfile:
    # Write header based on GPA flag
    if INCLUDE_GPA_IN_CSV:
      outfile.write("Major Group,Actual Major,Name,PUID,GPA\n")
    else:
      outfile.write("Major Group,Actual Major,Name,PUID\n")
    
    for major_group in sorted(students_by_major_group.keys()):
      major_students = students_by_major_group[major_group]
      major_students = sort_alphabetically(major_students.copy())
      
      for student in major_students:
        if INCLUDE_GPA_IN_CSV:
          outline = f"{student.major_group},{student.major},{student.name},{student.puid},{student.gpa:.2f}\n"
        else:
          outline = f"{student.major_group},{student.major},{student.name},{student.puid}\n"
        outfile.write(outline)


def generate_report(filename, seniors, juniors, sophomores, totseniors, totjuniors, totsophomores, use_sample_data=False, unfiltered_seniors=None, unfiltered_juniors=None, unfiltered_sophomores=None, bymajor=True):
  """Generate a comprehensive report file with all statistics"""
  with open(filename, "w") as outfile:
    # Header
    outfile.write("=" * 60 + "\n")
    if use_sample_data:
      outfile.write("HKN RECRUITMENT FILTER REPORT (SAMPLE DATA)\n")
    else:
      outfile.write("HKN RECRUITMENT FILTER REPORT\n")
    outfile.write(f"Generated on: {datetime.date.today()}\n")
    outfile.write(f"Current term: {current_term}\n")
    if not bymajor:
      outfile.write("NOTE: Major-based filtering DISABLED - using grade-level filtering only\n")
    outfile.write("=" * 60 + "\n\n")
    
    # Errors and Concerns section
    outfile.write("ERRORS AND CONCERNS\n")
    outfile.write("-" * 20 + "\n")
    
    if not ERROR_LOG:
      outfile.write("No errors encountered during processing.\n\n")
    else:
      outfile.write(f"Total errors logged: {len(ERROR_LOG)}\n\n")
      
      for i, error in enumerate(ERROR_LOG, 1):
        outfile.write(f"ERROR #{i}\n")
        outfile.write(f"  Type: {error['type']}\n")
        outfile.write(f"  Message: {error['message']}\n")
        if error['student_name']:
          outfile.write(f"  Student: {error['student_name']}\n")
        if error['sheet_name']:
          outfile.write(f"  Sheet: {error['sheet_name']}\n")
        outfile.write("\n")
    
    # Overall summary
    outfile.write("OVERALL SUMMARY\n")
    outfile.write("-" * 20 + "\n")
    outfile.write(f"Total students parsed: {totseniors + totjuniors + totsophomores}\n")
    outfile.write(f"Total seniors parsed: {totseniors} ({totseniors / (totseniors + totjuniors + totsophomores) * 100:.1f}%)\n")
    outfile.write(f"Total juniors parsed: {totjuniors} ({totjuniors / (totseniors + totjuniors + totsophomores) * 100:.1f}%)\n")
    outfile.write(f"Total sophomores parsed: {totsophomores} ({totsophomores / (totseniors + totjuniors + totsophomores) * 100:.1f}%)\n\n")

    # Calculate total by major group across all grade levels (only if bymajor is True)
    if bymajor:
      all_students = seniors + juniors + sophomores
      major_counts = {}
      for student in all_students:
        major_group = student.major_group if hasattr(student, 'major_group') else "Unknown"
        major_counts[major_group] = major_counts.get(major_group, 0) + 1
      
      # Calculate total unfiltered students by major group to get percentage of major selected
      if unfiltered_seniors and unfiltered_juniors and unfiltered_sophomores:
        all_unfiltered_students = unfiltered_seniors + unfiltered_juniors + unfiltered_sophomores
        unfiltered_major_counts = {}
        for student in all_unfiltered_students:
          major_group = student.major_group if hasattr(student, 'major_group') else "Unknown"
          unfiltered_major_counts[major_group] = unfiltered_major_counts.get(major_group, 0) + 1
      
      if major_counts:
        outfile.write("Total invitations by major group (all grades):\n")
        for major_group in sorted(major_counts.keys()):
          count = major_counts[major_group]
          if unfiltered_seniors and unfiltered_juniors and unfiltered_sophomores:
            # Show percentage of that major group that was selected
            total_in_major = unfiltered_major_counts.get(major_group, 0)
            percentage = (count / total_in_major * 100) if total_in_major > 0 else 0
            outfile.write(f"  {major_group}: {count} students ({percentage:.1f}% of {major_group} students)\n")
          else:
            # Fallback to original calculation
            total_parsed = totseniors + totjuniors + totsophomores
            percentage = (count / total_parsed * 100) if total_parsed > 0 else 0
            outfile.write(f"  {major_group}: {count} students ({percentage:.1f}%)\n")
        outfile.write("\n")
    
    # Detailed statistics for each grade level
    for grade_name, students, total in [("SENIORS", seniors, totseniors),
                                       ("JUNIORS", juniors, totjuniors),
                                       ("SOPHOMORES", sophomores, totsophomores)]:
      outfile.write(f"{grade_name} STATISTICS\n")
      outfile.write("-" * 20 + "\n")
      
      total_selected = len(students)
      percentage = (total_selected / total * 100) if total > 0 else 0
      
      outfile.write(f"Total {grade_name.lower()} selected: {total_selected} out of {total} ({percentage:.1f}%)\n")
      
      if bymajor:
        # Show breakdown by major when filtering by major
        students_by_major = group_students_by_major(students)
        if students_by_major:
          outfile.write("Breakdown by major:\n")
          for major in sorted(students_by_major.keys()):
            major_students = students_by_major[major]
            num_students = len(major_students)
            gpa_cutoff = get_gpa_cutoff(major_students)
            outfile.write(f"  {major}: {num_students} students, GPA cutoff: {gpa_cutoff:.2f}\n")
        else:
          outfile.write("No students selected for this grade level.\n")
      else:
        # Show overall GPA cutoff when not filtering by major
        if students:
          gpa_cutoff = get_gpa_cutoff(students)
          outfile.write(f"GPA cutoff: {gpa_cutoff:.2f}\n")
        else:
          outfile.write("No students selected for this grade level.\n")
      
      outfile.write("\n")
    
    # Filter criteria
    outfile.write("FILTER CRITERIA\n")
    outfile.write("-" * 15 + "\n")
    outfile.write("Top percentages by grade level:\n")
    outfile.write(f"  Seniors: {SENIOR_CUTOFF:.0%}\n")
    outfile.write(f"  Juniors: {JUNIOR_CUTOFF:.0%}\n")
    outfile.write(f"  Sophomores: {SOPHOMORE_CUTOFF:.0%}\n")
    outfile.write("Additional filters:\n")
    outfile.write("  - Minimum 10 ECE credits\n")
    if bymajor:
      outfile.write("  - Filtered by major (EE, CMPE, Other)\n")
      outfile.write("  - Sorted by GPA within each major\n")
    else:
      outfile.write("  - Sorted by GPA (major filtering disabled)\n")
    outfile.write("\n")
    
    # Sample data distributions (only shown when using sample data)
    if use_sample_data:
      outfile.write("SAMPLE DATA DISTRIBUTIONS\n")
      outfile.write("-" * 25 + "\n")
      outfile.write("Note: This report was generated using sample data with the following distributions:\n\n")
      outfile.write("GPA Distribution:\n")
      outfile.write(f"  - Normal distribution with mean={SAMPLE_GPA_MEAN}, std dev={SAMPLE_GPA_STD}\n")
      outfile.write("  - Clamped to range [0.0, 4.0]\n\n")
      outfile.write("Major Distribution:\n")
      outfile.write(f"  - EE: {SAMPLE_MAJOR_WEIGHTS[0]:.0%}\n")
      outfile.write(f"  - CMPE: {SAMPLE_MAJOR_WEIGHTS[1]:.0%}\n")
      outfile.write(f"  - Other: {SAMPLE_MAJOR_WEIGHTS[2]:.0%}\n\n")
      outfile.write("Grade Level Distribution:\n")
      outfile.write(f"  - Sophomores (3-4 semesters): {SAMPLE_GRADE_WEIGHTS[0]:.0%}\n")
      outfile.write(f"  - Juniors (5-6 semesters): {SAMPLE_GRADE_WEIGHTS[1]:.0%}\n")
      outfile.write(f"  - Seniors (7-10 semesters): {SAMPLE_GRADE_WEIGHTS[2]:.0%}\n\n")
      outfile.write("ECE Credits:\n")
      outfile.write(f"  - Base: {SAMPLE_ECE_CREDITS_PER_SEMESTER} credits per semester\n")
      outfile.write(f"  - Random variation: Normal(0, {SAMPLE_ECE_CREDITS_VARIATION})\n")
      outfile.write("  - Minimum: 0 credits\n\n")
      outfile.write("Names and PUIDs:\n")
      outfile.write("  - Names: Random combination of 50 first names × 50 last names\n")
      outfile.write("  - PUIDs: Random 9-digit numbers\n")
      outfile.write(f"  - Default sample size: {SAMPLE_DEFAULT_COUNT} students\n\n")
    
    # Other majors section (only shown when processing real data and OTHER_MAJORS is not empty)
    if not use_sample_data and OTHER_MAJORS:
      outfile.write("OTHER MAJORS ENCOUNTERED\n")
      outfile.write("-" * 24 + "\n")
      outfile.write("The following majors were classified as 'Other' during processing:\n\n")
      
      # Sort the majors alphabetically for consistent output
      sorted_other_majors = sorted(OTHER_MAJORS)
      for i, major in enumerate(sorted_other_majors, 1):
        outfile.write(f"  {i:2d}. {major}\n")
      
      outfile.write(f"\nTotal unique 'Other' majors found: {len(OTHER_MAJORS)}\n")
      outfile.write("Note: All non-EE and non-CMPE majors are grouped as 'Other' for filtering purposes.\n\n")
    
    # Output files
    outfile.write("OUTPUT FILES\n")
    outfile.write("-" * 12 + "\n")
    outfile.write("Generated CSV files:\n")
    outfile.write("  - seniors.csv\n")
    outfile.write("  - juniors.csv\n")
    outfile.write("  - sophomores.csv\n")
    if INCLUDE_GPA_IN_CSV:
      outfile.write("Note: Individual GPAs are included in CSV files.\n")
    else:
      outfile.write("Note: Individual GPAs are NOT included in CSV files (production mode).\n")


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
    try:
      self._get_identifying(sheet)
      self._get_gpa(sheet)
      self._get_year_and_credit_number(sheet)
      self._validate_data()
    except Exception as e:
      # Re-raise with more context
      raise Exception(f"{str(e)}")
    
    # Set major group for filtering and keep track of non-EE/CMPE majors
    if self.major in ["EE", "CMPE"]:
      self.major_group = self.major
    else:
      self.major_group = "Other"
      OTHER_MAJORS.add(self.major)

  def _get_identifying(self, sheet):
    try:
      name_puid_str = sheet.iloc[4, 1]
      if pd.isna(name_puid_str) or not str(name_puid_str).strip():
        raise ValueError("Name/PUID field is empty or missing")
      
      pattern = r"(.*?)\s+(\d+)"
      matches = re.search(pattern, str(name_puid_str))
      if not matches:
        raise ValueError(f"Could not parse name and PUID from: '{name_puid_str}'")
      
      self.name = matches.group(1).strip()
      self.puid = matches.group(2)
      
      if not self.name:
        raise ValueError("Student name is empty")
      if len(self.puid) < 5:  # Reasonable PUID length check
        raise ValueError(f"PUID appears too short: {self.puid}")
        
    except (IndexError, KeyError) as e:
      raise ValueError(f"Could not find name/PUID at expected location (row 5, col 2): {str(e)}")

  def _get_gpa(self, sheet):
    try:
      gpa_value = sheet.iloc[7, 9]
      if pd.isna(gpa_value):
        raise ValueError("GPA field is empty")
      
      self.gpa = float(gpa_value)
      
    except (IndexError, KeyError) as e:
      raise ValueError(f"Could not find GPA at expected location (row 8, col 10): {str(e)}")
    except (ValueError, TypeError) as e:
      raise ValueError(f"GPA value is not a valid number: {gpa_value} - {str(e)}")

  def _get_year_and_credit_number(self, sheet):
    try:
      startrow = 10
      maxrow = sheet.index.values[-1]
      num_fallspring_semesters = 0
      num_ece_credits = 0
      self.major = None  # Initialize major to None
      
      for row in range(startrow, maxrow + 1):
        try:
          rowval = sheet.iloc[row, 0]
          if not isinstance(rowval, str):
            continue
          if 'Fall' in rowval or 'Spring' in rowval:
            num_fallspring_semesters += 1
            # Get the most recent major
            if row < len(sheet) and len(sheet.columns) > 8:
              major_val = sheet.iloc[row, 8]
              if not pd.isna(major_val):
                self.major = str(major_val).strip()    

          if 'ECE' in rowval and len(sheet.columns) > 9:
            credits = sheet.iloc[row, 9]
            if not pd.isna(credits):
              try:
                num_ece_credits += float(credits)
              except (ValueError, TypeError):
                log_error("PARSING", f"Invalid ECE credit value: {credits}", student_name=getattr(self, 'name', 'Unknown'))
                
          if rowval == current_term:
            break
        except (IndexError, KeyError):
          # Skip rows that can't be accessed
          continue
      
      self.num_ece_credits = num_ece_credits
      self.num_nonsummer_semesters = num_fallspring_semesters
      
    except Exception as e:
      raise ValueError(f"Error parsing semester/credit data: {str(e)}")

  def _validate_data(self):
    """Validate parsed student data for reasonableness"""
    if not hasattr(self, 'name') or not self.name:
      raise ValueError("Student name is missing or empty")
    
    if not hasattr(self, 'puid') or not self.puid:
      raise ValueError("Student PUID is missing or empty")
    
    if not hasattr(self, 'gpa') or self.gpa is None:
      raise ValueError("Student GPA is missing")
    
    if self.gpa < 0 or self.gpa > 4:  # Allow some margin for different GPA scales
      log_error("VALIDATION", f"Unusual GPA value: {self.gpa}")
    
    if not hasattr(self, 'num_nonsummer_semesters') or self.num_nonsummer_semesters < 0:
      raise ValueError("Invalid semester count")
    
    if not hasattr(self, 'num_ece_credits') or self.num_ece_credits < 0:
      raise ValueError("Invalid ECE credits count")
    
  def __repr__(self):
    return "Student({}, major={}, major_group={}, sems={}, creds={}, gpa={})".format(
      self.name, self.major, self.major_group, self.num_nonsummer_semesters,
      self.num_ece_credits, self.gpa
    )


def generate_sample_students(num_students=SAMPLE_DEFAULT_COUNT):
  """Generate sample students for testing purposes"""
  import random
  
  # Sample data for generating realistic students
  first_names = ["John", "Jane", "Michael", "Sarah", "David", "Emily", "James", "Ashley", "Christopher", "Jessica",
                 "Matthew", "Amanda", "Joshua", "Melissa", "Daniel", "Michelle", "Andrew", "Kimberly", "Mark", "Amy",
                 "Steven", "Lisa", "Paul", "Angela", "Kenneth", "Heather", "Joshua", "Nicole", "Kevin", "Elizabeth",
                 "Brian", "Rebecca", "George", "Maria", "Edward", "Samantha", "Ronald", "Deborah", "Timothy", "Rachel",
                 "Jason", "Carolyn", "Jeffrey", "Janet", "Ryan", "Catherine", "Jacob", "Frances", "Gary", "Christine"]
  
  last_names = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez",
                "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin",
                "Lee", "Perez", "Thompson", "White", "Harris", "Sanchez", "Clark", "Ramirez", "Lewis", "Robinson",
                "Walker", "Young", "Allen", "King", "Wright", "Scott", "Torres", "Nguyen", "Hill", "Flores",
                "Green", "AdAMS", "Nelson", "Baker", "Hall", "Rivera", "Campbell", "Mitchell", "Carter", "Roberts"]
  
  majors = ["EE", "CMPE", "Other"]
  
  students = []
  
  for i in range(num_students):
    # Create a mock student object
    student = type('Student', (), {})()
    
    # Generate basic info
    student.name = f"{random.choice(first_names)} {random.choice(last_names)}"
    student.puid = str(random.randint(100000000, 999999999))  # 9-digit PUID
    
    # Generate realistic GPA using global parameters
    student.gpa = max(0.0, min(4.0, random.normalvariate(SAMPLE_GPA_MEAN, SAMPLE_GPA_STD)))
    student.gpa = round(student.gpa, 2)
    
    # Generate major based on global weights
    student.major = random.choices(majors, weights=SAMPLE_MAJOR_WEIGHTS)[0]
    
    # Set major group for filtering
    if student.major in ["EE", "CMPE"]:
      student.major_group = student.major
    else:
      student.major_group = "Other"
    
    # Generate semester count using global weights
    semester_category = random.choices(['sophomore', 'junior', 'senior'], weights=SAMPLE_GRADE_WEIGHTS)[0]
    
    if semester_category == 'sophomore':
      student.num_nonsummer_semesters = random.randint(3, 4)
    elif semester_category == 'junior':
      student.num_nonsummer_semesters = random.randint(5, 6)
    else:  # senior
      student.num_nonsummer_semesters = random.randint(7, 10)
    
    # Generate ECE credits using global parameters
    base_credits = student.num_nonsummer_semesters * SAMPLE_ECE_CREDITS_PER_SEMESTER # min 3 * 3.5 = 10.5
    variation = random.normalvariate(0, SAMPLE_ECE_CREDITS_VARIATION)
    student.num_ece_credits = max(0, base_credits + variation)
    student.num_ece_credits = round(student.num_ece_credits, 1)
    
    students.append(student)
  
  return students


if __name__ == "__main__":
  import sys
  
  # Parse command line arguments
  use_sample_data = "--sample" in sys.argv
  bymajor = "--no-major" not in sys.argv  # Default to True unless --no-major is specified
  
  if "--help" in sys.argv or "-h" in sys.argv:
    print("HKN Student Recruitment Filter")
    print("Usage: python hkn_student_parser.py [options]")
    print("Options:")
    print("  --sample      Use sample data instead of Excel files")
    print("  --no-major    Disable major-based filtering (use grade-level filtering only)")
    print("  --help, -h    Show this help message")
    sys.exit(0)
  
  # Run the main function with the parsed arguments
  main(use_sample_data=use_sample_data, bymajor=bymajor)
