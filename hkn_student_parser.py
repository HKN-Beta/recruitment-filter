import re
import os
import math
import tqdm
import datetime
import collections
import pandas as pd
import numpy as np
import warnings
import multiprocessing as mp

# Suppress pandas warnings for cleaner output
warnings.filterwarnings('ignore', category=UserWarning)

# === CONFIGURATION CONSTANTS ===
# Current term calculation
current_year = datetime.date.today().year
current_semester = 'Fall' if datetime.date.today().month > 6 else 'Spring'
current_term = f"{current_semester} {current_year}"

# File paths and output configuration
INCLUDE_GPA_IN_CSV = False  # Set to True to include individual GPAs in CSV files (for testing)
ACCEPTED_MAJORS = ["ECEB", "CMPE"]  # Accepted majors for filtering
REPORTS_DIR = "reports"
os.makedirs(REPORTS_DIR, exist_ok=True)

# Output file paths
SENIORS_FILE_PATH = f"{REPORTS_DIR}/seniors.csv"
JUNIORS_FILE_PATH = f"{REPORTS_DIR}/juniors.csv"
SOPHOMORES_FILE_PATH = f"{REPORTS_DIR}/sophomores.csv"
REPORT_FILE_PATH = f"{REPORTS_DIR}/report.log"
ERROR_LOG_FILE_PATH = f"{REPORTS_DIR}/errors.log"

# Cutoff percentages by grade level
SENIOR_CUTOFF = 0.30   # 30% for seniors
JUNIOR_CUTOFF = 0.25   # 25% for juniors
SOPHOMORE_CUTOFF = 0.20  # 20% for sophomores

# === SAMPLE DATA CONFIGURATION ===
SAMPLE_GPA_MEAN = 3.5
SAMPLE_GPA_STD = 0.5
SAMPLE_MAJOR_WEIGHTS = [0.46, 0.53, 0.01]  # [EE, CMPE, Other]
SAMPLE_GRADE_WEIGHTS = [0.32, 0.34, 0.34]  # [Sophomore, Junior, Senior]
SAMPLE_ECE_CREDITS_PER_SEMESTER = 3.5 # number of ECE classes per semester
SAMPLE_ECE_CREDITS_VARIATION = 2.0
SAMPLE_DEFAULT_COUNT = 2177

# === RUNTIME GLOBALS ===
ERROR_LOG = []  # Store errors for logging
OTHER_MAJORS = set()  # Track non-EE/CMPE majors encountered

# === COMPILED REGEX PATTERNS ===
FALL_SPRING_PATTERN = re.compile(r'Fall|Spring')
ECE_PATTERN = re.compile(r'ECE')
NAME_PUID_PATTERN = re.compile(r"(.*?)\s+(\d+)")


# === UTILITY FUNCTIONS ===

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


def collect_other_majors(students):
  """Collect other majors from a list of students for the global OTHER_MAJORS set"""
  for student in students:
    if hasattr(student, 'major') and student.major not in ACCEPTED_MAJORS:
      OTHER_MAJORS.add(student.major)


def classify_students_optimized(students):
  """Classify students in a single pass for better performance"""
  seniors = []
  juniors = []
  sophomores = []
  
  # Single pass classification
  for student in students:
    classification = student.classification
    if classification == "senior":
      seniors.append(student)
    elif classification == "junior":
      juniors.append(student)
    elif classification == "sophomore":
      sophomores.append(student)
    # Note: underclassmen are not included in any category
  
  return seniors, juniors, sophomores


def process_single_file(filename):
  """Process a single Excel file and return all students found in it"""
  students = []
  file_errors = []
  local_other_majors = set()  # Track other majors found in this process
  
  try:
    with pd.ExcelFile(filename, engine='calamine') as workbook:
      for sheet_name in workbook.sheet_names:
        try:
          # Parse only the first 10 columns
          sheet = workbook.parse(sheet_name, usecols='A:J')
          sheet.name = sheet_name
          
          student = Student(sheet)
          students.append(student)
          
          # Collect other majors from this student
          if hasattr(student, 'major') and student.major not in ACCEPTED_MAJORS:
            local_other_majors.add(student.major)
          
        except Exception as e:
          file_errors.append(("PARSING", f"Failed to parse student sheet: {str(e)}", None, sheet_name))
          
  except Exception as e:
    file_errors.append(("FILE", f"Failed to open file {filename}: {str(e)}", None, None))
    
  return students, file_errors, local_other_majors


# === MAIN PROCESSING FUNCTION ===


def main(use_sample_data=False, bymajor=True, use_multiprocessing=False, low_memory=False):
  seniors = list()
  juniors = list()
  sophomores = list()
  total_sheets = 0  # Initialize total_sheets for both sample and real data
  start = datetime.datetime.now()
  
  if use_sample_data:
    # Use sample data for testing
    print("Generating sample students for testing...")
    all_students = generate_sample_students()
    total_sheets = len(all_students)  # For sample data, count students as sheets
    
  else:
      # Choose between multiprocessing and sequential processing based on argument
      if use_multiprocessing:
        # Use multiprocessing at the file level for better performance
        filenames = os.listdir('.')
        xlsxfilter = lambda name: os.path.splitext(name)[1] == '.xlsx'
        workbook_filenames = list(filter(xlsxfilter, filenames))
        
        if len(workbook_filenames) == 0:
          raise RuntimeError("No .xlsx files in directory")
        
        print(f"Found {len(workbook_filenames)} Excel files to process")
        
        # Use multiprocessing to process files in parallel
        num_processes = min(mp.cpu_count(), len(workbook_filenames))
        print(f"Using {num_processes} processes for file processing")
        
        all_students = []
        
        with mp.Pool(processes=num_processes) as pool:
          # Process files in parallel with progress bar
          results = list(tqdm.tqdm(
            pool.imap(process_single_file, workbook_filenames),
            total=len(workbook_filenames),
            desc="Processing files"
          ))
        
        # Collect results and errors
        for students, file_errors, local_other_majors in results:
          all_students.extend(students)
          total_sheets += len(students)
          # Add errors to global error log
          for error_type, message, student_name, sheet_name in file_errors:
            log_error(error_type, message, student_name, sheet_name)
          # Merge other majors from this process
          OTHER_MAJORS.update(local_other_majors)
        
        # Classify students using optimized single-pass approach
        seniors, juniors, sophomores = classify_students_optimized(all_students)
      else:
        # Use optimized sequential processing for real Excel data
        wb_iter = FastWorkbookIterator()
        
        if low_memory:
          # Memory-efficient processing: classify students on-the-fly
          print("Using low-memory mode for large datasets...")
          seniors = []
          juniors = []
          sophomores = []
          parse_errors = []
          
          for sheet in tqdm.tqdm(wb_iter, desc="Processing sheets (low-memory)"):
            try:
              student = Student(sheet)
              # Immediate classification to avoid storing all students
              if student.classification == "senior":
                seniors.append(student)
              elif student.classification == "junior":
                juniors.append(student)
              elif student.classification == "sophomore":
                sophomores.append(student)
            except Exception as e:
              parse_errors.append((str(e), getattr(sheet, 'name', 'Unknown')))
          
          # Log parse errors
          for error_msg, sheet_name in parse_errors:
            log_error("PARSING", f"Failed to parse student sheet: {error_msg}", sheet_name=sheet_name)
            
          all_students = seniors + juniors + sophomores  # For compatibility with rest of code
          # Collect other majors for the global set
          collect_other_majors(all_students)
        else:
          # Standard processing with bulk validation
          all_students = []
          parse_errors = []
          
          # Bulk parsing without validation for speed
          for sheet in tqdm.tqdm(wb_iter, desc="Parsing sheets"):
            try:
              all_students.append(Student(sheet))
            except Exception as e:
              parse_errors.append((str(e), getattr(sheet, 'name', 'Unknown')))
          
          # Log parse errors
          for error_msg, sheet_name in parse_errors:
            log_error("PARSING", f"Failed to parse student sheet: {error_msg}", sheet_name=sheet_name)
          
          # Collect other majors for the global set
          collect_other_majors(all_students)
          
          # Classify all students using optimized single-pass approach
          seniors, juniors, sophomores = classify_students_optimized(all_students)
        
        total_sheets = len(wb_iter)
  end = datetime.datetime.now()
  
  # Handle classification for sample data mode
  if use_sample_data:
    # Classify sample students using optimized single-pass approach
    seniors, juniors, sophomores = classify_students_optimized(all_students)
    # Collect other majors for the global set
    collect_other_majors(all_students)

  processing_mode = "Multiprocessing" if use_multiprocessing else "Sequential"
  print(f"Parsing completed in {end - start} ({processing_mode} mode): {total_sheets/(end - start).total_seconds():.2f} sheets/sec")
  
  if use_sample_data:
    print(f"Generated {len(seniors) + len(juniors) + len(sophomores)} sample students")
  else:
    print(f"Successfully parsed {len(seniors) + len(juniors) + len(sophomores)} students from {total_sheets} sheets")
  totseniors = len(seniors) 
  totjuniors = len(juniors) 
  totsophomores = len(sophomores) 

  # Filter students based on bymajor parameter
  # Process each grade level with consolidated function
  filtered_seniors = process_grade_level(seniors, "seniors", totseniors, SENIOR_CUTOFF, bymajor)
  filtered_juniors = process_grade_level(juniors, "juniors", totjuniors, JUNIOR_CUTOFF, bymajor)
  filtered_sophomores = process_grade_level(sophomores, "sophomores", totsophomores, SOPHOMORE_CUTOFF, bymajor)

  # Generate report file
  generate_report(REPORT_FILE_PATH, filtered_seniors, filtered_juniors, filtered_sophomores, totseniors, totjuniors, totsophomores, use_sample_data, seniors, juniors, sophomores, bymajor)


# === STUDENT FILTERING FUNCTIONS ===


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
  """Filter students with at least 10 ECE credits"""
  return [student for student in students if student.num_ece_credits >= 10]

def take_only_top_percentile(students, top_percent):
  """Take top percentage of students (assumes already sorted by GPA)"""
  index_cutoff = math.ceil(len(students) * top_percent)
  return students[:index_cutoff]

def sort_students_by_gpa(students):
  """Sort students by GPA in descending order (modifies in-place)"""
  students.sort(key=lambda student: student.gpa, reverse=True)
  return students


# === OUTPUT AND REPORTING FUNCTIONS ===

def get_gpa_cutoff(students):
  if len(students) == 0:
    return float("NaN")
  return min(s.gpa for s in students)

def print_stats(class_string, students, totnum, by_major=False):
  """Print statistics for a class, optionally broken down by major"""
  if by_major:
    print(f"\n=== {class_string.upper()} STATISTICS ===")
    students_by_major = group_students_by_major(students)
    
    total_selected = len(students)
    print(f"Total {class_string} selected: {total_selected} out of {totnum}")
    
    for major, major_students in students_by_major.items():
      num_students = len(major_students)
      gpa_cutoff = get_gpa_cutoff(major_students)
      print(f"  {major}: {num_students} students, GPA cutoff: {gpa_cutoff:.2f}")
  else:
    # Original simple format
    fstr = "Number %10s Qualified: %3d  Out of: %3d  GPA cutoff: %.2f"
    nstudents = len(students)
    gpa = get_gpa_cutoff(students)
    print(fstr % (class_string, nstudents, totnum, gpa))

def write_student_list_to_file(filename, students, by_major=False):
  """Write student list to file, optionally organized by major group"""
  if by_major:
    students_by_major_group = group_students_by_major(students)
    
    with open(filename, "w") as outfile:
      # Write header based on GPA flag
      if INCLUDE_GPA_IN_CSV:
        outfile.write("Major Group,Actual Major,Last Name,First Name,PUID,GPA\n")
      else:
        outfile.write("Major Group,Actual Major,Last Name,First Name,PUID\n")

      for major_group in sorted(students_by_major_group.keys()):
        major_students = students_by_major_group[major_group]
        major_students = sort_alphabetically(major_students.copy())
        
        for student in major_students:
          if INCLUDE_GPA_IN_CSV:
            outline = f"{student.major_group},{student.major},{student.name},{student.puid},{student.gpa:.2f}\n"
          else:
            outline = f"{student.major_group},{student.major},{student.name},{student.puid}\n"
          outfile.write(outline)
  else:
    # Original simple format
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

def process_grade_level(students, grade_name, total_count, cutoff_percent, bymajor):
  """Process a single grade level: filter, print stats, and write to file"""
  # Determine output file path
  file_paths = {
    'seniors': SENIORS_FILE_PATH,
    'juniors': JUNIORS_FILE_PATH, 
    'sophomores': SOPHOMORES_FILE_PATH
  }
  
  # Filter students
  if bymajor:
    filtered_students = filter_students_by_major(students, cutoff_percent)
  else:
    filtered_students = filter_out_students(students, cutoff_percent)
  
  # Print statistics
  print_stats(grade_name, filtered_students, total_count, by_major=bymajor)
  
  # Write to file
  write_student_list_to_file(file_paths[grade_name], filtered_students, by_major=bymajor)
  
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


def write_error_log(filename):
  """Write detailed error information to a separate error log file"""
  with open(filename, "w") as error_file:
    error_file.write("=" * 60 + "\n")
    error_file.write("HKN RECRUITMENT FILTER - ERROR LOG\n")
    error_file.write(f"Generated on: {datetime.date.today()}\n")
    error_file.write(f"Current term: {current_term}\n")
    error_file.write("=" * 60 + "\n\n")
    
    if not ERROR_LOG:
      error_file.write("No errors encountered during processing.\n")
    else:
      error_file.write(f"Total errors logged: {len(ERROR_LOG)}\n\n")
      
      for i, error in enumerate(ERROR_LOG, 1):
        error_file.write(f"ERROR #{i}\n")
        error_file.write(f"  Type: {error['type']}\n")
        error_file.write(f"  Message: {error['message']}\n")
        if error['student_name']:
          error_file.write(f"  Student: {error['student_name']}\n")
        if error['sheet_name']:
          error_file.write(f"  Sheet: {error['sheet_name']}\n")
        error_file.write("\n")


def generate_report(filename, seniors, juniors, sophomores, totseniors, totjuniors, totsophomores, use_sample_data=False, unfiltered_seniors=None, unfiltered_juniors=None, unfiltered_sophomores=None, bymajor=True):
  """Generate a comprehensive report file with all statistics"""
  
  # Write detailed errors to separate error log file
  write_error_log(ERROR_LOG_FILE_PATH)
  
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
    
    # Simplified Errors and Concerns section
    outfile.write("ERRORS AND CONCERNS\n")
    outfile.write("-" * 20 + "\n")
    
    if not ERROR_LOG:
      outfile.write("No errors encountered during processing.\n\n")
    else:
      outfile.write(f"Total errors logged: {len(ERROR_LOG)}\n")
      outfile.write(f"Detailed error information can be found in: {ERROR_LOG_FILE_PATH}\n\n")
    
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
      for i, major in enumerate(OTHER_MAJORS, 1):
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


# === EXCEL PROCESSING CLASSES ===

class FastWorkbookIterator:
  """Optimized workbook iterator with faster Excel reading"""
  def __init__(self, path='.'):
    filenames = os.listdir(path)
    xlsxfilter = lambda name: os.path.splitext(name)[1] == '.xlsx'
    self.workbook_filenames = list(filter(xlsxfilter, filenames))
    if len(self.workbook_filenames) == 0:
      raise RuntimeError("No .xlsx files in directory")
    
    # Pre-scan for sheet names without loading the data
    self.sheet_info = []
    for filename in self.workbook_filenames:
      try:
        with pd.ExcelFile(filename, engine='calamine') as workbook:
          for sheet_name in workbook.sheet_names:
            self.sheet_info.append((filename, sheet_name))
      except Exception as e:
        print(f"Warning: Could not scan {filename}: {e}")
        continue
    
    self.current_index = 0
    self.current_workbook = None
    self.current_filename = None

  def __len__(self):
    return len(self.sheet_info)

  def __iter__(self):
    return self

  def __next__(self):
    if self.current_index >= len(self.sheet_info):
      raise StopIteration
    
    filename, sheet_name = self.sheet_info[self.current_index]
    
    # Only open a new file if we're switching files
    if self.current_filename != filename:
      if self.current_workbook:
        self.current_workbook.close()
      self.current_workbook = pd.ExcelFile(filename, engine='calamine')
      self.current_filename = filename
    
    # Parse only the specific sheet we need, using calamine engine and limiting to first 10 columns
    sheet = self.current_workbook.parse(sheet_name, engine='calamine', usecols='A:J')
    sheet.name = sheet_name  # Add sheet name for error reporting
    
    self.current_index += 1
    return sheet
  
  def __del__(self):
    if hasattr(self, 'current_workbook') and self.current_workbook:
      self.current_workbook.close()

# === STUDENT DATA PROCESSING CLASS ===

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
    
    # Set major group for filtering
    if self.major in ACCEPTED_MAJORS:
      self.major_group = self.major
    else:
      self.major_group = "Other"
      # Only add to global OTHER_MAJORS if we're not in multiprocessing mode
      # (multiprocessing mode will handle this separately in the main process)
      if mp.current_process().name == 'MainProcess':
        OTHER_MAJORS.add(self.major)

  def _get_identifying(self, sheet):
    try:
      # Use .iat for faster single value access
      name_puid_str = sheet.iat[4, 1]
      if pd.isna(name_puid_str) or not str(name_puid_str).strip():
        raise ValueError("Name/PUID field is empty or missing")
      
      # Use pre-compiled regex pattern
      matches = NAME_PUID_PATTERN.search(str(name_puid_str))
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
      # Use .iat for faster single value access
      gpa_value = sheet.iat[7, 9]
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
      maxrow = len(sheet) - 1
      num_fallspring_semesters = 0
      num_ece_credits = 0
      self.major = None  # Initialize major to None
      
      # More efficient approach: get the data we need in fewer operations
      if len(sheet.columns) > 0 and maxrow >= startrow:
        # Get the range of rows we need to examine
        end_row = min(maxrow + 1, len(sheet))
        
        # Pre-check if we have the columns we need
        has_major_col = len(sheet.columns) > 8
        has_credits_col = len(sheet.columns) > 9
        
        # Robust vectorized operations using pandas
        if end_row > startrow:
          # Extract the column data we need in bulk - more efficient slicing
          row_slice = slice(startrow, end_row)
          col0_data = sheet.iloc[row_slice, 0]
          
          # Convert to string only once and handle NaN values efficiently
          # Use .astype(str) which handles all data types, including NaN -> 'nan'
          col0_str = col0_data.astype(str).fillna('')
          
          # Early termination optimization: find current term and limit processing
          current_term_mask = col0_str == current_term
          if current_term_mask.any():
            # Get the first occurrence index and limit our processing
            current_term_idx = current_term_mask.idxmax() - startrow
            # Slice data up to current term for more efficient processing
            col0_str = col0_str.iloc[:current_term_idx + 1]
            # Update row_slice for consistent indexing
            row_slice = slice(startrow, startrow + current_term_idx + 1)
          
          # Use pre-compiled regex patterns for maximum efficiency
          fall_spring_mask = col0_str.str.contains(FALL_SPRING_PATTERN, na=False, regex=True)
          ece_mask = col0_str.str.contains(ECE_PATTERN, na=False, regex=True)
          
          # Count semesters with vectorized sum
          num_fallspring_semesters = int(fall_spring_mask.sum())
          
          # Streamlined most recent major extraction
          if has_major_col and fall_spring_mask.any():
            # Get major column values for the relevant row range
            major_series = sheet.iloc[row_slice, 8]
            # Filter to only fall/spring semesters and get the last valid one
            valid_majors = major_series[fall_spring_mask].dropna()
            if not valid_majors.empty:
              # Get the most recent (last) major
              self.major = str(valid_majors.iloc[-1]).strip()
          
          # Streamlined ECE credit summation
          if has_credits_col and ece_mask.any():
            # Get credits column values for the relevant row range
            credits_series = sheet.iloc[row_slice, 9]
            # Sum only credits from ECE rows after coercing to numeric
            # pd.to_numeric with errors='coerce' converts invalid values to NaN
            ece_credits_numeric = pd.to_numeric(credits_series[ece_mask], errors='coerce')
            # .sum() automatically ignores NaN values
            num_ece_credits = float(ece_credits_numeric.sum())
      
      self.num_ece_credits = num_ece_credits
      self.num_nonsummer_semesters = num_fallspring_semesters
      
      # Classify student by semester count
      if num_fallspring_semesters > 6:
        self.classification = "senior"
      elif num_fallspring_semesters > 4:
        self.classification = "junior"
      elif num_fallspring_semesters > 2:
        self.classification = "sophomore"
      else:
        self.classification = "underclassman"  # For students with 2 or fewer semesters
      
    except Exception as e:
      raise ValueError(f"Error parsing semester/credit data: {str(e)}")

  def _validate_data(self):
    """Validate parsed student data for reasonableness"""
    if not self.name:
      raise ValueError("Student name is missing or empty")
    
    if not self.puid:
      raise ValueError("Student PUID is missing or empty")
    
    if self.gpa is None:
      raise ValueError("Student GPA is missing")
    
    if self.gpa < 0 or self.gpa > 4:  # Allow some margin for different GPA scales
      log_error("VALIDATION", f"Unusual GPA value: {self.gpa}")
    
    if self.num_nonsummer_semesters < 0:
      raise ValueError("Invalid semester count")
    
    if self.num_ece_credits < 0:
      raise ValueError("Invalid ECE credits count")
    
    if not self.classification:
      raise ValueError("Student classification is missing")
    
  def is_senior(self):
    """Check if student is classified as a senior"""
    return self.classification == "senior"
  
  def is_junior(self):
    """Check if student is classified as a junior"""
    return self.classification == "junior"
  
  def is_sophomore(self):
    """Check if student is classified as a sophomore"""
    return self.classification == "sophomore"
  
  def is_underclassman(self):
    """Check if student is classified as an underclassman (<=2 semesters)"""
    return self.classification == "underclassman"
    
  def __repr__(self):
    return "Student({}, classification={}, major={}, major_group={}, sems={}, creds={}, gpa={})".format(
      self.name, self.classification, self.major, self.major_group, self.num_nonsummer_semesters,
      self.num_ece_credits, self.gpa
    )


# === SAMPLE DATA GENERATION ===

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
  
  majors = ACCEPTED_MAJORS + ["Other"]  # Include 'Other' for non-EE/CMPE majors
  
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
    if student.major in ACCEPTED_MAJORS:
      student.major_group = student.major
    else:
      student.major_group = "Other"
    
    # Generate semester count using global weights
    semester_category = random.choices(['sophomore', 'junior', 'senior'], weights=SAMPLE_GRADE_WEIGHTS)[0]
    
    if semester_category == 'sophomore':
      student.num_nonsummer_semesters = random.randint(3, 4)
      student.classification = 'sophomore'
    elif semester_category == 'junior':
      student.num_nonsummer_semesters = random.randint(5, 6)
      student.classification = 'junior'
    else:  # senior
      student.num_nonsummer_semesters = random.randint(7, 10)
      student.classification = 'senior'
    
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
  use_multiprocessing = "--multiprocessing" in sys.argv or "--mp" in sys.argv  # Enable multiprocessing with --multiprocessing or --mp
  low_memory = "--low-memory" in sys.argv or "--lm" in sys.argv  # Enable low-memory mode
  
  if "--help" in sys.argv or "-h" in sys.argv:
    print("HKN Student Recruitment Filter")
    print("Usage: python hkn_student_parser.py [options]")
    print("Options:")
    print("  --sample      Use sample data instead of Excel files")
    print("  --no-major    Disable major-based filtering (use grade-level filtering only)")
    print("  --multiprocessing, --mp    Enable multiprocessing for file-level parallel processing")
    print("  --low-memory, --lm         Enable low-memory mode for large datasets")
    print("  --help, -h    Show this help message")
    print("")
    print("Processing modes:")
    print("  1. Sequential: Direct Excel processing with calamine engine (default)")
    print("  2. Multiprocessing: File-level parallel processing (use --mp)")
    print("  3. Low-memory: Memory-efficient processing for large datasets (use --lm)")
    print("")
    print("Performance tips:")
    print("  - Use --mp for multiple large Excel files")
    print("  - Use --lm if you encounter memory issues with large datasets")
    print("  - Combine --mp and --lm for maximum efficiency with large multi-file datasets")
    sys.exit(0)
  
  # Run the main function with the parsed arguments
  main(use_sample_data=use_sample_data, bymajor=bymajor, use_multiprocessing=use_multiprocessing, low_memory=low_memory)
