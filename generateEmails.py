import requests
import time
import random
import csv
import os
import glob
from bs4 import BeautifulSoup
from tqdm import tqdm

def email_from_name(name):
    """
    Generates an email address from a person's name based on Purdue's email format.
    
    Args:
        name (str): The full name of the person.
    
    Returns:
        list: List of possible email addresses.
    """
    url = "https://www.purdue.edu/directory/"
    payload = {
        "SearchString": name
    }
    time.sleep(random.uniform(1, 2)) # this is to be compliant with robot.txt

    
    try:
        response = requests.post(url, data=payload)
        response.raise_for_status()
        html_content = response.text
    except requests.exceptions.RequestException as e:
        print(f"An error occurred during search: {e}")
        return []
    
    soup = BeautifulSoup(html_content, 'html.parser')
    results_section = soup.find('section', id='results')

    if not results_section:
        print("No 'results' section found in the HTML.")
        return []

    # Find all list items (<li>) within the results section, each representing a person
    person_entries = results_section.find_all('li')

    possible_emails = []

    for entry in person_entries:
        person_info = {}

        # Get the person's name (h2 tag within the table header)
        name_tag = entry.find('h2', class_='cn-name')
        if name_tag:
            person_info['Name'] = name_tag.get_text(strip=True)

        # Find the main table containing details (class="more")
        main_table = entry.find('table', class_='more')
        if main_table:
            # Iterate through table rows in tbody for initial details
            for row in main_table.find('tbody').find_all('tr'):
                header = row.find('th')
                value = row.find('td')
                if header and value:
                    key = header.get_text(strip=True).replace(' ', '_').replace('.', '').lower()
                    person_info[key] = value.get_text(strip=True)
        
        # Find the hidden table for additional details (class="hide")
        hidden_div = entry.find('div', class_='hide')
        if hidden_div:
            hidden_table = hidden_div.find('table')
            if hidden_table:
                for row in hidden_table.find_all('tr'):
                    header = row.find('th')
                    value = row.find('td')
                    if header and value:
                        key = header.get_text(strip=True).replace(' ', '_').replace('.', '').lower()
                        person_info[key] = value.get_text(strip=True)

        # Only add if we found some data and can verify the name match
        if person_info and 'alias' in person_info:
            directory_name = person_info.get('Name', '').strip()
            
            # Compare names to ensure we have the right person
            if name_matches(name, directory_name):
                possible_emails.append(person_info['alias'].lower() + "@purdue.edu")
    if len(possible_emails) > 1:
        print(f"Multiple emails found for {name}: {possible_emails}")

    return possible_emails


def name_matches(search_name, directory_name):
    """
    Check if the searched name matches the directory name.
    Handles different name formats and variations.
    
    Args:
        search_name (str): The name we searched for
        directory_name (str): The name returned from directory
    
    Returns:
        bool: True if names match, False otherwise
    """
    if not search_name or not directory_name:
        return False
    
    # Convert to lowercase for comparison
    search_lower = search_name.lower().strip()
    directory_lower = directory_name.lower().strip()
    
    # Direct match
    if search_lower == directory_lower:
        return True
    
    # Split names into parts for fuzzy matching
    search_parts = [part.strip() for part in search_lower.replace(',', ' ').split() if part.strip()]
    directory_parts = [part.strip() for part in directory_lower.replace(',', ' ').split() if part.strip()]
    
    # Check if all significant parts of search name are in directory name
    if len(search_parts) >= 2:  # First and last name minimum
        matches = 0
        for part in search_parts:
            if len(part) > 1:  # Skip single letters/initials
                for dir_part in directory_parts:
                    if part == dir_part or dir_part == part:
                        matches += 1
                        break
        
        # Consider it a match if at least 2 name parts match
        return matches >= min(2, len(search_parts))
    
    return False


def process_csv_files():
    """
    Process all CSV files in the reports folder and generate emails.csv
    """
    reports_dir = "reports"
    output_file = "emails.csv"
    
    # Check if reports directory exists
    if not os.path.exists(reports_dir):
        print(f"Error: {reports_dir} directory not found!")
        return
    
    # Find all CSV files in reports directory
    csv_files = glob.glob(os.path.join(reports_dir, "*.csv"))
    
    # Filter out the output file to avoid processing it
    csv_files = [f for f in csv_files if os.path.basename(f) != output_file]
    
    if not csv_files:
        print(f"No CSV files found in {reports_dir} directory!")
        return
    
    print(f"Found {len(csv_files)} CSV files to process:")
    for file in csv_files:
        print(f"  - {os.path.basename(file)}")
    
    # Prepare output data
    output_data = []
    
    # Count total rows for progress bar
    total_rows = 0
    for csv_file in csv_files:
        try:
            with open(csv_file, 'r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                total_rows += sum(1 for _ in reader)
        except:
            pass
    
    print(f"\nTotal students to process: {total_rows}")
    print(f"Expected Time: {total_rows * 1.5 / 60} minutes (assuming 3 seconds per student on average)")

    # Process each CSV file with progress bar
    with tqdm(total=total_rows, desc="Processing students", unit="student") as pbar:
        for csv_file in csv_files:
            source_file = os.path.basename(csv_file)
            pbar.set_description(f"Processing {source_file}")
            
            try:
                with open(csv_file, 'r', encoding='utf-8') as file:
                    reader = csv.DictReader(file)
                    
                    # Verify expected columns exist
                    expected_columns = ['Major Group', 'Actual Major', 'Last Name', 'First Name', 'PUID']
                    if not all(col in reader.fieldnames for col in expected_columns):
                        tqdm.write(f"Warning: {source_file} missing expected columns. Expected: {expected_columns}")
                        tqdm.write(f"Found: {reader.fieldnames}")
                        continue
                    
                    for row in reader:
                        # Extract data
                        actual_major = row.get('Actual Major', '').strip()
                        last_name = row.get('Last Name', '').strip()
                        first_name = row.get('First Name', '').strip()
                        
                        # Skip if essential data is missing
                        if not last_name and not first_name:
                            tqdm.write(f"  Skipping: Missing name data")
                            pbar.update(1)
                            continue
                        
                        # Concatenate name as "First Last" (changed order)
                        full_name = f"{first_name} {last_name}".strip()
                        pbar.set_postfix_str(f"Searching: {full_name}")
                        
                        # Get possible emails
                        try:
                            possible_emails = email_from_name(full_name)
                            
                            # Convert list of emails to string (semicolon separated)
                            emails_str = "; ".join(possible_emails) if possible_emails else "Not found"
                            
                            # Add to output data
                            output_data.append({
                                'Actual Major': actual_major,
                                'Source File': source_file,
                                'Last Name': last_name,
                                'First Name': first_name,
                                'Emails': emails_str
                            })
                            
                        except Exception as e:
                            tqdm.write(f"    Error processing {full_name}: {e}")
                            output_data.append({
                                'Actual Major': actual_major,
                                'Source File': source_file,
                                'Last Name': last_name,
                                'First Name': first_name,
                                'Emails': f"Error: {str(e)}"
                            })
                        
                        pbar.update(1)
                    
            except Exception as e:
                tqdm.write(f"Error reading {source_file}: {e}")
                continue
    
    # Write output CSV
    if output_data:
        try:
            with open(os.path.join(reports_dir, output_file), 'w', newline='', encoding='utf-8') as file:
                fieldnames = ['Actual Major', 'Source File', 'Last Name', 'First Name', 'Emails']
                writer = csv.DictWriter(file, fieldnames=fieldnames)
                
                writer.writeheader()
                writer.writerows(output_data)
            
            print(f"\nSuccess! Generated {output_file} with {len(output_data)} entries.")
            
        except Exception as e:
            print(f"Error writing output file: {e}")
    else:
        print("No data to write to output file.")


if __name__ == "__main__":
    process_csv_files()
