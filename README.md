# HKN Beta Chapter Recruitment Filter

An automated recruitment management system for HKN Beta Chapter that processes ECE student data to identify eligible candidates and generate email contact lists.

## Features

- **Student Data Processing**: Automatically processes Cognos ECE student reports to filter eligible candidates
- **Multi-Mode Processing**: Supports sequential, multiprocessing, and low-memory modes for different dataset sizes
- **Major-Based Filtering**: Intelligent filtering by major groups (EE, CMPE, Other) with customizable cutoff percentages
- **Email Generation**: Automated email lookup using Purdue Directory with intelligent name matching
- **Comprehensive Reporting**: Detailed reports with statistics, error logging, and major breakdowns
- **Automated Setup**: PowerShell script for complete environment setup and execution

## Quick Start

### Option 1: Automated Setup (Recommended)

```powershell
# Run the all-in-one setup script (Windows PowerShell)
.\run.ps1
```

### Option 2: Manual Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Run with Excel files in current directory
python hkn_student_parser.py

# Or test with sample data
python hkn_student_parser.py --sample
```

## Setup Requirements

- **Python 3.7+** with pip
- **Git** (for automated setup)
- **Cognos ECE student reports** in Excel format (.xlsx)

## Dependencies

All dependencies are automatically installed via `requirements.txt`:

- `pandas` - Data processing and Excel file handling
- `numpy` - Numerical operations
- `tqdm` - Progress bars for long-running operations
- `python-calamine` - Fast Excel reading engine
- `requests` - HTTP requests for email lookup
- `beautifulsoup4` - HTML parsing for directory scraping

## Usage

### Basic Operation

1. **Obtain Student Data**: Run Cognos 'record copy - PUID entry' report for all ECE students

   - Process in batches (Cognos limit: 1000 entries per report)
   - Save all Excel files with `.xlsx` extension

2. **Place Files**: Put all Excel files in the same directory as the script

3. **Run Processing**: Execute the main script

   ```bash
   python hkn_student_parser.py
   ```

4. **Review Output**: Check the generated files in the `reports/` directory

### Command Line Options

```bash
python hkn_student_parser.py [options]

Options:
  --sample                    Use sample data instead of Excel files
  --no-major                  Disable major-based filtering
  --multiprocessing, --mp     Enable parallel file processing
  --low-memory, --lm          Enable memory-efficient mode
  --help, -h                  Show help message
```

### Processing Modes

1. **Sequential** (default): Direct Excel processing with optimized single-threaded performance
2. **Multiprocessing** (`--mp`): Parallel file processing for multiple large Excel files
3. **Low-Memory** (`--lm`): Memory-efficient processing for very large datasets

### Advanced Usage

```bash
# For multiple large files with parallel processing
python hkn_student_parser.py --mp

# For very large datasets with memory constraints
python hkn_student_parser.py --lm

# Combined for maximum efficiency with large multi-file datasets
python hkn_student_parser.py --mp --lm

# Test without real data
python hkn_student_parser.py --sample

# Disable major-specific filtering (use overall GPA ranking)
python hkn_student_parser.py --no-major
```

## Email Generation

Generate email addresses for recruited students:

```bash
# Process CSV files in reports/ directory to generate emails
python generateEmails.py
```

Features:

- Automatic CSV file discovery in `reports/` directory
- Intelligent name matching with fuzzy logic
- Progress tracking with visual progress bars
- Purdue Directory integration with rate limiting
- Outputs consolidated `emails.csv` file

## Output Files

### Generated in `reports/` directory:

- **`seniors.csv`** - Eligible senior candidates
- **`juniors.csv`** - Eligible junior candidates
- **`sophomores.csv`** - Eligible sophomore candidates
- **`report.log`** - Comprehensive processing report with statistics
- **`errors.log`** - Detailed error information and debugging data
- **`emails.csv`** - Email addresses for recruited students (after email generation)

### File Formats

CSV files include columns based on filtering mode:

- **Major-based filtering**: `Major Group, Actual Major, Last Name, First Name, PUID`
- **Grade-level filtering**: `Last Name, First Name, PUID`
- **Email output**: `First Name, Last Name, Email`

## Configuration

### Filter Criteria (in `hkn_student_parser.py`):

```python
# Cutoff percentages by grade level
SENIOR_CUTOFF = 0.30      # Top 30% of seniors
JUNIOR_CUTOFF = 0.25      # Top 25% of juniors
SOPHOMORE_CUTOFF = 0.20   # Top 20% of sophomores

# Accepted majors for major-based filtering
ACCEPTED_MAJORS = ["ECEB", "CMPE"]  # EE and Computer Engineering

# Additional filters applied automatically:
# - Minimum 10 ECE credits
# - GPA-based ranking within each major group
```

### Major Groupings:

- **ECEB**: Electrical Engineering students
- **CMPE**: Computer Engineering students
- **Other**: All other majors (grouped together for filtering)

## Performance Optimization

### For Large Datasets:

- **Multiple Excel files**: Use `--mp` for parallel processing
- **Memory constraints**: Use `--lm` for reduced memory usage
- **Very large datasets**: Combine `--mp --lm` for optimal performance

### Performance Metrics:

- Typical processing speed: ~200 sheets/second
- Memory usage scales with dataset size
- Multiprocessing scales with available CPU cores

## Error Handling

- **Automatic error logging**: All errors saved to `reports/errors.log`
- **Graceful failure handling**: Processing continues despite individual student parsing errors
- **Detailed error context**: Includes student names, sheet names, and error types
- **Git repository handling**: Automatic safe directory configuration for various environments

## Troubleshooting

### Common Issues:

1. **No .xlsx files found**: Ensure Excel files are in the same directory as the script
2. **Memory errors**: Use `--lm` flag for large datasets
3. **Git ownership errors**: Automatically handled by the PowerShell setup script
4. **Permission errors**: Run PowerShell as Administrator if needed

### Getting Help:

- Run `python hkn_student_parser.py --help` for command options
- Check `reports/errors.log` for detailed error information
- Submit issues on GitHub with error messages and context
- Contact HKN Beta chapter representatives for assistance

## Development

### Project Structure:

```
recruitment-filter/
├── hkn_student_parser.py   # Main processing script
├── generateEmails.py       # Email generation utility
├── run.ps1                # Automated setup script
├── requirements.txt       # Python dependencies
├── reports/               # Output directory
│   ├── *.csv             # Student lists
│   ├── report.log        # Processing report
│   └── errors.log        # Error details
└── README.md             # This file
```

### Key Features Implementation:

- **Optimized Excel Processing**: Uses calamine engine for fast Excel reading
- **Vectorized Operations**: Pandas-based operations for maximum performance
- **Memory Management**: Configurable memory usage patterns
- **Error Recovery**: Robust error handling with detailed logging
- **Progress Tracking**: Visual progress bars for long operations

## License

This project is maintained by HKN Beta Chapter for internal recruitment management.
