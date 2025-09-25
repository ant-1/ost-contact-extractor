# OST Contact Extractor

A Python script to extract contact information from Microsoft Outlook OST/PST files and export them to CSV format.

## Overview

This tool helps you extract contact data from Outlook data files (OST/PST) and convert them into a more accessible CSV format. It's particularly useful for:
- Migrating contacts between email systems
- Backing up contact information
- Analyzing contact data
- Converting Outlook contacts to other formats

## Features

- ✅ Extract contacts from OST and PST files
- ✅ Export to CSV format with comprehensive contact fields
- ✅ Handle multiple contact folders
- ✅ Support for business and personal contact information
- ✅ Command-line interface for easy automation
- ✅ Robust error handling and logging

## Requirements

- Python 3.6 or higher
- `pst-utils` library

## Installation

1. Clone or download this repository
2. Install the required dependency:

```bash
pip install pst-utils
```

## Usage

### Basic Usage

Extract contacts from an OST file with default output filename:

```bash
python extract_contact.py your_file.ost
```

This will create a CSV file named `your_file_contacts.csv`

### Custom Output File

Specify a custom output filename:

```bash
python extract_contact.py your_file.ost my_contacts.csv
```

### Examples

```bash
# Extract from outlook.ost to outlook_contacts.csv
python extract_contact.py outlook.ost

# Extract from data.pst to business_contacts.csv
python extract_contact.py data.pst business_contacts.csv

# Extract from archive.ost to specific location
python extract_contact.py archive.ost /path/to/output/contacts.csv
```

## Extracted Contact Fields

The script extracts the following contact information when available:

### Basic Information
- Display Name
- Given Name (First Name)
- Surname (Last Name)
- Email Address
- Company Name
- Job Title
- Department

### Phone Numbers
- Business Telephone
- Home Telephone
- Mobile Telephone

### Business Address
- Business Address Street
- Business City
- Business State
- Business Postal Code
- Business Country

### Home Address
- Home Address Street
- Home City
- Home State
- Home Postal Code
- Home Country

## Output Format

The script generates a CSV file with UTF-8 encoding. Each row represents a contact, and columns contain the contact fields. Empty fields are included but left blank.

Example CSV output:
```csv
display_name,email_address,given_name,surname,company_name,job_title,business_telephone
John Doe,john.doe@company.com,John,Doe,Acme Corp,Manager,+1-555-0123
Jane Smith,jane.smith@email.com,Jane,Smith,,,+1-555-0456
```

## Error Handling

The script includes comprehensive error handling:
- Validates input file existence
- Handles corrupted or inaccessible OST/PST files
- Manages missing contact fields gracefully
- Provides informative error messages

## Troubleshooting

### Common Issues

**"pst_utils not available" error:**
```bash
pip install pst-utils
```

**"File not found" error:**
- Verify the OST/PST file path is correct
- Ensure you have read permissions for the file

**"No contacts found" message:**
- The OST/PST file may not contain contact folders
- Contacts might be stored in non-standard folder names
- The file might be corrupted or password-protected

### Supported File Types

- `.ost` - Outlook Offline Storage Table files
- `.pst` - Outlook Personal Storage Table files

### Limitations

- Password-protected files are not supported
- Some encrypted OST files may not be accessible
- Very large files may take significant time to process

## Script Structure

```
extract_contact.py
├── extract_contacts_from_ost()    # Main extraction function
├── extract_contact_info()         # Individual contact processing
├── save_contacts_to_csv()         # CSV export functionality
└── main()                         # Command-line interface
```

## Development

### Running Tests

To test the script with your OST file:

```bash
python extract_contact.py test_file.ost
```

### Customization

You can modify the script to:
- Add additional contact fields
- Change CSV output format
- Add filtering options
- Support other output formats (JSON, XML, etc.)

## Contributing

Feel free to submit issues, feature requests, or pull requests to improve this tool.

## License

This project is open source and available under the [MIT License](LICENSE).

## Support

If you encounter issues:
1. Check the troubleshooting section above
2. Verify your Python and pst-utils versions
3. Ensure the OST/PST file is not corrupted
4. Check file permissions

## Changelog

### v1.0.0
- Initial release
- Basic contact extraction functionality
- CSV export support
- Command-line interface