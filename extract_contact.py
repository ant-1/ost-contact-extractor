#!/usr/bin/env python3
"""
Extract contacts from OST/PST files
This script extracts contact information from Outlook data files
"""

import os
import sys
import csv
from pathlib import Path

try:
    from pst_utils import PstFile
except ImportError:
    print("Warning: pst_utils not available. Install with: pip install pst-utils")
    sys.exit(1)

def extract_contacts_from_ost(ost_file_path, output_csv=None):
    """
    Extract contacts from OST/PST file
    
    Args:
        ost_file_path (str): Path to the OST/PST file
        output_csv (str): Optional output CSV file path
    """
    
    if not os.path.exists(ost_file_path):
        print(f"Error: File {ost_file_path} not found")
        return
    
    try:
        # Open the PST/OST file
        with PstFile(ost_file_path) as pst:
            print(f"Processing: {ost_file_path}")
            
            contacts = []
            
            # Navigate through the folder structure to find contacts
            for folder in pst.folder_generator():
                folder_name = folder.name.lower() if folder.name else ""
                
                # Look for contact folders (common names)
                if any(keyword in folder_name for keyword in ['contact', 'address', 'people']):
                    print(f"Found contacts folder: {folder.name}")
                    
                    # Extract messages/contacts from this folder
                    for message in folder.messages:
                        contact_info = extract_contact_info(message)
                        if contact_info:
                            contacts.append(contact_info)
            
            # Save to CSV if output path provided
            if output_csv and contacts:
                save_contacts_to_csv(contacts, output_csv)
                print(f"Contacts saved to: {output_csv}")
            
            return contacts
            
    except Exception as e:
        print(f"Error processing file: {e}")
        return []

def extract_contact_info(message):
    """
    Extract contact information from a message object
    
    Args:
        message: Message object from PST file
    
    Returns:
        dict: Contact information dictionary
    """
    
    contact = {}
    
    try:
        # Basic contact fields
        contact['display_name'] = getattr(message, 'display_name', '')
        contact['given_name'] = getattr(message, 'given_name', '')
        contact['surname'] = getattr(message, 'surname', '')
        contact['email_address'] = getattr(message, 'email_address', '')
        contact['business_telephone'] = getattr(message, 'business_telephone_number', '')
        contact['home_telephone'] = getattr(message, 'home_telephone_number', '')
        contact['mobile_telephone'] = getattr(message, 'mobile_telephone_number', '')
        contact['company_name'] = getattr(message, 'company_name', '')
        contact['job_title'] = getattr(message, 'title', '')
        contact['department'] = getattr(message, 'department_name', '')
        
        # Address information
        contact['business_address'] = getattr(message, 'business_address_street', '')
        contact['business_city'] = getattr(message, 'business_address_city', '')
        contact['business_state'] = getattr(message, 'business_address_state', '')
        contact['business_postal_code'] = getattr(message, 'business_address_postal_code', '')
        contact['business_country'] = getattr(message, 'business_address_country', '')
        
        contact['home_address'] = getattr(message, 'home_address_street', '')
        contact['home_city'] = getattr(message, 'home_address_city', '')
        contact['home_state'] = getattr(message, 'home_address_state', '')
        contact['home_postal_code'] = getattr(message, 'home_address_postal_code', '')
        contact['home_country'] = getattr(message, 'home_address_country', '')
        
        # Only return contact if it has meaningful data
        if any([contact['display_name'], contact['email_address'], contact['given_name'], contact['surname']]):
            return contact
            
    except Exception as e:
        print(f"Error extracting contact info: {e}")
    
    return None

def save_contacts_to_csv(contacts, output_file):
    """
    Save contacts to CSV file
    
    Args:
        contacts (list): List of contact dictionaries
        output_file (str): Output CSV file path
    """
    
    if not contacts:
        print("No contacts to save")
        return
    
    # Get all unique field names
    fieldnames = set()
    for contact in contacts:
        fieldnames.update(contact.keys())
    
    fieldnames = sorted(list(fieldnames))
    
    try:
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(contacts)
        
        print(f"Successfully saved {len(contacts)} contacts to {output_file}")
        
    except Exception as e:
        print(f"Error saving to CSV: {e}")

def main():
    """Main function to handle command line arguments"""
    
    if len(sys.argv) < 2:
        print("Usage: python extract_contact.py <ost_file> [output_csv]")
        print("Example: python extract_contact.py data.ost contacts.csv")
        sys.exit(1)
    
    ost_file = sys.argv[1]
    output_csv = sys.argv[2] if len(sys.argv) > 2 else None
    
    # Generate default output filename if not provided
    if not output_csv:
        base_name = Path(ost_file).stem
        output_csv = f"{base_name}_contacts.csv"
    
    print(f"Extracting contacts from: {ost_file}")
    contacts = extract_contacts_from_ost(ost_file, output_csv)
    
    if contacts:
        print(f"Extracted {len(contacts)} contacts")
    else:
        print("No contacts found")

if __name__ == "__main__":
    main()