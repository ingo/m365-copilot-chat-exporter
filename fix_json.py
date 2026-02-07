#!/usr/bin/env python3
import json
import re
import sys

def clean_control_characters(text):
    """Remove or replace invalid control characters in JSON strings."""
    # Remove control characters except for whitespace (tab, newline, carriage return)
    # Control characters are ASCII 0-31 excluding tab(9), newline(10), carriage return(13)
    cleaned = ''
    for char in text:
        code = ord(char)
        if code < 32 and char not in ['\t', '\n', '\r']:
            # Skip invalid control characters
            continue
        cleaned += char
    return cleaned

def fix_json_file(input_file, output_file):
    """Read JSON file, clean it, and write to output."""
    try:
        # Read the file as raw text first
        with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
            content = f.read()
        
        print(f'Original file size: {len(content)} characters')
        
        # Clean control characters
        cleaned_content = clean_control_characters(content)
        print(f'Cleaned file size: {len(cleaned_content)} characters')
        print(f'Removed {len(content) - len(cleaned_content)} invalid characters')
        
        # Validate it's valid JSON
        try:
            data = json.loads(cleaned_content)
            print(f'✓ JSON is valid after cleaning')
            print(f'✓ Contains {len(data)} conversations')
        except json.JSONDecodeError as e:
            print(f'✗ Still has JSON errors: {e}')
            print(f'  Line: {e.lineno}, Column: {e.colno}')
            return False
        
        # Write cleaned content
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(cleaned_content)
        
        print(f'\n✓ Cleaned JSON written to {output_file}')
        return True
        
    except Exception as e:
        print(f'Error: {e}')
        return False

if __name__ == '__main__':
    input_file = 'copilot_conversations.json'
    output_file = 'copilot_conversations_fixed.json'
    
    if fix_json_file(input_file, output_file):
        print('\nSuccess! You can now use the fixed file.')
        print(f'To replace the original: mv {output_file} {input_file}')
    else:
        print('\nFailed to fix the JSON file.')
        sys.exit(1)
