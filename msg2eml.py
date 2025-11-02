#!/usr/bin/env python3
"""
MSG to EML Converter
Converts Outlook MSG files to EML format
"""

import argparse
import sys
from pathlib import Path
import extract_msg


def convert_msg_to_eml(msg_path, output_path=None):
    """
    Convert a single MSG file to EML format.
    
    Args:
        msg_path: Path to the input MSG file
        output_path: Path for the output EML file (optional)
    
    Returns:
        Path to the created EML file or None on error
    """
    try:
        msg_path = Path(msg_path)
        
        if not msg_path.exists():
            print(f"Error: File '{msg_path}' not found", file=sys.stderr)
            return None
        
        if not msg_path.suffix.lower() == '.msg':
            print(f"Warning: File '{msg_path}' doesn't have .msg extension", file=sys.stderr)
        
        # Open the MSG file
        msg = extract_msg.Message(str(msg_path))
        
        # Determine output path
        if output_path is None:
            output_path = msg_path.with_suffix('.eml')
        else:
            output_path = Path(output_path)
        
        # Create output directory if it doesn't exist
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Convert to email.message.Message object and save as EML
        email_msg = msg.asEmailMessage()
        with open(output_path, 'wb') as f:
            f.write(email_msg.as_bytes())
        
        msg.close()
        
        print(f"✓ Converted: {msg_path} -> {output_path}")
        return output_path
        
    except Exception as e:
        print(f"Error converting '{msg_path}': {str(e)}", file=sys.stderr)
        return None


def convert_directory(input_dir, output_dir=None, recursive=False):
    """
    Convert all MSG files in a directory.
    
    Args:
        input_dir: Directory containing MSG files
        output_dir: Directory for output EML files (optional)
        recursive: Whether to search subdirectories
    
    Returns:
        Tuple of (success_count, failure_count)
    """
    input_dir = Path(input_dir)
    
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Error: Directory '{input_dir}' not found", file=sys.stderr)
        return (0, 0)
    
    # Find all MSG files
    pattern = '**/*.msg' if recursive else '*.msg'
    msg_files = list(input_dir.glob(pattern))
    
    if not msg_files:
        print(f"No MSG files found in '{input_dir}'", file=sys.stderr)
        return (0, 0)
    
    print(f"Found {len(msg_files)} MSG file(s)")
    
    success_count = 0
    failure_count = 0
    
    for msg_file in msg_files:
        if output_dir:
            # Preserve directory structure if recursive
            rel_path = msg_file.relative_to(input_dir)
            output_path = Path(output_dir) / rel_path.with_suffix('.eml')
        else:
            output_path = msg_file.with_suffix('.eml')
        
        result = convert_msg_to_eml(msg_file, output_path)
        if result:
            success_count += 1
        else:
            failure_count += 1
    
    return (success_count, failure_count)


def main():
    parser = argparse.ArgumentParser(
        description='Convert Outlook MSG files to EML format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s message.msg                    # Convert single file to message.eml
  %(prog)s message.msg -o output.eml      # Convert to specific output file
  %(prog)s -d input_folder                # Convert all MSG files in folder
  %(prog)s -d input_folder -o output_dir  # Convert to output directory
  %(prog)s -d input_folder -r             # Recursively process subfolders
        """
    )
    
    parser.add_argument(
        'input',
        nargs='?',
        help='Input MSG file to convert'
    )
    
    parser.add_argument(
        '-o', '--output',
        help='Output EML file or directory'
    )
    
    parser.add_argument(
        '-d', '--directory',
        help='Process all MSG files in directory'
    )
    
    parser.add_argument(
        '-r', '--recursive',
        action='store_true',
        help='Process subdirectories recursively (use with -d)'
    )
    
    args = parser.parse_args()
    
    # Validate arguments
    if not args.input and not args.directory:
        parser.print_help()
        sys.exit(1)
    
    if args.input and args.directory:
        print("Error: Cannot specify both input file and directory", file=sys.stderr)
        sys.exit(1)
    
    # Process directory
    if args.directory:
        success, failure = convert_directory(
            args.directory,
            args.output,
            args.recursive
        )
        print(f"\nCompleted: {success} successful, {failure} failed")
        sys.exit(0 if failure == 0 else 1)
    
    # Process single file
    if args.input:
        result = convert_msg_to_eml(args.input, args.output)
        sys.exit(0 if result else 1)


if __name__ == '__main__':
    main()
