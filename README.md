# MSG to EML Converter

A simple console application that converts Outlook MSG files to EML format.

## Installation

1. Install Python 3.7 or higher
2. Create a virtual environment (recommended):

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

**Note:** If using a virtual environment, activate it first:

```bash
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

### Convert a single file

```bash
python msg2eml.py message.msg
```

This creates `message.eml` in the same directory.

### Convert with specific output name

```bash
python msg2eml.py message.msg -o output.eml
```

### Convert all MSG files in a directory

```bash
python msg2eml.py -d /path/to/folder
```

### Convert directory with output directory

```bash
python msg2eml.py -d input_folder -o output_folder
```

### Recursively process subdirectories

```bash
python msg2eml.py -d /path/to/folder -r
```

## Options

- `input` - Path to MSG file to convert
- `-o, --output` - Output EML file or directory
- `-d, --directory` - Process all MSG files in directory
- `-r, --recursive` - Process subdirectories recursively (use with -d)

## Examples

```bash
# Convert single file
python msg2eml.py email.msg

# Convert with custom output
python msg2eml.py email.msg -o converted/email.eml

# Convert all MSG files in current directory
python msg2eml.py -d .

# Convert all MSG files recursively
python msg2eml.py -d ./emails -r -o ./converted
```

## Requirements

- Python 3.7+
- extract-msg library (automatically installed from requirements.txt)

## How It Works

The application uses the `extract-msg` library to:

1. Parse Outlook MSG files (Microsoft's proprietary format)
2. Extract email headers, body, attachments, and metadata
3. Convert to standard EML format (RFC 822/MIME format)
4. Save as `.eml` files that can be opened by most email clients

The converted EML files are fully compatible with:

- Apple Mail
- Mozilla Thunderbird
- Gmail (via import)
- Most other email clients that support standard EML format

## Troubleshooting

### "No module named 'extract_msg'"

Make sure you've installed the dependencies and activated your virtual environment:

```bash
source venv/bin/activate
pip install -r requirements.txt
```

### "No MSG files found"

- Check that the input directory path is correct
- Ensure the directory contains `.msg` files
- Try using the `-r` flag to search subdirectories recursively

### Permission errors

Ensure you have read permissions for the input directory and write permissions for the output directory.

## License

MIT
