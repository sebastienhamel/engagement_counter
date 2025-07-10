# Engagement Counter

`engagement_counter.py` is a Python script designed count user engagement metrics from a Microsoft Teams meeting recording transcript in a .docx format. It provides a simple interface for processing intervention count, generating summaries, and exporting results in a .csv format.

## Features

- Reads engagement data from a .docx format downloaded from Clipchamp (i.e. meeting recording in Microsoft Teams)
- Calculates intervention count per speaker in a meeting. 
- Outputs summary report (date, speaker, intervention count)
- Autodetects the date from the .docx file
- Autodetects the .docx files

## Requirements

- Python 3.7+
- pandas
- python-docx

Install dependencies with:

```bash
pip install -r requirements.txt
```

## Usage

This script is made to be run directly in VS Code (or your favorite IDE)
1. Install the dependencies
2. Download your files from Clipchamp. Select the option _Download as .docx_
3. Create a folder called "transcripts" at the root of the project and copy the transcript files into it.
4. If you need to exclude a participant (like the instructor or a moderator), create a file called instructor_name.txt and insert the name of the person you need to exclude into it. 
5. Run the script

### Arguments

None

## Example

```bash
python engagement_counter.py
```

## License

MIT License

---
