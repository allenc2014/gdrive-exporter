# Google Drive Exporter

A powerful Python tool for exporting and synchronizing documents from Google Drive to your local filesystem. Supports Google Docs, native DOCX files, and other file types with intelligent conversion and organization.

## Why This Tool?

Google already provides native support for downloading individual Google Docs as Markdown. However, this tool was created to address specific limitations:

- **Recursive Processing**: Automatically process entire folder structures with hundreds of documents, rather than manual one-by-one downloads
- **Chapter Splitting**: Split large Google Docs into multiple Markdown files using `###chapter###` tags while preserving embedded images
- **Batch Operations**: Efficiently handle bulk exports with intelligent syncing and force refresh options
- **Organized Structure**: Maintain your Google Drive folder hierarchy in the local output

**Note**: This tool does not handle all Google Docs formatting features. Your mileage may vary depending on document complexity. For simple documents, Google's native export may be sufficient.

## Features

- **Google Drive Integration**: Navigate and browse your Google Drive folders interactively
- **Multiple File Types**: Handle Google Docs, native DOCX files, and other file formats
- **Smart Conversion**: Convert Google Docs to Markdown while preserving formatting, images, and structure
- **Chapter Splitting**: Automatically split documents by `###chapter###` tags into separate Markdown files
- **Tab Support**: Handle Google Docs with multiple tabs, exporting each tab as a separate Markdown file
- **Image Downloads**: Automatically download and reference embedded images
- **Recursive Processing**: Process entire folder structures with customizable options
- **Intelligent Syncing**: Only download newer versions (unless forced)
- **Dual Output**: Create both Markdown files and download original DOCX versions

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd gdrive-exporter
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up Google Drive API credentials:
   - Create a project in Google Cloud Console
   - Enable Google Drive API and Google Docs API
   - Create OAuth 2.0 credentials
   - Download the credentials JSON file and place it as `credentials.json` in the project directory

## Usage

### Command Line Arguments

```bash
python gdrive-exporter/main.py [OPTIONS]
```

#### Options:

- `-o, --output FOLDER`: Specify output folder (default: `./output`)
- `--path PATH`: Process a specific Google Drive path (e.g., `'/My Drive/Stories/Story1/Drafts'`)
- `--force`: Force reprocess all files, ignoring timestamp comparisons
- `--download-all`: Download all files in directory (not just Google Docs and DOCX)

### Usage Examples

#### Interactive Mode (Default)
```bash
python gdrive-exporter/main.py
```
Launches an interactive browser to navigate your Google Drive and select files/folders for export.

#### Direct Path Processing
```bash
python gdrive-exporter/main.py --path "/My Drive/Stories/Story1/Drafts"
```
Process a specific folder or file at the given Google Drive path.

#### Custom Output Directory
```bash
python gdrive-exporter/main.py -o "my_exports" --path "/My Drive/Documents"
```
Save exports to a custom directory.

#### Force Refresh
```bash
python gdrive-exporter/main.py --force --path "/My Drive/Stories"
```
Re-download and reprocess all files, even if they haven't changed.

#### Download All Files
```bash
python gdrive-exporter/main.py --download-all --path "/My Drive/Archive"
```
Download all file types in the specified folder, not just documents.

### Interactive Navigation

When running in interactive mode, you can:

- **Navigate folders**: Use `f#` commands (e.g., `f1` to open folder 1)
- **Export documents**: Use `d#` commands (e.g., `d2` to export document 2)
- **Folder options**: When selecting a folder, choose from:
  1. Open this folder
  2. Recursively export all Google Docs
  3. Recursively download all DOCX files
  4. Recursively process both (Docs → MD, DOCX → download)
  5. Recursively download ALL files
- **Navigate back**: Use `b` to go to the parent folder
- **Quit**: Use `q` to exit

## Output Structure

The tool preserves your Google Drive folder structure in the output directory:

```
output/
├── My Drive/
│   ├── Stories/
│   │   ├── Story1/
│   │   │   ├── Drafts/
│   │   │   │   ├── Chapter 1/
│   │   │   │   │   ├── Chapter 1.md
│   │   │   │   │   └── images/
│   │   │   │   │       └── Chapter 1_1.png
│   │   │   │   ├── Chapter 2.md
│   │   │   │   └── Novel.docx
│   │   │   └── Research.md
│   │   └── Other Story.docx
│   └── Notes.md
```

## Document Processing

### Google Docs
- Converted to Markdown with preserved formatting
- Documents containing `###chapter###` tags are automatically split into separate Markdown files at each tag
- Images downloaded and referenced
- Original DOCX version also downloaded

### Native DOCX Files
- Downloaded as-is without conversion
- Organized in the same folder structure

### Other Files
- Downloaded when using `--download-all` option
- Preserved in original format

## Logging

All operations are logged to `output/output.log` with timestamps and detailed information about:
- Files processed and skipped
- Download sizes and timestamps
- Error messages and warnings
- Navigation actions

## Requirements

- Python 3.7+
- Google APIs: `google-api-python-client`, `google-auth-httplib2`, `google-auth-oauthlib`
- Markdown processing: `markdownify`

See `requirements.txt` for the complete list of dependencies.

## Disclaimer

**Important Notes Before Use:**

- **Content Export**: This tool makes no claims that it will capture all content during export. Some complex formatting, embedded objects, or special Google Docs features may not be perfectly preserved.

- **Markdown Conversion**: The automated conversion to Markdown is not guaranteed to be perfect. Complex layouts, tables, and advanced formatting may require manual cleanup.

- **Test First**: Always test the tool on a small sample of documents before running it on large folders or critical content.

- **OAuth Token Expiration**: The current implementation does not include automatic OAuth token refresh logic. Tokens may expire during long export operations, requiring re-authentication.

- **Backup Your Data**: Always maintain backups of your original Google Drive content before using this export tool.

## Security

- Uses OAuth 2.0 for secure Google Drive access
- Credentials stored locally in token files
- No data sent to third-party services

## License

This project is licensed under the Apache License 2.0. See the [LICENSE](LICENSE) file for details.
