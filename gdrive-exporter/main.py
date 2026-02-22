#!/usr/bin/env python3
"""
Main entry point for the word2md application.
Handles command line parsing and orchestrates the conversion process.
"""

import argparse
import pathlib
import logging
from typing import Optional

from google_drive import GoogleDriveNavigator
from doc_converter import DocumentConverter
from auth import get_google_services


def setup_logging(output_root: pathlib.Path):
    """Setup logging configuration."""
    output_root.mkdir(parents=True, exist_ok=True)
    log_path = output_root / "output.log"
    logging.basicConfig(
        filename=str(log_path),
        level=logging.INFO,
        format="[%(asctime)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    # Also log to console (info+)
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(message)s"))
    logging.getLogger().addHandler(console)
    logging.info("=== Run started ===")


def process_drive_path(docs_service, drive_service, drive_navigator, doc_converter, target_path: str, output_root, creds, download_all: bool = False):
    """
    Process a specific Google Drive path (file or folder).
    
    Args:
        target_path: Full Google Drive path (e.g., '/My Drive/Stories/Dead Clocks/Drafts')
        download_all: If True, download all files in directory (not just Google Docs and DOCX)
    """
    from google_drive import sanitize_name
    
    # Remove leading "/" and split the path
    path_parts = target_path.strip("/").split("/")
    
    if not path_parts:
        logging.error("Invalid path provided")
        return
    
    logging.info(f"Looking for path: {target_path}")
    logging.info(f"Path parts: {path_parts}")
    
    current_folder_id = "root"
    
    # Navigate to the parent of the target
    for i, part in enumerate(path_parts[:-1]):
        logging.info(f"Searching for folder: '{part}' (level {i+1}/{len(path_parts)-1})")
        
        folders, docs, others = drive_navigator.list_drive_items(current_folder_id)
        
        # Find the matching folder
        found = False
        for folder in folders:
            if folder["name"] == part:
                current_folder_id = folder["id"]
                logging.info(f"Found folder: {folder['name']} (ID: {folder['id']})")
                found = True
                break
        
        if not found:
            logging.error(f"Folder not found: '{part}' at level {i+1}")
            logging.info("Available folders at this level:")
            for folder in folders:
                logging.info(f"  - {folder['name']}")
            return
    
    # Now look for the target (file or folder)
    target_name = path_parts[-1]
    logging.info(f"Looking for target: '{target_name}'")
    
    folders, docs, others = drive_navigator.list_drive_items(current_folder_id)
    
    # Check if it's a folder
    for folder in folders:
        if folder["name"] == target_name:
            logging.info(f"Found target folder: {folder['name']} (ID: {folder['id']})")
            # Process folder recursively
            drive_path_parts = path_parts
            mode = "all" if download_all else "both"
            drive_navigator.process_folder_recursive(
                docs_service=docs_service,
                folder_id=folder["id"],
                drive_path_parts=drive_path_parts,
                output_root=output_root,
                mode=mode,  # Process both docs and docx, or all files if download_all is True
                creds=creds,
                doc_converter=doc_converter
            )
            return
    
    # Check if it's a Google Doc
    for doc in docs:
        if doc["name"] == target_name:
            logging.info(f"Found target Google Doc: {doc['name']} (ID: {doc['id']})")
            # Process single document
            doc["_creds"] = creds
            doc_converter.export_doc_tabs(doc, path_parts)
            return
    
    # Check if it's a native DOCX file or other file
    for other in others:
        if other["name"] == target_name:
            if other.get("mimeType") == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                logging.info(f"Found target DOCX file: {other['name']} (ID: {other['id']})")
                # Download DOCX file
                doc_converter.download_raw_docx_if_newer(other, path_parts)
                return
            else:
                logging.info(f"Found target file: {other['name']} (ID: {other['id']}, type: {other.get('mimeType', 'unknown')})")
                # Download any file
                doc_converter.download_raw_file_if_newer(other, path_parts)
                return
    
    logging.error(f"Target not found: '{target_name}'")
    logging.info("Available items at this level:")
    for folder in folders:
        logging.info(f"  FOLDER: {folder['name']}")
    for doc in docs:
        logging.info(f"  DOC:    {doc['name']}")
    for other in others:
        logging.info(f"  FILE:   {other['name']} ({other.get('mimeType', 'unknown')})")


def main():
    """Main entry point for the application."""
    parser = argparse.ArgumentParser(
        description="Export Google Docs tabs and DOCX from Drive, preserving structure."
    )
    parser.add_argument("-o", "--output", help="Output folder (default: ./output)")
    parser.add_argument(
        "--path",
        help="Full Google Drive path to file or folder (e.g., '/My Drive/Stories/Dead Clocks/Drafts')"
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Force reprocess all files, ignoring timestamp comparisons"
    )
    parser.add_argument(
        "--download-all",
        action="store_true",
        help="Download all files in directory (not just Google Docs and DOCX)"
    )
    
    args = parser.parse_args()

    output_root = pathlib.Path(args.output or "output")
    setup_logging(output_root)

    try:
        # Initialize services
        docs_service, drive_service = get_google_services()
        creds = docs_service._http.credentials  # type: ignore
        
        # Initialize modules
        drive_navigator = GoogleDriveNavigator(drive_service)
        doc_converter = DocumentConverter(docs_service, drive_service, output_root, force=args.force)
        
        if args.path:
            # Direct path mode - process specific file/folder
            logging.info(f"Processing path: {args.path}")
            process_drive_path(docs_service, drive_service, drive_navigator, doc_converter, args.path, output_root, creds, download_all=args.download_all)
        else:
            # Interactive mode
            logging.info("Navigate your Google Drive. You can export single docs or recurse on folders.")
            drive_navigator.interactive_with_recursive(docs_service, drive_service, output_root, creds, doc_converter)

    except Exception as e:
        logging.error(f"Error: {e}")
        raise
    finally:
        logging.info("=== Run finished ===")


if __name__ == "__main__":
    main()
