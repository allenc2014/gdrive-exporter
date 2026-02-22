"""
Google Drive navigation and document selection module.
Handles browsing Google Drive folders and selecting documents for conversion.
"""

import logging
from typing import List, Tuple, Optional, Dict, Any
from datetime import datetime, timezone

from auth import call_with_retry


def log_info(msg: str):
    """Log info message."""
    logging.info(msg)


def log_warn(msg: str):
    """Log warning message."""
    logging.warning(msg)


def sanitize_name(name: str) -> str:
    """Sanitize filename for local storage."""
    import re
    name = name.strip()
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name)
    return name or "untitled"


def parse_drive_time(ts: str) -> datetime:
    """Parse Drive timestamp to datetime object."""
    # Example: "2026-02-21T18:12:34.123Z"
    return datetime.fromisoformat(ts.replace("Z", "+00:00")).astimezone(timezone.utc)


class GoogleDriveNavigator:
    """Handles Google Drive navigation and document selection."""
    
    def __init__(self, drive_service):
        self.drive_service = drive_service
    
    def list_drive_items(self, folder_id: str = "root") -> Tuple[List[Dict], List[Dict], List[Dict]]:
        """
        List items in a Drive folder.
        
        Returns:
            Tuple of (folders, docs, others)
        """
        query = f"'{folder_id}' in parents and trashed = false"
        results = call_with_retry(
            self.drive_service.files().list,
            q=query,
            fields="files(id, name, mimeType, modifiedTime)"
        ).execute()

        files = results.get("files", [])
        folders = [f for f in files if f["mimeType"] == "application/vnd.google-apps.folder"]
        docs = [f for f in files if f["mimeType"] == "application/vnd.google-apps.document"]
        others = [f for f in files if f not in folders and f not in docs]
        return folders, docs, others

    def navigate_drive(self) -> Optional[Dict]:
        """
        Interactive Drive navigation for single document selection.
        
        Returns:
            Selected document dict or None if quit
        """
        path_stack = ["root"]
        name_stack = ["My Drive"]

        while True:
            folder_id = path_stack[-1]
            folders, docs, _ = self.list_drive_items(folder_id)

            print("\nCurrent folder:", " / ".join(name_stack))

            print("\nFolders:")
            for i, f in enumerate(folders, start=1):
                print(f"  F{i}. {f['name']}")

            print("\nGoogle Docs:")
            for i, d in enumerate(docs, start=1):
                print(f"  D{i}. {d['name']}")

            print("\nCommands:")
            print("  f#  → open folder")
            print("  d#  → select Google Doc")
            print("  b   → go back")
            print("  q   → quit")

            cmd = input("\nEnter command: ").strip().lower()

            if cmd == "q":
                return None

            if cmd == "b":
                if len(path_stack) > 1:
                    path_stack.pop()
                    name_stack.pop()
                continue

            if cmd.startswith("f"):
                idx = cmd[1:]
                if idx.isdigit() and 1 <= int(idx) <= len(folders):
                    folder = folders[int(idx) - 1]
                    path_stack.append(folder["id"])
                    name_stack.append(folder["name"])
                else:
                    print("Invalid folder.")
                continue

            if cmd.startswith("d"):
                idx = cmd[1:]
                if idx.isdigit() and 1 <= int(idx) <= len(docs):
                    return docs[int(idx) - 1]
                else:
                    print("Invalid document.")
                continue

            print("Unknown command.")

    def folder_options_menu(self) -> str:
        """Display folder options menu and return user choice."""
        print("\nFolder Options:")
        print("  1. Open this folder")
        print("  2. Recursively export all Google Docs in this folder")
        print("  3. Recursively download all DOCX files in this folder")
        print("  4. Recursively process BOTH (Docs → MD, DOCX → download)")
        print("  5. Recursively download ALL files in this folder")
        print("  0. Go back")
        choice = input("Choose: ").strip()
        return choice

    def process_folder_recursive(self, docs_service, folder_id: str, drive_path_parts: List[str], 
                                output_root, mode: str, creds, doc_converter):
        """
        Recursively process a folder.
        
        Args:
            mode: 'docs' for Google Docs only, 'docx' for DOCX only, 'both' for both, 'all' for all files
        """
        folders, docs, others = self.list_drive_items(folder_id)

        # Process docs
        if mode in ("docs", "both", "all"):
            for d in docs:
                d["_creds"] = creds  # inject creds for export_media
                doc_converter.export_doc_tabs(d, drive_path_parts + [d["name"]])

        # Process other files
        if mode in ("docx", "both"):
            for f in others:
                mime = f.get("mimeType", "")
                if mime == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
                    doc_converter.download_raw_docx_if_newer(f, drive_path_parts + [f["name"]])
                else:
                    log_warn(f"SKIPPED: /{'/'.join(drive_path_parts + [f['name']])} (unsupported type={mime})")
        
        # Process all files (including all other file types)
        if mode == "all":
            for f in others:
                doc_converter.download_raw_file_if_newer(f, drive_path_parts + [f["name"]])

        # Recurse into subfolders
        for folder in folders:
            sub_path = drive_path_parts + [folder["name"]]
            self.process_folder_recursive(
                docs_service,
                folder["id"],
                sub_path,
                output_root,
                mode,
                creds,
                doc_converter
            )

    def interactive_with_recursive(self, docs_service, drive_service, output_root, creds, doc_converter):
        """
        Interactive Drive navigation with recursive processing options.
        """
        path_stack = ["root"]
        name_stack = ["My Drive"]

        while True:
            folder_id = path_stack[-1]
            folders, docs, others = self.list_drive_items(folder_id)

            print("\nCurrent folder:", " / ".join(name_stack))

            print("\nFolders:")
            for i, f in enumerate(folders, start=1):
                print(f"  F{i}. {f['name']}")

            print("\nGoogle Docs:")
            for i, d in enumerate(docs, start=1):
                print(f"  D{i}. {d['name']}")

            if others:
                print("\nOther files:")
                for i, o in enumerate(others, start=1):
                    print(f"  O{i}. {o['name']} ({o['mimeType']})")

            print("\nCommands:")
            print("  f#  → open folder / folder options")
            print("  d#  → export single Google Doc now")
            print("  b   → go back")
            print("  q   → quit")

            cmd = input("\nEnter command: ").strip().lower()

            if cmd == "q":
                return

            if cmd == "b":
                if len(path_stack) > 1:
                    path_stack.pop()
                    name_stack.pop()
                continue

            if cmd.startswith("f"):
                idx = cmd[1:]
                if idx.isdigit() and 1 <= int(idx) <= len(folders):
                    folder = folders[int(idx) - 1]
                    # Folder options
                    choice = self.folder_options_menu()
                    if choice == "1":
                        path_stack.append(folder["id"])
                        name_stack.append(folder["name"])
                    elif choice == "2":
                        log_info(f"RECURSIVE DOCS from /{'/'.join(name_stack + [folder['name']])}")
                        self.process_folder_recursive(
                            docs_service,
                            folder["id"],
                            name_stack + [folder["name"]],
                            output_root,
                            mode="docs",
                            creds=creds,
                            doc_converter=doc_converter
                        )
                    elif choice == "3":
                        log_info(f"RECURSIVE DOCX from /{'/'.join(name_stack + [folder['name']])}")
                        self.process_folder_recursive(
                            docs_service,
                            folder["id"],
                            name_stack + [folder["name"]],
                            output_root,
                            mode="docx",
                            creds=creds,
                            doc_converter=doc_converter
                        )
                    elif choice == "4":
                        log_info(f"RECURSIVE BOTH from /{'/'.join(name_stack + [folder['name']])}")
                        self.process_folder_recursive(
                            docs_service,
                            folder["id"],
                            name_stack + [folder["name"]],
                            output_root,
                            mode="both",
                            creds=creds,
                            doc_converter=doc_converter
                        )
                    elif choice == "5":
                        log_info(f"RECURSIVE ALL FILES from /{'/'.join(name_stack + [folder['name']])}")
                        self.process_folder_recursive(
                            docs_service,
                            folder["id"],
                            name_stack + [folder["name"]],
                            output_root,
                            mode="all",
                            creds=creds,
                            doc_converter=doc_converter
                        )
                    elif choice == "0":
                        pass
                    else:
                        print("Invalid choice.")
                else:
                    print("Invalid folder.")
                continue

            if cmd.startswith("d"):
                idx = cmd[1:]
                if idx.isdigit() and 1 <= int(idx) <= len(docs):
                    d = docs[int(idx) - 1]
                    d["_creds"] = creds
                    log_info(f"SINGLE DOC export: /{'/'.join(name_stack + [d['name']])}")
                    doc_converter.export_doc_tabs(d, name_stack + [d["name"]])
                else:
                    print("Invalid document.")
                continue

            print("Unknown command.")
