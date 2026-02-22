"""
Document conversion module.
Handles conversion of Google Docs to Markdown and DOCX downloads.
"""

import logging
import pathlib
from typing import List, Optional
from datetime import datetime, timezone

from markdownify import markdownify as md
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

from auth import call_with_retry


def log_info(msg: str):
    """Log info message."""
    logging.info(msg)


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


def get_local_mtime(path: pathlib.Path) -> Optional[datetime]:
    """Get local file modification time."""
    if not path.exists():
        return None
    return datetime.fromtimestamp(path.stat().st_mtime, tz=timezone.utc)


class DocumentConverter:
    """Handles conversion of Google Docs to Markdown and DOCX downloads."""
    
    def __init__(self, docs_service, drive_service, output_root: pathlib.Path, force: bool = False):
        self.docs_service = docs_service
        self.drive_service = drive_service
        self.output_root = output_root
        self.force = force
        self.image_counter = {}  # Track image numbers per document
        self.chapter_counter = 1  # Track chapter numbers across entire document
        self.current_document = None
        self.current_document_id = None

    def get_document(self, document_id: str) -> dict:
        if document_id != self.current_document_id:
            self.current_document = self.docs_service.documents().get(
                documentId=document_id,
                includeTabsContent=True).execute()
            self.current_document_id = document_id
        return self.current_document
    
    def download_image(self, inline_object_element: dict, doc_folder: pathlib.Path, doc_name: str, document_id: str, tab_name: str = None, section_title: str = None) -> str:
        """Download an embedded image and return the markdown reference."""
        try:
            # Get the inline object ID from the correct location
            inline_object_data = inline_object_element.get("inlineObjectElement", {})
            inline_object_id = inline_object_data.get("inlineObjectId")
            if not inline_object_id:
                return "[IMAGE]"
            
            logging.info(f"Attempting to download image: {inline_object_id}")
            
            # Create images folder in the correct location
            # For chapters within tabs, images should go in the chapter folder, not the parent tab folder
            if tab_name and section_title:  # If this is a chapter within a tab
                images_folder = doc_folder / sanitize_name(tab_name) / sanitize_name(section_title) / "images"
                images_relative_folder_path = f"{sanitize_name(section_title)}/images"
            elif section_title:  # If this is a chapter at document level
                images_folder = doc_folder / sanitize_name(section_title) / "images"
                images_relative_folder_path = f"{sanitize_name(section_title)}/images"
            else:
                images_folder = doc_folder / "images"
                images_relative_folder_path = "images"
            
            images_folder.mkdir(parents=True, exist_ok=True)
            
            # Create a unique key for this document/tab combination
            key = f"{doc_name}_{tab_name}" if tab_name else doc_name
            if key not in self.image_counter:
                self.image_counter[key] = 0
            
            self.image_counter[key] += 1
            image_number = self.image_counter[key]
            
            # Generate filename with chapter name if available, otherwise document name
            if section_title:  # If this is a chapter within a tab
                sanitized_section_title = sanitize_name(section_title)
                image_filename = f"{sanitized_section_title}_{image_number}.png"
            elif tab_name:  # If this is at tab level
                sanitized_tab_name = sanitize_name(tab_name)
                image_filename = f"{sanitized_tab_name}_{image_number}.png"
            else:
                # Document level
                sanitized_doc_name = sanitize_name(doc_name)
                image_filename = f"{sanitized_doc_name}_{image_number}.png"
            
            image_path = images_folder / image_filename
            
            # if a document has already been retrieved before, there is no need to re-download it
            if image_path.exists():
                logging.info(f"Image already exists: {image_path}")
                return f"![{image_filename}](<{images_relative_folder_path}/{image_filename}>)"
            
            # Get inline objects from the correct nested path in tab structure
            document = self.get_document(document_id)
            
            inline_objects = {}
            tabs = document.get("tabs", [])
            for tab in tabs:
                document_tab = tab.get("documentTab", {})
                tab_inline_objects = document_tab.get("inlineObjects", {})
                inline_objects.update(tab_inline_objects)
                if inline_object_id in tab_inline_objects:
                    break
            
            if inline_object_id not in inline_objects:
                logging.warning(f"Inline object {inline_object_id} not found in document tabs")
                return "[IMAGE]"
            
            inline_obj = inline_objects[inline_object_id]
            
            # Access embedded object through inlineObjectProperties
            inline_props = inline_obj.get("inlineObjectProperties", {})
            embedded_object = inline_props.get("embeddedObject")
            
            if not embedded_object:
                logging.warning(f"No embedded object found for {inline_object_id}")
                return "[IMAGE]"
            
            # Download the image content using the content URI
            image_properties = embedded_object.get("imageProperties", {})
            content_uri = image_properties.get("contentUri")
            
            if content_uri:
                # Use the documents service HTTP client to download the image
                response = self.docs_service._http.request(content_uri)
                if response[0].status == 200:
                    with open(image_path, "wb") as f:
                        f.write(response[1])
                    
                    logging.info(f"Successfully downloaded image: {image_filename}")
                    # Return relative markdown reference
                    return f"![{image_filename}](<{images_relative_folder_path}/{image_filename}>)"
                else:
                    logging.warning(f"Failed to download image {inline_object_id}: HTTP {response[0].status}")
            else:
                logging.warning(f"No content URI found for image {inline_object_id}")
            
            return "[IMAGE]"
            
        except Exception as e:
            logging.warning(f"Failed to download image {inline_object_id if 'inline_object_id' in locals() else 'unknown'}: {e}")
            return "[IMAGE]"

    def create_heading1_files(self, sections: list, doc_folder: pathlib.Path, doc_name: str, tab_name: str = None) -> str:
        """Create separate files for each H1 section and return index content."""
        index_content = []
        section_files = []
        
        # Create a subfolder for sections if we have a tab name
        if tab_name:
            sections_folder = doc_folder / sanitize_name(tab_name)
        else:
            sections_folder = doc_folder / sanitize_name(doc_name)
        
        sections_folder.mkdir(exist_ok=True)
        
        for i, section in enumerate(sections):
            if section["title"] and section["title"].strip():  # Check for non-empty title
                title = section["title"].strip()
                
                # Handle incomplete chapter titles (e.g., just "Chapter")
                if title.lower() == "chapter" or (title.lower().startswith("chapter") and len(title) <= 10):
                    title = f"Chapter {self.chapter_counter}"
                    self.chapter_counter += 1
                    logging.info(f"Auto-numbered incomplete chapter as: {title}")
                
                # Create section file
                section_filename = sanitize_name(title) + ".md"
                section_filepath = sections_folder / section_filename
                section_title = title
                
                # Write section content (excluding the H1 heading since it's in the filename)
                section_content = '\n\n'.join(section["content"][1:])  # Skip first element (the H1)
                with open(section_filepath, "w", encoding="utf-8") as f:
                    f.write(f"# {title}\n\n{section_content}")
                
                # Add to index
                if tab_name:  # Chapter within a tab
                    relative_path = f"{sanitize_name(tab_name)}/{sanitize_name(title)}/{section_filename}"
                else:  # Chapter at document level
                    relative_path = f"{sanitize_name(title)}/{section_filename}"
                index_content.append(f"- [{title}]({relative_path})")
                section_files.append(section_filename)
                
                logging.info(f"Created section file: {section_filepath}")
            else:
                # This is the content before the first H1, add to index as introduction
                intro_content = '\n\n'.join(section["content"])
                if intro_content.strip():  # Only add if there's actual content
                    index_content.insert(0, intro_content)
                    logging.info(f"Added introduction content with {len(section['content'])} items")
        
        # Return the index content
        return '\n\n'.join(index_content)

    def extract_text_from_body(self, body: dict, doc_folder: pathlib.Path, doc_name: str, document_id: str, tab_name: str = None) -> str:
        """Extract text with formatting from Google Docs body structure."""
        text_parts = []
        heading1_sections = []  # Store H1 sections for splitting (only for ###chapter###)
        
        def process_element(element):
            """Recursively process document elements to extract formatted text."""
            if "paragraph" in element:
                para = element["paragraph"]
                para_text = []
                
                for el in para.get("elements", []):
                    if "textRun" in el:
                        text_run = el["textRun"]
                        content = text_run.get("content", "")
                        
                        # Filter out control characters except newlines and tabs
                        content = ''.join(char for char in content if ord(char) >= 32 or char in '\n\t')
                        
                        # Clean up whitespace but preserve essential formatting
                        content = content.replace('\r\n', '\n').replace('\r', '\n')
                        
                        # Apply text formatting
                        text_style = text_run.get("textStyle", {})
                        
                        # Debug: Log text style info for first few paragraphs
                        if len(text_parts) < 5:  # Only log first few to avoid spam
                            font_family = text_style.get("weightedFontFamily", {}).get("fontFamily", "")
                            font_size = text_style.get("fontSize", {}).get("magnitude", 0)
                            is_bold = text_style.get("bold", False)
                            # logging.info(f"Text style debug - Font: {font_family}, Size: {font_size}, Bold: {is_bold}, Content: {content[:50]}...")
                        
                        # Check for Chapter separator (most reliable method)
                        if "###chapter###" in content.lower():
                            # Found a chapter separator, extract the title from the same line
                            import re
                            # Match pattern like ###chapter### Chapter 2: The Tower of Broken Time
                            match = re.search(r'###chapter###\s*(.+)', content, re.IGNORECASE)
                            if match:
                                title = match.group(1).strip()
                                # If title is just "Chapter" or incomplete, try to get more context
                                if title.lower() == "chapter" or len(title) < 5:
                                    logging.warning(f"Incomplete chapter title detected: '{title}' - content: {content}")
                                    # Use a generic numbered chapter based on section count
                                    # We'll handle this in the section creation logic
                                logging.info(f"Detected chapter separator with title: '{title}'")
                                return {
                                    "type": "heading1",
                                    "content": title,
                                    "raw_content": content
                                }
                            else:
                                # If no title found on same line, log the content for debugging
                                logging.warning(f"Chapter separator found but no title extracted from: '{content}'")
                        
                        # Check for subtitle style (convert to H2 ##)
                        font_family = text_style.get("weightedFontFamily", {}).get("fontFamily", "")
                        font_size = text_style.get("fontSize", {}).get("magnitude", 0)
                        if font_family == "Subtitle" or (font_size >= 14 and font_size < 18):
                            content = f"## {content.strip()}"
                            logging.info(f"Detected subtitle style: {content[:50]}...")
                        
                        # Check for Heading 2 style (convert to H3 ###)
                        elif font_family == "Heading 2" or (font_size >= 12 and font_size < 14):
                            content = f"### {content.strip()}"
                            logging.info(f"Detected heading 2 style: {content[:50]}...")
                        
                        # Check for Heading 1 style and convert to markdown # (no longer used for splitting)
                        is_heading1 = False
                        
                        # Method 1: Check font family
                        font_family = text_style.get("weightedFontFamily", {}).get("fontFamily", "")
                        if font_family == "Heading 1":
                            is_heading1 = True
                        
                        # Method 2: Check for large font size (common for H1)
                        font_size = text_style.get("fontSize", {}).get("magnitude", 0)
                        if font_size >= 18:  # Reduced threshold for better detection
                            is_heading1 = True
                        
                        # Method 3: Check for bold + large size combination
                        if text_style.get("bold") and font_size >= 16:
                            is_heading1 = True
                        
                        # Method 4: Check if content starts with # (markdown style)
                        if content.strip().startswith('#'):
                            is_heading1 = True
                        
                        if is_heading1:
                            # Convert H1 style to markdown heading (no longer used for splitting)
                            heading_text = content.strip('# ').strip()
                            if heading_text:  # Only convert if title is not empty
                                logging.info(f"Converted H1 to markdown: {heading_text} (font: {font_family}, size: {font_size})")
                                content = f"# {heading_text}"
                            else:
                                logging.info(f"Skipping empty H1 (font: {font_family}, size: {font_size})")
                                # Treat as regular content if title is empty
                                pass
                        
                        # Bold
                        if text_style.get("bold"):
                            content = f"**{content}**"
                        
                        # Italic
                        if text_style.get("italic"):
                            content = f"*{content}*"
                        
                        # Underline (convert to emphasis in markdown)
                        if text_style.get("underline"):
                            content = f"__{content}__"
                        
                        # Strikethrough
                        if text_style.get("strikethrough"):
                            content = f"~~{content}~~"
                        
                        para_text.append(content)
                    
                    elif "horizontalRule" in el:
                        para_text.append("\n---\n")
                    
                    elif "inlineObjectElement" in el:
                        # Download and reference the image
                        # Extract section title for folder structure - use the current chapter being tracked
                        image_ref = self.download_image(el, doc_folder, doc_name, document_id, tab_name, section_title=current_chapter)
                        para_text.append(image_ref)
                
                if para_text:
                    # Join the paragraph text and clean up formatting issues
                    paragraph_content = ''.join(para_text)
                    
                    # Remove newlines within formatted text that cause splits
                    # Look for patterns like **text\n** and fix them
                    import re
                    paragraph_content = re.sub(r'(\*\*|__|\*\*|~~|_|\*)(.*?)\n(\1)', r'\1\2\3', paragraph_content)
                    
                    # Clean up any remaining trailing whitespace
                    paragraph_content = paragraph_content.rstrip()
                    
                    if paragraph_content:
                        return paragraph_content
            
            elif "table" in element:
                # Process table
                table = element["table"]
                table_rows = []
                
                for row in table.get("tableRows", []):
                    row_cells = []
                    for cell in row.get("tableCells", []):
                        cell_text = []
                        for cell_content in cell.get("content", []):
                            cell_text.extend(process_element(cell_content))
                        row_cells.append('|'.join(cell_text))
                    if row_cells:
                        table_rows.append('|' + '|'.join(row_cells) + '|')
                
                if table_rows:
                    # Add table header separator
                    if len(table_rows) > 0:
                        col_count = len(table_rows[0].split('|')) - 1
                        separator = '|' + '|'.join(['---'] * col_count) + '|'
                        table_rows.insert(1, separator)
                    return '\n'.join(table_rows)
            
            elif "list" in element:
                # Process list
                list_element = element["list"]
                list_items = []
                
                for list_item in list_element.get("listItems", []):
                    item_text = []
                    for content in list_item.get("content", []):
                        item_text.extend(process_element(content))
                    
                    if item_text:
                        # Determine list type and level
                        bullet = "- "  # Default bullet
                        # You could enhance this to check list properties for ordered/unordered
                        list_items.append(f"{bullet}{''.join(item_text)}")
                
                text_parts.extend(list_items)
            
            return text_parts
        
        # Process all elements and collect H1 sections
        current_section = {"title": None, "content": []}
        current_chapter = None  # Track which chapter we're currently processing
        
        for element in body.get("content", []):
            result = process_element(element)
            
            if result:
                if isinstance(result, dict) and result.get("type") == "heading1":
                    # Found a Heading 1, save previous section and start new one
                    if current_section["title"] or current_section["content"]:
                        heading1_sections.append(current_section)
                        logging.info(f"Added section: {current_section['title'] or 'Introduction'} with {len(current_section['content'])} content items")
                    current_section = {
                        "title": result["content"],
                        "content": [result["raw_content"]]  # Include H1 in the content
                    }
                    current_chapter = result["content"]  # Track current chapter for image context
                    logging.info(f"Started new section: {result['content']}")
                else:
                    # Regular content, add to current section
                    current_section["content"].append(result)
        
        # Don't forget the last section
        if current_section["title"] or current_section["content"]:
            heading1_sections.append(current_section)
            logging.info(f"Added final section: {current_section['title'] or 'Introduction'} with {len(current_section['content'])} content items")
        
        logging.info(f"Total sections found: {len(heading1_sections)}")
        
        # If we have H1 sections, create separate files and index
        if len(heading1_sections) > 1:
            return self.create_heading1_files(heading1_sections, doc_folder, doc_name, tab_name)
        else:
            # No H1 sections or only one, return normal content
            # Join all content from the single section
            if heading1_sections:
                result = '\n\n'.join(heading1_sections[0]["content"])
            else:
                result = '\n\n'.join(text_parts)
            return result

    def export_doc_tabs(self, drive_file: dict, drive_path_parts: List[str]):
        """
        Export a Google Doc with tabs to Markdown files.
        
        Args:
            drive_file: Drive file metadata
            drive_path_parts: Path parts for organizing output
        """
        doc_id = drive_file["id"]
        doc_name = drive_file["name"]
        doc_modified = parse_drive_time(drive_file["modifiedTime"])

        # Local parent folder for this doc - drive_path_parts already includes the document name
        doc_folder = self.output_root.joinpath(*[sanitize_name(p) for p in drive_path_parts])
        doc_folder.mkdir(parents=True, exist_ok=True)

        # Fetch full doc with tabs content
        doc = call_with_retry(
            self.docs_service.documents().get,
            documentId=doc_id,
            includeTabsContent=True
        ).execute()

        self.current_document = doc
        self.current_document_id = doc_id

        tabs = doc.get("tabs", [])

        if tabs:
            # For docs with tabs, we compare doc_modified vs any existing .md in doc_folder
            if not self.force:
                existing_mtimes = [
                    get_local_mtime(p)
                    for p in doc_folder.glob("*.md")
                ]
                latest_local = max([t for t in existing_mtimes if t is not None], default=None)

                if latest_local and latest_local >= doc_modified:
                    msg = f"UP-TO-DATE DOC (tabs): /{'/'.join(drive_path_parts)}"
                    log_info(msg)
                    return
            

            # Export each tab as Markdown
            log_info(f"EXPORTED DOC (tabs): /{'/'.join(drive_path_parts)}")
            for tab in tabs:
                props = tab.get("tabProperties", {})
                title = props.get("title", "Untitled Tab")
                doc_tab = tab.get("documentTab", {})
                body = doc_tab.get("body", {})
                tab_name = props.get("title", "Untitled Tab")

                text = self.extract_text_from_body(body, doc_folder, doc_name, doc_id, tab_name)
                # No need for markdownify since we're already handling formatting
                md_text = text

                filename = sanitize_name(title) + ".md"
                filepath = doc_folder / filename

                with open(filepath, "w", encoding="utf-8") as f:
                    f.write(f"# {title}\n\n{md_text}")

            # Also download the full DOCX version for tabbed documents
            docx_path = doc_folder / (sanitize_name(doc_name) + ".docx")
            self.download_docx_if_newer(drive_file, docx_path, drive_path_parts)

        else:
            # No tabs: export whole doc as MD + download DOCX
            md_path = doc_folder / (sanitize_name(doc_name) + ".md")
            docx_path = doc_folder / (sanitize_name(doc_name) + ".docx")

            if not self.force:
                latest_local = max(
                    [t for t in [get_local_mtime(md_path), get_local_mtime(docx_path)] if t is not None],
                    default=None
                )

                if latest_local and latest_local >= doc_modified:
                    msg = f"UP-TO-DATE DOC (no tabs): /{'/'.join(drive_path_parts)}"
                    log_info(msg)
                else:
                    # Export full body as MD
                    body = doc.get("body", {})
                    text = self.extract_text_from_body(body, doc_folder, doc_name, doc_id)
                    # No need for markdownify since we're already handling formatting
                    md_text = text

                    with open(md_path, "w", encoding="utf-8") as f:
                        f.write(f"# {doc_name}\n\n{md_text}")

                    log_info(f"EXPORTED DOC (no tabs → MD): /{'/'.join(drive_path_parts)} (size={md_path.stat().st_size} bytes)")
            else:
                # Force export full body as MD
                body = doc.get("body", {})
                text = self.extract_text_from_body(body, doc_folder, doc_name, doc_id)
                # No need for markdownify since we're already handling formatting
                md_text = text

                with open(md_path, "w", encoding="utf-8") as f:
                    f.write(f"# {doc_name}\n\n{md_text}")

                log_info(f"EXPORTED DOC (no tabs → MD): /{'/'.join(drive_path_parts)} (size={md_path.stat().st_size} bytes)")

            # Always ensure DOCX is downloaded if remote is newer
            self.download_docx_if_newer(drive_file, docx_path, drive_path_parts)

    def download_docx_if_newer(self, drive_file: dict, local_path: pathlib.Path, drive_path_parts: List[str], creds=None):
        """Download DOCX version of a Google Doc if newer than local version."""
        drive_modified = parse_drive_time(drive_file["modifiedTime"])
        local_mtime = get_local_mtime(local_path)

        if not self.force and local_mtime and local_mtime >= drive_modified:
            log_info(f"UP-TO-DATE DOCX: /{'/'.join(drive_path_parts)}.docx")
            return

        # Download DOCX using the existing drive service
        request = self.drive_service.files().export_media(
            fileId=drive_file["id"],
            mimeType="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = call_with_retry(downloader.next_chunk)

        local_path.parent.mkdir(parents=True, exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(fh.getvalue())

        size = local_path.stat().st_size
        log_info(f"DOWNLOADED DOCX: /{'/'.join(drive_path_parts)}.docx (size={size} bytes)")

    def download_raw_docx_if_newer(self, drive_file: dict, drive_path_parts: List[str]):
        """
        Download native DOCX files from Drive if newer than local version.
        
        Args:
            drive_file: Drive file metadata (native DOCX file)
            drive_path_parts: Path parts for organizing output
        """
        # For native DOCX files in Drive (mimeType = application/vnd.openxmlformats-officedocument.wordprocessingml.document)
        drive_modified = parse_drive_time(drive_file["modifiedTime"])
        filename = sanitize_name(drive_file["name"])
        if not filename.lower().endswith(".docx"):
            filename += ".docx"

        local_path = self.output_root.joinpath(*[sanitize_name(p) for p in drive_path_parts], filename)
        local_mtime = get_local_mtime(local_path)

        if not self.force and local_mtime and local_mtime >= drive_modified:
            log_info(f"UP-TO-DATE DOCX: /{'/'.join(drive_path_parts)}")
            return

        request = self.drive_service.files().get_media(fileId=drive_file["id"])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = call_with_retry(downloader.next_chunk)

        local_path.parent.mkdir(parents=True, exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(fh.getvalue())

        size = local_path.stat().st_size
        log_info(f"DOWNLOADED DOCX: /{'/'.join(drive_path_parts)} (size={size} bytes)")

    def download_raw_file_if_newer(self, drive_file: dict, drive_path_parts: List[str]):
        """
        Download any file from Drive if newer than local version.
        
        Args:
            drive_file: Drive file metadata (any file type)
            drive_path_parts: Path parts for organizing output (excluding filename)
        """
        drive_modified = parse_drive_time(drive_file["modifiedTime"])
        filename = sanitize_name(drive_file["name"])
        
        # For individual files, use only the directory path parts (exclude the filename)
        if len(drive_path_parts) > 0 and drive_path_parts[-1] == filename:
            # This is an individual file, use parent directory path
            local_path = self.output_root.joinpath(*[sanitize_name(p) for p in drive_path_parts[:-1]], filename)
            log_path = "/".join(drive_path_parts[:-1] + [filename])
        else:
            # This is part of a directory processing, use full path
            local_path = self.output_root.joinpath(*[sanitize_name(p) for p in drive_path_parts], filename)
            log_path = "/".join(drive_path_parts + [filename])
        
        local_mtime = get_local_mtime(local_path)

        if not self.force and local_mtime and local_mtime >= drive_modified:
            log_info(f"UP-TO-DATE FILE: /{log_path}")
            return

        request = self.drive_service.files().get_media(fileId=drive_file["id"])
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)

        done = False
        while not done:
            status, done = call_with_retry(downloader.next_chunk)

        local_path.parent.mkdir(parents=True, exist_ok=True)
        with open(local_path, "wb") as f:
            f.write(fh.getvalue())

        size = local_path.stat().st_size
        log_info(f"DOWNLOADED FILE: /{log_path} (size={size} bytes)")
