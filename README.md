# Photo Metadata Auditor (EXIF Scanner)

A digital forensics utility designed for photography contests and image verification. This tool batch-scans image files to extract embedded metadata (EXIF), helping organizers verify image authenticity, original capture dates, and equipment usage.

## Features

* **Deep Metadata Extraction:** Scans both standard EXIF tags (Make, Model, DateTime) and internal image info dictionary (Software, Comments).
* **Cross-Format Support:** Works with JPG, PNG, TIFF, and WebP files.
* **Automated Reporting:** Generates a detailed Excel spreadsheet with a summary of findings for each file.
* **Integrity Audit:** Quickly identifies images that have been stripped of metadata (potentially edited or screenshot-based submissions).
* **Smart Formatting:** Automatically sanitizes filenames for cleaner reports and auto-adjusts Excel column widths.

## Requirements

* **Python 3.x**
* **Pillow (PIL):** For image opening and metadata parsing.
* **Pandas & Openpyxl:** For data structuring and Excel report generation.

## Installation & Usage

1. **Clone the repo** and install dependencies:
   ```bash
   pip install Pillow pandas openpyxl
