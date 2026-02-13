PDF Combiner - Instructions

Getting Started
Use the tabs at the top of the screen to switch between selecting input files, configuring output settings and options, and showing these instructions.

Adding Files
Click "Add PDFs/Images..." to select PDF and image files to combine.
Select one or multiple files to combine using the file browser.
Supported file formats are: PDF, JPG, JPEG, PNG, BMP, GIF, TIFF.
Images will be automatically converted to PDF during the merge.
The same file cannot be added twice. Only one copy will be retained on the list.
The list of files you select will be displayed, with PDFs in black text and images in blue text.

Organizing Files in the File List
Click a file to select it and use the move up/down buttons to reorder it.
Use Click and Ctrl+Click to select more than one file if you want to "Delete Selected" files.
Hover over a file row to see a preview and full path and name.

Saving and Loading Lists
Click "Load/Save List" to manage file lists of PDFs/images you want to convert in bulk.
Save Current List: Export your current file list to a .pdflist file. This is useful for reusing the same set of files later.
Load Previously Saved List: Import a previously saved list. Choose to append files to your current list or replace it entirely. Duplicate files are automatically skipped. Saved lists preserve file properties such as rotation, page ranges, and reverse settings.

Sorting
Click column headers (Filename, File Size, Date) to sort the list. Sorting resets any custom ordering. Click again to reverse the sort order. An up arrow means ascending order; a down arrow means descending order.

File Properties
Page rotation: Set 0°, 90°, 180°, or 270° clockwise for each file. For images, rotation is applied during conversion to PDF.
Pages: Specify which pages to include using:
  "All" or leave empty to include all pages.
  Single page: 5
  Range: 1-10
  Multiple ranges: 1-3,5,7-9
For images, page selection is not available.
Rev: Check to reverse the page order for that file. For images, this option is not available.

Output Settings
Enter the filename for the combined PDF.
Click "Browse" to choose where to save the combined PDF.
Add filename bookmarks: Creates PDF bookmarks from each source file name; existing bookmarks will be removed when this option is selected.
Insert breaker pages: Adds a separator page before each file showing which file follows.
Scale all pages to uniform size: Makes all pages the same size; results may vary with mixed page sizes.
Ignore blank pages: Skips blank pages in source files when combining.
Compression/Quality: Choose a level to reduce file size; higher compression results in smaller files but lower quality.

Metadata and Watermark
Add PDF metadata: Includes Title, Author, Subject, and Keywords in the combined PDF.
Add watermark to pages: Overlays text on all pages. You can set the text, opacity, font size, and rotation angle.

Combining PDFs and Images
At least two files must be selected to enable combining.
Images and PDFs can be mixed in any order.
Images are converted to PDF automatically during the merge.
Click "Combine PDFs" to merge the files.
Review the summary and click "Proceed."
The combined PDF will be created at your chosen location.

Status Bar
The bottom status bar shows the number of files and their total size.

Memory Considerations
Each PDF is loaded entirely into memory, so RAM can be a bottleneck. Try to keep individual PDF sizes under 100 Mb for reliable performance.

Number of PDFs to Combine
There is no hard-coded limit, but more than 100 files may be combined depending on their sizes. The app processes files sequentially, so performance depends on available RAM, the combined size of all source PDFs, and the accumulated pages in the PdfWriter object before writing to disk.

Combined Output Size
The combined PDF can be as large as your disk space and available RAM allow. Large combined PDFs (1000+ pages or multiple large files) may cause slow progress, high memory usage during compression, and long write times.

Real-World Guidelines
Source PDFs: Keep each file under 100 Mb for smooth operation.
Number of files: Combining 2–50 files is very reliable; 50–100+ may slow down depending on size.

Key Factors Affecting Performance
PyPDF2 slows down with very large files.
Compression increases RAM and CPU usage.
Page scaling and transformations add memory overhead per page.
Available system RAM limits the total size you can process.