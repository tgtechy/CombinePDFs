# PDF Combiner - Instructions

## Getting Started
- Use the tabs at the top of the screen to switch between selecting input files, configuring output settings and options and showing these instructions.

## Adding Files

- Click **"Add PDFs/Images..."** button to select PDF and image files to combine
- Select one or multiple files to combine using the file browser
- Supported file formats are: PDF, JPG, JPEG, PNG, BMP, GIF, TIFF
  - Images will be automatically converted to PDF during the merge
- The same file cannot be added twice, a notice will be displayed if this occurs
- The list of files you select will be displayed, with PDFs in black text and images in blue text

## Organizing Files in the File List

- Drag files up or down to change the order they'll be combined
- Click a file to select it, Ctrl+Click on other files to select more than one file
- Hover over a file row to see the full path in the status bar at the bottom of the screen

## Saving and Loading Lists

- Click **"Load/Save List"** to manage file lists
- **Save Current List:** Export your current file list to a `.pdflist` file
  - Useful for reusing the same set of files later
- **Load Previously Saved List:** Import a previously saved list
  - Choose to append files to your current list or replace it entirely
  - Duplicate files are automatically skipped
  - Shows a count of valid files loaded and any duplicate files skipped
- Saved lists preserve file properties like rotation, page ranges, and page reverse settings

## Sorting

- Click column headers (Filename, File Size, Date) to sort the list
  - Note that this resets any custom ordering
- Click again to reverse the sort order (arrows show sort direction)
- An up arrow (▲) means the sort is in ascending order, a down arrow (▼) means it is in descending order

## File Properties

- **Page rotation:** Set 0°, 90°, 180°, or 270° (clockwise) for each individual file
  - For images, rotation is applied during conversion to PDF
- **Pages:** Specify which pages to include in the combined PDF using:
  - `"All"` or leave empty to include all pages
  - Single page: `"5"` (without the quotes)
  - Range: `"1-10"` (without the quotes)
  - Multiple ranges: `"1-3,5,7-9"` (without the quotes)
  - For images, this selection is not available
- **Rev:** Check to reverse the page order for that file
  - For images, this selection is not available

## Output Settings

- Enter the filename for the combined PDF
- Click **"Browse"** to choose where to save the combined PDF
- Check **"Add filename bookmarks"** to create PDF bookmarks from each source file's name in the combined PDF, existing bookmarks will be removed when selecting this option.
- Check **"Insert breaker pages"** to add a separator page before each file showing which file follows
- Check **"Scale all pages to uniform size"** to make all pages the same size (this may produce unpredictable results with varying page sizes)
- Check **"Ignore blank pages"** to skip blank pages in source files when combining
- Select **Compression/Quality** level to reduce file size (higher compression = smaller file but lower quality)

## Metadata & Watermark

- Check **"Add PDF metadata"** to include Title, Author, Subject, and Keywords in the combined PDF
- Check **"Add watermark to pages"** to overlay text on all pages
  - Set text, opacity, font size, and rotation angle

## Combining PDFs and Images

- At least 2 files are required to be selected to enable combining
- Images and PDFs can be mixed in any order
- Images will be converted to PDF automatically during the merge process
- Click the **"Combine PDFs"** button to merge the files
- Review the summary and click **"Proceed"**
- The combined PDF will be created at your chosen location

## Preview

- Hover over a file in the list to see a thumbnail of its first page (for PDFs) or the image itself
- The thumbnail also shows the full filename.

## Status Bar

- The bottom status bar shows the full path of the file you're currently hovering over

### Memory Considerations

- Each PDF is loaded entirely into memory, so RAM can be a bottleneck
- Try to keep individual PDF sizes under 1 GB for reliable performance

### Number of PDFs to Combine

- There is no hard-coded limit, but more than 100 files can be combined depending on their sizes
- The app processes files sequentially, so it's mainly constrained by:
  - Total available RAM (all pages accumulate in a PdfWriter object before writing to disk)
  - Combined size of all source PDFs

### Combined Output Size

- Can theoretically be as large as your disk space and available RAM
- If you're generating a combined PDF with 1000+ pages and/or multiple large files, you may experience:
  - Slow progress updates
  - Memory strain during the compression phase (if compression is enabled)
  - Long write times

### Real-World Guidelines

- Source PDFs: Keep each file under 1 GB for smooth operation
- Number of files: 2-50 files to combine is very reliable; 50-100+ will slow down based on their size

### Key Factors Affecting Performance

- The PDF engine (PyPDF2) efficiency handles several GB sized files but it slows down with size
- Compression consumes RAM and processing time
- Page scaling/transformations add memory overhead per page
- Available system RAM limits the aggregate size you can process