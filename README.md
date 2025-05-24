#Folder Slide Counter

A simple Python utility that traverses a user-specified directory (and its subfolders), identifies all PowerPoint files (both `.pptx` and legacy `.ppt`), and reports the number of slides in each.

Key Features

-Recursive Directory Walk: Automatically explores all subdirectories to find every PowerPoint file.
-Dual Format Support:

  -`.pptx` files are handled via the `python-pptx` library for direct slide enumeration.
  -`.ppt` files are processed through the Windows COM interface (requires PowerPoint to be installed) to accurately count slides in legacy presentations.
-Per-File Reporting: Displays each filename alongside its individual slide count as it’s processed.
-Running Total: Maintains and prints a cumulative tally of slides across all discovered presentations.
-Robust Error Handling: Catches and reports any exceptions encountered while opening or reading a file, ensuring that a single problematic file won’t halt the entire script.

Dependencies

-python-pptx: For reading `.pptx` files.
-pywin32 (win32com): To automate PowerPoint via COM for `.ppt` slide counts.
-Windows PowerPoint (for COM automation)

How It Works

1. User Prompt: On launch, the script asks the user to enter the path to the target folder.
2. File Discovery: It walks every directory under that path, checking file extensions.
3. Slide Counting:

    For each `.pptx`, it loads the file with `Presentation()` and counts the `slides` collection.
    For each `.ppt`, it opens the file invisibly in PowerPoint via COM and reads its `Slides.Count`.
4. Reporting: After each file, the script prints the file’s slide count and updates the running total.
5. Summary: Once all files are processed, it displays the grand total of slides found.

This utility is ideal for anyone who needs a quick audit of slide volumes across many presentations—useful for preparing batch exports, estimating recording time, or simply inventorying corporate slide decks.

