import os
from pptx import Presentation
import win32com.client

def count_slides_in_pptx(pptx_path):
    """Counts the number of slides in a .pptx file"""
    try:
        prs = Presentation(pptx_path)
        return len(prs.slides)
    except Exception as e:
        print(f"Error processing {pptx_path}: {e}")
        return 0

def count_slides_in_ppt(ppt_path):
    """Counts the number of slides in a .ppt file using PowerPoint COM"""
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)
        slide_count = presentation.Slides.Count
        presentation.Close()
        return slide_count
    except Exception as e:
        print(f"Error processing {ppt_path}: {e}")
        return 0

def process_ppt_files(directory):
    """Processes all .pptx and .ppt files in the given directory and counts slides"""
    total_slides = 0
    file_slide_count = {}

    # Walk through the directory and its subdirectories
    for root, dirs, files in os.walk(directory):
        for file in files:
            ppt_path = os.path.join(root, file)
            if file.endswith('.pptx'):
                slide_count = count_slides_in_pptx(ppt_path)
            elif file.endswith('.ppt'):
                slide_count = count_slides_in_ppt(ppt_path)
            else:
                continue

            file_slide_count[file] = slide_count
            total_slides += slide_count
            print(f"{file} -> {slide_count} slides {total_slides}")

    print(f"\nTotal slides across all presentations: {total_slides}")
    return file_slide_count, total_slides

if __name__ == "__main__":
    # Prompt user for directory input
    directory = input("Enter the directory path: ").strip()
    process_ppt_files(directory)
