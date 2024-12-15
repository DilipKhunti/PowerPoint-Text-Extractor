import os
import glob
import zipfile
from bs4 import BeautifulSoup

def extract_ppt_xml(ppt_file_path, output_dir):
    try:
        # Change the PPT file extension to .zip
        zip_file_path = ppt_file_path.replace(".pptx", ".zip")

        # Rename the PPT file to .zip and extract it
        os.rename(ppt_file_path, zip_file_path)

        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)

        # After extracting, rename the .zip file back to .pptx
        os.rename(zip_file_path, ppt_file_path)

        print(f"Extracted {ppt_file_path} content to {output_dir}")
    except Exception as e:
        print(f"Error processing {ppt_file_path}: {e}")
        
def extract_text_from_slides(slides_dir, output_file_path):
    try:
        # Get all slide XML files in the slides directory
        slide_files = glob.glob(os.path.join(slides_dir, "slide*.xml"))
        
        with open(output_file_path, "w", encoding="utf-8") as out_file:
            for slide_number, slide_file in enumerate(sorted(slide_files), start=1):
                # Read and parse each slide
                with open(slide_file, "r", encoding="utf-8") as file:
                    xml_content = file.read()

                soup = BeautifulSoup(xml_content, "xml")
                text_elements = soup.find_all("a:t")
                extracted_text = " ".join([t.get_text() for t in text_elements])

                # Write slide content to output file with separator
                out_file.write(f"--- Slide {slide_number} ---\n")
                out_file.write(extracted_text)
                out_file.write("\n\n")

        print(f"Extracted text from all slides has been saved to {output_file_path}")
    except Exception as e:
        print(f"Error extracting text from slides: {e}")

# Relative paths
ppt_file_dir = "ppts"
output_base_dir = "extracts"
output_txt_dir = "outputs"

# Ensure output directory for text files exists
if not os.path.exists(output_txt_dir):
    os.makedirs(output_txt_dir)

# Get all .pptx files in the directory
pptx_files = glob.glob(os.path.join(ppt_file_dir, "*.pptx"))

# Process each PPT file
for index, ppt_file_path in enumerate(pptx_files, start=1):
    # Extract the PPT content
    output_dir = os.path.join(output_base_dir, f"ppt{index}")
    extract_ppt_xml(ppt_file_path, output_dir)

    # Extract text from slides
    slides_dir = os.path.join(output_dir, "ppt/slides")
    output_file_path = os.path.join(output_txt_dir, f"output_ppt{index}.txt")
    extract_text_from_slides(slides_dir, output_file_path)
