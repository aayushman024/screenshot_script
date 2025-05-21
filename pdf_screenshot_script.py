import os
import argparse
import fitz  # PyMuPDF
from PIL import Image
import tempfile
import time

def create_dir_if_not_exists(directory):
    """Create directory if it doesn't exist."""
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"Created directory: {directory}")

def take_pdf_screenshots(input_pdf, output_dir, dpi=300):
    """Take screenshots of each page in a PDF file and save them as images."""
    print(f"Processing PDF: {input_pdf}")
    
    # Open the PDF file
    doc = fitz.open(input_pdf)
    total_pages = len(doc)
    print(f"Total pages: {total_pages}")
    
    # Create a list to store the paths of the saved images
    image_paths = []
    
    # Process each page
    for page_num in range(total_pages):
        # Get the page
        page = doc.load_page(page_num)
        
        # Set the resolution (pixels per inch)
        zoom = dpi / 72  # 72 is the default resolution in PyMuPDF
        
        # Create a matrix for zooming
        mat = fitz.Matrix(zoom, zoom)
        
        # Render page to an image
        pix = page.get_pixmap(matrix=mat)
        
        # Define output image path
        output_image = os.path.join(output_dir, f"page_{page_num + 1:03}.png")
        
        # Save the image
        pix.save(output_image)
        
        # Add the image path to the list
        image_paths.append(output_image)
        
        print(f"Saved screenshot of page {page_num + 1}/{total_pages} to {output_image}")
    
    # Close the PDF
    doc.close()
    
    return image_paths

def combine_images_to_pdf(image_paths, output_pdf):
    """Combine multiple images into a single PDF file."""
    print(f"Combining {len(image_paths)} images into PDF: {output_pdf}")
    
    # Open the first image to get dimensions
    first_img = Image.open(image_paths[0])
    
    # Create a list to store all images
    images = []
    
    # Open all images (except the first one which we already opened)
    for img_path in image_paths[1:]:
        img = Image.open(img_path)
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        images.append(img)
    
    # Convert first image to RGB if needed
    if first_img.mode == 'RGBA':
        first_img = first_img.convert('RGB')
    
    # Save the combined PDF
    first_img.save(
        output_pdf,
        save_all=True,
        append_images=images,
        resolution=100.0,
        quality=95,
        optimize=True
    )
    
    print(f"Successfully created PDF: {output_pdf}")

def process_pdf(input_pdf, output_pdf, temp_dir=None, dpi=300, cleanup=True):
    """Process a PDF by taking screenshots of each page and combining them."""
    # Create a temporary directory if not provided
    if temp_dir is None:
        temp_dir = tempfile.mkdtemp()
        print(f"Created temporary directory: {temp_dir}")
    else:
        create_dir_if_not_exists(temp_dir)
    
    try:
        # Take screenshots of each page
        image_paths = take_pdf_screenshots(input_pdf, temp_dir, dpi)
        
        # Combine the images into a PDF
        combine_images_to_pdf(image_paths, output_pdf)
        
        # Clean up temporary image files if requested
        if cleanup:
            print("Cleaning up temporary files...")
            for img_path in image_paths:
                os.remove(img_path)
            
            # Try to remove the temporary directory if it's empty
            try:
                os.rmdir(temp_dir)
                print(f"Removed temporary directory: {temp_dir}")
            except OSError:
                print(f"Directory {temp_dir} not empty or not removable.")
    
    except Exception as e:
        print(f"Error processing PDF: {e}")
        raise

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Take screenshots of PDF pages and combine them into a new PDF.')
    parser.add_argument('input_pdf', help='Input PDF file path')
    parser.add_argument('output_pdf', help='Output PDF file path')
    parser.add_argument('--temp_dir', help='Directory to store temporary image files')
    parser.add_argument('--dpi', type=int, default=300, help='DPI (resolution) for screenshots')
    parser.add_argument('--keep_temp', action='store_true', help='Keep temporary image files')
    
    args = parser.parse_args()
    
    # Process the PDF
    process_pdf(
        args.input_pdf,
        args.output_pdf,
        args.temp_dir,
        args.dpi,
        not args.keep_temp
    )

if __name__ == "__main__":
    start_time = time.time()
    main()
    print(f"Total processing time: {time.time() - start_time:.2f} seconds")