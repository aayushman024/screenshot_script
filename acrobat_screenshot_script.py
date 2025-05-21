import os
import time
import argparse
import subprocess
import tempfile
import pyautogui
import pygetwindow as gw
from PIL import Image

def create_dir_if_not_exists(directory):
    """Create directory if it doesn't exist."""
    if not os.path.exists(directory):
        os.makedirs(directory)
        print(f"Created directory: {directory}")

def open_pdf_with_acrobat(pdf_path):
    """Open a PDF file with Adobe Acrobat."""
    print(f"Opening PDF with Adobe Acrobat: {pdf_path}")
    
    # Use absolute path for the PDF file
    abs_pdf_path = os.path.abspath(pdf_path)
    
    try:
        # Try to find Acrobat executable
        acrobat_paths = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
        ]
        
        acrobat_path = None
        for path in acrobat_paths:
            if os.path.exists(path):
                acrobat_path = path
                break
        
        if acrobat_path is None:
            print("Adobe Acrobat executable not found. Please specify the path.")
            return False
        
        # Open PDF with Acrobat
        subprocess.Popen([acrobat_path, abs_pdf_path])
        
        # Give Acrobat time to open
        time.sleep(5)
        
        # Try to find the Acrobat window
        acrobat_windows = [win for win in gw.getAllWindows() if "Acrobat" in win.title or "PDF" in win.title]
        
        if not acrobat_windows:
            print("Cannot find Adobe Acrobat window. Please ensure the PDF is open.")
            return False
        
        # Activate the Acrobat window
        acrobat_windows[0].activate()
        time.sleep(1)
        
        # Press Escape to close any dialogs that might appear
        pyautogui.press('escape')
        time.sleep(0.5)
        
        # Toggle fullscreen mode (F11 or Ctrl+L for full screen in Acrobat)
        pyautogui.hotkey('ctrl', 'l')
        time.sleep(1)
        
        # Enter presentation mode (Ctrl+Shift+H for Acrobat DC)
        # If this doesn't work, try alternative key combinations
        pyautogui.hotkey('ctrl', 'shift', 'h')
        time.sleep(1)
        
        # Success
        return True
    
    except Exception as e:
        print(f"Error opening PDF with Acrobat: {e}")
        return False

def goto_first_page():
    """Navigate to the first page of the PDF."""
    # Press Home key to go to first page
    pyautogui.press('home')
    time.sleep(0.5)

def count_pages():
    """Try to determine the number of pages in the PDF.
    Since we can't directly get this from Acrobat via automation,
    we'll need to rely on user input or other means."""
    return None  # We'll handle this differently

def capture_screenshots(output_dir, delay=1):
    """Capture screenshots of each page in the PDF.
    This function will need to be controlled by the user to determine when to stop."""
    print("Starting screenshot capture process...")
    
    # List to store screenshot paths
    screenshot_paths = []
    
    # Go to first page
    goto_first_page()
    time.sleep(1)
    
    current_page = 1
    has_more_pages = True
    
    while has_more_pages:
        # Take a screenshot
        screenshot_path = os.path.join(output_dir, f"page_{current_page:03}.png")
        pyautogui.screenshot(screenshot_path)
        screenshot_paths.append(screenshot_path)
        
        print(f"Captured screenshot of page {current_page} to {screenshot_path}")
        
        # Try to go to next page (Right arrow or Page Down)
        last_screenshot = screenshot_path
        pyautogui.press('right')
        # Wait for page transition
        time.sleep(delay)
        
        # Take another screenshot to compare
        temp_screenshot = os.path.join(output_dir, "temp.png")
        pyautogui.screenshot(temp_screenshot)
        
        # Compare screenshots to check if we reached the end
        if compare_images(last_screenshot, temp_screenshot):
            # Images are the same, meaning we're probably at the last page
            os.remove(temp_screenshot)
            has_more_pages = False
        else:
            # Images are different, we moved to a new page
            os.remove(temp_screenshot)
            current_page += 1
    
    # Exit fullscreen (Escape key)
    pyautogui.press('escape')
    time.sleep(0.5)
    
    return screenshot_paths

def compare_images(image1_path, image2_path, threshold=1000):
    """Compare two images to check if they are essentially the same."""
    from PIL import Image, ImageChops
    import math
    
    # Open images
    image1 = Image.open(image1_path)
    image2 = Image.open(image2_path)
    
    # Ensure images are the same size
    if image1.size != image2.size:
        return False
    
    # Get difference
    diff = ImageChops.difference(image1, image2)
    
    # Calculate the mean of the difference
    stat = diff.convert("L").getdata()
    diff_value = sum(stat)
    
    # Return True if the difference is below threshold
    return diff_value < threshold

def combine_images_to_pdf(image_paths, output_pdf):
    """Combine multiple images into a single PDF file."""
    if not image_paths:
        print("No images to combine.")
        return
    
    print(f"Combining {len(image_paths)} images into PDF: {output_pdf}")
    
    # Open the first image
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

def close_acrobat():
    """Close Adobe Acrobat."""
    # Try to find and close Acrobat window
    acrobat_windows = [win for win in gw.getAllWindows() if "Acrobat" in win.title or "PDF" in win.title]
    
    for window in acrobat_windows:
        try:
            window.close()
        except:
            pass
    
    # Allow time for closing
    time.sleep(1)

def process_pdf_with_acrobat(input_pdf, output_pdf, temp_dir=None, delay=1, cleanup=True):
    """Process a PDF using Adobe Acrobat for opening, capturing screenshots, and combining."""
    # Create a temporary directory if not provided
    if temp_dir is None:
        temp_dir = tempfile.mkdtemp()
        print(f"Created temporary directory: {temp_dir}")
    else:
        create_dir_if_not_exists(temp_dir)
    
    try:
        # Open PDF with Acrobat
        if not open_pdf_with_acrobat(input_pdf):
            return False
        
        # Capture screenshots
        print("\nCapturing screenshots of each page...")
        image_paths = capture_screenshots(temp_dir, delay)
        
        # Combine the images into a PDF
        combine_images_to_pdf(image_paths, output_pdf)
        
        # Close Acrobat
        close_acrobat()
        
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
        
        return True
    
    except Exception as e:
        print(f"Error processing PDF: {e}")
        # Try to close Acrobat on error
        close_acrobat()
        return False

def main():
    # Parse command line arguments
    parser = argparse.ArgumentParser(description='Take screenshots of PDF pages using Adobe Acrobat and combine them into a new PDF.')
    parser.add_argument('input_pdf', help='Input PDF file path')
    parser.add_argument('output_pdf', nargs='?', help='Output PDF file path (optional)')
    parser.add_argument('--temp_dir', help='Directory to store temporary image files')
    parser.add_argument('--delay', type=float, default=1.0, help='Delay between page turns (seconds)')
    parser.add_argument('--keep_temp', action='store_true', help='Keep temporary image files')
    
    args = parser.parse_args()
    
    # If output_pdf is not provided, generate a name based on input file
    if args.output_pdf is None:
        input_base = os.path.basename(args.input_pdf)
        input_name = os.path.splitext(input_base)[0]
        output_dir = os.path.dirname(args.input_pdf) or '.'
        args.output_pdf = os.path.join(output_dir, f"{input_name}_screenshot.pdf")
    
    # Process the PDF
    start_time = time.time()
    
    success = process_pdf_with_acrobat(
        args.input_pdf,
        args.output_pdf,
        args.temp_dir,
        args.delay,
        not args.keep_temp
    )
    
    if success:
        print(f"\nTotal processing time: {time.time() - start_time:.2f} seconds")
        print(f"\nOutput PDF saved to: {args.output_pdf}")
    else:
        print("\nPDF processing failed.")

if __name__ == "__main__":
    main()