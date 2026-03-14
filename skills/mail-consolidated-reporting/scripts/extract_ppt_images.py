import win32com.client
from PIL import ImageGrab
import os
import time

def extract_ppt_images(ppt_path, output_dir):
    """
    Extracts images from PPT slides using the clipboard (DRM bypass).
    """
    os.makedirs(output_dir, exist_ok=True)
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True # Required for Copy()
    try:
        presentation = ppt_app.Presentations.Open(ppt_path, False, False, False)
        for i, slide in enumerate(presentation.Slides):
            for j, shape in enumerate(slide.Shapes):
                if shape.Type == 13 or shape.Type == 11: # Picture/Linked
                    shape.Copy()
                    time.sleep(1) # Clipboard buffer
                    img = ImageGrab.grabclipboard()
                    if img and hasattr(img, 'save'):
                        img.save(os.path.join(output_dir, f"slide_{i+1}.png"), "PNG")
        presentation.Close()
    finally:
        ppt_app.Quit()

if __name__ == "__main__":
    # Example usage
    # extract_ppt_images("path/to/ppt", "output/dir")
    pass
