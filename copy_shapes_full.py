import win32com.client
from PIL import ImageGrab
import os
import time

ppt_file = r"C:\Users\307984\.openclaw\workspace\temp_attachments\AI_space.pptx"
out_dir = r"C:\Users\307984\.openclaw\workspace\temp_attachments\clipboard_images_full"
os.makedirs(out_dir, exist_ok=True)

ppt_app = win32com.client.Dispatch("PowerPoint.Application")
ppt_app.Visible = True
presentation = ppt_app.Presentations.Open(ppt_file, False, False, False)

for i, slide in enumerate(presentation.Slides):
    for j, shape in enumerate(slide.Shapes):
        if shape.Type == 13 or shape.Type == 11:
            shape.Copy()
            time.sleep(1) # wait longer for clipboard
            try:
                img = ImageGrab.grabclipboard()
                if img and hasattr(img, 'save'):
                    img.save(os.path.join(out_dir, f"slide_{i+1}.png"), "PNG")
                    print(f"Saved slide {i+1}")
            except Exception as e:
                print(f"Error on slide {i+1}: {e}")

presentation.Close()
ppt_app.Quit()
