import win32com.client
import os

ppt_file = r"C:\Users\307984\.openclaw\workspace\temp_attachments\AI.pptx"
out_dir = r"C:\Users\307984\.openclaw\workspace\temp_attachments\extracted_shapes"
os.makedirs(out_dir, exist_ok=True)

try:
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    presentation = ppt_app.Presentations.Open(ppt_file, False, False, False)
    
    for i, slide in enumerate(presentation.Slides):
        for j, shape in enumerate(slide.Shapes):
            if shape.Type == 13 or shape.Type == 11 or shape.Type == 14: # Picture/Linked/Placeholder
                out_path = os.path.join(out_dir, f"slide_{i+1}_shape_{j+1}.jpg")
                shape.Export(out_path, 2) # 2 = JPG
    presentation.Close()
    ppt_app.Quit()
    print("Shape export complete")
except Exception as e:
    print(f"Error: {e}")
