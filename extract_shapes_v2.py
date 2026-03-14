import win32com.client
import os

ppt_file = r"C:\Users\307984\.openclaw\workspace\temp_attachments\AI.pptx"
out_dir = r"C:\Users\307984\.openclaw\workspace\temp_attachments\extracted_shapes"

ppt_app = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt_app.Presentations.Open(ppt_file, False, False, False)

for i, slide in enumerate(presentation.Slides):
    for j, shape in enumerate(slide.Shapes):
        if shape.Type == 13 or shape.Type == 11 or shape.Type == 14:
            out_path = os.path.join(out_dir, f"slide_{i+1}_shape_{j+1}_v2.png")
            shape.Export(out_path, 2)  # Wait, doc says for PP Shape.Export(PathName, Filter, ...) filter is ppShapeFormatPNG = 2. Wait, no. Filter is a pbPictureFormat? 
            # Actually, `shape.Export(out_path, 2)` uses ppShapeFormatPNG in PowerPoint, BUT ppt_app might be using string "PNG".
            shape.Export(out_path, "PNG") 

presentation.Close()
ppt_app.Quit()
print("Shape export complete")
