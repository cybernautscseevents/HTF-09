import os
import win32com.client

def export_pptx(pptx_path, pdf_path, preview_dir):
    abs_pptx = os.path.abspath(pptx_path)
    abs_pdf = os.path.abspath(pdf_path)
    abs_preview = os.path.abspath(preview_dir)
    os.makedirs(abs_preview, exist_ok=True)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    # Open presentation (ReadOnly=True, Untitled=False, WithWindow=False)
    deck = powerpoint.Presentations.Open(abs_pptx, True, False, False)

    try:
        # Save as PDF (Format type 32 is ppSaveAsPDF)
        deck.SaveAs(abs_pdf, 32)
        print("Exported PDF to:", abs_pdf)

        # Export each slide as PNG image
        for i, slide in enumerate(deck.Slides):
            slide_img_path = os.path.join(abs_preview, f"slide_{i+1:02d}.png")
            slide.Export(slide_img_path, "PNG", 1920, 1080)
            print(f"Exported Slide {i+1} to:", slide_img_path)
    finally:
        deck.Close()
        powerpoint.Quit()

if __name__ == "__main__":
    export_pptx(
        "SCAMSHIELD_Hackatopia_Final_Presentation.pptx",
        "SCAMSHIELD_Hackatopia_Final_Presentation.pdf",
        "presentation_preview"
    )
