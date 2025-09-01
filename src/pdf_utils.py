from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from pypdf import PdfWriter, PdfReader
import os
from src.log import log_event

def generate_summary_page(tickets, month, board_id, output_path, attachment_map=None):
    c = canvas.Canvas(output_path, pagesize=A4)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, 800, f"Resolved Tickets — {month}")
    c.setFont("Helvetica", 12)
    c.drawString(50, 780, f"Board: {board_id}")
    c.drawString(50, 765, f"Generated on: {os.path.basename(output_path)}")
    c.drawString(50, 750, f"Total Tickets: {len(tickets)}")
    
    # Add deduplication info if available
    if attachment_map:
        unique_attachments = len(attachment_map)
        total_attachments = sum(len(t.attachments) for t in tickets)
        c.drawString(50, 735, f"Unique Attachments: {unique_attachments} (saved {total_attachments - unique_attachments} duplicates)")
    
    c.setFont("Helvetica-Bold", 12)
    header_y = 710 if attachment_map else 720
    c.drawString(30, header_y, "#")
    c.drawString(60, header_y, "Ticket #")
    c.drawString(200, header_y, "Open Date")
    c.drawString(300, header_y, "Close Date")
    c.drawString(420, header_y, "Attachments")
    y = header_y - 20
    c.setFont("Helvetica", 11)
    for idx, t in enumerate(tickets, 1):
        c.drawString(30, y, str(idx))
        c.drawString(60, y, t.item_name)
        c.drawString(200, y, t.open_date)
        c.drawString(300, y, t.close_date or "")
        c.drawString(420, y, str(len(t.attachments)))
        y -= 18
        if y < 50:
            c.showPage()
            c.setFont("Helvetica-Bold", 12)
            c.drawString(30, 780, "#")
            c.drawString(60, 780, "Ticket #")
            c.drawString(200, 780, "Open Date")
            c.drawString(300, 780, "Close Date")
            c.drawString(420, 780, "Attachments")
            y = 760
            c.setFont("Helvetica", 11)
    
    # Add shared attachments reference page if available
    if attachment_map:
        shared_attachments = [(asset.name, tickets) for asset, tickets in attachment_map.values() if len(tickets) > 1]
        if shared_attachments:
            c.showPage()
            c.setFont("Helvetica-Bold", 16)
            c.drawString(50, 800, "Shared Attachments Reference")
            c.setFont("Helvetica", 12)
            c.drawString(50, 780, f"The following {len(shared_attachments)} files appear in multiple tickets:")
            y = 750
            c.setFont("Helvetica", 11)
            for filename, ticket_list in shared_attachments:
                c.drawString(50, y, f"• {filename}")
                y -= 15
                for ticket in ticket_list:
                    c.drawString(70, y, f"  - {ticket}")
                    y -= 12
                y -= 5  # Extra space between files
                if y < 80:
                    c.showPage()
                    y = 780
    
    c.save()
    log_event(action="summary_page", status="success")
    return output_path

def merge_pdfs(summary_pdf, item_pdfs, output_path):
    writer = PdfWriter()
    writer.append(PdfReader(summary_pdf))
    for pdf in item_pdfs:
        writer.append(PdfReader(pdf))
    with open(output_path, "wb") as f:
        writer.write(f)
    log_event(action="merge", status="success")
    return output_path
