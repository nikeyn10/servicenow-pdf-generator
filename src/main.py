import argparse
import os
import yaml
from dotenv import load_dotenv
from src.monday_client import MondayClient
from src.filters import get_month_range, get_status_index
from src.files import download_asset, dedupe_assets
from src.convert import to_pdf
from src.pdf_utils import generate_summary_page, merge_pdfs
from src.excel_utils import create_monthly_excel_summary
from src.models import Asset, TicketRow
from src.log import log_event

def verify_all_tickets_found(client, config, month, found_tickets):
    """Double-check that we found all resolved tickets for the month"""
    print(f"\n=== DOUBLE-CHECK: Verifying all resolved tickets for {month} ===")
    
    # Get ALL items with full pagination to ensure we didn't miss any
    items_data = client.get_items_page(config['board']['id'], '', '', 0, limit=500)
    all_items = items_data['boards'][0]['items_page']['items']
    cursor = items_data['boards'][0]['items_page']['cursor']
    
    while cursor:
        next_page = client.next_items_page(cursor, limit=500) 
        next_items = next_page['next_items_page']['items']
        cursor = next_page['next_items_page']['cursor']
        all_items.extend(next_items)
    
    print(f"Total items fetched for verification: {len(all_items)}")
    
    resolved_for_month = []
    for item in all_items:
        status_label = None
        open_date_str = None
        for cv in item.get('column_values', []):
            if cv['id'] == config['board']['columns']['status']:
                status_label = cv.get('text')
            if cv['id'] == config['board']['columns']['date_filter']:
                open_date_str = cv.get('text')
        
        if (status_label and status_label.lower() == config['board']['status_label_required'].lower() 
            and open_date_str and month in open_date_str):
            # Only count tickets that have attachments (same filter as main processing)
            assets = item.get('assets', [])
            if len(assets) > 0:
                resolved_for_month.append({
                    'id': item['id'],
                    'name': item['name'],
                    'date': open_date_str,
                    'status': status_label,
                    'attachments': len(assets)
                })
    
    found_ids = {t['id'] for t in found_tickets}
    all_resolved_ids = {t['id'] for t in resolved_for_month}
    
    print(f"Found in processing: {len(found_tickets)} tickets")
    print(f"Found in verification: {len(resolved_for_month)} tickets")
    
    missing = all_resolved_ids - found_ids
    if missing:
        print(f"⚠️  WARNING: Missing tickets: {missing}")
        for ticket in resolved_for_month:
            if ticket['id'] in missing:
                print(f"  - {ticket['name']} ({ticket['date']}) - {ticket['status']} - {ticket['attachments']} attachments")
    else:
        print("✅ All resolved tickets found - no missing tickets")
    
    return len(missing) == 0

def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description="ServiceNow Monthly PDF Generator")
    parser.add_argument("--month", required=True)
    parser.add_argument("--config", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--downloads", required=True)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    with open(args.config) as f:
        config = yaml.safe_load(f)

    client = MondayClient(config)
    first_day, last_day = get_month_range(args.month)
    status_col = client.get_status_column(config['board']['id'])
    settings_str = status_col['boards'][0]['columns'][0]['settings_str']
    status_idx = get_status_index(settings_str, config['board']['status_label_required'])
    max_items = config.get('run', {}).get('max_items', None)
    items_data = client.get_items_page(config['board']['id'], first_day, last_day, status_idx, limit=500)
    items = items_data['boards'][0]['items_page']['items']
    cursor = items_data['boards'][0]['items_page']['cursor']
    all_items = items[:]
    while cursor:
        next_page = client.next_items_page(cursor, limit=500)
        next_items = next_page['next_items_page']['items']
        cursor = next_page['next_items_page']['cursor']
        all_items.extend(next_items)

    # Filter client-side for month and status
    filtered_items = []
    for item in all_items:
        status_label = None
        open_date_str = None
        for cv in item.get('column_values', []):
            if cv['id'] == config['board']['columns']['status']:
                status_label = cv.get('text')
            if cv['id'] == config['board']['columns']['date_filter']:
                open_date_str = cv.get('text')
        
        # Check if status matches "Resolved" 
        if status_label and status_label.lower() == config['board']['status_label_required'].lower():
            # More precise date matching - only exact month match
            if open_date_str and args.month in open_date_str:
                # Only include tickets that have attachments
                assets = item.get('assets', [])
                if len(assets) > 0:
                    filtered_items.append(item)
                    print(f"DEBUG: Found ticket {item['name']} with date {open_date_str} for month {args.month} ({len(assets)} attachments)")
                else:
                    print(f"DEBUG: Skipping ticket {item['name']} - no attachments")
                # Only apply max_items limit in dry-run mode for testing
                if args.dry_run and max_items and len(filtered_items) >= max_items:
                    break

    ticket_rows = []
    attachment_map = {}  # Maps filename -> (asset, [ticket_names])
    
    for item in filtered_items:
        assets = [Asset(**a) for a in item.get('assets', [])]
        assets = dedupe_assets(assets)
        open_date = None
        close_date = None
        for cv in item.get('column_values', []):
            if cv['id'] == config['board']['columns']['date_filter']:
                open_date = cv['text']
            if cv['id'] == config['board']['columns']['close_date']:
                close_date = cv['text']
        
        # Track which tickets use which attachments (by filename)
        for asset in assets:
            filename = asset.name
            if filename not in attachment_map:
                attachment_map[filename] = (asset, [])
            attachment_map[filename][1].append(item['name'])
        
        ticket_rows.append(TicketRow(
            item_id=item['id'],
            item_name=item['name'],
            open_date=open_date or "",
            close_date=close_date,
            attachments=assets
        ))

    if args.dry_run:
        print(f"Found {len(ticket_rows)} tickets for {args.month}.")
        print(f"Found {len(attachment_map)} unique attachments.")
        
        # Show shared attachments
        shared_attachments = [(asset.name, len(tickets)) for asset, tickets in attachment_map.values() if len(tickets) > 1]
        if shared_attachments:
            print("\nShared attachments:")
            for name, count in shared_attachments:
                print(f"  {count}x: {name}")
        
        for t in ticket_rows:
            print(f"Ticket: {t.item_name}, Attachments: {[a.name for a in t.attachments]}")
        
        # Double-check verification
        verify_all_tickets_found(client, config, args.month, ticket_rows)
        return

    # Download and convert unique attachments only
    item_pdfs = []
    unique_attachments = [asset for asset, tickets in attachment_map.values()]
    
    print(f"Processing {len(unique_attachments)} unique attachments (instead of {sum(len(t.attachments) for t in ticket_rows)} total)...")
    
    for asset in unique_attachments:
        # Use first ticket that has this attachment for download directory
        first_ticket = [t for t in ticket_rows if any(a.name == asset.name for a in t.attachments)][0]
        item_dir = os.path.join(args.downloads, args.month, first_ticket.item_id)
        os.makedirs(item_dir, exist_ok=True)
        converted_dir = os.path.join(item_dir, "converted")
        os.makedirs(converted_dir, exist_ok=True)
        
        file_path = download_asset(asset, item_dir)
        if file_path:  # Only process if download succeeded
            ext = asset.file_extension
            pdf_path = to_pdf(file_path, ext, converted_dir)
            if pdf_path and os.path.exists(pdf_path):  # Only add if conversion succeeded
                item_pdfs.append(pdf_path)
                log_event(item_id=first_ticket.item_id, asset_id=asset.id, action="pdf_added", status="success")
            else:
                log_event(item_id=first_ticket.item_id, asset_id=asset.id, action="pdf_add_failed", status="fail")

    # Summary page
    summary_pdf = os.path.join(args.out, f"{args.month}-summary.pdf")
    generate_summary_page(ticket_rows, args.month, config['board']['id'], summary_pdf, attachment_map)

    # Debug: Print what PDFs we're merging
    print(f"Summary PDF: {summary_pdf}")
    print(f"Item PDFs to merge: {item_pdfs}")
    for pdf in item_pdfs:
        print(f"  - {pdf} (exists: {os.path.exists(pdf) if pdf else False})")

    # Merge
    output_pdf = os.path.join(args.out, f"{args.month}-Resolved-Tickets.pdf")
    merge_pdfs(summary_pdf, item_pdfs, output_pdf)
    print(f"PDF generated: {output_pdf}")
    
    # Generate Excel summary
    if not args.dry_run:
        # Prepare ticket data for Excel
        excel_tickets = []
        shared_attachments = {}
        
        for ticket in ticket_rows:
            # Build attachment list for this ticket
            ticket_attachments = [asset.name for asset in ticket.attachments]
            
            excel_tickets.append({
                'name': ticket.item_name,
                'id': ticket.item_id,
                'date_opened': ticket.open_date,
                'status': 'Resolved',  # We know all these are resolved
                'summary': '',  # Could extract from item data if needed
                'attachments': ticket_attachments
            })
        
        # Build shared attachments mapping for Excel
        for filename, (asset, ticket_names) in attachment_map.items():
            if len(ticket_names) > 1:  # Only shared files
                shared_attachments[filename] = ticket_names
        
        # Generate Excel file
        try:
            excel_path = create_monthly_excel_summary(
                tickets=excel_tickets,
                month=args.month,
                output_dir=args.out,
                shared_attachments=shared_attachments,
                unique_attachments_count=len(unique_attachments)
            )
            print(f"Excel summary generated: {excel_path}")
        except Exception as e:
            print(f"Warning: Could not generate Excel summary: {e}")
    
    # Verify all tickets found
    verify_all_tickets_found(client, config, args.month, filtered_items)


if __name__ == "__main__":
    main()
