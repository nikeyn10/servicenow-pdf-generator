# ServiceNow PDF Generator

A comprehensive Python-based tool for generating monthly PDF reports and Excel summaries from ServiceNow tickets via Monday.com integration. This tool automates the process of collecting, processing, and organizing ServiceNow ticket data with attachments into professional PDF reports and detailed Excel spreadsheets.

## 🚀 Features

- **📄 PDF Generation**: Automated monthly PDF reports with ticket summaries and attachments
- **📊 Excel Reports**: Comprehensive Excel files with multiple worksheets including:
  - Summary statistics and metrics
  - Detailed ticket information
  - Shared attachment analysis
- **🔄 Attachment Processing**: Intelligent deduplication and conversion of ticket attachments
- **🗂️ Project Management**: Integration with Monday.com for ticket data retrieval
- **🧹 Maintenance Tools**: Automated cleanup scripts and project organization
- **⚙️ Flexible Configuration**: YAML-based configuration for easy customization

## 📁 Project Structure

```
servicenow-pdf-generator/
├── src/                    # Source code modules
│   ├── main.py            # Main application entry point
│   ├── monday_client.py   # Monday.com API integration
│   ├── excel_utils.py     # Excel generation utilities
│   ├── pdf_utils.py       # PDF processing utilities
│   ├── convert.py         # File conversion utilities
│   ├── files.py           # File management utilities
│   ├── filters.py         # Data filtering utilities
│   ├── models.py          # Data models
│   ├── queries.py         # GraphQL queries
│   └── log.py             # Logging utilities
├── config/                # Configuration files
│   └── config.yml         # Main configuration
├── scripts/               # Utility scripts
│   └── cleanup.py         # Project maintenance script
├── archive/               # Archived files
├── data/                  # Data directory (downloads)
├── output/                # Generated reports
│   └── merged/            # Final PDF and Excel files
├── requirements.txt       # Python dependencies
├── .gitignore            # Git ignore rules
└── README.md             # This file
```

## 🛠️ Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/nikeyn10/servicenow-pdf-generator.git
   cd servicenow-pdf-generator
   ```

2. **Create and activate virtual environment:**
   ```bash
   python3 -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

4. **Configure environment variables:**
   Create a `.env` file with your Monday.com API credentials:
   ```env
   MONDAY_API_TOKEN=your_monday_api_token_here
   ```

5. **Configure the application:**
   Edit `config/config.yml` with your specific board IDs and column mappings.

## 🚀 Usage

### Generate Monthly Reports

Generate PDF and Excel reports for a specific month:

```bash
python3 -m src.main --month "2025-07" --config config/config.yml --out output/merged --downloads data/downloads
```

### Command Line Options

- `--month`: Target month in YYYY-MM format (e.g., "2025-07")
- `--config`: Path to configuration file
- `--out`: Output directory for generated reports
- `--downloads`: Directory containing downloaded attachments
- `--dry-run`: Preview processing without generating files

### Example Commands

```bash
# Generate July 2025 reports
python3 -m src.main --month "2025-07" --config config/config.yml --out output/merged --downloads data/downloads

# Dry run to preview June 2025 data
python3 -m src.main --month "2025-06" --config config/config.yml --out output/merged --downloads data/downloads --dry-run

# Run project cleanup
python3 scripts/cleanup.py
```

## 📊 Output Files

The tool generates two main types of output files:

### PDF Reports
- `YYYY-MM-Resolved-Tickets.pdf` - Complete PDF with all ticket attachments
- `YYYY-MM-summary.pdf` - Summary page with statistics

### Excel Reports
- `YYYY-MM-Resolved-Tickets-Summary.xlsx` - Multi-sheet Excel workbook containing:
  - **Summary Sheet**: Key metrics and statistics
  - **Tickets Detail**: Complete ticket information with attachment lists
  - **Shared Attachments**: Analysis of attachments used by multiple tickets

## ⚙️ Configuration

The main configuration is stored in `config/config.yml`. Key sections include:

- **Board Configuration**: Monday.com board ID and column mappings
- **Status Filters**: Ticket status requirements
- **Output Settings**: File naming and organization preferences

## 🧹 Maintenance

The project includes automated maintenance tools:

```bash
# Run cleanup script to organize files and remove temporary data
python3 scripts/cleanup.py
```

The cleanup script:
- Removes temporary and system files
- Archives old CSV reports
- Organizes project structure
- Reports storage usage

## 🔧 Development

### Prerequisites
- Python 3.11+
- Homebrew (macOS):
  - `brew install python libreoffice imagemagick ghostscript poppler wkhtmltopdf`

### Project Dependencies

- **openpyxl**: Excel file generation
- **PyPDF2**: PDF processing
- **requests**: API communication
- **python-dotenv**: Environment variable management
- **PyYAML**: Configuration file parsing

### Adding New Features

1. Create new modules in the `src/` directory
2. Update configuration in `config/config.yml` as needed
3. Add new dependencies to `requirements.txt`
4. Update this README with usage instructions

## 📈 Recent Enhancements

- ✅ **Project Housekeeping**: Automated cleanup and organization
- ✅ **Excel Integration**: Comprehensive Excel summary generation
- ✅ **Multi-sheet Workbooks**: Detailed analysis across multiple worksheets
- ✅ **Attachment Deduplication**: Smart handling of shared attachments
- ✅ **Maintenance Automation**: Scripts for ongoing project hygiene

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🆘 Support

For support and questions:
1. Check the existing documentation
2. Create an issue in the GitHub repository
3. Review the configuration files for setup guidance

## 📊 Statistics

- **Languages**: Python
- **Framework**: Monday.com GraphQL API
- **Output Formats**: PDF, Excel (XLSX)
- **Maintenance**: Automated cleanup and organization tools
