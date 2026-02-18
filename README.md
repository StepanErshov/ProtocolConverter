# Convert2DBC - Automotive Protocol Converter Tool

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://cmdtool.streamlit.app/)

A comprehensive web-based tool for converting between automotive communication protocols (CAN DBC, Excel, LIN LDF) with validation and analysis capabilities.

## 🚗 Overview

Convert2DBC is a powerful Streamlit application designed for automotive engineers and developers working with vehicle communication protocols. It provides bidirectional conversion between:

- **CAN DBC ↔ Excel** - Convert CAN database files to/from Excel format
- **Excel → LIN LDF** - Convert Excel files to LIN Description Format
- **Protocol Validation** - Validate CAN, LIN, and Ethernet communication data

## ✨ Features

### 🔄 Bidirectional Conversions
- **DBC to Excel**: Convert CAN database files to structured Excel format
- **Excel to DBC**: Convert Excel matrices to CAN database format
- **Excel to LDF**: Convert Excel files to LIN Description Format
- **Real-time Preview**: Preview data before conversion
- **Batch Processing**: Handle multiple files efficiently

### 📊 Data Validation & Analysis
- **CAN Validation**: Validate CAN message and signal parameters
- **LIN Validation**: Check LIN frame and signal configurations
- **Ethernet Validation**: Validate Ethernet communication data
- **Error Reporting**: Detailed error and warning messages
- **Data Quality Checks**: Ensure compliance with automotive standards

### 🎨 User Interface
- **Modern Web Interface**: Clean, responsive Streamlit-based UI
- **File Upload/Download**: Easy file management with drag-and-drop
- **Progress Tracking**: Real-time conversion progress indicators
- **History Management**: Track conversion history and results
- **Customizable Output**: Configure output file names and versions

### 🔧 Technical Capabilities
- **CAN FD Support**: Full support for CAN FD protocol
- **Multiple ECU Support**: Handle complex multi-ECU configurations
- **Signal Mapping**: Automatic signal-to-ECU mapping
- **Format Preservation**: Maintain Excel formatting and styles
- **Cross-Platform**: Works on Windows, Linux, and macOS

## 🛠️ Installation

### Prerequisites
- Python 3.9 or higher
- Git

### Local Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/Convert2DBC.git
   cd Convert2DBC
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**:
   ```bash
   streamlit run main.py
   ```

### Cloud Deployment
The application is also available as a hosted service:
- **Streamlit Cloud**: [https://cmdtool.streamlit.app/](https://cmdtool.streamlit.app/)

## 📖 Usage Guide

### 1. DBC to Excel Conversion

**Purpose**: Convert CAN database files to Excel format for analysis and documentation.

**Steps**:
1. Navigate to the "DBC to Excel" page
2. Upload your `.dbc` file
3. Configure output settings (version, filename)
4. Preview the data structure
5. Convert and download the Excel file

**Features**:
- Automatic ECU column detection
- Message and signal organization
- Preserved formatting from template
- Version control integration

### 2. Excel to DBC Conversion

**Purpose**: Convert Excel matrices to CAN database format for ECU development.

**Steps**:
1. Navigate to the "Excel to DBC" page
2. Upload your Excel file
3. Validate the data structure
4. Configure conversion settings
5. Generate and download the DBC file

**Required Excel Columns**:
```
Message ID, Message Name, Signal Name, Start Byte, Start Bit, Bit Length, 
Factor, Offset, Initial Value(Hex), Invalid Value(Hex), Min Value, Max Value, 
Unit, Receiver, Byte Order, Data Type, Msg Cycle Time, Msg Send Type, 
Description, Msg Length, Signal Value Description, Senders
```

### 3. Excel to LIN LDF Conversion

**Purpose**: Convert Excel files to LIN Description Format for LIN network configuration.

**Steps**:
1. Navigate to the "Excel to LDF" page
2. Upload your Excel file
3. Configure LIN-specific settings
4. Generate and download the LDF file

### 4. Protocol Validation

**Purpose**: Validate communication data for compliance and correctness.

**Available Validators**:
- **CAN Validator**: Check CAN message and signal parameters
- **LIN Validator**: Validate LIN frame configurations
- **Ethernet Validator**: Verify Ethernet communication data

## 📋 Supported Formats

### Input Formats
- **DBC Files**: CAN database files (.dbc)
- **Excel Files**: Microsoft Excel (.xlsx, .xls)
- **LDF Files**: LIN Description Format (.ldf)

### Output Formats
- **Excel Files**: Structured Excel format with formatting
- **DBC Files**: Standard CAN database format
- **LDF Files**: LIN Description Format
- **Validation Reports**: Detailed error and warning reports

## 🔍 Data Validation Rules

### CAN Validation
- Message ID format and range checking
- Signal bit allocation validation
- Byte order verification (Intel/Motorola)
- Cycle time and length constraints
- ECU sender/receiver validation

### LIN Validation
- Frame ID range checking
- Signal mapping validation
- Schedule table verification
- Master/slave configuration checks

### Excel Format Validation
- Required column presence
- Data type consistency
- Value range validation
- Format compliance checking

## 🏗️ Project Structure

```
Convert2DBC/
├── main.py                 # Main Streamlit application
├── pages/                  # Streamlit page modules
│   ├── DBC_2_Xlsx.py      # DBC to Excel converter
│   ├── Xlsx_2_DBC.py      # Excel to DBC converter
│   ├── Xls_2_LDF.py       # Excel to LDF converter
│   ├── CANValidator.py    # CAN protocol validator
│   ├── LINValidator.py    # LIN protocol validator
│   └── ETHValidator.py    # Ethernet validator
├── dbc2xlsx.py            # DBC to Excel conversion logic
├── xlsx2dbc.py            # Excel to DBC conversion logic
├── xlsx2ldf.py            # Excel to LDF conversion logic
├── requirements.txt       # Python dependencies
├── test.xlsx             # Template file for formatting
└── README.md             # This file
```

## 🧪 Testing

### Local Testing
```bash
# Run the application locally
streamlit run main.py

# Test specific conversions
python dbc2xlsx.py
python xlsx2dbc.py --input test.xlsx --output test.dbc
```

### Sample Files
The project includes sample files for testing:
- `test.xlsx` - Template Excel file
- `output.xlsx` - Sample output file

## 🔧 Configuration

### Environment Variables
- `STREAMLIT_SERVER_PORT`: Custom port for local development
- `STREAMLIT_SERVER_ADDRESS`: Custom server address

### Database Configuration
The application uses SQLite for storing conversion history:
- Database file: `can_database.db`
- Tables: `dbc_converted_files`, `xlsx_converted_files`, `ldf_converted_files`

## 🚀 Deployment

### Streamlit Cloud
1. Connect your GitHub repository to Streamlit Cloud
2. Set the main file path to `main.py`
3. Deploy automatically on push

### Docker Deployment
```dockerfile
FROM python:3.9-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
EXPOSE 8501
CMD ["streamlit", "run", "main.py"]
```

## 🤝 Contributing

We welcome contributions! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Setup
```bash
# Install development dependencies
pip install -r requirements.txt
pip install black flake8 pytest

# Run code formatting
black .

# Run linting
flake8 .

# Run tests
pytest
```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **cantools**: CAN database parsing and generation
- **pandas**: Data manipulation and Excel handling
- **openpyxl**: Excel file processing
- **streamlit**: Web application framework
- **ldfparser**: LIN Description Format parsing

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/StepanErshov/Convert2DBC/issues)
- **Documentation**: [Wiki](https://github.com/StepanErshov/Convert2DBC/wiki)
- **Email**: SFS_stepan@mail.ru

## 🔄 Version History

- **v1.2.0** (Current): Added DBC to Excel conversion, improved validation
- **v1.1.0**: Added LIN LDF support, enhanced UI
- **v1.0.0**: Initial release with Excel to DBC conversion

---

**Made with ❤️ for the automotive industry**
