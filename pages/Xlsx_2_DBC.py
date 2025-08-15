import streamlit as st
import pandas as pd
from xlsx2dbc import ExcelToDBCConverter
import os
from datetime import datetime
import re
from sqlalchemy import text

conn = st.connection(
    "can_db", type="sql", dialect="sqlite", database="/tmp/can_database.db"
)

with conn.session as s:
    s.execute(
        text(
            """
        CREATE TABLE IF NOT EXISTS converted_files (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            original_filename TEXT NOT NULL,
            dbc_filename TEXT NOT NULL,
            version TEXT NOT NULL,
            conversion_date TEXT NOT NULL,
            file_size INTEGER NOT NULL,
            user_id TEXT NOT NULL
        )
    """
        )
    )
    s.commit()

# st.set_page_config(
#     page_title="Excel to DBC Converter",
#     page_icon=":car:",
#     layout="wide"
# )

st.markdown(
    """
    <style>
    .main {
        background-color: #f5f5f5;
    }
    .stButton>button {
        background-color: #4CAF50;
        color: white;
        border-radius: 5px;
        padding: 10px 24px;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .stFileUploader>div>div>div>button {
        background-color: #2196F3;
        color: white;
    }
    .stTextInput>div>div>input {
        border-radius: 5px;
    }
    .title {
        color: #2c3e50;
    }
    .error-box {
        background-color: #ffebee;
        border-left: 5px solid #f44336;
        padding: 10px;
        margin: 10px 0;
        border-radius: 5px;
    }
    .warning-box {
        background-color: #fff8e1;
        border-left: 5px solid #ffc107;
        padding: 10px;
        margin: 10px 0;
        border-radius: 5px;
    }
    .success-box {
        background-color: #e8f5e9;
        border-left: 5px solid #4caf50;
        padding: 10px;
        margin: 10px 0;
        border-radius: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def extract_version_date(filename):
    pattern = r"_V(\d+\.\d+\.\d+)_(\d{8})\."
    match = re.search(pattern, filename)
    if match:
        return match.group(1), match.group(2)
    return None, None


def generate_base_name(input_filename):
    base_name = os.path.splitext(input_filename)[0]
    base_name = re.sub(r"_V\d+\.\d+\.\d+_\d{8}$", "", base_name)
    return base_name


def generate_default_output_filename(input_filename, new_version=None):
    base_name = generate_base_name(input_filename)
    current_date = datetime.now().strftime("%Y%m%d")

    if new_version is None:
        version, _ = extract_version_date(input_filename)
        new_version = version if version else "1.0.0"

    return f"{base_name}_V{new_version}_{current_date}.dbc"


def display_errors(errors):
    if errors:
        with st.expander("⚠️ Validation Errors", expanded=True):
            for error in errors:
                st.markdown(
                    f'<div class="error-box">{error}</div>', unsafe_allow_html=True
                )


def display_warnings(warnings):
    if warnings:
        with st.expander("⚠️ Validation Warnings", expanded=False):
            for warning in warnings:
                st.markdown(
                    f'<div class="warning-box">{warning}</div>', unsafe_allow_html=True
                )


def validate_input_data(uploaded_file):
    errors = []
    warnings = []

    try:
        df = pd.read_excel(uploaded_file, sheet_name="Matrix")

        required_columns = [
            "Msg ID\n报文标识符",
            "Msg Name\n报文名称",
            "Signal Name\n信号名称",
            "Start Byte\n起始字节",
            "Start Bit\n起始位",
            "Bit Length (Bit)\n信号长度",
            "Byte Order\n排列格式\n(Intel/Motorola)",
            "Data Type\n数据类型",
            "Msg Length (Byte)\n报文长度",
        ]

        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            errors.append(f"Missing required columns: {', '.join(missing_columns)}")

        for msg_id in df["Msg ID\n报文标识符"].dropna().unique():
            try:
                if isinstance(msg_id, str):
                    if msg_id.startswith("0x"):
                        int(msg_id, 16)
                    else:
                        int(msg_id)
                else:
                    int(msg_id)
            except ValueError:
                errors.append(f"Invalid message ID: {msg_id}")

        for _, group in df.groupby("Msg ID\n报文标识符"):
            msg_length = group["Msg Length (Byte)\n报文长度"].iloc[0]
            for _, row in group.iterrows():
                start_byte = row["Start Byte\n起始字节"]
                start_bit = row["Start Bit\n起始位"]
                bit_length = row["Bit Length (Bit)\n信号长度"]

                if start_byte >= msg_length:
                    errors.append(
                        f"Signal '{row['Signal Name\n信号名称']}' is outside message bounds "
                        f"(byte {start_byte} >= message length {msg_length})"
                    )

                if start_bit >= 8:
                    errors.append(
                        f"Invalid start bit {start_bit} in signal '{row['Signal Name\n信号名称']}'"
                    )

                if bit_length <= 0:
                    errors.append(
                        f"Invalid length {bit_length} in signal '{row['Signal Name\n信号名称']}'"
                    )

                if start_byte * 8 + start_bit + bit_length > msg_length * 8:
                    errors.append(
                        f"Signal '{row['Signal Name\n信号名称']}' exceeds message bounds"
                    )

        for _, row in df.iterrows():
            data_type = str(row["Data Type\n数据类型"])
            bit_length = row["Bit Length (Bit)\n信号长度"]

            if "Float" in data_type and bit_length not in [32, 64]:
                errors.append(
                    f"Invalid length {bit_length} for float type in signal '{row['Signal Name\n信号名称']}'"
                )

            if "Signed" in data_type and bit_length < 2:
                errors.append(
                    f"Invalid length {bit_length} for signed type in signal '{row['Signal Name\n信号名称']}'"
                )

        bus_users = [
            col
            for col in df.columns
            if any(val in ["S", "R"] for val in df[col].dropna().unique())
            and col != "Unit\n单位"
        ]

        for msg_id, group in df.groupby("Msg ID\n报文标识符"):
            senders = []
            receivers = []

            for bus_user in bus_users:
                if bus_user in group.columns:
                    if "S" in group[bus_user].values:
                        senders.append(bus_user)
                    if "R" in group[bus_user].values:
                        receivers.append(bus_user)

            if not senders:
                warnings.append(f"Message {msg_id} has no senders")

            if not receivers:
                warnings.append(f"Message {msg_id} has no receivers")

        for _, row in df.iterrows():
            if pd.notna(row["Initial Value (Hex)\n初始值"]):
                try:
                    int(row["Initial Value (Hex)\n初始值"], 16)
                except ValueError:
                    errors.append(
                        f"Invalid initial value {row['Initial Value (Hex)\n初始值']} "
                        f"for signal '{row['Signal Name\n信号名称']}'"
                    )

    except Exception as e:
        errors.append(f"Error reading the Excel file: {str(e)}")

    return errors, warnings


def main():
    st.markdown(
        '<h1 class="title">📊 Excel to DBC Converter</h1>', unsafe_allow_html=True
    )
    st.markdown(
        "Upload your Excel file containing CAN data to convert it to a DBC file."
    )

    col1, col2 = st.columns([3, 1])

    with col1:
        uploaded_file = st.file_uploader(
            "Choose an Excel file", type=["xlsx"], key="file_uploader"
        )

        if uploaded_file is not None:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Matrix")
                st.subheader("Data Preview")
                st.dataframe(
                    df.head().style.set_properties(
                        **{
                            "background-color": "#f0f2f6",
                            "color": "#2c3e50",
                            "border": "1px solid #dfe6e9",
                        }
                    )
                )

                errors, warnings = validate_input_data(uploaded_file)
                display_errors(errors)
                display_warnings(warnings)

                if errors:
                    st.error(
                        "Cannot convert due to validation errors. Please fix the issues in your Excel file."
                    )
                    return

            except Exception as e:
                st.error(f"Error reading the Excel file: {str(e)}")
                return

    with col2:
        if uploaded_file is not None:
            st.subheader("Output Settings")

            version, _ = extract_version_date(uploaded_file.name)
            default_version = version if version else "1.0.0"

            new_version = st.text_input(
                "DBC Version",
                value=default_version,
                help="Enter the version number in format X.X.X",
            )

            base_name = generate_base_name(uploaded_file.name)
            default_output_name = generate_default_output_filename(
                uploaded_file.name, new_version
            )

            custom_filename = st.text_input(
                "Output DBC file name",
                value=default_output_name,
                help="You can customize the output file name",
            )

            if not custom_filename.lower().endswith(".dbc"):
                custom_filename += ".dbc"

            st.markdown("**Final DBC file name:**")
            st.code(custom_filename)

            if st.button("Convert to DBC", key="convert_button"):
                with st.spinner("Converting... Please wait"):
                    try:
                        converter = ExcelToDBCConverter(uploaded_file)
                        success = converter.convert(custom_filename)

                        if success:
                            st.markdown(
                                f'<div class="success-box">Conversion completed successfully!</div>',
                                unsafe_allow_html=True,
                            )

                            file_size = os.path.getsize(custom_filename)
                            current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                            with conn.session as s:
                                s.execute(
                                    text(
                                        """
                                        INSERT INTO converted_files (
                                            original_filename, 
                                            dbc_filename, 
                                            version, 
                                            conversion_date, 
                                            file_size,
                                            user_id
                                        ) VALUES (:original, :dbc, :version, :date, :size, :user)
                                    """
                                    ),
                                    {
                                        "original": uploaded_file.name,
                                        "dbc": custom_filename,
                                        "version": new_version,
                                        "date": current_date,
                                        "size": file_size,
                                        "user": st.session_state.get(
                                            "keycloak", {}
                                        ).get("username", "unknown"),
                                    },
                                )
                                s.commit()

                            with open(custom_filename, "rb") as f:
                                bytes_data = f.read()
                                st.download_button(
                                    label="Download DBC File",
                                    data=bytes_data,
                                    file_name=custom_filename,
                                    mime="application/octet-stream",
                                    key="download_button",
                                )

                        st.subheader("Conversion History")
                        with conn.session as s:
                            result = s.execute(
                                text(
                                    """
                                SELECT original_filename, dbc_filename, version, conversion_date, file_size
                                FROM converted_files
                                ORDER BY conversion_date DESC
                                LIMIT 10
                            """
                                )
                            )
                            history_df = pd.DataFrame(
                                result.fetchall(), columns=result.keys()
                            )

                        if not history_df.empty:
                            history_df["file_size"] = history_df["file_size"].apply(
                                lambda x: (
                                    f"{x/1024:.2f} KB"
                                    if x < 1024 * 1024
                                    else f"{x/(1024*1024):.2f} MB"
                                )
                            )
                            st.dataframe(history_df)

                    except Exception as e:
                        st.error(f"An error occurred: {str(e)}")
                        st.error(f"Full error: {repr(e)}")


conn.session.close()

if __name__ == "__main__":
    main()
