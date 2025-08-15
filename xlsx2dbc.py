import cantools
from cantools.database.can.formats.arxml.message_specifics import (
    AutosarMessageSpecifics,
)
from cantools.database.can.formats.arxml.node_specifics import AutosarNodeSpecifics
import cantools.database
import cantools.database.conversion
import pandas as pd
from cantools.database.can.formats.dbc import DbcSpecifics
from cantools.database.can.attribute import Attribute
from cantools.database.can.attribute_definition import AttributeDefinition
from cantools.database.can import Node
import re
import os
import argparse
from typing import Optional, Dict


class ValueDescriptionParser:
    @staticmethod
    def parse(desc_str: str) -> Optional[Dict[int, str]]:
        """Convert multi-line hex descriptions to single-line decimal format"""
        if not isinstance(desc_str, str) or not desc_str.strip():
            return None

        descriptions = {}
        try:
            desc_str = " ".join(desc_str.replace("\r", "\n").split())
            parts = re.split(r"(0x[0-9a-fA-F]+)\s*:\s*", desc_str)
            if len(parts) > 1:
                for i in range(1, len(parts), 2):
                    hex_val = parts[i]
                    text = parts[i + 1].split(";")[0].split("~")[0].strip()
                    text = re.sub(r"[^a-zA-Z0-9_\- ]", "", text)
                    if hex_val and text:
                        try:
                            dec_val = int(hex_val, 16)
                            descriptions[dec_val] = text
                        except ValueError:
                            continue
            else:
                for item in desc_str.split(";"):
                    item = item.strip()
                    if ":" in item:
                        val_part, text = item.split(":", 1)
                        val_part = val_part.strip()
                        text = text.strip()
                        if val_part.startswith("0x"):
                            try:
                                dec_val = int(val_part, 16)
                                descriptions[dec_val] = text
                            except ValueError:
                                continue

            range_matches = re.finditer(
                r"(0x[0-9a-fA-F]+)\s*~\s*(0x[0-9a-fA-F]+)\s*:\s*([^;]+)", desc_str
            )
            for match in range_matches:
                start = int(match.group(1), 16)
                end = int(match.group(2), 16)
                text = match.group(3).strip()
                for val in range(start, end + 1):
                    descriptions[val] = text

            return dict(sorted(descriptions.items())) if descriptions else None

        except Exception as e:
            print(f"Error parsing value descriptions '{desc_str}': {str(e)}")
            return None


class ExcelToDBCConverter:

    def __init__(self, excel_path: str):
        self.excel_path = excel_path
        self.diag_messages = []  # For diagnostic messages (0x7...)
        self.nm_messages = []  # For network management messages (0x5...)
        self.normal_messages = []  # For normal messages

        self.attr_def_dbname = AttributeDefinition(
            name="DBName", default_value="", type_name="STRING"
        )

        self.attr_def_bus_type = AttributeDefinition(
            name="BusType", default_value="CAN", type_name="STRING"
        )

        self.db = cantools.database.can.Database(
            version=ExcelToDBCConverter.get_file_info(excel_path.name)["version"],
            sort_signals=None,
            strict=False,
        )

        self.db.dbc = DbcSpecifics(
            attributes={
                "DBName": Attribute(
                    value=str(self.excel_path.name).split(".xlsx")[0],
                    definition=self.attr_def_dbname,
                ),
                "BusType": Attribute(
                    value=ExcelToDBCConverter.get_file_info(excel_path.name)[
                        "protocol"
                    ],
                    definition=self.attr_def_bus_type,
                ),
            }
        )

        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine="openpyxl",
        )

        self.bus_users = [
            col
            for col in df.columns
            if any(val in ["S", "R"] for val in df[col].dropna().unique())
            and col != "Unit\n单位"
        ]

        self._initialize_nodes()
        self._initialize_attr()

    def _initialize_nodes(self):
        self.db.nodes.extend([Node(name=bus_name) for bus_name in self.bus_users])

    def _initialize_attr(self):
        self.attr_def_manufacturer = AttributeDefinition(
            name="Manufacturer", default_value="", type_name="STRING"
        )
        self.attr_def_nm_type = AttributeDefinition(
            name="NmType", default_value="", type_name="STRING"
        )
        self.attr_def_nm_base_addr = AttributeDefinition(
            name="NmBaseAddress",
            default_value=1280,
            type_name="HEX",
            minimum=1280,
            maximum=1407,
        )
        self.attr_def_nm_msg_cnt = AttributeDefinition(
            name="NmMessageCount",
            default_value=128,
            type_name="INT",
            minimum=0,
            maximum=255,
        )

        self.attr_def_node_layer_modules = AttributeDefinition(
            name="NodeLayerModules",
            kind="BU_",
            default_value="CANoeILNLVector.dll",
            type_name="STRING",
        )
        self.attr_def_il_used = AttributeDefinition(
            name="ILUsed",
            kind="BU_",
            default_value="No",
            type_name="ENUM",
            choices=["No", "Yes"],
        )
        self.attr_def_diag_station_addr = AttributeDefinition(
            name="DiagStationAddress",
            kind="BU_",
            default_value=0,
            type_name="HEX",
            minimum=0,
            maximum=255,
        )
        self.attr_def_nm_node = AttributeDefinition(
            name="NmNode",
            kind="BU_",
            default_value="Not",
            type_name="ENUM",
            choices=["Not", "Yes"],
        )
        self.attr_def_nm_station_addr = AttributeDefinition(
            name="NmStationAddress",
            kind="BU_",
            default_value=0,
            type_name="HEX",
            minimum=0,
            maximum=65535,
        )
        self.attr_def_nm_can = AttributeDefinition(
            name="NmCAN",
            kind="BU_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=2,
        )

        self.attr_def_msg_send_type = AttributeDefinition(
            name="GenMsgSendType",
            kind="BO_",
            default_value="Cyclic",
            type_name="ENUM",
            choices=["Cyclic", "Event", "IfActive", "CE", "CA", "NoMsgSendType"],
        )
        self.attr_def_msg_cycle_time = AttributeDefinition(
            name="GenMsgCycleTime",
            kind="BO_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_msg_cycle_time_fast = AttributeDefinition(
            name="GenMsgCycleTimeFast",
            kind="BO_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_msg_nr_repetition = AttributeDefinition(
            name="GenMsgNrOfRepetition",
            kind="BO_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_msg_delay_time = AttributeDefinition(
            name="GenMsgDelayTime",
            kind="BO_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_msg_cycle_time_active = AttributeDefinition(
            name="GenMsgCycleTimeActive",
            kind="BO_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_msg_il_support = AttributeDefinition(
            name="GenMsgILSupport",
            kind="BO_",
            default_value="No",
            type_name="ENUM",
            choices=["No", "Yes"],
        )
        self.attr_def_msg_start_delay = AttributeDefinition(
            name="GenMsgStartDelayTime",
            kind="BO_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=65535,
        )
        self.attr_def_nm_message = AttributeDefinition(
            name="NmMessage",
            kind="BO_",
            default_value="No",
            type_name="ENUM",
            choices=["No", "Yes"],
        )
        self.attr_def_diag_state = AttributeDefinition(
            name="DiagState",
            kind="BO_",
            default_value="No",
            type_name="ENUM",
            choices=["No", "Yes"],
        )
        self.attr_def_diag_request = AttributeDefinition(
            name="DiagRequest",
            kind="BO_",
            default_value="No",
            type_name="ENUM",
            choices=["No", "Yes"],
        )
        self.attr_def_diag_response = AttributeDefinition(
            name="DiagResponse",
            kind="BO_",
            default_value="No",
            type_name="ENUM",
            choices=["No", "Yes"],
        )

        self.attr_def_sig_send_type = AttributeDefinition(
            name="GenSigSendType",
            kind="SG_",
            default_value="Cyclic",
            type_name="ENUM",
            choices=[
                "Cyclic",
                "OnChange",
                "OnWrite",
                "IfActive",
                "OnChangeWithRepetition",
                "OnWriteWithRepetition",
                "IfActiveWithRepetition",
                "NoSigSendType",
                "OnChangeAndIfActive",
                "OnChangeAndIfActiveWithRepetition",
                "CA",
                "CE",
                "Event",
            ],
        )
        self.attr_def_sig_start_value = AttributeDefinition(
            name="GenSigStartValue",
            kind="SG_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_sig_inactive_value = AttributeDefinition(
            name="GenSigInactiveValue",
            kind="SG_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_sig_invalid_value = AttributeDefinition(
            name="GenSigInvalidValue",
            kind="SG_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=0,
        )
        self.attr_def_sig_timeout_value = AttributeDefinition(
            name="GenSigTimeoutValue",
            kind="SG_",
            default_value=0,
            type_name="INT",
            minimum=0,
            maximum=1000000000,
        )

        self.db.dbc.attribute_definitions["DBName"] = self.attr_def_dbname
        self.db.dbc.attribute_definitions["BusType"] = self.attr_def_bus_type
        self.db.dbc.attribute_definitions["Manufacturer"] = self.attr_def_manufacturer
        self.db.dbc.attribute_definitions["NmType"] = self.attr_def_nm_type
        self.db.dbc.attribute_definitions["NmBaseAddress"] = self.attr_def_nm_base_addr
        self.db.dbc.attribute_definitions["NmMessageCount"] = self.attr_def_nm_msg_cnt

        self.db.dbc.attribute_definitions["NodeLayerModules"] = (
            self.attr_def_node_layer_modules
        )
        self.db.dbc.attribute_definitions["ILUsed"] = self.attr_def_il_used
        self.db.dbc.attribute_definitions["DiagStationAddress"] = (
            self.attr_def_diag_station_addr
        )
        self.db.dbc.attribute_definitions["NmNode"] = self.attr_def_nm_node
        self.db.dbc.attribute_definitions["NmStationAddress"] = (
            self.attr_def_nm_station_addr
        )
        self.db.dbc.attribute_definitions["NmCAN"] = self.attr_def_nm_can

        self.db.dbc.attribute_definitions["GenMsgSendType"] = (
            self.attr_def_msg_send_type
        )
        self.db.dbc.attribute_definitions["GenMsgCycleTime"] = (
            self.attr_def_msg_cycle_time
        )
        self.db.dbc.attribute_definitions["GenMsgCycleTimeFast"] = (
            self.attr_def_msg_cycle_time_fast
        )
        self.db.dbc.attribute_definitions["GenMsgNrOfRepetition"] = (
            self.attr_def_msg_nr_repetition
        )
        self.db.dbc.attribute_definitions["GenMsgDelayTime"] = (
            self.attr_def_msg_delay_time
        )
        self.db.dbc.attribute_definitions["GenMsgCycleTimeActive"] = (
            self.attr_def_msg_cycle_time_active
        )
        self.db.dbc.attribute_definitions["GenMsgILSupport"] = (
            self.attr_def_msg_il_support
        )
        self.db.dbc.attribute_definitions["GenMsgStartDelayTime"] = (
            self.attr_def_msg_start_delay
        )
        self.db.dbc.attribute_definitions["NmMessage"] = self.attr_def_nm_message
        self.db.dbc.attribute_definitions["DiagState"] = self.attr_def_diag_state
        self.db.dbc.attribute_definitions["DiagRequest"] = self.attr_def_diag_request
        self.db.dbc.attribute_definitions["DiagResponse"] = self.attr_def_diag_response

        self.db.dbc.attribute_definitions["GenSigSendType"] = (
            self.attr_def_sig_send_type
        )
        self.db.dbc.attribute_definitions["GenSigStartValue"] = (
            self.attr_def_sig_start_value
        )
        self.db.dbc.attribute_definitions["GenSigInactiveValue"] = (
            self.attr_def_sig_inactive_value
        )
        self.db.dbc.attribute_definitions["GenSigInvalidValue"] = (
            self.attr_def_sig_invalid_value
        )
        self.db.dbc.attribute_definitions["GenSigTimeoutValue"] = (
            self.attr_def_sig_timeout_value
        )

    def _load_excel_data(self) -> pd.DataFrame:
        df = pd.read_excel(
            self.excel_path,
            sheet_name="Matrix",
            keep_default_na=True,
            engine="openpyxl",
        )

        df_history = pd.read_excel(
            self.excel_path,
            sheet_name="History",
            keep_default_na=True,
            engine="openpyxl",
        )

        all_revisions = df_history["Revision Management\n版本管理"].apply(
            lambda x: x.split("版本")[-1] if pd.notna(x) else x
        )

        df_history = df_history.reindex(df.index)

        senders = []
        receivers = []

        for _, row in df.iterrows():
            row_senders = []
            row_receivers = []

            for bus_user in self.bus_users:
                if bus_user in df.columns:
                    if pd.notna(row[bus_user]) and row[bus_user] == "S":
                        row_senders.append(bus_user)
                    elif pd.notna(row[bus_user]) and row[bus_user] == "R":
                        row_receivers.append(bus_user)

            senders.append(",".join(row_senders) if row_senders else "Vector__XXX")
            receivers.append(
                ",".join(row_receivers) if row_receivers else "Vector__XXX"
            )

        df["Msg Cycle Time (ms)\n报文周期时间"] = (
            pd.to_numeric(df["Msg Cycle Time (ms)\n报文周期时间"], errors="coerce")
            .fillna(0)
            .astype(int)
        )
        df["Msg Cycle Time Fast(ms)\n报文发送的快速周期"] = (
            pd.to_numeric(
                df["Msg Cycle Time Fast(ms)\n报文发送的快速周期"], errors="coerce"
            )
            .fillna(0)
            .astype(int)
        )
        df["Msg Nr. Of Reption\n报文快速发送的次数"] = (
            pd.to_numeric(df["Msg Nr. Of Reption\n报文快速发送的次数"], errors="coerce")
            .fillna(0)
            .astype(int)
        )
        df["Msg Delay Time(ms)\n报文延时时间"] = (
            pd.to_numeric(df["Msg Delay Time(ms)\n报文延时时间"], errors="coerce")
            .fillna(0)
            .astype(int)
        )

        new_df = pd.DataFrame(
            {
                "Message ID": df["Msg ID\n报文标识符"].ffill(),
                "Message Name": df["Msg Name\n报文名称"].ffill(),
                "Signal Name": df["Signal Name\n信号名称"],
                "Cycle Type": df["Msg Cycle Time (ms)\n报文周期时间"].ffill(),
                "Msg Time Fast": df[
                    "Msg Cycle Time Fast(ms)\n报文发送的快速周期"
                ].ffill(),
                "Msg Reption": df["Msg Nr. Of Reption\n报文快速发送的次数"].ffill(),
                "Msg Delay": df["Msg Delay Time(ms)\n报文延时时间"].ffill(),
                "Start Byte": df["Start Byte\n起始字节"],
                "Start Bit": df["Start Bit\n起始位"],
                "Length": df["Bit Length (Bit)\n信号长度"],
                "Factor": df["Resolution\n精度"],
                "Offset": df["Offset\n偏移量"],
                "Initinal": df["Initial Value (Hex)\n初始值"],
                "Invalid": df["Invalid Value(Hex)\n无效值"],
                "Min": df["Signal Min. Value (phys)\n物理最小值"],
                "Max": df["Signal Max. Value (phys)\n物理最大值"],
                "Unit": df["Unit\n单位"],
                "Receiver": receivers,
                "Byte Order": df["Byte Order\n排列格式\n(Intel/Motorola)"],
                "Data Type": df["Data Type\n数据类型"],
                "Message Type": df["Msg Type\n报文类型"].ffill(),
                "Send Type": df["Msg Send Type\n报文发送类型"].ffill(),
                "Description": df["Signal Description\n信号描述"],
                "Msg Length": df["Msg Length (Byte)\n报文长度"].ffill(),
                "Signal Value Description": df["Signal Value Description\n信号值描述"],
                "Senders": senders,
                "Signal Send Type": df["Signal Send Type\n信号发送类型"],
                "Inactive value": df["Inactive Value (Hex)\n非使能值"],
            }
        )

        consistent_fields = [
            "Message ID",
            "Message Name",
            "Cycle Type",
            "Send Type",
            "Msg Length",
            "Message Type",
            "Msg Time Fast",
            "Msg Reption",
            "Msg Delay",
        ]

        for field in consistent_fields:
            new_df[field] = new_df.groupby("Message Name")[field].transform("first")

        new_df["Send Type"] = (
            new_df["Send Type"].astype(str).str.replace("Cycle", "Cyclic")
        )
        new_df["Signal Send Type"] = (
            new_df["Signal Send Type"].astype(str).str.replace("Cycle", "Cyclic")
        )

        new_df["Unit"] = new_df["Unit"].astype(str)
        new_df["Unit"] = new_df["Unit"].str.replace("Ω", "Ohm", regex=False)
        new_df["Unit"] = new_df["Unit"].str.replace("℃", "degC", regex=False)

        new_df = new_df.dropna(subset=["Signal Name"])
        new_df["Is Signed"] = new_df["Data Type"].str.contains("Signed", na=False)

        return new_df, all_revisions

    def _create_signal(self, row: pd.Series) -> Optional[cantools.database.can.Signal]:
        try:
            comment = str(row["Description"]) if pd.notna(row["Description"]) else ""
            comment = re.sub(r"[\u4e00-\u9fff]+", "", comment)
            comment = str.replace(comment, "/", "")
            comment = str.replace(comment, "\n", "")
            unit = str(row["Unit"]) if pd.notna(row["Unit"]) else ""
            unit = str.replace(unit, "nan", "")
            byte_order = (
                "big_endian" if row["Byte Order"] == "Motorola MSB" else "little_endian"
            )

            is_float = (
                "Float" in str(row["Data Type"])
                if pd.notna(row["Data Type"])
                else False
            )

            value_descriptions = None
            if pd.notna(row["Signal Value Description"]):
                value_descriptions = ValueDescriptionParser.parse(
                    row["Signal Value Description"]
                )

            receivers = []
            if pd.notna(row["Receiver"]):
                if isinstance(row["Receiver"], str):
                    receivers = row["Receiver"].split(",")
                else:
                    receivers = [str(row["Receiver"])]

            raw_invalid = (
                int(int(row["Invalid"], 16)) if pd.notna(row["Invalid"]) else 0
            )

            send_type_map = {
                "Cyclic": 0,
                "OnChange": 1,
                "OnWrite": 2,
                "IfActive": 3,
                "OnChangeWithRepetition": 4,
                "OnWriteWithRepetition": 5,
                "IfActiveWithRepetition": 6,
                "NoSigSendType": 7,
                "OnChangeAndIfActive": 8,
                "OnChangeAndIfActiveWithRepetition": 9,
                "CA": 10,
                "CE": 11,
                "Event": 12,
            }

            signal_send_type = (
                str(row["Signal Send Type"])
                if str(row["Signal Send Type"])
                else "Cyclic"
            )
            send_type_int = send_type_map.get(signal_send_type, 0)

            attr_sig_inv_val = Attribute(
                value=raw_invalid, definition=self.attr_def_sig_invalid_value
            )
            attr_sig_send_type = Attribute(
                value=send_type_int, definition=self.attr_def_sig_send_type
            )
            attr_sig_inact_val = Attribute(
                value=(
                    int(row["Inactive value"]) if pd.notna(row["Inactive value"]) else 0
                ),
                definition=self.attr_def_sig_inactive_value,
            )

            signal = cantools.database.can.Signal(
                name=str(row["Signal Name"]),
                start=int(row["Start Bit"]),
                length=int(row["Length"]),
                byte_order=byte_order,
                is_signed=bool(row["Is Signed"]),
                raw_initial=int(
                    int(row["Initinal"], 16) if int(row["Initinal"], 16) else 0
                ),
                raw_invalid=(
                    int(int(row["Invalid"], 16)) if pd.notna(row["Invalid"]) else None
                ),
                dbc_specifics=DbcSpecifics(
                    attributes={
                        "GenSigInvalidValue": attr_sig_inv_val,
                        "GenSigSendType": attr_sig_send_type,
                        "GenSigInactiveValue": attr_sig_inact_val,
                    }
                ),
                conversion=cantools.database.conversion.LinearConversion(
                    scale=(
                        int(row["Factor"])
                        if pd.notna(row["Factor"]) and row["Factor"].is_integer()
                        else (float(row["Factor"]) if pd.notna(row["Factor"]) else 1.0)
                    ),
                    offset=(
                        int(row["Offset"])
                        if pd.notna(row["Offset"]) and row["Offset"].is_integer()
                        else (float(row["Offset"]) if pd.notna(row["Offset"]) else 0.0)
                    ),
                    is_float=is_float,
                ),
                minimum=(
                    int(row["Min"])
                    if pd.notna(row["Min"]) and float(row["Min"]).is_integer()
                    else (float(row["Min"]) if pd.notna(row["Min"]) else None)
                ),
                maximum=(
                    int(row["Max"])
                    if pd.notna(row["Max"]) and float(row["Max"]).is_integer()
                    else (float(row["Max"]) if pd.notna(row["Max"]) else None)
                ),
                unit=unit,
                comment=comment,
                receivers=receivers,
                is_multiplexer=False,
            )

            if value_descriptions:
                signal.choices = value_descriptions

            return signal

        except Exception as e:
            print(f"Error creating signal {row['Signal Name']}: {str(e)}")
            return None

    def _create_message(self, msg_id: str, msg_name: str, group: pd.DataFrame) -> bool:
        try:
            frame_id = (
                int(msg_id, 16)
                if isinstance(msg_id, str) and msg_id.startswith("0x")
                else int(msg_id)
            )

            signals = []
            for _, row in group.iterrows():
                signal = self._create_signal(row)
                if signal:
                    signals.append(signal)

            if not signals:
                return False

            senders = []
            if pd.notna(group["Senders"].iloc[0]):
                if isinstance(group["Senders"].iloc[0], str):
                    senders = group["Senders"].iloc[0].split(",")
                else:
                    senders = [str(group["Senders"].iloc[0])]

            # autosar_specifics = AutosarMessageSpecifics()
            # autosar_specifics=autosar_specifics,

            send_type = (
                group["Send Type"].iloc[0]
                if pd.notna(group["Send Type"].iloc[0])
                else None
            )

            send_type_map = {
                "Cyclic": 0,
                "Event": 1,
                "IfActive": 2,
                "CE": 3,
                "CA": 4,
                "NoMsgSendType": 5,
            }

            send_type_str = (
                group["Send Type"].iloc[0]
                if pd.notna(group["Send Type"].iloc[0])
                else "Cyclic"
            )

            mtf = (
                int(group["Msg Time Fast"].iloc[0])
                if pd.notna(group["Msg Time Fast"].iloc[0])
                else 0
            )
            mor = (
                int(group["Msg Reption"].iloc[0])
                if pd.notna(group["Msg Reption"].iloc[0])
                else 0
            )
            mdt = (
                int(group["Msg Delay"].iloc[0])
                if pd.notna(group["Msg Delay"].iloc[0])
                else 0
            )
            send_type_int = send_type_map.get(send_type_str, 0)

            attr_msg_send_type = Attribute(
                value=send_type_int, definition=self.attr_def_msg_send_type
            )
            attr_msg_time_fast = Attribute(
                value=mtf, definition=self.attr_def_msg_cycle_time_fast
            )
            attr_msg_rep = Attribute(
                value=mor, definition=self.attr_def_msg_nr_repetition
            )
            attr_msg_del = Attribute(value=mdt, definition=self.attr_def_msg_delay_time)

            message = cantools.database.can.Message(
                frame_id=frame_id,
                name=str(msg_name),
                length=int(group["Msg Length"].iloc[0]),
                signals=signals,
                senders=senders,
                send_type=send_type,
                cycle_time=(
                    int(group["Cycle Type"].iloc[0])
                    if pd.notna(group["Cycle Type"].iloc[0])
                    else None
                ),
                dbc_specifics=DbcSpecifics(
                    attributes={
                        "GenMsgSendType": attr_msg_send_type,
                        "GenMsgCycleTimeFast": attr_msg_time_fast,
                        "GenMsgNrOfRepetition": attr_msg_rep,
                        "GenMsgDelayTime": attr_msg_del,
                    }
                ),
                # autosar_specifics=AutosarMessageSpecifics(attr_msg_send_type),
                is_extended_frame=False,
                header_byte_order="big_endian",
                protocol=ExcelToDBCConverter.get_file_info(self.excel_path.name)[
                    "protocol"
                ],
                is_fd=(
                    True
                    if ExcelToDBCConverter.get_file_info(self.excel_path.name)[
                        "protocol"
                    ]
                    == "CANFD"
                    else False
                ),
                bus_name=ExcelToDBCConverter.get_file_info(self.excel_path.name)[
                    "domain_name"
                ],
                comment=None,
                sort_signals=None,
            )

            if msg_id.startswith("0x7") and "DiagReq_" in message.name:
                message.dbc.attributes = {
                    "DiagRequest": Attribute(
                        value=1, definition=self.attr_def_diag_request
                    )
                }
            elif msg_id.startswith("0x7") and "DiagResp_" in message.name:
                message.dbc.attributes = {
                    "DiagResponse": Attribute(
                        value=1, definition=self.attr_def_diag_response
                    )
                }
            elif msg_id.startswith("0x7") and "DiagState_" in message.name:
                message.dbc.attributes = {
                    "DiagState": Attribute(value=1, definition=self.attr_def_diag_state)
                }
            elif msg_id.startswith("0x5") and "NM_" in message.name:
                message.dbc.attributes = {
                    "NmMessage": Attribute(value=1, definition=self.attr_def_nm_message)
                }
                self.nm_messages.append(message)
            else:
                self.normal_messages.append(message)

            self.db.messages.append(message)

            return True

        except Exception as e:
            print(f"Error creating message {msg_name}: {str(e)}")
            return False

    def _validate_excel_structure(self, df: pd.DataFrame) -> bool:
        required_columns = [
            "Msg ID\n报文标识符",
            "Msg Name\n报文名称",
            "Signal Name\n信号名称",
            "Start Byte\n起始字节",
            "Start Bit\n起始位",
            "Bit Length (Bit)\n信号长度",
            "Byte Order\n排列格式(Intel/Motorola)",
            "Data Type\n数据类型",
            "Msg Length (Byte)\n报文长度",
        ]

        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            print(
                f"Ошибка: В файле отсутствуют обязательные столбцы: {missing_columns}"
            )
            return False
        return True

    def _validate_message_ids(self, df: pd.DataFrame) -> bool:
        valid = True
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
                print(f"Ошибка: Некорректный ID сообщения: {msg_id}")
                valid = False
        return valid

    def _validate_signal_positions(self, df: pd.DataFrame) -> bool:
        valid = True
        for _, group in df.groupby("Msg ID\n报文标识符"):
            msg_length = group["Msg Length (Byte)\n报文长度"].iloc[0]
            for _, row in group.iterrows():
                start_byte = row["Start Byte\n起始字节"]
                start_bit = row["Start Bit\n起始位"]
                bit_length = row["Bit Length (Bit)\n信号长度"]

                if start_byte >= msg_length:
                    print(
                        f"Ошибка: Сигнал {row['Signal Name\n信号名称']} выходит за пределы сообщения (байт {start_byte} >= длины сообщения {msg_length})"
                    )
                    valid = False

                if start_bit >= 8:
                    print(
                        f"Ошибка: Некорректный стартовый бит {start_bit} в сигнале {row['Signal Name\n信号名称']}"
                    )
                    valid = False

                if bit_length <= 0:
                    print(
                        f"Ошибка: Некорректная длина {bit_length} в сигнале {row['Signal Name\n信号名称']}"
                    )
                    valid = False

                if start_byte * 8 + start_bit + bit_length > msg_length * 8:
                    print(
                        f"Ошибка: Сигнал {row['Signal Name\n信号名称']} выходит за пределы сообщения"
                    )
                    valid = False
        return valid

    def _validate_data_types(self, df: pd.DataFrame) -> bool:
        valid = True
        for _, row in df.iterrows():
            data_type = str(row["Data Type\n数据类型"])
            bit_length = row["Bit Length (Bit)\n信号长度"]

            if "Float" in data_type and bit_length not in [32, 64]:
                print(
                    f"Ошибка: Некорректная длина {bit_length} для типа float в сигнале {row['Signal Name\n信号名称']}"
                )
                valid = False

            if "Signed" in data_type and bit_length < 2:
                print(
                    f"Ошибка: Некорректная длина {bit_length} для signed типа в сигнале {row['Signal Name\n信号名称']}"
                )
                valid = False
        return valid

    def _validate_senders_receivers(self, df: pd.DataFrame) -> bool:
        valid = True
        for msg_id, group in df.groupby("Msg ID\n报文标识符"):
            senders = []
            receivers = []

            for bus_user in self.bus_users:
                if bus_user in group.columns:
                    if "S" in group[bus_user].values:
                        senders.append(bus_user)
                    if "R" in group[bus_user].values:
                        receivers.append(bus_user)

            if not senders:
                print(f"Предупреждение: Сообщение {msg_id} не имеет отправителей")

            if not receivers:
                print(f"Предупреждение: Сообщение {msg_id} не имеет получателей")

        return valid

    def _validate_initial_values(self, df: pd.DataFrame) -> bool:
        valid = True
        for _, row in df.iterrows():
            if pd.notna(row["Initial Value (Hex)\n初始值"]):
                try:
                    int(row["Initial Value (Hex)\n初始值"], 16)
                except ValueError:
                    print(
                        f"Ошибка: Некорректное начальное значение {row['Initial Value (Hex)\n初始值']} для сигнала {row['Signal Name\n信号名称']}"
                    )
                    valid = False
        return valid

    def validate_input_data(self) -> bool:
        try:
            df = pd.read_excel(
                self.excel_path,
                sheet_name="Matrix",
                keep_default_na=True,
                engine="openpyxl",
            )

            checks = [
                self._validate_excel_structure(df),
                self._validate_message_ids(df),
                self._validate_signal_positions(df),
                self._validate_data_types(df),
                self._validate_senders_receivers(df),
                self._validate_initial_values(df),
            ]

            return all(checks)
        except Exception as e:
            print(f"Ошибка при проверке входных данных: {str(e)}")
            return False

    def get_file_info(file_name: str):
        file_start = "ATOM_CAN_Matrix_"
        file_start1 = "ATOM_CANFD_Matrix_"
        file_name_only = os.path.splitext(os.path.basename(file_name))[0]
        if file_name_only.startswith(file_start1):
            protocol = "CANFD"
            start_index = 0
            parts = file_name_only[len(file_start1) :].split("_")
        elif file_name_only.startswith(file_start):
            protocol = "CAN"
            start_index = 0
            parts = file_name_only[len(file_start) :].split("_")
        else:
            protocol = ""
        if not (
            file_name_only.startswith(file_start)
            or file_name_only.startswith(file_start1)
        ):
            return None
        start_index = file_name_only.find(file_start1)
        if start_index != -1:
            parts = file_name_only[start_index + len(file_start1) :].split("_")
        else:
            parts = file_name_only[len(file_start) :].split("_")
        domain_name = parts.pop(0)
        version_string = parts.pop(0)
        if version_string.startswith("V"):
            version = version_string[1:]
            versions = version.split(".")
            if len(versions) != 3:
                return None
        else:
            version = ""
        file_date = parts.pop(0)
        if len(parts) > 0:
            if parts[0] == "internal":
                parts.pop(0)
            device_name = "_".join(parts)
        else:
            device_name = ""

        return {
            "version": version,
            "date": file_date,
            "device_name": device_name,
            "domain_name": domain_name,
            "protocol": protocol,
        }

    def convert(self, output_path: str = "output.dbc") -> bool:
        """Main method convert"""
        try:
            if not self.validate_input_data():
                print("Ошибка: Входные данные не прошли проверку")
                return False
            df, _ = self._load_excel_data()
            grouped = df.groupby(["Message ID", "Message Name"])

            for (msg_id, msg_name), group in grouped:
                self._create_message(msg_id, msg_name, group)

            # revision_lines = [f"Revision:{rev}" for rev in all_revisions]
            # global_comment = 'CM_ "' + ",\n".join(revision_lines) + '" ;\n'

            cantools.database.dump_file(self.db, output_path)

            # with open(output_path, "a", encoding="utf-8") as f:
            #     f.write("\n")
            #     f.write(global_comment)

            print(f"DBC-file successfully created: {output_path}")
            return True

        except Exception as e:
            print(f"Error during conversion: {str(e)}")
            return False


def main():
    parser = argparse.ArgumentParser(description="Convert Excel-files to DBC-files")
    parser.add_argument("--input", required=True, help="Path to Excel-file")
    parser.add_argument("--output", default="output.dbc", help="Output name DBC-file")
    args = parser.parse_args()

    converter = ExcelToDBCConverter(args.input)
    if converter.convert(args.output):
        print("Conversion completed successfully")
    else:
        print("Conversion failed")


if __name__ == "__main__":
    main()
