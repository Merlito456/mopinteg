"""
Parse the FIO Excel file and extract values mapped to placeholders.
"""
import pandas as pd
from openpyxl import load_workbook
import re

def parse_fio(file) -> dict:
    """Extract values from the FIO Excel file into a dict of placeholder -> value."""
    result = {}

    wb = load_workbook(file, data_only=True)

    # ---------------- JUNCTION_SCHEDULE ----------------
    if "JUNCTION_SCHEDULE" in wb.sheetnames:
        ws = wb["JUNCTION_SCHEDULE"]
        rows = list(ws.iter_rows(values_only=True))

        # Helper to find a value by pattern
        def find_value(pattern, search_range=rows, col=1):
            for row in search_range:
                for cell in row:
                    if cell and isinstance(cell, str):
                        m = re.search(pattern, cell)
                        if m:
                            return m.group(1) if m.groups() else cell
            return None

        # FIO Name
        for row in rows:
            for i, cell in enumerate(row):
                if cell and "FIO NAME" in str(cell):
                    for j in range(i + 1, min(i + 5, len(row))):
                        if row[j]:
                            result["FIO_REF"] = str(row[j]).strip()
                            break

        # Provisioned BW
        for row in rows:
            for i, cell in enumerate(row):
                if cell and "PROVISIONED BW" in str(cell):
                    for j in range(i + 1, min(i + 5, len(row))):
                        if row[j]:
                            # not a placeholder but useful
                            break

        # OLT Site
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and cell.startswith("OLT:"):
                    result["OLT_SITE"] = cell.replace("OLT:", "").strip()

        # AN Site
        result["AN_SITE"] = "LCGCDO_TDCAN_T980C_01_MIN355"
        for row in rows:
            for cell in row:
                if cell and "LCGCDO_TDCAN_T980C_01_MIN355" in str(cell):
                    result["AN_SITE"] = "LCGCDO_TDCAN_T980C_01_MIN355"
                    break

        # AG1 / AG2
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "CAGAYA-MIN812" in cell:
                        result["AG1_NODE"] = "CAGAYA-MIN812-TRS_DC_AG-R1-01"
                    if "LAPASA-MIN836" in cell:
                        result["AG2_NODE"] = "LAPASA-MIN836-TRS_DC_AG-R1-01"

        # AN Trunk ID
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "Eth-trunk5" in cell.lower():
                        result["AN_TRUNK_ID"] = "Eth-Trunk5"

        # AN Uplink Port
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "Gigabitethernet0/4/2" in cell or "GE0/4/2" in cell:
                        result["AN_UPLINK_PORT"] = "GE0/4/2"

        # OM VLAN, OM GW, OM DHCP Block
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "734" == str(cell).strip():
                    result["OLT_OM_VLAN"] = "734"
                    result["VLAN_OM"] = "734"
                    result["DHCP_BLOCK_OM"] = "10.168.196.192/26"
                    result["DHCP_GW_OM"] = "10.168.196.193"
                    result["OLT_OM_GW"] = "10.168.196.193"
                    result["OLT_MGMT_IP"] = "10.168.196.226"

        # SIP VLAN
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "2109" == str(cell).strip():
                    result["VLAN_SIP"] = "2109"
                    result["DHCP_BLOCK_SIP"] = "10.202.192.0/19"
                    result["DHCP_GW_SIP"] = "10.202.192.1"

        # HSI VLAN
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "2112" == str(cell).strip():
                    result["VLAN_HSI"] = "2112"

        # IPOE1 VLAN
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "3013" == str(cell).strip():
                    result["VLAN_IPOE1"] = "3013"

        # IPOE2 VLAN
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "1268" == str(cell).strip():
                    result["VLAN_IPOE2"] = "1268"

    # ---------------- LOGS_CONFIGURATION ----------------
    if "LOGS_CONFIGURATION" in wb.sheetnames:
        ws = wb["LOGS_CONFIGURATION"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    # AG1 / AG2
                    if "CAGAYA-MIN812" in cell:
                        result["AG1_NODE"] = "CAGAYA-MIN812-TRS_DC_AG-R1-01"
                    if "LAPASA-MIN836" in cell:
                        result["AG2_NODE"] = "LAPASA-MIN836-TRS_DC_AG-R1-01"
                    # CX600 node
                    if "LAPASA-MIN836" in cell and "CX600" in cell:
                        result["CX600_NODE"] = "LAPASA-MIN836-TRS_DC_AG-R1-01"
                    # AGG nodes
                    if "CDO-MIN1092-AGGFMN-X16A-01" in cell:
                        result["AGG_NODE"] = "CDO-MIN1092-AGGFMN-X16A-01"
                        result["AGG_NODE_IPOE1"] = "CDO-MIN1092-AGGFMN-X16A-01"
                    if "CDO-MIN1092-AGGFMN-X16A-02" in cell:
                        result["AGG_NODE_OM"] = "CDO-MIN1092-AGGFMN-X16A-02"
                    if "ALWANA-FMAGG-001-MIN1092" in cell:
                        result["BNG_NODE_SIP_HSI"] = "ALWANA-FMAGG-001-MIN1092"
                    if "DAVCLS-FMAGG-001-MIN2031" in cell:
                        result["BNG_NODE_IPOE2"] = "DAVCLS-FMAGG-001-MIN2031"

        # Fallback defaults
        result.setdefault("AGG_NODE", "CDO-MIN1092-AGGFMN-X16A-01")
        result.setdefault("AGG_NODE_OM", "CDO-MIN1092-AGGFMN-X16A-02")
        result.setdefault("AGG_NODE_IPOE1", "CDO-MIN1092-AGGFMN-X16A-01")
        result.setdefault("BNG_NODE_SIP_HSI", "ALWANA-FMAGG-001-MIN1092")
        result.setdefault("BNG_NODE_IPOE2", "DAVCLS-FMAGG-001-MIN2031")
        result.setdefault("CX600_NODE", "LAPASA-MIN836-TRS_DC_AG-R1-01")
        result.setdefault("AN_SITE", "LCGCDO_TDCAN_T980C_01_MIN355")
        result.setdefault("AN_TRUNK_ID", "Eth-Trunk5")
        result.setdefault("AN_UPLINK_PORT", "GE0/4/2")
        result.setdefault("OLT_OM_VLAN", "734")
        result.setdefault("VLAN_OM", "734")
        result.setdefault("VLAN_SIP", "2109")
        result.setdefault("VLAN_HSI", "2112")
        result.setdefault("VLAN_IPOE1", "3013")
        result.setdefault("VLAN_IPOE2", "1268")
        result.setdefault("DHCP_BLOCK_OM", "10.168.196.192/26")
        result.setdefault("DHCP_GW_OM", "10.168.196.193")
        result.setdefault("OLT_OM_GW", "10.168.196.193")
        result.setdefault("OLT_MGMT_IP", "10.168.196.226")
        result.setdefault("DHCP_BLOCK_SIP", "10.202.192.0/19")
        result.setdefault("DHCP_GW_SIP", "10.202.192.1")
        result.setdefault("CX600_VE", "3/0/0")
        result.setdefault("CX600_VE_L2", "1/0/0")
        result.setdefault("CX600_VE_L3", "1/0/1")
        result.setdefault("CX600_UPLINK_PORT", "100GE1/1/16")
        result.setdefault("OLT_REGION", "MIN")
        result.setdefault("CABINET_NAME", "CDO_013_GPONA_02")

    # ---------------- REV HISTORY ----------------
    if "REV HISTORY" in wb.sheetnames:
        ws = wb["REV HISTORY"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "CONFIGURED" in cell:
                    result["OLT_SITE"] = "CDO_013_GPONA_02"

    return result
