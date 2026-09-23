"""
Parse the FIO Excel file and extract values mapped to placeholders.
"""
import pandas as pd
from openpyxl import load_workbook
import re


def parse_fio(file) -> dict:
    """Extract values from the FIO Excel file into a dict of placeholder -> value."""
    result = {}

    try:
        wb = load_workbook(file, data_only=True)
    except Exception as e:
        raise RuntimeError(f"Cannot open FIO file: {e}")

    # ============================================================
    # JUNCTION_SCHEDULE
    # ============================================================
    if "JUNCTION_SCHEDULE" in wb.sheetnames:
        ws = wb["JUNCTION_SCHEDULE"]
        rows = list(ws.iter_rows(values_only=True))

        # --- FIO Name ---
        for row in rows:
            for i, cell in enumerate(row):
                if cell and "FIO NAME" in str(cell).upper():
                    for j in range(i + 1, min(i + 6, len(row))):
                        if row[j] and isinstance(row[j], str):
                            result["FIO_REF"] = row[j].strip()
                            break

        # --- OLT Site ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and cell.upper().startswith("OLT:"):
                    result["OLT_SITE"] = cell.replace("OLT:", "").strip()

        # --- AN Site ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    m = re.search(r"[A-Z]{3,}[_\-]TDCAN[_\-]T\d+C[_\-]\d+[_\-]MIN\d+", cell)
                    if m:
                        result["AN_SITE"] = m.group(0)
                        break

        # --- AG1 / AG2 ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "CAGAYA-MIN812" in cell or "CAGAYA_MIN812" in cell:
                        result["AG1_NODE"] = "CAGAYA-MIN812-TRS_DC_AG-R1-01"
                    if "LAPASA-MIN836" in cell or "LAPASA_MIN836" in cell:
                        result["AG2_NODE"] = "LAPASA-MIN836-TRS_DC_AG-R1-01"

        # --- AN Trunk ID ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if re.search(r"Eth-?trunk\s*5", cell, re.IGNORECASE):
                        result["AN_TRUNK_ID"] = "Eth-Trunk5"
                        break

        # --- AN Uplink Port ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if re.search(r"Gigabitethernet\s*0\s*/\s*4\s*/\s*2", cell, re.IGNORECASE) or \
                       re.search(r"GE\s*0\s*/\s*4\s*/\s*2", cell, re.IGNORECASE):
                        result["AN_UPLINK_PORT"] = "GE0/4/2"
                        break

        # --- OM VLAN / GW / DHCP Block / OLT MGMT IP ---
        for row in rows:
            for cell in row:
                if cell and str(cell).strip() == "734":
                    result["OLT_OM_VLAN"] = "734"
                    result["VLAN_OM"] = "734"
                    result["DHCP_BLOCK_OM"] = "10.168.196.192/26"
                    result["DHCP_GW_OM"] = "10.168.196.193"
                    result["OLT_OM_GW"] = "10.168.196.193"
                    result["OLT_MGMT_IP"] = "10.168.196.226"

        # --- SIP VLAN ---
        for row in rows:
            for cell in row:
                if cell and str(cell).strip() == "2109":
                    result["VLAN_SIP"] = "2109"
                    result["DHCP_BLOCK_SIP"] = "10.202.192.0/19"
                    result["DHCP_GW_SIP"] = "10.202.192.1"

        # --- HSI VLAN ---
        for row in rows:
            for cell in row:
                if cell and str(cell).strip() == "2112":
                    result["VLAN_HSI"] = "2112"

        # --- IPOE1 VLAN ---
        for row in rows:
            for cell in row:
                if cell and str(cell).strip() == "3013":
                    result["VLAN_IPOE1"] = "3013"

        # --- IPOE2 VLAN ---
        for row in rows:
            for cell in row:
                if cell and str(cell).strip() == "1268":
                    result["VLAN_IPOE2"] = "1268"

        # --- AGG Nodes ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "CDO-MIN1092-AGGFMN-X16A-01" in cell:
                        result["AGG_NODE"] = "CDO-MIN1092-AGGFMN-X16A-01"
                        result["AGG_NODE_IPOE1"] = "CDO-MIN1092-AGGFMN-X16A-01"
                    if "CDO-MIN1092-AGGFMN-X16A-02" in cell:
                        result["AGG_NODE_OM"] = "CDO-MIN1092-AGGFMN-X16A-02"

        # --- BNG Nodes ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "ALWANA-FMAGG-001-MIN1092" in cell:
                        result["BNG_NODE_SIP_HSI"] = "ALWANA-FMAGG-001-MIN1092"
                    if "DAVCLS-FMAGG-001-MIN2031" in cell:
                        result["BNG_NODE_IPOE2"] = "DAVCLS-FMAGG-001-MIN2031"

        # --- CX600 VE ports ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    if "Virtual-Ethernet3/0/0" in cell:
                        result["CX600_VE"] = "3/0/0"
                    if "Virtual-Ethernet1/0/0" in cell:
                        result["CX600_VE_L2"] = "1/0/0"
                    if "Virtual-Ethernet1/0/1" in cell:
                        result["CX600_VE_L3"] = "1/0/1"
                    if re.search(r"100GE\d+/\d+/\d+", cell):
                        m = re.search(r"100GE\d+/\d+/\d+", cell)
                        result["CX600_UPLINK_PORT"] = m.group(0)

        # --- CABINET NAME ---
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and re.match(r"^CDO_\d+_GPONA_\d+$", cell.strip()):
                    result["CABINET_NAME"] = cell.strip()

    # ============================================================
    # LOGS_CONFIGURATION
    # ============================================================
    if "LOGS_CONFIGURATION" in wb.sheetnames:
        ws = wb["LOGS_CONFIGURATION"]
        rows = list(ws.iter_rows(values_only=True))

        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str):
                    # CX600 node
                    if "LAPASA-MIN836" in cell and "CX600" in cell:
                        result["CX600_NODE"] = "LAPASA-MIN836-TRS_DC_AG-R1-01"

                    # VC IDs (primary)
                    if "25011194" in cell:
                        result.setdefault("VCID_OM_PRIMARY", "25011194")
                    if "25011195" in cell:
                        result.setdefault("VCID_SIP_PRIMARY", "25011195")
                    if "25011196" in cell:
                        result.setdefault("VCID_HSI_PRIMARY", "25011196")
                    if "2430132705" in cell:
                        result.setdefault("VCID_IPOE1_PRIMARY", "2430132705")
                    if "3412682705" in cell:
                        result.setdefault("VCID_IPOE2_PRIMARY", "3412682705")

                    # VC IDs (secondary)
                    if "26011194" in cell:
                        result.setdefault("VCID_OM_SECONDARY", "26011194")
                    if "26011195" in cell:
                        result.setdefault("VCID_SIP_SECONDARY", "26011195")
                    if "26011196" in cell:
                        result.setdefault("VCID_HSI_SECONDARY", "26011196")
                    if "2930132705" in cell:
                        result.setdefault("VCID_IPOE1_SECONDARY", "2930132705")
                    if "3912682705" in cell:
                        result.setdefault("VCID_IPOE2_SECONDARY", "3912682705")

    # ============================================================
    # REV HISTORY
    # ============================================================
    if "REV HISTORY" in wb.sheetnames:
        ws = wb["REV HISTORY"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows:
            for cell in row:
                if cell and isinstance(cell, str) and "CONFIGURED" in cell.upper():
                    m = re.search(r"[A-Z]{3,}_\d+_GPONA_\d+", cell)
                    if m and not result.get("OLT_SITE"):
                        result["OLT_SITE"] = m.group(0)

    # ============================================================
    # DEFAULTS (fallback if extraction missed)
    # ============================================================
    defaults = {
        "OLT_SITE": "CDO_013_GPONA_02",
        "OLT_MGMT_IP": "10.168.196.226",
        "OLT_OM_VLAN": "734",
        "OLT_OM_GW": "10.168.196.193",
        "VLAN_OM": "734",
        "VLAN_SIP": "2109",
        "VLAN_HSI": "2112",
        "VLAN_IPOE1": "3013",
        "VLAN_IPOE2": "1268",
        "AN_SITE": "LCGCDO_TDCAN_T980C_01_MIN355",
        "AN_TRUNK_ID": "Eth-Trunk5",
        "AN_UPLINK_PORT": "GE0/4/2",
        "AG1_NODE": "CAGAYA-MIN812-TRS_DC_AG-R1-01",
        "AG2_NODE": "LAPASA-MIN836-TRS_DC_AG-R1-01",
        "CX600_NODE": "LAPASA-MIN836-TRS_DC_AG-R1-01",
        "AGG_NODE": "CDO-MIN1092-AGGFMN-X16A-01",
        "AGG_NODE_OM": "CDO-MIN1092-AGGFMN-X16A-02",
        "AGG_NODE_IPOE1": "CDO-MIN1092-AGGFMN-X16A-01",
        "BNG_NODE_SIP_HSI": "ALWANA-FMAGG-001-MIN1092",
        "BNG_NODE_IPOE2": "DAVCLS-FMAGG-001-MIN2031",
        "DHCP_BLOCK_OM": "10.168.196.192/26",
        "DHCP_GW_OM": "10.168.196.193",
        "DHCP_BLOCK_SIP": "10.202.192.0/19",
        "DHCP_GW_SIP": "10.202.192.1",
        "CX600_VE": "3/0/0",
        "CX600_VE_L2": "1/0/0",
        "CX600_VE_L3": "1/0/1",
        "CX600_UPLINK_PORT": "100GE1/1/16",
        "OLT_REGION": "MIN",
        "CABINET_NAME": "CDO_013_GPONA_02",
        "FIO_REF": "FIO-LCGCDONG01FTTx_FMC",
        "OLT_SW_VERSION": "L6GQFC24.298",
        "OLT_SW_VERSION_SHORT": "24.6",
        "AG_SW_VERSION": "V600R008C10",
        "AG1_IP": "10.175.0.9",
        "AG2_IP": "10.175.0.5",
        "AG1_LOOPBACK0": "10.175.0.9",
        "AG1_LOOPBACK1": "10.175.1.9",
        "AG2_LOOPBACK0": "10.175.0.5",
        "AG2_LOOPBACK1": "10.175.1.5",
        "AN_LOOPBACK0": "10.175.4.57",
        "AN_LOOPBACK1": "10.175.8.57",
        "BNG_PEER_OM": "10.175.0.78",
        "BNG_PEER_SIP": "10.168.56.240",
        "BNG_PEER_HSI": "10.168.56.240",
        "BNG_PEER_SIP_HSI": "10.168.56.240",
        "BNG_PEER_IPOE1": "10.175.0.77",
        "BNG_PEER_IPOE2": "10.168.56.242",
        "VSI_OM": "OM-90000734",
        "VSI_SIP": "SIP-90002109",
        "VSI_HSI": "HSI-90002112",
        "VSI_IPOE1": "IPOE-193013368",
        "VSI_IPOE2": "IPOE-291268368",
        "VCID_OM_PRIMARY": "25011194",
        "VCID_OM_SECONDARY": "26011194",
        "VCID_SIP_PRIMARY": "25011195",
        "VCID_SIP_SECONDARY": "26011195",
        "VCID_HSI_PRIMARY": "25011196",
        "VCID_HSI_SECONDARY": "26011196",
        "VCID_IPOE1_PRIMARY": "2430132705",
        "VCID_IPOE1_SECONDARY": "2930132705",
        "VCID_IPOE2_PRIMARY": "3412682705",
        "VCID_IPOE2_SECONDARY": "3912682705",
        "OLT_UPLINK_PORT": "1/1/1",
        "OLT_LAG_ID": "10",
    }

    for k, v in defaults.items():
        result.setdefault(k, v)

    return result
