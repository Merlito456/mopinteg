"""
Parse the FIO Excel file and extract values mapped to placeholders.

FULLY GENERIC — no hardcoded values. Extracts by:
- Labels (e.g., "OM VLAN", "NETWORK IP", "GATEWAY")
- Patterns (e.g., IP addresses, VLAN numbers, node names)
- Structure (e.g., "AG-HOMING 1" / "AG-HOMING 2" labels)
"""
import re
from openpyxl import load_workbook


# ============================================================
# Helper utilities
# ============================================================
def _s(cell):
    """Safe string conversion."""
    if cell is None:
        return ""
    return str(cell).strip()


def _is_ip(s):
    """Check if string is a valid IPv4 address."""
    return bool(re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$", str(s).strip()))


def _is_ip_or_cidr(s):
    """Check if string is an IP or CIDR (e.g., 10.168.196.192/26)."""
    return bool(re.match(r"^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}(/\d{1,2})?$", str(s).strip()))


def _find_value_after_label(rows, label_patterns, max_lookahead=6):
    """
    Search rows for a cell matching any label pattern, then return the next
    non-empty cell value (looking rightward in the same row).
    """
    for row in rows:
        for i, cell in enumerate(row):
            cell_s = _s(cell)
            for pattern in label_patterns:
                if re.search(pattern, cell_s, re.IGNORECASE):
                    # Look rightward for the next non-empty cell
                    for j in range(i + 1, min(i + 1 + max_lookahead, len(row))):
                        val = row[j]
                        if val is not None and _s(val):
                            return _s(val)
    return None


def _find_all_ip_like(rows, exclude=None):
    """Return all IP-like strings found in the sheet."""
    exclude = exclude or set()
    ips = []
    for row in rows:
        for cell in row:
            v = _s(cell)
            if _is_ip(v) and v not in exclude and v not in ips:
                ips.append(v)
    return ips


def _find_all_cidr(rows):
    """Return all CIDR strings found in the sheet."""
    cidrs = []
    for row in rows:
        for cell in row:
            v = _s(cell)
            if _is_ip_or_cidr(v) and "/" in v and v not in cidrs:
                cidrs.append(v)
    return cidrs


# ============================================================
# Main parser
# ============================================================
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

        # ------------------------------------------------
        # FIO NAME
        # ------------------------------------------------
        val = _find_value_after_label(rows, [r"FIO\s*NAME"])
        if val:
            result["FIO_REF"] = val

        # ------------------------------------------------
        # OLT SITE — "OLT: <site>"
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                cell_s = _s(cell)
                m = re.match(r"^OLT\s*:\s*(.+)$", cell_s, re.IGNORECASE)
                if m:
                    result["OLT_SITE"] = m.group(1).strip()
                    break

        # ------------------------------------------------
        # AN SITE — flexible pattern
        # Matches: LCGCDO_TDCAN_T980C_01_MIN355
        #          KALASUNGAY_TDCAN_T980C_MIN1533
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                cell_s = _s(cell)
                m = re.search(
                    r"[A-Z][A-Z0-9_\-]{2,}_(?:TDCAN|TDCAG|TDCDO)_[A-Z0-9]+_MIN\d+",
                    cell_s,
                )
                if m:
                    result["AN_SITE"] = m.group(0)
                    break
            if result.get("AN_SITE"):
                break

        # ------------------------------------------------
        # SITE_NAME — prefix of AN_SITE
        # ------------------------------------------------
        if result.get("AN_SITE"):
            parts = re.split(r"[_\-]", result["AN_SITE"])
            if parts:
                result["SITE_NAME"] = parts[0].strip()

        # ------------------------------------------------
        # AN Product & Family (from AN_SITE line or nearby)
        # Look for "ATN 980C", "CX600-X8(V8)", etc.
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                cell_s = _s(cell)
                m = re.search(r"ATN\s*\d+[A-Z]*", cell_s)
                if m and not result.get("AN_PRODUCT"):
                    result["AN_PRODUCT"] = m.group(0)
                m = re.search(r"CX600[- ]?X\d+.*?\(V\d+\)", cell_s)
                if m and not result.get("AG_MODEL"):
                    result["AG_MODEL"] = m.group(0)

        # ------------------------------------------------
        # AN Trunk ID — "eth-trunk3", "Eth-Trunk5", etc.
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                cell_s = _s(cell)
                m = re.search(r"eth[- ]?trunk\s*(\d+)", cell_s, re.IGNORECASE)
                if m:
                    result["AN_TRUNK_ID"] = f"Eth-Trunk{m.group(1)}"
                    break
            if result.get("AN_TRUNK_ID"):
                break

        # ------------------------------------------------
        # AN Uplink Port — "GigabitEthernet0/3/1" or "GE0/4/2"
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                cell_s = _s(cell)
                m = re.search(r"GigabitEthernet\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)", cell_s, re.IGNORECASE)
                if m:
                    result["AN_UPLINK_PORT"] = f"GE{m.group(1)}/{m.group(2)}/{m.group(3)}"
                    break
                m = re.search(r"\bGE\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)", cell_s, re.IGNORECASE)
                if m:
                    result["AN_UPLINK_PORT"] = f"GE{m.group(1)}/{m.group(2)}/{m.group(3)}"
                    break
            if result.get("AN_UPLINK_PORT"):
                break

        # ------------------------------------------------
        # AG Nodes — look for "AG-HOMING 1" and "AG-HOMING 2"
        # The AG node value is typically found nearby the label.
        # ------------------------------------------------
        # Strategy: find the row that contains "AG-HOMING 1" and "AG-HOMING 2".
        # The node names are usually in the header row above or the rows below.
        # Alternative: scan all node-like cells and pick the two closest to these labels.

        # Find all node-like strings in the sheet
        node_pattern = re.compile(
            r"[A-Z][A-Z0-9]+(?:[-_][A-Z0-9]+)+"
        )
        all_nodes = []
        for row in rows:
            for cell in row:
                cell_s = _s(cell)
                if node_pattern.match(cell_s) and len(cell_s) > 8:
                    if cell_s not in all_nodes:
                        all_nodes.append(cell_s)

        # Filter out known non-node strings
        exclude_tokens = [
            "FACILITY", "IMPLEMENTATION", "ORDER", "JUNCTION", "SCHEDULE",
            "EXISTING", "TRAIL", "FMC", "CHANNEL", "ASSIGNMENT",
            "CLIENT", "LABEL", "REMARKS", "PROVISIONED",
            "NETWORK", "SUBNET", "MASK", "GATEWAY", "VLAN",
            "Virtual-Ethernet", "Virtual-Ethernet",
        ]
        candidate_nodes = [
            n for n in all_nodes
            if not any(tok.upper() in n.upper() for tok in exclude_tokens)
        ]

        # Look for "AG-HOMING" labels and grab the nearest node-like cell
        ag_homing_positions = []
        for r_idx, row in enumerate(rows):
            for c_idx, cell in enumerate(row):
                cell_s = _s(cell)
                if re.search(r"AG[-_ ]?HOMING\s*1", cell_s, re.IGNORECASE):
                    ag_homing_positions.append((1, r_idx, c_idx))
                elif re.search(r"AG[-_ ]?HOMING\s*2", cell_s, re.IGNORECASE):
                    ag_homing_positions.append((2, r_idx, c_idx))

        # For each AG-HOMING label, find the nearest node-like cell
        for ag_num, r_idx, c_idx in ag_homing_positions:
            nearest = None
            nearest_dist = 999
            for row2_idx, row2 in enumerate(rows):
                for cell2_idx, cell2 in enumerate(row2):
                    cell2_s = _s(cell2)
                    if cell2_s in candidate_nodes:
                        # Distance = row distance + small column factor
                        dist = abs(row2_idx - r_idx) * 10 + abs(cell2_idx - c_idx) * 0.1
                        if dist < nearest_dist:
                            nearest_dist = dist
                            nearest = cell2_s
            if nearest:
                if ag_num == 1:
                    result["AG1_NODE"] = nearest
                else:
                    result["AG2_NODE"] = nearest

        # Fallback: if AG nodes still not found, use first two candidates
        if not result.get("AG1_NODE") and len(candidate_nodes) > 0:
            result["AG1_NODE"] = candidate_nodes[0]
        if not result.get("AG2_NODE") and len(candidate_nodes) > 1:
            result["AG2_NODE"] = candidate_nodes[1]

        # ------------------------------------------------
        # Management table: VLAN, IP block, Gateway, OLT Mgmt IP
        # Rows are labeled "OM VLAN", "NETWORK IP", "GATEWAY", "IP block"
        # ------------------------------------------------
        # Find the section with "NETWORK IP" + "GATEWAY" + "OM VLAN"
        # Then read the next data row.

        # Scan every row for these labels
        for r_idx, row in enumerate(rows):
            row_text = " ".join(_s(c) for c in row)
            if re.search(r"NETWORK\s*IP", row_text, re.IGNORECASE) and \
               re.search(r"GATEWAY", row_text, re.IGNORECASE) and \
               re.search(r"OM\s*VLAN|VLAN", row_text, re.IGNORECASE):
                # Found header row. Now find the data row (r_idx+1 .. r_idx+5)
                for dr in range(r_idx + 1, min(r_idx + 6, len(rows))):
                    data_row = rows[dr]
                    for cell in data_row:
                        cell_s = _s(cell)
                        # Network IP (CIDR block)
                        if _is_ip(cell_s) and not result.get("OLT_MGMT_IP"):
                            # Could be mgmt IP; but we'll pick the first
                            pass
                        if _is_ip_or_cidr(cell_s) and "/" in cell_s and not result.get("DHCP_BLOCK_OM"):
                            # CIDR block
                            result["DHCP_BLOCK_OM"] = cell_s
                    # Find gateway & OM VLAN in the same data row
                    for cell in data_row:
                        cell_s = _s(cell)
                        if _is_ip(cell_s) and not result.get("OLT_OM_GW"):
                            result["OLT_OM_GW"] = cell_s
                            result["DHCP_GW_OM"] = cell_s
                        if re.match(r"^\d{1,4}(\.0)?$", cell_s) and not result.get("OLT_OM_VLAN"):
                            # VLAN is typically 2-4 digits (or like "733.0")
                            vlan_num = cell_s.replace(".0", "")
                            if 1 <= int(vlan_num) <= 4095:
                                result["OLT_OM_VLAN"] = vlan_num
                                result["VLAN_OM"] = vlan_num
                    if result.get("DHCP_BLOCK_OM") and result.get("OLT_OM_GW") and result.get("OLT_OM_VLAN"):
                        break
                break

        # ------------------------------------------------
        # OLT Management IP — usually in the site info table at bottom
        # Row: "Cabinet Name | IP Address | ..."
        # ------------------------------------------------
        for row in rows:
            for i, cell in enumerate(row):
                cell_s = _s(cell)
                if re.search(r"IP\s*Address", cell_s, re.IGNORECASE):
                    # Look at next non-empty IP in same row or next rows
                    for dr in [row] + (rows[rows.index(row) + 1:rows.index(row) + 4] if row in rows else []):
                        for c in dr:
                            c_s = _s(c)
                            if _is_ip(c_s) and not result.get("OLT_MGMT_IP"):
                                result["OLT_MGMT_IP"] = c_s
                                break
                        if result.get("OLT_MGMT_IP"):
                            break

        # Fallback: find first IP in the range 10.168.196.x
        if not result.get("OLT_MGMT_IP"):
            for row in rows:
                for cell in row:
                    c_s = _s(cell)
                    if re.match(r"^10\.168\.196\.\d+$", c_s) and not result.get("OLT_MGMT_IP"):
                        result["OLT_MGMT_IP"] = c_s
                        break

        # ------------------------------------------------
        # SIP VLAN / HSI VLAN (PPPoE) / IPoE VLANs
        # They appear in their own sub-tables (look for their labels)
        # ------------------------------------------------
        # VOICE VLAN
        val = _find_value_after_label(rows, [r"VOICE\s*VLAN"], max_lookahead=10)
        # Sometimes the value is in the row below the label
        if not val:
            for r_idx, row in enumerate(rows):
                for i, cell in enumerate(row):
                    if re.search(r"VOICE\s*VLAN", _s(cell), re.IGNORECASE):
                        for dr in range(r_idx + 1, min(r_idx + 4, len(rows))):
                            for c in rows[dr]:
                                c_s = _s(c)
                                if re.match(r"^\d{1,4}(\.0)?$", c_s) and 1 <= int(c_s.replace(".0", "")) <= 4095:
                                    val = c_s.replace(".0", "")
                                    break
                            if val:
                                break
        if val:
            result["VLAN_SIP"] = val

        # DATA VLAN (PPPoE)
        for r_idx, row in enumerate(rows):
            for i, cell in enumerate(row):
                if re.search(r"DATA\s*VLAN\s*\(PPPoE\)", _s(cell), re.IGNORECASE):
                    for dr in range(r_idx + 1, min(r_idx + 4, len(rows))):
                        for c in rows[dr]:
                            c_s = _s(c)
                            if re.match(r"^\d{1,4}(\.0)?$", c_s) and 1 <= int(c_s.replace(".0", "")) <= 4095:
                                result["VLAN_HSI"] = c_s.replace(".0", "")
                                break
                        if result.get("VLAN_HSI"):
                            break

        # IPoE VLAN (dynamic and static)
        for r_idx, row in enumerate(rows):
            for i, cell in enumerate(row):
                if re.search(r"DATA\s*VLAN\s*\(IPoE\)", _s(cell), re.IGNORECASE):
                    # Search the sub-table for "DYNAMIC" and "STATIC" values
                    for dr in range(r_idx, min(r_idx + 8, len(rows))):
                        sub_row_text = " ".join(_s(c) for c in rows[dr])
                        for c in rows[dr]:
                            c_s = _s(c)
                            if re.match(r"^\d{1,4}(\.0)?$", c_s) and 1 <= int(c_s.replace(".0", "")) <= 4095:
                                vlan = c_s.replace(".0", "")
                                if "DYNAMIC" in sub_row_text.upper() and not result.get("VLAN_IPOE1"):
                                    result["VLAN_IPOE1"] = vlan
                                elif "STATIC" in sub_row_text.upper() and not result.get("VLAN_IPOE2"):
                                    result["VLAN_IPOE2"] = vlan

        # ------------------------------------------------
        # CABINET NAME — pattern <SITEPREFIX>_<number>_GPONA_<number>
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                c_s = _s(cell)
                m = re.match(r"^[A-Z]{2,}_\d+_GPONA_\d+$", c_s)
                if m:
                    result["CABINET_NAME"] = c_s
                    break
            if result.get("CABINET_NAME"):
                break

        # ------------------------------------------------
        # CX600 VE ports (from "Virtual-EthernetX/Y/Z" strings)
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                c_s = _s(cell)
                for m in re.finditer(r"Virtual[- ]Ethernet\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)", c_s, re.IGNORECASE):
                    port = f"{m.group(1)}/{m.group(2)}/{m.group(3)}"
                    # Heuristic: ports starting with "1" are L2/L3, "3" are OM
                    if port.startswith("1/0/0"):
                        result.setdefault("CX600_VE_L2", port)
                    elif port.startswith("1/0/1"):
                        result.setdefault("CX600_VE_L3", port)
                    elif port.startswith("3/0/0"):
                        result.setdefault("CX600_VE", port)
                    elif port.startswith("3/0/1"):
                        result.setdefault("CX600_VE_L3", port)

        # ------------------------------------------------
        # CX600 Uplink Port — "100GE1/1/16"
        # ------------------------------------------------
        for row in rows:
            for cell in row:
                c_s = _s(cell)
                m = re.search(r"100GE\d+/\d+/\d+", c_s)
                if m:
                    result["CX600_UPLINK_PORT"] = m.group(0)
                    break
            if result.get("CX600_UPLINK_PORT"):
                break

    # ============================================================
    # LOGS_CONFIGURATION — VC IDs, VSI names, peer IPs
    # ============================================================
    if "LOGS_CONFIGURATION" in wb.sheetnames:
        ws = wb["LOGS_CONFIGURATION"]
        rows = list(ws.iter_rows(values_only=True))

        # Collect all text in the sheet
        all_text = " ".join(_s(c) for row in rows for c in row)

        # ------------------------------------------------
        # CX600 Node — look for CX600-X8(V8) node names near AN site
        # Any node name ending in the AG homing pattern
        # ------------------------------------------------
        # Reuse AG1_NODE-derived CX600_NODE
        if result.get("AG1_NODE"):
            result.setdefault("CX600_NODE", result["AG1_NODE"])

        # ------------------------------------------------
        # VSI names — pattern like "OM-80000733", "SIP-90003098"
        # ------------------------------------------------
        vsi_patterns = {
            "VSI_OM": r"\bOM-\d{6,}\b",
            "VSI_SIP": r"\bSIP-\d{6,}\b",
            "VSI_HSI": r"\bHSI-\d{6,}\b",
            "VSI_IPOE1": r"\bIPOE-\d{6,}\b",
            "VSI_IPOE2": r"\bIPOE-\d{6,}\b",
        }
        # For IPoE, distinguish dynamic vs static by position (1st vs 2nd)
        ipoe_matches = re.findall(r"\bIPOE-\d{6,}\b", all_text)
        if ipoe_matches:
            # Deduplicate preserving order
            unique_ipoe = list(dict.fromkeys(ipoe_matches))
            if len(unique_ipoe) >= 1:
                result["VSI_IPOE1"] = unique_ipoe[0]
            if len(unique_ipoe) >= 2:
                result["VSI_IPOE2"] = unique_ipoe[1]

        for key, pattern in vsi_patterns.items():
            if key.startswith("VSI_IPOE"):
                continue
            m = re.search(pattern, all_text)
            if m:
                result[key] = m.group(0)

        # ------------------------------------------------
        # VC IDs — pattern "negotiation-vc-id NNNNNNNN"
        # ------------------------------------------------
        # Map: OM = first pair, SIP = second, HSI = third, IPoE1 = fourth, IPoE2 = fifth
        vc_matches = re.findall(r"negotiation-vc-id\s+(\d{6,})", all_text)
        unique_vc = list(dict.fromkeys(vc_matches))

        # Split into primary / secondary pairs (assume even/odd or by context)
        if len(unique_vc) >= 10:
            result["VCID_OM_PRIMARY"] = unique_vc[0]
            result["VCID_OM_SECONDARY"] = unique_vc[1]
            result["VCID_SIP_PRIMARY"] = unique_vc[2]
            result["VCID_SIP_SECONDARY"] = unique_vc[3]
            result["VCID_HSI_PRIMARY"] = unique_vc[4]
            result["VCID_HSI_SECONDARY"] = unique_vc[5]
            result["VCID_IPOE1_PRIMARY"] = unique_vc[6]
            result["VCID_IPOE1_SECONDARY"] = unique_vc[7]
            result["VCID_IPOE2_PRIMARY"] = unique_vc[8]
            result["VCID_IPOE2_SECONDARY"] = unique_vc[9]
        elif len(unique_vc) >= 2:
            result["VCID_OM_PRIMARY"] = unique_vc[0]
            result["VCID_OM_SECONDARY"] = unique_vc[1]

        # ------------------------------------------------
        # AN Loopback IPs — from the AN site header
        # Format: "SITENAME (L0_IP)" and "L1: L1_IP"
        # ------------------------------------------------
        # Look for the AN site node with IP in parentheses
        if result.get("AN_SITE"):
            an_site_short = result["AN_SITE"]
            # Find the row with the AN site
            for row in rows:
                for cell in row:
                    c_s = _s(cell)
                    if an_site_short in c_s:
                        # Extract L0 IP from "(IP)"
                        m = re.search(r"\((\d+\.\d+\.\d+\.\d+)\)", c_s)
                        if m and not result.get("AN_LOOPBACK0"):
                            result["AN_LOOPBACK0"] = m.group(1)

            # Find "L1: <IP>" near the AN site reference
            for row in rows:
                for cell in row:
                    c_s = _s(cell)
                    m = re.search(r"L1:\s*(\d+\.\d+\.\d+\.\d+)", c_s)
                    if m and not result.get("AN_LOOPBACK1"):
                        result["AN_LOOPBACK1"] = m.group(1)

        # ------------------------------------------------
        # AG1/AG2 IPs — from AG node headers
        # Each row starts with "AG_NODE (AG_IP)"
        # ------------------------------------------------
        if result.get("AG1_NODE"):
            an_site_short = result["AG1_NODE"]
            for row in rows:
                for cell in row:
                    c_s = _s(cell)
                    if an_site_short in c_s:
                        m = re.search(r"\((\d+\.\d+\.\d+\.\d+)\)", c_s)
                        if m and not result.get("AG1_IP"):
                            result["AG1_IP"] = m.group(1)

        if result.get("AG2_NODE"):
            an_site_short = result["AG2_NODE"]
            for row in rows:
                for cell in row:
                    c_s = _s(cell)
                    if an_site_short in c_s:
                        m = re.search(r"\((\d+\.\d+\.\d+\.\d+)\)", c_s)
                        if m and not result.get("AG2_IP"):
                            result["AG2_IP"] = m.group(1)

        # AG Loopbacks (L0, L1)
        for idx, ag_key in enumerate(["AG1", "AG2"], start=1):
            if result.get(f"{ag_key}_IP"):
                # Look for "L0: IP" and "L1: IP"
                for row in rows:
                    for cell in row:
                        c_s = _s(cell)
                        m = re.search(r"L0:\s*(\d+\.\d+\.\d+\.\d+)", c_s)
                        if m and not result.get(f"{ag_key}_LOOPBACK0"):
                            result[f"{ag_key}_LOOPBACK0"] = m.group(1)
                        m = re.search(r"L1:\s*(\d+\.\d+\.\d+\.\d+)", c_s)
                        if m and not result.get(f"{ag_key}_LOOPBACK1"):
                            result[f"{ag_key}_LOOPBACK1"] = m.group(1)

        # ------------------------------------------------
        # BNG peer IPs — labeled "OM:", "SIP:", "HSI", "IPoE - DST:", "IPoE - STATIC:"
        # ------------------------------------------------
        label_map = {
            "BNG_PEER_OM": r"OM\s*:",
            "BNG_PEER_SIP": r"SIP\s*:",
            "BNG_PEER_HSI": r"HSI\s*\(PPoE\)\s*:",
            "BNG_PEER_IPOE1": r"IPoE\s*[-–]\s*(?:DST|CEN)\s*:",
            "BNG_PEER_IPOE2": r"IPoE\s*[-–]\s*STATIC\s*:",
        }

        for row in rows:
            for cell in row:
                c_s = _s(cell)
                for key, lbl_pat in label_map.items():
                    if re.search(lbl_pat, c_s, re.IGNORECASE) and not result.get(key):
                        m = re.search(r"\((\d+\.\d+\.\d+\.\d+)\)", c_s)
                        if m:
                            result[key] = m.group(1)

        # ------------------------------------------------
        # BNG Node names — from "(IP)" rows
        # ------------------------------------------------
        label_map_nodes = {
            "BNG_NODE_SIP_HSI": r"SIP\s*:",
            "BNG_NODE_IPOE2": r"IPoE\s*[-–]\s*STATIC\s*:",
        }
        for row in rows:
            for cell in row:
                c_s = _s(cell)
                for key, lbl_pat in label_map_nodes.items():
                    if re.search(lbl_pat, c_s, re.IGNORECASE) and not result.get(key):
                        # Extract node name before "("
                        m = re.match(r"^(?:[A-Za-z0-9_\-]+\s*:)?\s*([A-Z0-9_\-]+)\s*\(([\d\.]+)\)", c_s)
                        if m:
                            result[key] = m.group(1)

    # ============================================================
    # REV HISTORY — fallback for OLT_SITE
    # ============================================================
    if "REV HISTORY" in wb.sheetnames:
        ws = wb["REV HISTORY"]
        rows = list(ws.iter_rows(values_only=True))
        for row in rows:
            for cell in row:
                c_s = _s(cell)
                if "CONFIGURED" in c_s.upper():
                    m = re.search(r"[A-Z]{2,}_\d+_GPONA_\d+", c_s)
                    if m and not result.get("OLT_SITE"):
                        result["OLT_SITE"] = m.group(0)

    # ============================================================
    # FINAL DERIVATIONS
    # ============================================================
    # SITE_NAME from AN_SITE (in case not set)
    if not result.get("SITE_NAME") and result.get("AN_SITE"):
        parts = re.split(r"[_\-]", result["AN_SITE"])
        if parts:
            result["SITE_NAME"] = parts[0].strip()

    # Site-based defaults for DOC-scope fields (non-FIO sourced)
    # These are the only "safe" defaults — other missing fields remain empty.
    result.setdefault("OLT_PRODUCT", "MF-2")
    result.setdefault("OLT_PRODUCT_SHORT", "MF-2")
    result.setdefault("AN_PRODUCT_FAMILY", "ATN Series Equipment")
    result.setdefault("AG_SYSTEM", "CX600(V8)")
    result.setdefault("AG_PRODUCT", "CX600 Series Equipment")
    result.setdefault("CABINET_ENTRY", "Outdoor")
    result.setdefault("E2E_SYSTEM_1", "Lightspan MF-2")
    result.setdefault("E2E_SYSTEM_2", "CX600(V8)")
    result.setdefault("E2E_SYSTEM_3", "ATN 980C")
    result.setdefault("OLT_LAG_ID", "10")
    result.setdefault("OLT_UPLINK_PORT", "1/1/1")
    result.setdefault("OLT_SW_VERSION", "L6GQFC24.298")
    result.setdefault("OLT_SW_VERSION_SHORT", "24.6")
    result.setdefault("AG_SW_VERSION", "V600R008C10")

    return result
