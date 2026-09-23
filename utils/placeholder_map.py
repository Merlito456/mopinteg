"""
Central mapping of all placeholders in the MOP Integration Template.
Each entry defines:
- key: placeholder name (without braces)
- label: human-readable label
- source: 'FIO' | 'EWP' | 'DOC' | 'VAULT'
- default: default value (optional)
- group: UI grouping
"""

PLACEHOLDER_MAP = [
    # ============================================================
    # DOCUMENT METADATA
    # ============================================================
    {"key": "CIRCUIT_ID", "label": "Circuit / Project ID", "source": "DOC", "group": "Document Metadata", "default": "B2C VMDU FTH"},
    {"key": "VERSION", "label": "MOP Version", "source": "DOC", "group": "Document Metadata", "default": "1.1"},
    {"key": "DATE", "label": "Document Date", "source": "DOC", "group": "Document Metadata", "default": "2026-07-31"},
    {"key": "DATE_PRIMARY", "label": "Primary Implementation Date", "source": "DOC", "group": "Document Metadata", "default": "July 31, 2026"},
    {"key": "DATE_SECONDARY", "label": "Secondary Implementation Date", "source": "DOC", "group": "Document Metadata", "default": "August 1, 2026"},

    # ============================================================
    # NETWORK ELEMENTS
    # ============================================================
    {"key": "OLT_PRODUCT", "label": "OLT Product", "source": "EWP", "group": "Network Elements", "default": "Lightspan MF-2"},
    {"key": "OLT_PRODUCT_SHORT", "label": "OLT Product (Short)", "source": "DOC", "group": "Network Elements", "default": "MF-2"},
    {"key": "OLT_SITE", "label": "OLT Site", "source": "FIO", "group": "Network Elements"},
    {"key": "OLT_MGMT_IP", "label": "OLT Management IP", "source": "FIO", "group": "Network Elements"},
    {"key": "OLT_OM_VLAN", "label": "OLT OM VLAN", "source": "FIO", "group": "Network Elements"},
    {"key": "OLT_OM_GW", "label": "OLT OM Gateway", "source": "FIO", "group": "Network Elements"},
    {"key": "OLT_UPLINK_PORT", "label": "OLT Uplink Port", "source": "EWP", "group": "Network Elements", "default": "1/1/1"},
    {"key": "OLT_LAG_ID", "label": "OLT LAG ID", "source": "EWP", "group": "Network Elements", "default": "10"},
    {"key": "OLT_SW_VERSION", "label": "OLT SW Version", "source": "EWP", "group": "Network Elements"},
    {"key": "OLT_SW_VERSION_SHORT", "label": "OLT SW Version (Short)", "source": "EWP", "group": "Network Elements"},
    {"key": "AN_SITE", "label": "AN Site", "source": "FIO", "group": "Network Elements"},
    {"key": "AN_PRODUCT", "label": "AN Product", "source": "EWP", "group": "Network Elements", "default": "ATN 980C"},
    {"key": "AN_PRODUCT_FAMILY", "label": "AN Product Family", "source": "EWP", "group": "Network Elements", "default": "ATN Series Equipment"},
    {"key": "AN_TRUNK_ID", "label": "AN Trunk ID", "source": "FIO", "group": "Network Elements"},
    {"key": "AN_UPLINK_PORT", "label": "AN Uplink Port", "source": "FIO", "group": "Network Elements"},
    {"key": "AN_LOOPBACK0", "label": "AN Loopback0", "source": "EWP", "group": "Network Elements"},
    {"key": "AN_LOOPBACK1", "label": "AN Loopback1", "source": "EWP", "group": "Network Elements"},
    {"key": "AG1_NODE", "label": "AG1 Node", "source": "FIO", "group": "Network Elements"},
    {"key": "AG2_NODE", "label": "AG2 Node", "source": "FIO", "group": "Network Elements"},
    {"key": "AG_SYSTEM", "label": "AG System", "source": "EWP", "group": "Network Elements", "default": "CX600(V8)"},
    {"key": "AG_PRODUCT", "label": "AG Product", "source": "EWP", "group": "Network Elements", "default": "CX600 Series Equipment"},
    {"key": "AG_MODEL", "label": "AG Model", "source": "EWP", "group": "Network Elements", "default": "CX600-X8(V8)"},
    {"key": "AG_SW_VERSION", "label": "AG SW Version", "source": "EWP", "group": "Network Elements"},
    {"key": "AG1_IP", "label": "AG1 IP", "source": "EWP", "group": "Network Elements"},
    {"key": "AG2_IP", "label": "AG2 IP", "source": "EWP", "group": "Network Elements"},
    {"key": "AG1_LOOPBACK0", "label": "AG1 Loopback0", "source": "EWP", "group": "Network Elements"},
    {"key": "AG1_LOOPBACK1", "label": "AG1 Loopback1", "source": "EWP", "group": "Network Elements"},
    {"key": "AG2_LOOPBACK0", "label": "AG2 Loopback0", "source": "EWP", "group": "Network Elements"},
    {"key": "AG2_LOOPBACK1", "label": "AG2 Loopback1", "source": "EWP", "group": "Network Elements"},

    # ============================================================
    # VLAN & VSI
    # ============================================================
    {"key": "VLAN_OM", "label": "OM VLAN", "source": "FIO", "group": "VLAN & VSI"},
    {"key": "VLAN_SIP", "label": "SIP VLAN", "source": "FIO", "group": "VLAN & VSI"},
    {"key": "VLAN_HSI", "label": "HSI VLAN", "source": "FIO", "group": "VLAN & VSI"},
    {"key": "VLAN_IPOE1", "label": "IPOE1 VLAN", "source": "FIO", "group": "VLAN & VSI"},
    {"key": "VLAN_IPOE2", "label": "IPOE2 VLAN", "source": "FIO", "group": "VLAN & VSI"},
    {"key": "VSI_OM", "label": "OM VSI Name", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VSI_SIP", "label": "SIP VSI Name", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VSI_HSI", "label": "HSI VSI Name", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VSI_IPOE1", "label": "IPOE1 VSI Name", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VSI_IPOE2", "label": "IPOE2 VSI Name", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "BNG_PEER_OM", "label": "BNG Peer IP (OM)", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "BNG_PEER_SIP", "label": "BNG Peer IP (SIP)", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "BNG_PEER_HSI", "label": "BNG Peer IP (HSI)", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "BNG_PEER_SIP_HSI", "label": "BNG Peer IP (SIP/HSI L2)", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "BNG_PEER_IPOE1", "label": "BNG Peer IP (IPOE1)", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "BNG_PEER_IPOE2", "label": "BNG Peer IP (IPOE2)", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_OM_PRIMARY", "label": "VCID OM Primary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_OM_SECONDARY", "label": "VCID OM Secondary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_SIP_PRIMARY", "label": "VCID SIP Primary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_SIP_SECONDARY", "label": "VCID SIP Secondary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_HSI_PRIMARY", "label": "VCID HSI Primary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_HSI_SECONDARY", "label": "VCID HSI Secondary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_IPOE1_PRIMARY", "label": "VCID IPOE1 Primary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_IPOE1_SECONDARY", "label": "VCID IPOE1 Secondary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_IPOE2_PRIMARY", "label": "VCID IPOE2 Primary", "source": "EWP", "group": "VLAN & VSI"},
    {"key": "VCID_IPOE2_SECONDARY", "label": "VCID IPOE2 Secondary", "source": "EWP", "group": "VLAN & VSI"},

    # ============================================================
    # NODE NAMES
    # ============================================================
    {"key": "AGG_NODE", "label": "AGG Node", "source": "FIO", "group": "Node Names"},
    {"key": "AGG_NODE_OM", "label": "AGG Node (OM)", "source": "FIO", "group": "Node Names"},
    {"key": "AGG_NODE_IPOE1", "label": "AGG Node (IPOE1)", "source": "FIO", "group": "Node Names"},
    {"key": "BNG_NODE_SIP_HSI", "label": "BNG Node (SIP/HSI)", "source": "FIO", "group": "Node Names"},
    {"key": "AGG_NODE_SIP_HSI_L2", "label": "AGG Node (SIP/HSI L2)", "source": "EWP", "group": "Node Names"},
    {"key": "BNG_NODE_IPOE2", "label": "BNG Node (IPOE2)", "source": "FIO", "group": "Node Names"},
    {"key": "OLT_REGION", "label": "OLT Region", "source": "FIO", "group": "Node Names", "default": "MIN"},
    {"key": "CORE_NODE_1", "label": "Core Node 1", "source": "EWP", "group": "Node Names"},
    {"key": "CORE_NODE_2", "label": "Core Node 2", "source": "EWP", "group": "Node Names"},
    {"key": "BNG_NODE", "label": "BNG Node", "source": "EWP", "group": "Node Names"},
    {"key": "CX600_NODE", "label": "CX600 Node", "source": "EWP", "group": "Node Names"},

    # ============================================================
    # CX600 PORTS
    # ============================================================
    {"key": "CX600_VE", "label": "CX600 Virtual-Ethernet", "source": "FIO", "group": "CX600 Ports"},
    {"key": "CX600_VE_L2", "label": "CX600 VE L2", "source": "FIO", "group": "CX600 Ports"},
    {"key": "CX600_VE_L3", "label": "CX600 VE L3", "source": "FIO", "group": "CX600 Ports"},
    {"key": "CX600_UPLINK_PORT", "label": "CX600 Uplink Port", "source": "FIO", "group": "CX600 Ports"},

    # ============================================================
    # IP & DHCP
    # ============================================================
    {"key": "OM_MSAN_VPN", "label": "OM MSAN VPN Instance", "source": "DOC", "group": "IP & DHCP", "default": "OM_MSAN"},
    {"key": "DHCP_GW_SIP", "label": "SIP DHCP Gateway", "source": "FIO", "group": "IP & DHCP"},
    {"key": "DHCP_BLOCK_SIP", "label": "SIP DHCP Block", "source": "FIO", "group": "IP & DHCP"},
    {"key": "DHCP_GW_OM", "label": "OM DHCP Gateway", "source": "FIO", "group": "IP & DHCP"},
    {"key": "DHCP_BLOCK_OM", "label": "OM DHCP Block", "source": "FIO", "group": "IP & DHCP"},
    {"key": "ALTIPLANO_SERVER_IP", "label": "Altiplano Server IP", "source": "DOC", "group": "IP & DHCP", "default": "10.203.214.40"},
    {"key": "DNS_IP", "label": "DNS IP", "source": "DOC", "group": "IP & DHCP", "default": "202.126.40.5"},
    {"key": "OLT_LT_IP", "label": "OLT LT IP", "source": "DOC", "group": "IP & DHCP", "default": "10.189.232.35"},
    {"key": "OLT_LT_PORT", "label": "OLT LT SSH Port", "source": "DOC", "group": "IP & DHCP", "default": "923"},
    {"key": "OLT_IHB_IP", "label": "OLT IHB IP", "source": "DOC", "group": "IP & DHCP", "default": "169.254.1.3"},
    {"key": "OLT_IHB_PORT", "label": "OLT IHB SSH Port", "source": "DOC", "group": "IP & DHCP", "default": "2222"},
    {"key": "OLT_LEMI_IP", "label": "OLT LEMI IP", "source": "DOC", "group": "IP & DHCP", "default": "169.254.0.1"},
    {"key": "OLT_LEMI_PORT", "label": "OLT LEMI Port", "source": "DOC", "group": "IP & DHCP", "default": "2222"},
    {"key": "LAPTOP_IP", "label": "Laptop IP", "source": "DOC", "group": "IP & DHCP", "default": "169.254.0.111/24"},
    {"key": "IHB_CLI_USER", "label": "IHB CLI User", "source": "DOC", "group": "IP & DHCP", "default": "cli_user@169.254.1.251"},

    # ============================================================
    # SERVICES & TEST
    # ============================================================
    {"key": "SERVICE_TYPE", "label": "Service Type", "source": "DOC", "group": "Services & Test", "default": "Wireline voice and data"},
    {"key": "PRODUCT_TYPE", "label": "Product Type", "source": "DOC", "group": "Services & Test", "default": "Voice and Broadband"},
    {"key": "MVNO_IR", "label": "MVNO/IR/Others", "source": "DOC", "group": "Services & Test", "default": "NA"},
    {"key": "TEST_BROWSE_URLS", "label": "Test Browse URLs", "source": "DOC", "group": "Services & Test", "default": "yahoo, google, facebook, speedtest.net"},
    {"key": "TEST_VIDEO_URL", "label": "Test Video URL", "source": "DOC", "group": "Services & Test", "default": "youtube"},

    # ============================================================
    # SITE & SYSTEM
    # ============================================================
    {"key": "CABINET_NAME", "label": "Cabinet Name", "source": "FIO", "group": "Site & System"},
    {"key": "CABINET_ENTRY", "label": "Cabinet Entry", "source": "EWP", "group": "Site & System", "default": "Outdoor"},
    {"key": "SITE_KEYS", "label": "Site Keys Required", "source": "DOC", "group": "Site & System", "default": "Yes"},
    {"key": "KNL_GENERATOR", "label": "KNL Generator Required", "source": "DOC", "group": "Site & System", "default": "Yes"},
    {"key": "E2E_SYSTEM_1", "label": "E2E System 1", "source": "EWP", "group": "Site & System", "default": "Lightspan MF-2"},
    {"key": "E2E_SYSTEM_2", "label": "E2E System 2", "source": "EWP", "group": "Site & System", "default": "CX600(V8)"},
    {"key": "E2E_SYSTEM_3", "label": "E2E System 3", "source": "EWP", "group": "Site & System", "default": "ATN 980C"},

    # ============================================================
    # REFERENCE DOCS
    # ============================================================
    {"key": "FIO_REF", "label": "FIO Reference", "source": "FIO", "group": "Reference Docs"},
    {"key": "FIO_DESCRIPTION", "label": "FIO Description", "source": "DOC", "group": "Reference Docs", "default": "Attach FIO for preconfigurations"},
    {"key": "TTX_FILE", "label": "TTx File", "source": "DOC", "group": "Reference Docs", "default": "TTx_FMC.xlsx"},
    {"key": "COMMISSIONING_SCRIPT", "label": "Commissioning Script", "source": "DOC", "group": "Reference Docs", "default": "Commissioning script for SW 24.6_up"},
    {"key": "ALTIPLANO_UPLOAD_DOC", "label": "Altiplano Upload Doc", "source": "DOC", "group": "Reference Docs", "default": "MF-2 NE Uploading via Altiplano Web GUI"},
    {"key": "DEVICE_CONFIG_DOC", "label": "Device Config Doc", "source": "DOC", "group": "Reference Docs", "default": "Device-config-mf Intent Configuration"},
    {"key": "UPLINK_CONFIG_DOC", "label": "Uplink Config Doc", "source": "DOC", "group": "Reference Docs", "default": "Uplink-connection Intent Configuration"},
    {"key": "L2_INFRA_DOC", "label": "L2 Infra Doc", "source": "DOC", "group": "Reference Docs", "default": "L2-infra Intent Configuration (VLAN)"},
    {"key": "FIBER_INTENT_DOC", "label": "Fiber Intent Doc", "source": "DOC", "group": "Reference Docs", "default": "Fiber Intent Configuration"},
    {"key": "ONT_INTENT_DOC", "label": "ONT Intent Doc", "source": "DOC", "group": "Reference Docs", "default": "ONT Intent Configuration"},
    {"key": "L2_USER_DOC", "label": "L2 User Doc", "source": "DOC", "group": "Reference Docs", "default": "L2-user Intent Configuration"},
    {"key": "POST_VALIDATION_TEMPLATE", "label": "Post Validation Template", "source": "DOC", "group": "Reference Docs", "default": "POST VALIDATION TEMPLATE"},
    {"key": "SO_REF", "label": "SO Reference", "source": "DOC", "group": "Reference Docs", "default": ""},
    {"key": "MOP_REF", "label": "MOP Reference", "source": "DOC", "group": "Reference Docs", "default": ""},

    # ============================================================
    # CREDENTIALS (VAULT)
    # ============================================================
    {"key": "OLT_ROOT_PW", "label": "OLT Root Password", "source": "VAULT", "group": "Credentials", "default": "<****>"},
    {"key": "IHB_CLI_PW", "label": "IHB CLI Password", "source": "VAULT", "group": "Credentials", "default": "<****cli>"},
]


# ============================================================
# HELPERS
# ============================================================
def get_fio_keys():
    return [p["key"] for p in PLACEHOLDER_MAP if p["source"] == "FIO"]


def get_ewp_keys():
    return [p["key"] for p in PLACEHOLDER_MAP if p["source"] == "EWP"]


def get_doc_keys():
    return [p["key"] for p in PLACEHOLDER_MAP if p["source"] == "DOC"]


def get_vault_keys():
    return [p["key"] for p in PLACEHOLDER_MAP if p["source"] == "VAULT"]


def get_groups():
    return sorted(set(p["group"] for p in PLACEHOLDER_MAP))


def get_by_group(group):
    return [p for p in PLACEHOLDER_MAP if p["group"] == group]


def get_all_keys():
    return [p["key"] for p in PLACEHOLDER_MAP]
