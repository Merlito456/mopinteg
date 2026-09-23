"""
Gemini prompt template and response parser.
"""
import re
import json


GEMINI_PROMPT_TEMPLATE = """You are an expert network engineer reading a Facility Implementation Order (FIO) and an Engineering Work Plan (EWP) image for a Nokia Lightspan MF-2 OLT integration project.

TASK:
Extract ALL of the following data points and return them as a single valid JSON object.

FIELDS TO EXTRACT:

{
  "CIRCUIT_ID": "B2C VMDU FTH",
  "OLT_PRODUCT": "Lightspan MF-2",
  "OLT_SITE": "e.g., CDO_013_GPONA_02",
  "OLT_MGMT_IP": "e.g., 10.168.196.226",
  "OLT_OM_VLAN": "e.g., 734",
  "OLT_OM_GW": "e.g., 10.168.196.193",
  "OLT_UPLINK_PORT": "e.g., 1/1/1",
  "OLT_LAG_ID": "e.g., 10",
  "OLT_SW_VERSION": "e.g., L6GQFC24.298",
  "OLT_SW_VERSION_SHORT": "e.g., 24.6",
  "AN_SITE": "e.g., LCGCDO_TDCAN_T980C_01_MIN355",
  "AN_PRODUCT": "e.g., ATN 980C",
  "AN_PRODUCT_FAMILY": "e.g., ATN Series Equipment",
  "AN_TRUNK_ID": "e.g., Eth-Trunk5",
  "AN_UPLINK_PORT": "e.g., GE0/4/2",
  "AN_LOOPBACK0": "e.g., 10.175.4.57",
  "AN_LOOPBACK1": "e.g., 10.175.8.57",
  "AG1_NODE": "e.g., CAGAYA-MIN812-TRS_DC_AG-R1-01",
  "AG2_NODE": "e.g., LAPASA-MIN836-TRS_DC_AG-R1-01",
  "AG_SYSTEM": "e.g., CX600(V8)",
  "AG_PRODUCT": "e.g., CX600 Series Equipment",
  "AG_MODEL": "e.g., CX600-X8(V8)",
  "AG_SW_VERSION": "e.g., V600R008C10",
  "AG1_IP": "e.g., 10.175.0.9",
  "AG2_IP": "e.g., 10.175.0.5",
  "AG1_LOOPBACK0": "e.g., 10.175.0.9",
  "AG1_LOOPBACK1": "e.g., 10.175.1.9",
  "AG2_LOOPBACK0": "e.g., 10.175.0.5",
  "AG2_LOOPBACK1": "e.g., 10.175.1.5",
  "VLAN_OM": "e.g., 734",
  "VLAN_SIP": "e.g., 2109",
  "VLAN_HSI": "e.g., 2112",
  "VLAN_IPOE1": "e.g., 3013",
  "VLAN_IPOE2": "e.g., 1268",
  "VSI_OM": "e.g., OM-90000734",
  "VSI_SIP": "e.g., SIP-90002109",
  "VSI_HSI": "e.g., HSI-90002112",
  "VSI_IPOE1": "e.g., IPOE-193013368",
  "VSI_IPOE2": "e.g., IPOE-291268368",
  "BNG_PEER_OM": "e.g., 10.175.0.78",
  "BNG_PEER_SIP": "e.g., 10.168.56.240",
  "BNG_PEER_HSI": "e.g., 10.168.56.240",
  "BNG_PEER_SIP_HSI": "e.g., 10.168.56.240",
  "BNG_PEER_IPOE1": "e.g., 10.175.0.77",
  "BNG_PEER_IPOE2": "e.g., 10.168.56.242",
  "VCID_OM_PRIMARY": "e.g., 25011194",
  "VCID_OM_SECONDARY": "e.g., 26011194",
  "VCID_SIP_PRIMARY": "e.g., 25011195",
  "VCID_SIP_SECONDARY": "e.g., 26011195",
  "VCID_HSI_PRIMARY": "e.g., 25011196",
  "VCID_HSI_SECONDARY": "e.g., 26011196",
  "VCID_IPOE1_PRIMARY": "e.g., 2430132705",
  "VCID_IPOE1_SECONDARY": "e.g., 2930132705",
  "VCID_IPOE2_PRIMARY": "e.g., 3412682705",
  "VCID_IPOE2_SECONDARY": "e.g., 3912682705",
  "AGG_NODE": "e.g., CDO-MIN1092-AGGFMN-X16A-01",
  "AGG_NODE_OM": "e.g., CDO-MIN1092-AGGFMN-X16A-02",
  "AGG_NODE_IPOE1": "e.g., CDO-MIN1092-AGGFMN-X16A-01",
  "AGG_NODE_SIP_HSI_L2": "e.g., MINALWANA-FMAGG-002-MIN1092",
  "BNG_NODE_SIP_HSI": "e.g., ALWANA-FMAGG-001-MIN1092",
  "BNG_NODE_IPOE2": "e.g., DAVCLS-FMAGG-001-MIN2031",
  "OLT_REGION": "e.g., MIN",
  "CORE_NODE_1": "e.g., CDO-PS-PTR-002",
  "CORE_NODE_2": "e.g., CDO-PS-PTR-001/002",
  "CX600_NODE": "e.g., LAPASA-MIN836-TRS_DC_AG-R1-01",
  "CX600_VE": "e.g., 3/0/0",
  "CX600_VE_L2": "e.g., 1/0/0",
  "CX600_VE_L3": "e.g., 1/0/1",
  "CX600_UPLINK_PORT": "e.g., 100GE1/1/16",
  "DHCP_BLOCK_SIP": "e.g., 10.202.192.0/19",
  "DHCP_GW_SIP": "e.g., 10.202.192.1",
  "DHCP_BLOCK_OM": "e.g., 10.168.196.192/26",
  "DHCP_GW_OM": "e.g., 10.168.196.193",
  "CABINET_ENTRY": "e.g., Outdoor",
  "CABINET_NAME": "e.g., CDO_013_GPONA_02",
  "E2E_SYSTEM_1": "e.g., Lightspan MF-2",
  "E2E_SYSTEM_2": "e.g., CX600(V8)",
  "E2E_SYSTEM_3": "e.g., ATN 980C",
  "FIO_REF": "e.g., FIO-LCGCDONG01FTTx_FMC"
}

INSTRUCTIONS:
1. Extract each field from the FIO Excel and the EWP image.
2. If a value is not found, set it to null.
3. Return ONLY a valid JSON object — no markdown fences, no explanation.
4. For node names, use the exact format from the source (hyphens, underscores preserved).
5. For VSI names, use the exact pattern (e.g., OM-90000734, IPOE-193013368).
6. For IPs, use the exact format found.

FILES ATTACHED:
- FIO: <paste FIO data or upload file>
- EWP: <upload EWP image>

OUTPUT: A single JSON object with the 80+ fields above.
"""


def parse_ai_response(response_text: str) -> dict:
    """
    Parse Gemini's response into a dict.
    Handles:
    - Pure JSON
    - JSON inside ```json ... ``` fences
    - JSON inside ``` ... ``` fences
    - JSON with leading/trailing text
    """
    if not response_text or not response_text.strip():
        raise ValueError("Empty response")

    text = response_text.strip()

    # Strip markdown code fences
    fence_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if fence_match:
        text = fence_match.group(1)

    # Try direct JSON parse
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        # Try to find the first { ... last }
        start = text.find("{")
        end = text.rfind("}")
        if start == -1 or end == -1 or end <= start:
            raise ValueError("No JSON object found in response")
        json_str = text[start:end + 1]
        try:
            data = json.loads(json_str)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON: {e}")

    # Filter out None values
    data = {k: v for k, v in data.items() if v is not None and v != ""}

    # Ensure all values are strings
    data = {k: str(v) for k, v in data.items()}

    return data
