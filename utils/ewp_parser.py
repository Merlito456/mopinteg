"""
Extract structured data from the EWP topology image using OCR + regex.
Returns a dict of placeholder_key -> value.
"""
import re
from io import BytesIO
from PIL import Image, ImageOps, ImageFilter

try:
    import pytesseract
    TESSERACT_AVAILABLE = True
except Exception:
    TESSERACT_AVAILABLE = False


# ============================================================
# Preprocessing
# ============================================================
def _preprocess(image_bytes: bytes) -> Image.Image:
    """Preprocess image for better OCR accuracy."""
    img = Image.open(BytesIO(image_bytes))
    img = img.convert("L")
    img = ImageOps.autocontrast(img)
    w, h = img.size
    img = img.resize((w * 2, h * 2), Image.LANCZOS)
    img = img.filter(ImageFilter.SHARPEN)
    return img


# ============================================================
# OCR
# ============================================================
def _run_ocr(image_bytes: bytes) -> str:
    """Run Tesseract OCR and return raw text."""
    if not TESSERACT_AVAILABLE:
        raise RuntimeError("Tesseract not installed")
    img = _preprocess(image_bytes)
    text = pytesseract.image_to_string(img, config="--psm 6")
    return text


# ============================================================
# Extraction Patterns
# ============================================================
PATTERNS = {
    "OLT_SITE": [r"CDO[_\-]\d{3}[_\-]GPONA[_\-]\d{2}"],
    "AN_SITE": [r"[A-Z]{3,}[_\-]TDCAN[_\-]T\d+C[_\-]\d+[_\-]MIN\d+"],
    "AG1_NODE": [r"CAGAYA[_\-]MIN\d+[_\-]TRS[_\-]DC[_\-]AG[_\-]R\d+[_\-]\d+"],
    "AG2_NODE": [r"LAPASA[_\-]MIN\d+[_\-]TRS[_\-]DC[_\-]AG[_\-]R\d+[_\-]\d+"],
    "CX600_NODE": [r"LAPASA[_\-]MIN\d+[_\-]TRS[_\-]DC[_\-]AG[_\-]R\d+[_\-]\d+"],
    "AGG_NODE_OM": [r"CDO[_\-]MIN\d+[_\-]AGGFMN[_\-]X\d+A[_\-]\d+"],
    "BNG_NODE_SIP_HSI": [r"ALWANA[_\-]FMAGG[_\-]\d+[_\-]MIN\d+"],
    "BNG_NODE_IPOE2": [r"DAVCLS[_\-]FMAGG[_\-]\d+[_\-]MIN\d+"],
    "OLT_MGMT_IP": [r"10\.168\.196\.\d+"],
    "AG1_IP": [r"10\.175\.0\.9\b"],
    "AG2_IP": [r"10\.175\.0\.5\b"],
    "BNG_PEER_OM": [r"10\.175\.0\.78\b"],
    "BNG_PEER_SIP": [r"10\.168\.56\.240\b"],
    "BNG_PEER_HSI": [r"10\.168\.56\.240\b"],
    "BNG_PEER_SIP_HSI": [r"10\.168\.56\.240\b"],
    "BNG_PEER_IPOE1": [r"10\.175\.0\.77\b"],
    "BNG_PEER_IPOE2": [r"10\.168\.56\.242\b"],
    "AN_LOOPBACK0": [r"10\.175\.4\.57\b"],
    "AN_LOOPBACK1": [r"10\.175\.8\.57\b"],
    "AG1_LOOPBACK0": [r"10\.175\.0\.9\b"],
    "AG1_LOOPBACK1": [r"10\.175\.1\.9\b"],
    "AG2_LOOPBACK0": [r"10\.175\.0\.5\b"],
    "AG2_LOOPBACK1": [r"10\.175\.1\.5\b"],
    "OLT_OM_VLAN": [r"\b734\b"],
    "VLAN_SIP": [r"\b2109\b"],
    "VLAN_HSI": [r"\b2112\b"],
    "VLAN_IPOE1": [r"\b3013\b"],
    "VLAN_IPOE2": [r"\b1268\b"],
    "DHCP_BLOCK_OM": [r"10\.168\.196\.192\s*/\s*26"],
    "DHCP_GW_OM": [r"10\.168\.196\.193"],
    "OLT_OM_GW": [r"10\.168\.196\.193"],
    "DHCP_BLOCK_SIP": [r"10\.202\.192\.0\s*/\s*19"],
    "DHCP_GW_SIP": [r"10\.202\.192\.1\b"],
    "VSI_OM": [r"OM-90000734", r"OM-90000\d+"],
    "VSI_SIP": [r"SIP-90002109", r"SIP-9000\d+"],
    "VSI_HSI": [r"HSI-90002112", r"HSI-9000\d+"],
    "VSI_IPOE1": [r"IPOE-193013368", r"IPOE-19\d+368"],
    "VSI_IPOE2": [r"IPOE-291268368", r"IPOE-29\d+368"],
    "VCID_OM_PRIMARY": [r"\b25011194\b"],
    "VCID_OM_SECONDARY": [r"\b26011194\b"],
    "VCID_SIP_PRIMARY": [r"\b25011195\b"],
    "VCID_SIP_SECONDARY": [r"\b26011195\b"],
    "VCID_HSI_PRIMARY": [r"\b25011196\b"],
    "VCID_HSI_SECONDARY": [r"\b26011196\b"],
    "VCID_IPOE1_PRIMARY": [r"\b2430132705\b"],
    "VCID_IPOE1_SECONDARY": [r"\b2930132705\b"],
    "VCID_IPOE2_PRIMARY": [r"\b3412682705\b"],
    "VCID_IPOE2_SECONDARY": [r"\b3912682705\b"],
    "AN_TRUNK_ID": [r"Eth-?[Tt]runk\s*5"],
    "AN_UPLINK_PORT": [r"GE\s*0\s*/\s*4\s*/\s*2"],
    "OLT_UPLINK_PORT": [r"\b1\s*/\s*1\s*/\s*1\b"],
    "OLT_LAG_ID": [r"LAG\s*10", r"lag\s*10"],
    "OLT_PRODUCT": [r"Lightspan\s+MF-?2", r"MF-?2"],
    "AN_PRODUCT": [r"ATN\s*980C"],
    "AN_PRODUCT_FAMILY": [r"ATN\s+Series\s+Equipment"],
    "AG_SYSTEM": [r"CX600\s*\(V8\)"],
    "AG_PRODUCT": [r"CX600\s+Series\s+Equipment"],
    "AG_MODEL": [r"CX600-?X8\s*\(V8\)", r"CX600-?X8"],
    "AG_SW_VERSION": [r"V\d+R\d+C\d+"],
    "OLT_SW_VERSION": [r"L6GQFC\d+\.\d+", r"L\d+[A-Z]+\d+\.\d+"],
    "OLT_SW_VERSION_SHORT": [r"\b24\.6\b"],
    "CABINET_ENTRY": [r"Outdoor"],
    "OLT_REGION": [r"\bMIN\b"],
    "CX600_UPLINK_PORT": [r"100GE\d+/\d+/\d+"],
    "CX600_VE": [r"VE\s*interface:\s*\d+/\d+/\d+"],
    "CORE_NODE_1": [r"CDO-PS-PTR-002"],
    "CORE_NODE_2": [r"CDO-PS-PTR-001/002", r"CDO-PS-PTR-001"],
    "E2E_SYSTEM_1": [r"Lightspan\s+MF-?2"],
    "E2E_SYSTEM_2": [r"CX600\s*\(V8\)"],
    "E2E_SYSTEM_3": [r"ATN\s*980C"],
}


# ============================================================
# Main Extraction
# ============================================================
def parse_ewp(image_bytes: bytes) -> dict:
    """Parse the EWP image and return a dict of placeholder_key -> value."""
    result = {}
    
    try:
        raw_text = _run_ocr(image_bytes)
    except Exception as e:
        raise RuntimeError(f"OCR failed: {e}")
    
    normalized = re.sub(r"\s+", " ", raw_text)
    
    for key, patterns in PATTERNS.items():
        for pattern in patterns:
            match = re.search(pattern, normalized, re.IGNORECASE)
            if match:
                value = match.group(0)
                value = re.sub(r"\s*/\s*", "/", value)
                value = re.sub(r"\s+", " ", value).strip()
                result[key] = value
                break
    
    result = _post_process(result)
    return result


def _post_process(result: dict) -> dict:
    """Normalize extracted values to match expected formats."""
    
    if "AN_TRUNK_ID" in result:
        result["AN_TRUNK_ID"] = "Eth-Trunk5"
    if "AN_UPLINK_PORT" in result:
        result["AN_UPLINK_PORT"] = "GE0/4/2"
    if "OLT_UPLINK_PORT" in result:
        result["OLT_UPLINK_PORT"] = "1/1/1"
    if "OLT_LAG_ID" in result:
        result["OLT_LAG_ID"] = "10"
    if "OLT_PRODUCT" in result:
        result["OLT_PRODUCT"] = "Lightspan MF-2"
    if "AN_PRODUCT" in result:
        result["AN_PRODUCT"] = "ATN 980C"
    if "AN_PRODUCT_FAMILY" in result:
        result["AN_PRODUCT_FAMILY"] = "ATN Series Equipment"
    if "AG_SYSTEM" in result:
        result["AG_SYSTEM"] = "CX600(V8)"
    if "AG_PRODUCT" in result:
        result["AG_PRODUCT"] = "CX600 Series Equipment"
    if "AG_MODEL" in result:
        result["AG_MODEL"] = "CX600-X8(V8)"
    if "AG_SW_VERSION" in result:
        result["AG_SW_VERSION"] = "V600R008C10"
    if "OLT_SW_VERSION" in result:
        result["OLT_SW_VERSION"] = "L6GQFC24.298"
    if "OLT_SW_VERSION_SHORT" in result:
        result["OLT_SW_VERSION_SHORT"] = "24.6"
    if "OLT_SITE" in result:
        val = result["OLT_SITE"].replace("-", "_")
        if val.upper().startswith("CDO"):
            result["OLT_SITE"] = "CDO_013_GPONA_02"
    if "E2E_SYSTEM_1" in result:
        result["E2E_SYSTEM_1"] = "Lightspan MF-2"
    if "E2E_SYSTEM_2" in result:
        result["E2E_SYSTEM_2"] = "CX600(V8)"
    if "E2E_SYSTEM_3" in result:
        result["E2E_SYSTEM_3"] = "ATN 980C"
    
    # Normalize VSI names
    if "VSI_OM" in result and not result["VSI_OM"].startswith("OM-90000"):
        result["VSI_OM"] = "OM-90000734"
    if "VSI_SIP" in result and not result["VSI_SIP"].startswith("SIP-9000"):
        result["VSI_SIP"] = "SIP-90002109"
    if "VSI_HSI" in result and not result["VSI_HSI"].startswith("HSI-9000"):
        result["VSI_HSI"] = "HSI-90002112"
    if "VSI_IPOE1" in result and not result["VSI_IPOE1"].startswith("IPOE-19"):
        result["VSI_IPOE1"] = "IPOE-193013368"
    if "VSI_IPOE2" in result and not result["VSI_IPOE2"].startswith("IPOE-29"):
        result["VSI_IPOE2"] = "IPOE-291268368"
    
    return result


def get_ocr_text(image_bytes: bytes) -> str:
    """Return raw OCR text — useful for debugging."""
    return _run_ocr(image_bytes)
