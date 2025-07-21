import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from fuzzywuzzy import fuzz

def truncate_after_sri_lanka(addr: str) -> str:
    part, sep, _ = addr.partition("Sri Lanka")
    return (part + sep).strip() if sep else addr.strip()

def extract_wo_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    customer = delivery = ""
    lines = text.split("\n")
    for i, ln in enumerate(lines):
        if "Deliver To:" in ln:
            customer = lines[i - 1].strip() if i > 0 else ""
            delivery = re.sub(r"Deliver To:\s*", "", ln).strip()
            break
    codes = re.findall(r"Product Code[:\s]*([A-Z]+\s*\d+(?:\s*/\s*[A-Z]+\s*\d+)*)", text)
    all_codes = [c.strip() for item in codes for c in item.split("/")]
    return {"customer_name": customer, "delivery_address": delivery, "product_codes": list(set(all_codes))}

def extract_po_fields(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    lines = [ln.strip() for ln in text.split("\n")]
    capture = False
    address_lines = []
    blank_count = 0

    for ln in lines:
        if "Delivery Location:" in ln:
            capture = True
            continue

        if capture:
            if "Forwarder:" in ln:
                break

            if not ln:
                blank_count += 1
                if blank_count >= 2:
                    break
                continue
            else:
                blank_count = 0

            address_lines.append(ln)

    raw_addr = " ".join(address_lines)
    matches = re.findall(r".*Sri Lanka.*", text, re.IGNORECASE)
    unique = [raw_addr] + [m for m in matches if m != raw_addr]
    seen = []
    for a in unique:
        if a and a not in seen:
            seen.append(a)

    sri = [a for a in seen if "sri lanka" in a.lower()]
    chosen = max(sri, key=len) if sri else seen[0] if seen else raw_addr
    final_addr = truncate_after_sri_lanka(chosen)

    po_codes = re.findall(r"(LB\s*\d+)", text)

    return {
        "delivery_location": final_addr,
        "product_codes": po_codes,
        "all_found_addresses": seen
    }

def extract_po_details(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    po_items = []

    for i, line in enumerate(lines):
        item_match = re.match(r'^(\d+)\s+([A-Z0-9]+)\s+(\d+)\s+(\d+\.\d+)\s+PCS', line)
        if item_match:
            item_no, item_code, _, qty_str = item_match.groups()
            quantity = int(float(qty_str))
            colour = size = ""
            style_2 = ""

            for j in range(i+1, min(i+10, len(lines))):
                ln = lines[j]

                if not colour and "Colour/Size/Destination:" in ln:
                    cs = ln.split(":", 1)[1].strip()
                    size_keywords = ["XS", "S", "M", "L", "XL", "XXL", "XXXL"]
                    parts = [p.strip() for p in cs.split("/") if p.strip()]

                    if parts and parts[0].upper() in size_keywords:
                        size = parts[0].upper()
                        if len(parts) > 1:
                            # Take the first word only (up to first space)
                            colour = parts[1].strip().split()[0].strip().upper()
                    else:
                        # fallback to previous logic
                        colour = re.match(r'^(\S+)', cs).group(1) if re.match(r'^(\S+)', cs) else ""
                        size_match = re.search(r'/\s*([^/]+)\s*/', cs)
                        if size_match:
                            size = size_match.group(1).strip()

                if not style_2:
                    match = re.search(r"(\d{6,})\s*$", ln)
                    if match:
                        style_2 = match.group(1)

            po_items.append({
                "Item_Number": item_no,
                "Item_Code": item_code,
                "Quantity": quantity,
                "Colour_Code": (colour or "").strip().upper(),
                "Size": (size or "").strip().upper(),
                "Style 2": style_2
            })

    return po_items



def extract_wo_items_table(pdf_file):
    import re
    items = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables() or []:
                for row in table or []:
                    if row and len(row) >= 6:
                        first = (row[0] or "").strip()
                        if re.match(r"^\d{8}$", first):
                            colour = (row[1] or "").strip().upper()
                            size_full = ""
                            for col in row[2:]:
                                if col and "/" in str(col):
                                    size_full = str(col).strip()
                                    break
                            qty = 0
                            for col in reversed(row):
                                if col and str(col).strip().isdigit():
                                    qty = int(str(col).strip())
                                    break
                            if qty > 0:
                                size_part = size_full.split("/")[0].strip().upper() if "/" in size_full else size_full.strip().upper()
                                items.append({
                                    "Style": first,
                                    "WO Colour Code": colour,
                                    "Size 1": size_part,
                                    "Quantity": qty
                                })
    return items

def enhanced_quantity_matching(wo_items, po_details, tolerance=0):
    matched, mismatched = [], []
    used = set()
    for wo in wo_items:
        wq = wo["Quantity"]
        ws = wo.get("Size 1", "").strip().upper()
        wc = wo.get("WO Colour Code", "").strip().upper()
        wstyle = wo.get("Style", "").strip()

        found_exact = False
        for idx, po in enumerate(po_details):
            if idx in used:
                continue
            pq = po["Quantity"]
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()

            if wq == pq and ws == ps and wc == pc and wstyle == pstyle:
                matched.append({
                    **{"Style": wstyle, "WO Size": ws, "PO Size": ps, "WO Colour Code": wc, "PO Colour Code": pc,
                       "WO Qty": wq, "PO Qty": pq, "Style 2": pstyle},
                    **{"Qty Match": "Yes", "Size Match": "Yes", "Colour Match": "Yes", "Style Match": "Yes",
                       "Diff": 0, "Status": "✅ Full Match", "PO Item Code": po.get("Item_Code", "")}
                })
                used.add(idx)
                found_exact = True
                break

        if found_exact:
            continue

        best_idx, best_diff = None, float("inf")
        for idx, po in enumerate(po_details):
            if idx in used:
                continue
            diff = abs(po["Quantity"] - wq)
            if diff <= tolerance and diff < best_diff:
                best_idx = idx
                best_diff = diff

        if best_idx is not None:
            po = po_details[best_idx]
            used.add(best_idx)
            pq = po["Quantity"]
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()

            qty_match = "Yes" if wq == pq else "No"
            size_match = "Yes" if ws == ps else "No"
            colour_match = "Yes" if wc == pc else "No"
            style_match = "Yes" if wstyle == pstyle else "No"

            full = all([qty_match == "Yes", size_match == "Yes", colour_match == "Yes", style_match == "Yes"])

            matched.append({
                **{"Style": wstyle, "WO Size": ws, "PO Size": ps, "WO Colour Code": wc, "PO Colour Code": pc,
                   "WO Qty": wq, "PO Qty": pq, "Style 2": pstyle},
                **{"Qty Match": qty_match, "Size Match": size_match, "Colour Match": colour_match, "Style Match": style_match,
                   "Diff": pq - wq,
                   "Status": "✅ Full Match" if full else "❌ Partial Match",
                   "PO Item Code": po.get("Item_Code", "")}
            })
        else:
            mismatched.append({
                **{"Style": wstyle, "WO Size": ws, "PO Size": "", "WO Colour Code": wc, "PO Colour Code": "",
                   "WO Qty": wq, "PO Qty": None, "Style 2": ""},
                **{"Qty Match": "No", "Size Match": "No", "Colour Match": "No", "Style Match": "No",
                   "Diff": "", "Status": "❌ No PO Qty", "PO Item Code": ""}
            })

    for idx, po in enumerate(po_details):
        if idx not in used:
            ps = po.get("Size", "").strip().upper()
            pc = po.get("Colour_Code", "").strip().upper()
            pstyle = po.get("Style 2", "").strip()
            mismatched.append({
                **{"Style": "N/A", "WO Size": "N/A", "PO Size": ps, "WO Colour Code": "N/A", "PO Colour Code": pc,
                   "WO Qty": None, "PO Qty": po["Quantity"], "Style 2": pstyle},
                **{"Qty Match": "No", "Size Match": "No", "Colour Match": "No", "Style Match": "No",
                   "Diff": "", "Status": "❌ Extra PO Item", "PO Item Code": po.get("Item_Code", "")}
            })

    return matched, mismatched

def compare_addresses(wo, po):
    ns = fuzz.token_sort_ratio(wo["customer_name"], po["delivery_location"])
    as_ = fuzz.token_sort_ratio(wo["delivery_address"], po["delivery_location"])
    comb = max(ns, as_)
    return {"WO Name": wo["customer_name"], "WO Addr": wo["delivery_address"], "PO Addr": po["delivery_location"],
            "Name %": ns, "Addr %": as_, "Overall %": comb, "Status": "✅ Match" if comb > 85 else "⚠️ Review"}

def compare_codes(wo_codes, po_codes):
    res = []
    for wc in wo_codes:
        found = False
        for pc in po_codes:
            pts = fuzz.ratio(wc, pc)
            res.append({
                "WO Code": wc,
                "PO Code": pc,
                "Match %": pts,
                "Status": "✅ Match" if pts == 100 else "⚠️ Partial" if pts > 80 else "❌ No"
            })
            if pts == 100:
                found = True
        if not found:
            res.append({
                "WO Code": wc,
                "PO Code": "",
                "Match %": "",
                "Status": "❌ WO code not found in PO"
            })

    has_mismatch = any(r["Status"].startswith("❌") for r in res)
    return res if has_mismatch else res[:max(1, len(res) // 2)]

# [all your import and function definitions remain the same, unchanged]
# ...
# --- Streamlit UI ---
st.set_page_config(page_title="WO ↔ PO Comparator", layout="wide")

st.title("📄 Customer Care System")
st.subheader("🔁 PO vs WO Comparison Dashboard")

with st.sidebar:
    st.header("⚙️ Settings")
    method = st.selectbox("Select Matching Method:",
                          ["Enhanced Matching (with PO Color/Size)", "Smart Matching (Exact)", "Smart Matching with Tolerance"])
    wo_file = st.file_uploader("📤 Upload WO PDF", type="pdf")
    po_file = st.file_uploader("📤 Upload PO PDF", type="pdf")

if wo_file and po_file:
    with st.spinner("🔄 Processing files..."):
        wo = extract_wo_fields(wo_file)
        po = extract_po_fields(po_file)
        wo_items = extract_wo_items_table(wo_file)
        po_details = extract_po_details(po_file)
        addr_res = compare_addresses(wo, po)
        code_res = compare_codes(wo["product_codes"], po["product_codes"])

        if "Enhanced" in method:
            matched, mismatched = enhanced_quantity_matching(wo_items, po_details)
        else:
            matched, mismatched = [], []

    st.success("✅ Comparison Completed")

    st.markdown("---")
    st.subheader("📍 Address Comparison")
    st.dataframe(pd.DataFrame([addr_res]), use_container_width=True)

    st.subheader("🔢 Product Code Comparison")
    st.dataframe(pd.DataFrame(code_res), use_container_width=True)

    st.subheader("✅ Matched WO/PO Items")
    if matched:
        st.dataframe(pd.DataFrame(matched), use_container_width=True)
    else:
        st.info("No matched items found or matching method not selected.")

    # ✅ Updated Match Check Logic
    address_ok = addr_res.get("Overall %", 0) == 100
    codes_df = pd.DataFrame(code_res)
    codes_ok = not codes_df.empty and all(codes_df["Match %"] == 100)
    matched_df = pd.DataFrame(matched)
    status_ok = not matched_df.empty and all(matched_df["Status"] == "✅ Full Match")

    if address_ok and codes_ok and status_ok:
        st.success("🎉 All items are FULLY MATCHED between PO and WO!")

    st.subheader("❗ Mismatched or Extra Items")
    if mismatched:
        st.dataframe(pd.DataFrame(mismatched), use_container_width=True)
    else:
        st.success("No mismatched or extra PO/WO items found.")

    st.subheader("🧾 Work Order (WO) Items Table")
    st.dataframe(pd.DataFrame(wo_items), use_container_width=True)

    st.subheader("📦 Purchase Order (PO) Details")
    st.dataframe(pd.DataFrame(po_details), use_container_width=True)

st.markdown("<br><hr><center><b style='color:#888'>Created by Razz... </b></center>",
            unsafe_allow_html=True)