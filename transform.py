import re
import pandas as pd

SRC = "/mnt/user-data/uploads/LoopNet_Waller_Tomball_Navasota_Bridgeland_LoopNet_Brochures.xlsx"

df = pd.read_excel(SRC, sheet_name="Properties")

def parse_city_state_zip(addr):
    if not isinstance(addr, str):
        return "", "", ""
    m = re.match(r"^(.*),\s*([A-Z]{2})\s*(\d{5})?\s*$", addr.strip())
    if not m:
        return addr.strip(), "", ""
    city, state, zip_ = m.group(1).strip(), m.group(2).strip(), (m.group(3) or "").strip()
    return city, state, zip_

def parse_attachments(cell):
    """'Actual PDF Brochure / Attachment Links' cell -> list of (label, url)."""
    if not isinstance(cell, str) or not cell.strip():
        return []
    pairs = []
    # Entries are separated by literal \n (real newlines after excel read)
    for line in cell.split("\n"):
        line = line.strip()
        if not line:
            continue
        if ":" in line:
            label, url = line.split(":", 1)
            # rejoin in case the URL itself contained an early colon (http://) -
            # split(":",1) handles that fine since label has no colon
            label = label.strip()
            url = url.strip()
            if url.startswith("http"):
                pairs.append((label, url))
            else:
                # label contained "http" oddly split; try regex fallback
                m = re.search(r"(https?://\S+)", line)
                if m:
                    pairs.append((line[: m.start()].strip(" :"), m.group(1)))
        else:
            m = re.search(r"(https?://\S+)", line)
            if m:
                pairs.append(("Attachment", m.group(1)))
    return pairs

projects_rows = []
attachments_rows = []

for i, row in df.iterrows():
    pid = f"P{i+1:04d}"
    city, state, zip_ = parse_city_state_zip(row.get("Address"))
    street = str(row.get("Property") or "").strip()
    full_address = street

    year_built = row.get("Year Built")
    year_built = "" if pd.isna(year_built) else str(int(year_built))

    projects_rows.append({
        "project_id": pid,
        "area": row.get("Area") or "",
        "status": row.get("Status") or "",
        "project_name": street or f"Property {pid}",
        "address": full_address,
        "city": city,
        "state": state,
        "zipcode": zip_,
        "property_type": row.get("Property Type") or "",
        "size": row.get("Size") or "",
        "price": row.get("Price/Rent") or "",
        "year_built": year_built,
        "loopnet_url": row.get("LoopNet URL") or "",
        "description": row.get("Notes") or "",
    })

    for label, url in parse_attachments(row.get("Actual PDF Brochure / Attachment Links")):
        attachments_rows.append({"project_id": pid, "label": label, "url": url})

projects_df = pd.DataFrame(projects_rows)
attachments_df = pd.DataFrame(attachments_rows)

projects_df.to_csv("/home/claude/loopnet_ui2/projects.csv", index=False)
attachments_df.to_csv("/home/claude/loopnet_ui2/attachments.csv", index=False)

print("projects:", len(projects_df), "attachments:", len(attachments_df))
print(attachments_df.head(20).to_string())
