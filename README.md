# Project Search & Brochure Portal (v2)

Loaded from your own compiled research: `LoopNet_Waller_Tomball_Navasota_
Bridgeland_LoopNet_Brochures.xlsx`, covering **72 properties** across
Waller, Tomball, Navasota, and Bridgeland/Cypress, TX.

## Run it

```bash
pip install -r requirements.txt
streamlit run app.py
```

## What's in the data

- `data/projects.csv` — one row per property: area, for-sale/for-lease
  status, address, type, size, price/rent, year built, and the LoopNet
  listing URL.
- `data/attachments.csv` — one row per attachment. **13 real, direct PDF
  links** (brochures, surveys, plats, floor plans, flyers, etc.) were
  captured for 7 of the 72 properties in your source file — the rest
  don't have a brochure URL exposed in what you compiled, so they show
  the property details and a link to the LoopNet listing itself instead.

## How search works

Type any city, zip code, area name, or property type into the search box,
or use the **Area** / **Status** (For Sale / For Lease) dropdowns to
narrow results. Every result shows full property details plus:

- **All captured attachments**, each with its own download button (fetched
  live from LoopNet's own asset CDN URL at click time — these are the
  exact links from your workbook, not re-scraped)
- A link back to the live LoopNet listing/search page
- A checkbox to select it for the email panel at the bottom, which
  attaches every PDF for the properties you pick and sends via your SMTP
  settings

## Extending the data

To add more properties or attachments later, edit the two CSVs directly,
or re-run `transform.py` against an updated version of your Excel
workbook — it's a straightforward pandas script matching the exact column
names LoopNet's export/research format uses (`Area`, `Status`, `Property`,
`Address`, `Property Type`, `Size`, `Price/Rent`, `Year Built`, `LoopNet
URL`, `Actual PDF Brochure / Attachment Links`), so you can rerun it
whenever you compile a new batch of listings yourself.
