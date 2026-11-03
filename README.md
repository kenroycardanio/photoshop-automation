# Photoshop CS6 — Multi-ID print sheet automation

Fill **4 or 5 student IDs on one PSD** from a **CSV** (default). Excel (`.xlsx`) is optional.

**LRN tip:** in Google Sheets / Excel set the LRN column to **Plain text** before download/export, so digits don’t become `1.27E+11`.

If you pick `.xlsx` and Excel isn’t installed, Python + `excel_to_rows.py` is used as a fallback.

## PSD structure (recommended)

Use **one group per ID slot**. Same layer names inside each group:

```
📁 id1                    📁 id2                    📁 id3 …
   📝 lrn                     📝 lrn
   📝 gender                  📝 gender
   📝 dob                     📝 dob
   📝 validity                📝 validity
   📝 studentname             📝 studentname
   📝 guardian                📝 guardian
   📝 contact                 📝 contact
   📝 address                 📝 address
   🖼 smart (later)           🖼 smart
```

Groups: **`id1`**, **`id2`**, **`id3`**, **`id4`**, **`id5`** (lowercase).

## How rows map to groups

| You run (start → end) | CSV rows | Groups filled |
|----------------------|----------|---------------|
| 1 → 5 | rows 1–5 | `id1` … `id5` |
| 6 → 10 | rows 6–10 | `id1` … `id5` again (next sheet) |
| 1 → 4 | rows 1–4 | `id1` … `id4` only (4-up) |

## Layer names inside each group

| Field | Layer name | CSV column |
|-------|------------|------------|
| LRN | `lrn` | Learner's Reference Number (LRN) |
| Gender | `gender` | Gender |
| DOB | `dob` | Birthday |
| Validity | `validity` | ID VALIDITY |
| Student name | `studentname` | Full Name | ≤15=100% · ~21ch (Kreanna Maen G. Alesna)=**80%** |
| Guardian | `guardian` | IN CASE OF EMERGENCY PARENTS/ GUARDIAN | ≤15 chars 100% · longer names **80%** horizontal scale |
| Contact | `contact` | Contact Number |
| Address | `address` | Address | Point text — tracking first |
| Student photo | `smart` | **picture path** (column AB) | Fits inside PSD `smart` layer box |
| QR code | `qr` | **qr path** (column AC) | Fits inside PSD `qr` layer box |

## Image fit (smart / qr)

The **PSD layer is the frame**. Before placing, the script reads `smart` or `qr` **position and size** from your template, then:

1. Places the local image file  
2. Scales to **fit inside** that box (centered, no cropping)  
3. Keeps the same layer name and stack position  

Tip: size your `smart` layer in the PSD exactly how you want the photo slot (e.g. 2×2). If the layer was clipped to a shape, the clip is kept.

**Rule:** group or layer missing → skipped, no error.

**Fallback:** if no `id1` group, script tries old flat names like `lrn1`, `gender1`.

## How to run

1. Open your 4-up or 5-up template PSD.
2. **File → Scripts → Browse…** → `batch_fill.jsx`
3. Pick CSV, set start/end row (e.g. 1–5 or 1–4).
4. Check sheet → Save → next batch.

See `Photoshop_Layer_Codes.csv` for the full reference table.

---

## NFT client — `back_fill_NFT.jsx`

Same workflow (CSV + start/end row → `id1`…`id5`). Sheet: **NFT CARDS**.

| Layer (in `id1`) | Layer (in `id2`) | CSV column |
|------------------|------------------|------------|
| *(ignored)* | | **`#`** (row number, column 1) |
| `cardnumber` | `cardnumber` | **Card Number** (column 2) |
| `qrcardnumber` | `qrcardnumber` | QR Card Number |
| `name` | `name` | NAME |
| `membersince` | `membersince` | MEMBER SINCE |
| `validuntil` | `validuntil` | Valid Until |
| `smartA1` | `smartA2` | **PictureA** |
| `smartB1` | `smartB2` | **PictureB** |
| `smartC1`…`E1` | `smartC2`…`E2` | PictureC…E |
| `dataA1` | `dataA2` | **DataA** |
| `dataB1`…`J1` | `dataB2`…`J2` | DataB…J |

Columns are matched **by header name**, not position — so `#` as column 1 is fine; Card Number stays correct as column 2.
