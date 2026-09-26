Negative Stock Adjustment Tool
Streamlit app for preparing iTrade negative-stock adjustments (bundle break,
wrong sale) for Shams Al Madina Hypermarket.
Run locally
    pip install -r requirements.txt
    streamlit run app.py

Deploy on Streamlit Community Cloud
Push to a private GitHub repo.
share.streamlit.io -> New app -> main file `app.py`.
Settings -> Sharing -> private, add only the people who should see it.
The default is public to anyone with the link.
Saving runs and sheets (optional)
MongoDB Atlas M0 is free and gives 512 MB. Create a cluster and a database
user, then add under Settings -> Secrets:
    [mongo]
    uri = "mongodb+srv://user:password@cluster0.xxxxx.mongodb.net/?retryWrites=true&w=majority"
    db  = "stockadj"

Atlas -> Network Access must allow 0.0.0.0/0, because Streamlit Cloud has no
fixed outbound IP. The password is then the only barrier: make it long, and
give that user access to this one database only.
With it configured:
every matching run is saved automatically (settings, candidates, skipped combos)
generated sheets and import files are saved when you press Save on Build sheets
the Archive tab lists both, and any past file can be downloaded again
Without it the app still works, but nothing is kept after you close it.
Item templates page
The three request sheets that go to the purchaser: creation, activation
and description update. Pick the type, fill the header, type the items into
the grid, press Generate. Out comes a PDF in the same layout as the sheets
already in use, and a copy goes into the database.
GP% is never typed. It is worked out as `(RSP - COST) / RSP * 100`, on the
raw figures with no VAT adjustment, and shown live under the grid before you
generate. That is the one number on these sheets people get wrong.
Blank rows are dropped, so you can leave spare lines in the grid.
On a creation sheet an empty barcode prints as NEED BARCODE, which is what
the existing sheets do — a gap looks like a mistake.
Warnings, not blocks: a loss-making RSP or an empty cost is flagged, but the
sheet still generates. A creation with no barcode yet is normal.
Each sheet gets a reference — `CRE-260926-1`, `ACT-…`, `DSC-…` — counted from
what is already in the database, so numbering survives a reboot.
Saved sheets lists every one. Any of them can be downloaded again, or
loaded back into the grid to be corrected and reissued.
Signatory defaults (Prepared by / Concerned purchaser / Approved by) come from
the Settings page.
Using your own Excel template
The built-in layout is a rebuild from screenshots — close, but not your file.
On the Layouts tab you can upload the real `.xlsx` for any of the three
types. The app then writes rows straight into that workbook, so the logo,
column widths and formatting are yours, and you get the workbook back plus a
PDF of it.
What it works out for itself, and shows you to correct before saving:
which row holds the column headings, and which column is which field
where DATE / VENDOR / Main Grp / REASON / Remark go
where the three signature lines are
how many ruled rows the table has room for — counted from the ruling,
not the blank run, so overflow rows never land on the signature block
Nothing is saved until you press Save, and a Test it with sample rows
button gives you the filled sheet to look at first. If two fields are mapped
to the same column, or COST / RSP / Description are unmapped, saving is
blocked rather than producing a sheet that looks right and says the wrong
thing.
Notes:
Barcodes are written as text, so a 13-digit code keeps its leading zero
instead of becoming `1.00071E+10`.
If the GP% column in your template already holds a formula, that formula is
copied down rather than overwritten with a number.
Upload a blank copy. Any rows already in it are treated as the space
your items go into.
PDF from an uploaded template needs LibreOffice on the server. Uncomment
`libreoffice-calc` in `packages.txt` and reboot the app. It is a few hundred
MB and lengthens the first boot, so skip it if you only use the built-in
layout — that makes PDFs directly, with no converter. The Layouts tab says
which of the two you are in.
Layout of the repo
    app.py              the router — always the main file
    views/*.py          the pages
    auth.py             the PIN gate

Do not rename `views/` to `pages/`. Streamlit treats a folder called
`pages/` as an old-style multipage app: it builds its own sidebar navigation
from it and never runs `app.py`. That loses the top navigation, loses the wide
layout (`set_page_config` lives in `app.py`), and used to skip the PIN gate
entirely. Each page now calls `require_pin()` itself as well, so the gate holds
either way, but the folder must stay named `views/`. `app.py` prints an error
if a `pages/` folder reappears, and `test_app.py` fails on it.
Tabs
Overview - totals, value by category, concentration, dead barcodes
Candidates - pick a section, run matching, review. Selection is by value
threshold, no per-row ticking
Verify - export a per-section sheet for staff, upload it back once signed
Build sheets - generate all sheets at once, or one at a time to print. Each
gives an .xlsx and a .txt for iTrade bulk import
Manual pair - enter a pair by hand (repacking, or anything matched by eye)
Archive - saved batches and run history
Import file format
    SML,3113318000000,36.44,-12
    SML,70177178017,18.22,24

Prefix, barcode, cost, quantity. Outer line negative, single line positive.
No header, no quotes. Cost keeps up to 7 decimals so uneven conversions still
net to zero.
Conversion
400GM packs from 1 KG = 2.5, 18KG bag to loose KG = 18, wrong sale = 1.
Single qty = |outer qty| x conv, single cost = outer cost / conv.
