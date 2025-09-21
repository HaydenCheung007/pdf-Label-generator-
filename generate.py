import pandas as pd
from reportlab.graphics.barcode import code128
from reportlab.pdfgen import canvas
from pylibdmtx.pylibdmtx import encode
from pylibdmtx import pylibdmtx
from PIL import Image
import io
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageOps
from textwrap import wrap

# Load Excel input
df = pd.read_excel("input data.xlsx", sheet_name="Carton")
ef = pd.read_excel("input data.xlsx", sheet_name="Pallet")
sf = pd.read_excel("input data.xlsx", sheet_name="Select label type")
df.columns = df.columns.str.strip()

page_width = 15 * 28.35   # 15 cm → points
page_height = 10 * 28.35  # 10 cm → points

# PDF setup
num_items = len(df)
last_item_number = df.iloc[-1]['Item Number']
last_item_SerTIN = df.iloc[-1]['Ser./TIN']

if sf.iloc[0,1] == "Carton":
    pdf_file = f"CartonLabel_{last_item_number}_{last_item_SerTIN}.pdf"
else:
    pdf_file = f"PalletLabel_{ef.iloc[-1]['Item Number']}_{ef.iloc[-1]['Ser./TIN']}.pdf"
c = canvas.Canvas(f"{pdf_file}", pagesize=(page_width, page_height))

def draw_Carton(row, c):
    # Starting coordinates
    x_start = 5
    y_start = page_height - 15  # top margin

    # Item Number
    c.setFont("Helvetica-Bold", 12.5)
    c.drawString(x_start + 5, y_start, f"Carton")
    c.setFont("Helvetica", 10.5)
    c.drawString(x_start + 55, y_start, f"     Item Number: {row['Item Number']}")
    # Example GS1-128 data
    # (01) GTIN, (10) Batch/Lot, (17) Expiry Date
    gs1_data = f"{row['Item Number']}"

    # Generate Item number barcode
    ItemImg  = code128.Code128(
    gs1_data,
    barHeight=35,  # height of the barcode
    barWidth=1   # width of a single module
)
    ItemImg.drawOn(c, -13, 225)


    c.rect(0, y_start - 50, 250, 63, stroke=1, fill=0)# large
    c.rect(0, y_start - 6, 70, 19, stroke=1, fill=0) # small

    # Item Description
    c.setFont("Helvetica", 11.5)
    c.drawString(x_start, y_start -61, f"Item Description:")
    y_pos = y_start - 62
    c.setFont("Helvetica", 10)
    text = f"{row['Item Description']}"
    max_chars_per_line = 41
    lines = wrap(text, max_chars_per_line)
    for line in lines:
        c.drawString(x_start,  y_pos - 10, line)
        y_pos -= 11

    c.rect(0, y_start - 85, 250, 35, stroke=1, fill=0)

    # Revision No
    c.setFont("Helvetica", 12)
    rev_no = int(row['Revision no.'])
    rev_str = f"{rev_no:02d}"
    c.drawString(x_start, y_start -98, f"Revision: {rev_str}")
    # Generate Item number barcode
    gs1_data = f"{row['Revision no.']}"
    ItemImg  = code128.Code128(
    gs1_data,
    barHeight=35,  # height of the barcode
    barWidth=1   # width of a single module
    )
    ItemImg.drawOn(c, -13, 130)#draw barcode
    c.rect(0, y_start - 145, 155, 60, stroke=1, fill=0)

    # Mfg Date
    mfg_date = row['Mfg. Date']

    # If it's a pandas Timestamp or datetime, format it
    mfg_date_str = mfg_date.strftime("%m/%d/%Y")  # or "%Y-%m-%d"
    c.drawString(160, y_start -98, f"  Mfg. Date: ")
    c.drawString(165, y_start -115, mfg_date_str)
    c.rect(155, y_start - 145, 95, 60, stroke=1, fill=0)

    # Quantity
    c.drawString(x_start + 5, y_start - 158, f"  Quantity: ")
    c.drawString(x_start + 45, y_start - 175, f"{row['Quantity']}")
    c.rect(0, y_start - 190, 95, 45, stroke=1, fill=0)

    # PO Number
    c.drawString(100, y_start - 158, f"  PO Number: ")
    c.drawString(100, y_start - 175, f"  {row['PO Number']}")
    c.rect(95, y_start - 190, 155, 45, stroke=1, fill=0)

    # Ser./TIN
    c.setFont("Helvetica", 11.5)
    c.drawString(x_start, y_start -205, f"Ser./TIN: {row['Ser./TIN']}")
    gs1_data = f"{row['Ser./TIN']}"
    c.setFont("Helvetica", 10.5)
    # Generate Item number barcode
    ItemImg  = code128.Code128(
    gs1_data,
    barHeight=35,  # height of the barcode
    barWidth=1   # width of a single module
    )
    ItemImg.drawOn(c, -13, 15)#draw barcode
    c.rect(0, 5, 250, 73.5, stroke=1, fill=0)

    # Deviations
    y_pos = y_start - 180
    c.setFont("Helvetica", 12.5)
    c.drawString(255,  y_start -160, f"Deviations: ")
    
    deviation_value = row['Deviations']
    if pd.isna(deviation_value):
        text = " "
    else:
        text = str(deviation_value)   # convert to string only if not NaN

    max_chars_per_line = 26
    lines = wrap(text, max_chars_per_line)
    for line in lines:
        c.drawString(255,  y_pos, line)
        y_pos -= 12
    c.rect(250, 5, 174, 118.5, stroke=1, fill=0)

    # MIL code
    # barcode
    # Generate DataMatrix
    data = f"{row['MIL CODE']}"
    encoded = encode(data.encode("utf-8"))
    MILimg = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
    #img.save("datamatrix.png")
    MILimg = MILimg.resize((MILimg.width * 50, MILimg.height * 50), Image.NEAREST)
    img_reader = ImageReader(MILimg)
    #MIL text
    c.drawImage(img_reader, 261, 125, width=160, height=160, mask='auto')
    c.saveState()                         # save current state
    c.translate(265, 185)         # move origin to where you want text
    c.rotate(90)                          # rotate coordinate system 90 degrees
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0, 0, "MIL CODE")   # draw at new origin
    c.restoreState()                      # restore normal orientation

    c.rect(250, y_start - 145, 174, 158, stroke=1, fill=0)

#pallet pallet pallet pallet pallet pallet pallet pallet pallet pallet pallet pallet pallet

def draw_Pallet(row, c):
    # Starting coordinates
    x_start = 5
    y_start = page_height - 15  # top margin

    # Item Number
    c.setFont("Helvetica-Bold", 12.5)
    c.drawString(x_start + 5, y_start, f"Pallet")
    c.setFont("Helvetica", 10.5)
    c.drawString(x_start + 55, y_start, f"     Item Number: {row['Item Number']}")
    # Example GS1-128 data
    # (01) GTIN, (10) Batch/Lot, (17) Expiry Date
    gs1_data = f"{row['Item Number']}"

    # Generate Item number barcode
    ItemImg  = code128.Code128(
    gs1_data,
    barHeight=35,  # height of the barcode
    barWidth=1   # width of a single module
)
    ItemImg.drawOn(c, -13, 225)


    c.rect(0, y_start - 50, 250, 63, stroke=1, fill=0)# large
    c.rect(0, y_start - 6, 70, 19, stroke=1, fill=0) # small

    # Item Description
    c.setFont("Helvetica", 11.5)
    c.drawString(x_start, y_start -61, "Item Description:")
    y_pos = y_start - 62
    c.setFont("Helvetica", 10)
    text = f"{row['Item Description']}"
    max_chars_per_line = 41
    lines = wrap(text, max_chars_per_line)
    for line in lines:
        c.drawString(x_start,  y_pos - 10, line)
        y_pos -= 11

    c.rect(0, y_start - 85, 250, 35, stroke=1, fill=0)

    # Revision No
    c.setFont("Helvetica", 12)
    rev_no = int(row['Revision no.'])
    rev_str = f"{rev_no:02d}"
    c.drawString(x_start, y_start -98, f"Revision: {rev_str}")
    gs1_data = f"{row['Revision no.']}"

    # Generate Item number barcode
    ItemImg  = code128.Code128(
    gs1_data,
    barHeight=35,  # height of the barcode
    barWidth=1   # width of a single module
    )
    ItemImg.drawOn(c, -13, 130)#draw barcode
    c.rect(0, y_start - 145, 155, 60, stroke=1, fill=0)

    # Pallet Weight
    c.drawString(160, y_start -98, "  Pallet Weight: ")
    c.setFont("Helvetica", 11)

    y_pos = y_start - 110
    text = str(row['Pallet Weight'])

    # Split at "/" but keep track of where we are
    parts = text.split("/")

    for i, part in enumerate(parts):
        if i < len(parts) - 1:
            line = part.strip() + "/"   # add the slash back
        else:
            line = part.strip()         # last part → no slash
        c.drawString(160, y_pos - 5, line)
        y_pos -= 11

    #c.drawString(160, y_start -110, f"{row['Pallet Weight']}")
    c.setFont("Helvetica", 12)
    c.rect(155, y_start - 145, 95, 60, stroke=1, fill=0)

    # Quantity
    c.drawString(x_start + 5, y_start - 158, f"  Quantity: ")
    c.drawString(x_start + 5, y_start - 175, f"  Units: ")
    c.drawString(x_start + 45, y_start - 175, f"{row['Quantity']}")
    c.rect(0, y_start - 190, 95, 45, stroke=1, fill=0)

    # PO Number
    c.drawString(100, y_start - 158, f"  PO Number: ")
    c.drawString(100, y_start - 175, f"  {row['PO Number']}")
    c.rect(95, y_start - 190, 155, 45, stroke=1, fill=0)

    # Ser./TIN
    c.setFont("Helvetica", 11.5)
    c.drawString(x_start, y_start -205, f"Ser./TIN: {row['Ser./TIN']}")
    gs1_data = f"{row['Ser./TIN']}"
    c.setFont("Helvetica", 10.5)
    # Generate Item number barcode
    ItemImg  = code128.Code128(
    gs1_data,
    barHeight=35,  # height of the barcode
    barWidth=1   # width of a single module
    )
    ItemImg.drawOn(c, -13, 15)#draw barcode
    c.rect(0, 5, 250, 73.5, stroke=1, fill=0)

    # Information
    #upper info
    c.setFont("Helvetica", 9)
    name = row['Supplier Name']
    if pd.isna(name): 
        name = " "    
    c.drawString(252, 113.5, f"Supplier: {name}")
    #supplier address 
    address = row['Supplier Address']
    if pd.isna(address): 
        address = " "    
    c.drawString(252, 101.5, f"{address}")
    #City, State, Zip, Country
    address2 = row['Supplier Address 2']
    if pd.isna(address2): 
        address2 = " "    
    c.drawString(252,89.5, f"{address2}")

    CSZC_ = row['Supplier City , State, Zip, Country']
    if pd.isna(CSZC_): 
        CSZC_ = " "    
    c.drawString(252, 77.5, f"{CSZC_}")
    #lower info
    entity = row['Ship to Entity Name']
    if pd.isna(entity): 
        entity = " " 
    c.drawString(252, 51.75, f"Ship to: {entity}")

    ad1 = row['Ship to address 1']
    if pd.isna(ad1): 
        ad1 = " " 
    c.drawString(252, 39.75, f"{ad1}")
    tdaddress2 = row['Ship to address 2']
    if pd.isna(tdaddress2): 
        tdaddress2 = " "    
    c.drawString(252, 27.75, f"{tdaddress2}")

    SCSZC = row['Ship to City, State, Zip, Country']
    if pd.isna(SCSZC): 
        SCSZC = " " 
    c.drawString(252, 15.75, f"{SCSZC}")

    c.rect(250, 61.75, 174, 61.75, stroke=1, fill=0) #lining
    c.rect(250, 5, 174, 118.5, stroke=1, fill=0)

    # MIL code
    # barcode
    # Generate DataMatrix
    data = f"{row['MIL CODE']}"
    encoded = encode(data.encode("utf-8"))
    MILimg = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
    #img.save("datamatrix.png")
    MILimg = MILimg.resize((MILimg.width * 50, MILimg.height * 50), Image.NEAREST)
    img_reader = ImageReader(MILimg)
    #MIL text
    c.drawImage(img_reader, 261, 125, width=160, height=160, mask='auto')
    c.saveState()                         # save current state
    c.translate(265, 185)         # move origin to where you want text
    c.rotate(90)                          # rotate coordinate system 90 degrees
    c.setFont("Helvetica-Bold", 11)
    c.drawString(0, 0, "MIL CODE")   # draw at new origin
    c.restoreState()                      # restore normal orientation

    c.rect(250, y_start - 145, 174, 158, stroke=1, fill=0)



# Loop through each row and draw label

if sf.iloc[0,1] == "Carton":
    for index, row in df.iterrows():
        draw_Carton(row, c)
        c.showPage()  # each label on a new page
else: #pallet
    for index, row in ef.iterrows():
        draw_Pallet(row, c)
        c.showPage()


c.save()
print(f"Labels PDF saved: {pdf_file}")