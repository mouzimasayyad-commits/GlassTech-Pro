from flask import Flask, render_template, request, jsonify, send_file
import os
import ezdxf
import math
import openpyxl
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# In-Memory Order Database
ORDER_DATABASE = [
    {
        "id": "ORD-101",
        "customer_name": "Rohan Malhotra",
        "shopkeeper_name": "Sharma Glass House",
        "shop_detail": "Shop #12, MG Road Market",
        "glass_type": "V-Groove Mirror (Diamond Grid)",
        "dimensions": "1200 x 2400 mm",
        "sent_date": "2026-08-15",
        "rec_date": "2026-08-18",
        "status": "Done",
        "pending_days": 0,
        "total_amount": 12500,
        "pending_payment": 0
    },
    {
        "id": "ORD-102",
        "customer_name": "Ananya Interior",
        "shopkeeper_name": "Verma Glass Traders",
        "shop_detail": "Plot 45, Industrial Area",
        "glass_type": "Beveled Wall Paneling",
        "dimensions": "1860 x 2500 mm",
        "sent_date": "2026-08-19",
        "rec_date": "In Processing",
        "status": "In Process",
        "pending_days": 2,
        "total_amount": 28400,
        "pending_payment": 8000
    }
]

# Helper function to convert any unit to standard Millimeters
def convert_to_mm(value, unit):
    val = float(value or 0)
    if unit == 'inch':
        return val * 25.4
    elif unit == 'feet':
        return val * 304.8
    elif unit == 'cm':
        return val * 10.0
    return val  # Default mm

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/orders', methods=['GET'])
def get_orders():
    return jsonify(ORDER_DATABASE)

@app.route('/api/add-order', methods=['POST'])
def add_order():
    data = request.json
    order_id = f"ORD-{len(ORDER_DATABASE) + 101}"
    
    unit = data.get("unit", "mm")
    w_mm = convert_to_mm(data.get('width'), unit)
    h_mm = convert_to_mm(data.get('height'), unit)
    
    new_order = {
        "id": order_id,
        "customer_name": data.get("customer_name"),
        "shopkeeper_name": data.get("shopkeeper_name"),
        "shop_detail": data.get("shop_detail"),
        "glass_type": data.get("glass_type"),
        "dimensions": f"{round(w_mm,1)} x {round(h_mm,1)} mm ({data.get('width')}x{data.get('height')} {unit})",
        "sent_date": data.get("sent_date", datetime.today().strftime('%Y-%m-%d')),
        "rec_date": "Pending",
        "status": "In Process",
        "pending_days": 1,
        "total_amount": float(data.get("amount", 0)),
        "pending_payment": float(data.get("pending_amount", 0))
    }
    ORDER_DATABASE.insert(0, new_order)
    return jsonify({"message": "Order registered successfully!", "order": new_order})

# Export Database to Dynamic Excel Spreadsheet (.xlsx)
@app.route('/api/export-excel', methods=['GET'])
def export_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Glass Factory Ledger"

    # Header Row
    headers = [
        "Order ID", "Customer Name", "Shopkeeper / Merchant", "Shop Details", 
        "Glass / Mirror Type", "Dimensions", "Sent Date", "Received Date", 
        "Status", "Pending Days", "Total Amount (INR)", "Pending Amount (INR)"
    ]
    ws.append(headers)

    for order in ORDER_DATABASE:
        ws.append([
            order["id"], order["customer_name"], order["shopkeeper_name"], 
            order["shop_detail"], order["glass_type"], order["dimensions"], 
            order["sent_date"], order["rec_date"], order["status"], 
            order["pending_days"], order["total_amount"], order["pending_payment"]
        ])

    file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'Glass_Factory_Orders.xlsx')
    wb.save(file_path)
    return send_file(file_path, as_attachment=True)

# Full Automated Vector DXF Generator for CNC Cutting & Corel Import
@app.route('/api/generate-cad', methods=['POST'])
def generate_cad():
    data = request.json
    unit = data.get('unit', 'mm')
    
    width = convert_to_mm(data.get('width', 1860), unit)
    height = convert_to_mm(data.get('height', 2500), unit)
    pattern = data.get('pattern', 'vgroove_diamond')
    cols = int(data.get('cols', 4))
    rows = int(data.get('rows', 4))
    bevel = convert_to_mm(data.get('bevel', 40), unit)

    doc = ezdxf.new('R2010')
    msp = doc.modelspace()

    # Color 1 = Red (Boundary Glass Cut), Color 2 = Yellow (V-Groove / Bevel Engraving)
    doc.layers.add(name="GLASS_OUTER_CUT", color=1)
    doc.layers.add(name="V_GROOVE_ENGRAVE", color=2)

    # 1. Outer Frame Cutline
    msp.add_lwpolyline(
        [(0, 0), (width, 0), (width, height), (0, height), (0, 0)],
        dxfattribs={'layer': 'GLASS_OUTER_CUT'}
    )

    # 2. Pattern Vector Generation
    if pattern == 'vgroove_diamond':
        # Diamond V-Groove Matrix (Image 5)
        step_x = width / cols
        step_y = height / rows

        for i in range(cols + rows + 1):
            # Positively sloped lines (/)
            x1 = 0
            y1 = i * step_y
            x2 = i * step_x
            y2 = 0
            if y1 > height:
                x1 = (y1 - height) * (step_x / step_y)
                y1 = height
            if x2 > width:
                y2 = (x2 - width) * (step_y / step_x)
                x2 = width
            msp.add_line((x1, y1), (x2, y2), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})

            # Negatively sloped lines (\)
            x3 = 0
            y3 = height - (i * step_y)
            x4 = i * step_x
            y4 = height
            if y3 < 0:
                x3 = (-y3) * (step_x / step_y)
                y3 = 0
            if x4 > width:
                y4 = height - ((x4 - width) * (step_y / step_x))
                x4 = width
            msp.add_line((x3, y3), (x4, y4), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})

    elif pattern == 'chevron_slant':
        # Chevron / Slanted Panels (Image 4)
        col_w = width / cols
        for c in range(1, cols):
            msp.add_line((c * col_w, 0), (c * col_w, height), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})
        
        step_y = height / rows
        for c in range(cols):
            x_left = c * col_w
            x_right = (c + 1) * col_w
            for r in range(rows + 1):
                y_base = r * step_y
                y_offset = y_base + (step_y * 0.5 if c % 2 == 0 else -step_y * 0.5)
                msp.add_line((x_left, max(0, min(height, y_base))), 
                             (x_right, max(0, min(height, y_offset))), 
                             dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})

    elif pattern == 'bevel_grid':
        # Beveled Wall Paneling (Image 1)
        cell_w = width / cols
        cell_h = height / rows
        for r in range(rows):
            for c in range(cols):
                cx, cy = c * cell_w, r * cell_h
                msp.add_lwpolyline([(cx, cy), (cx+cell_w, cy), (cx+cell_w, cy+cell_h), (cx, cy+cell_h), (cx, cy)], dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})
                if cell_w > bevel * 2 and cell_h > bevel * 2:
                    icx, icy = cx + bevel, cy + bevel
                    icw, ich = cell_w - (bevel * 2), cell_h - (bevel * 2)
                    msp.add_lwpolyline([(icx, icy), (icx+icw, icy), (icx+icw, icy+ich), (icx, icy+ich), (icx, icy)], dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})
                    msp.add_line((cx, cy), (icx, icy), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})
                    msp.add_line((cx+cell_w, cy), (icx+icw, icy), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})
                    msp.add_line((cx, cy+cell_h), (icx, icy+ich), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})
                    msp.add_line((cx+cell_w, cy+cell_h), (icx+icw, icy+ich), dxfattribs={'layer': 'V_GROOVE_ENGRAVE'})

    filename = f"Glass_Design_{pattern}_{int(width)}x{int(height)}mm.dxf"
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    doc.saveas(filepath)

    return send_file(filepath, as_attachment=True)

# AI Assistant Bot
@app.route('/api/chatbot', methods=['POST'])
def chatbot():
    msg = request.json.get("message", "").lower()
    if "vgroove" in msg or "v-groove" in msg:
        reply = "V-Groove mirrors feature deep decorative CNC engraving lines (typically 3mm to 10mm width) in diamond, chevron, or custom geometric patterns."
    elif "unit" in msg or "inch" in msg:
        reply = "Our system supports Inches, Feet, CM, and MM! You can enter measurements in any unit, and our engine automatically outputs 1:1 scale DXF files."
    elif "excel" in msg:
        reply = "You can click '📥 Export to Excel Spreadsheet' in the Business Ledger section to download your customer, shopkeeper, and payment records."
    else:
        reply = "I am your Factory AI Assistant. Ask me about V-Groove Mirrors, Beveled Glass, Excel Ledger Exports, or Unit Conversions!"
    return jsonify({"reply": reply})

if __name__ == '__main__':
    app.run(debug=True)