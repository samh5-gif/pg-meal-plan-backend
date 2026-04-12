"""
Project GAIN — Meal Plan PDF Generator
Backend: Flask API
"""

import os, re, json, io, base64, textwrap
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
from PIL import Image as PILImage

app = Flask(__name__)
CORS(app)

# ── Ambiguous ingredient rules ────────────────────────────────────────────────
# Items that can't practically be weighed by grams and need a human note
AMBIGUOUS_EXCLUSIONS = [
    'egg white', 'protein water', 'egg powder', 'dried egg', 'passionfruit',
]

AMBIGUOUS_PATTERNS = [
    # Whole eggs only — not egg whites
    (r'\bboiled eggs?\b',    'boiled egg',    'e.g. "1 large boiled egg (~60g)"'),
    (r'\bscrambled eggs?\b', 'scrambled egg', 'e.g. "2 large eggs scrambled"'),
    (r'\bfried eggs?\b',     'fried egg',     'e.g. "1 large fried egg"'),
    (r'\bpoached eggs?\b',   'poached egg',   'e.g. "1 large poached egg"'),
    (r'^eggs?$',               'egg',           'e.g. "2 large eggs" or "2 medium eggs"'),
    # Fruits (whole pieces)
    (r'\bapple\b',   'apple',        'e.g. "1 large apple" or "1 medium apple"'),
    (r'\bbanana\b',  'banana',       'e.g. "1 medium banana" or "1 small banana"'),
    (r'\bpeach\b',   'peach',        'e.g. "1 large peach" or "1 medium peach"'),
    (r'\bplum\b',    'plum',         'e.g. "1 large plum" or "2 small plums"'),
    (r'\borange\b',  'orange',       'e.g. "1 large orange"'),
    (r'\bpear\b',    'pear',         'e.g. "1 medium pear"'),
    (r'^mango',        'mango',        'e.g. "half a large mango"'),
    (r'\bkiwi\b',    'kiwi',         'e.g. "2 kiwis"'),
    (r'\bgrapes?\b', 'grapes',       'e.g. "a small handful of grapes"'),
    # Veg hard to weigh precisely
    (r'\bcarrots?\b',      'carrot',       'e.g. "1 medium carrot" or "2 small carrots"'),
    (r'\bcourgette\b',     'courgette',    'e.g. "half a medium courgette"'),
    (r'\bavocado\b',       'avocado',      'e.g. "half an avocado"'),
    (r'\bonion\b',         'onion',        'e.g. "half a medium onion"'),
    (r'\bsweet potato\b',  'sweet potato', 'e.g. "1 medium sweet potato"'),
    (r'\bpotato\b',        'potato',       'e.g. "1 medium potato"'),
]

# Fruits specifically (for the fruit note)
FRUIT_NAMES = {
    'apple', 'banana', 'peach', 'plum', 'orange', 'pear', 'mango',
    'kiwi', 'grapes', 'blueberries', 'blueberry', 'strawberries',
    'strawberry', 'raspberries', 'raspberry', 'blackberries', 'blackberry',
    'melon', 'watermelon', 'pineapple', 'cherry', 'cherries', 'grape',
    'lemon', 'lime', 'grapefruit', 'fig', 'date', 'apricot', 'nectarine',
    'pomegranate', 'passion fruit', 'passionfruit',
}

def check_ambiguous(qty_g, food_name):
    """Return ambiguity info if ingredient needs clarification, else None."""
    name_lower = food_name.lower()
    # Skip excluded items
    for excl in AMBIGUOUS_EXCLUSIONS:
        if excl in name_lower:
            return None
    for pattern, label, hint in AMBIGUOUS_PATTERNS:
        if re.search(pattern, name_lower, re.IGNORECASE):
            return {
                'food': food_name,
                'qty_g': qty_g,
                'label': label,
                'hint': hint,
            }
    return None

def is_fruit(food_name):
    name_lower = food_name.lower()
    for fruit in FRUIT_NAMES:
        if fruit in name_lower:
            return True
    return False

# ── Excel parsing ─────────────────────────────────────────────────────────────

def parse_excel(file_bytes, filename):
    """
    Parse the meal plan Excel.
    Returns:
      client_name: str
      days: list of {
        day_num: int,
        total_kcal, total_prot, total_fat, total_carb: float,
        meals: list of {
          meal_num: int,
          meal_label: str,   # "Meal 1" etc
          kcal, prot, fat, carb: float,
          ingredients: list of {food, qty_g, qty_label (None initially)}
        }
      }
    """
    # Client name from filename — strip extension, replace underscores/hyphens
    stem = os.path.splitext(filename)[0]
    # Try to extract name from patterns like "Meal_Plan_-_John_Smith__1_"
    # Remove common prefixes
    stem = re.sub(r'(?i)meal[\s_-]*plan[\s_-]*[-_]*', '', stem)
    stem = re.sub(r'[\s_-]+\d+[\s_-]*$', '', stem)  # trailing numbers
    stem = re.sub(r'[_]+', ' ', stem).strip(' -_()[]')
    # Capitalise each word
    client_name = ' '.join(w.capitalize() for w in stem.split() if w)
    if not client_name:
        client_name = 'Client'

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

    days = []
    day_num = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        name_lower = sheet_name.lower()

        # Only process "Meal Plan Day N" sheets
        if 'meal plan day' not in name_lower and 'meal' not in name_lower:
            continue
        if 'shopping' in name_lower or 'food' in name_lower.replace('foodlist',''):
            continue

        day_num += 1

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        # Find header row
        header_row = None
        for i, row in enumerate(rows):
            row_str = [str(c).lower() if c else '' for c in row]
            if any('meal' in c for c in row_str) and any('food' in c or 'item' in c for c in row_str):
                header_row = i
                break
        if header_row is None:
            continue

        headers = [str(c).lower().strip() if c else '' for c in rows[header_row]]

        # Identify column indices
        def col(keywords):
            for i, h in enumerate(headers):
                if any(k in h for k in keywords):
                    return i
            return None

        meal_col  = col(['meal'])
        food_col  = col(['food', 'item', 'ingredient'])
        qty_col   = col(['quantity', 'qty', 'amount', 'weight'])
        kcal_col  = col(['calori', 'kcal', 'energy'])
        prot_col  = col(['protein'])
        fat_col   = col(['fat'])
        carb_col  = col(['carb'])
        total_macros_col = col(['total macros'])
        daily_col = col(['daily total'])

        # Parse daily totals from the "Daily Totals" column if present
        daily_kcal = daily_prot = daily_fat = daily_carb = 0.0
        if daily_col is not None:
            for row in rows[header_row+1:]:
                cell_val = str(row[daily_col]).strip() if row[daily_col] is not None else ''
                if 'calori' in cell_val.lower() or 'kcal' in cell_val.lower():
                    # Next cell might have value
                    pass
                # Look for Total macros label + values in same region
            # Simpler: scan for the summary block
            for i, row in enumerate(rows[header_row+1:], header_row+1):
                if row[total_macros_col] if total_macros_col is not None else False:
                    label = str(row[total_macros_col]).lower()
                    val_cell = row[daily_col] if daily_col < len(row) else None
                    try:
                        val = float(val_cell) if val_cell is not None else 0
                    except:
                        val = 0
                    if 'calori' in label or 'kcal' in label:
                        daily_kcal = val
                    elif 'protein' in label:
                        daily_prot = val
                    elif 'fat' in label:
                        daily_fat = val
                    elif 'carb' in label:
                        daily_carb = val

        # Parse meal rows
        meals_dict = {}  # meal_label -> {kcal,prot,fat,carb,ingredients}

        for row in rows[header_row+1:]:
            if not any(row):
                continue
            meal_val = str(row[meal_col]).strip() if (meal_col is not None and row[meal_col] is not None) else ''
            if not meal_val or meal_val.lower() in ('none', 'nan', 'meal'):
                continue
            # Must start with "Meal"
            if not re.match(r'meal\s*\d+', meal_val, re.IGNORECASE):
                continue

            food_val  = str(row[food_col]).strip()  if (food_col  is not None and row[food_col]  is not None) else ''
            qty_raw   = row[qty_col]  if qty_col  is not None else None
            kcal_raw  = row[kcal_col] if kcal_col is not None else None
            prot_raw  = row[prot_col] if prot_col is not None else None
            fat_raw   = row[fat_col]  if fat_col  is not None else None
            carb_raw  = row[carb_col] if carb_col is not None else None

            def to_f(v):
                try: return float(v)
                except: return 0.0

            qty_g  = to_f(qty_raw)
            kcal_v = to_f(kcal_raw)
            prot_v = to_f(prot_raw)
            fat_v  = to_f(fat_raw)
            carb_v = to_f(carb_raw)

            if not food_val or food_val.lower() == 'nan':
                continue

            if meal_val not in meals_dict:
                meals_dict[meal_val] = {
                    'meal_label': meal_val,
                    'kcal': 0.0, 'prot': 0.0, 'fat': 0.0, 'carb': 0.0,
                    'ingredients': [],
                }

            meals_dict[meal_val]['kcal'] += kcal_v
            meals_dict[meal_val]['prot'] += prot_v
            meals_dict[meal_val]['fat']  += fat_v
            meals_dict[meal_val]['carb'] += carb_v
            meals_dict[meal_val]['ingredients'].append({
                'food': food_val,
                'qty_g': qty_g,
                'qty_label': None,  # filled in after clarification
            })

        # Sort meals
        def meal_sort_key(m):
            nums = re.findall(r'\d+', m)
            return int(nums[0]) if nums else 0

        meals_sorted = sorted(meals_dict.values(), key=lambda m: meal_sort_key(m['meal_label']))

        # Round meal macros
        for m in meals_sorted:
            m['kcal'] = round(m['kcal'])
            m['prot'] = round(m['prot'], 1)
            m['fat']  = round(m['fat'],  1)
            m['carb'] = round(m['carb'], 1)

        # Derive daily totals from sum if not parsed
        if daily_kcal == 0:
            daily_kcal = round(sum(m['kcal'] for m in meals_sorted))
            daily_prot = round(sum(m['prot'] for m in meals_sorted), 1)
            daily_fat  = round(sum(m['fat']  for m in meals_sorted), 1)
            daily_carb = round(sum(m['carb'] for m in meals_sorted), 1)

        days.append({
            'day_num': day_num,
            'sheet': sheet_name,
            'total_kcal': round(daily_kcal),
            'total_prot': round(daily_prot, 1),
            'total_fat':  round(daily_fat,  1),
            'total_carb': round(daily_carb, 1),
            'meals': meals_sorted,
        })

    return client_name, days


def find_ambiguous_ingredients(days):
    """Return list of items needing clarification."""
    ambiguous = []
    seen = set()
    for day in days:
        for meal in day['meals']:
            for ing in meal['ingredients']:
                result = check_ambiguous(ing['qty_g'], ing['food'])
                if result:
                    key = ing['food'].lower().strip()
                    if key not in seen:
                        seen.add(key)
                        result['key'] = key
                        ambiguous.append(result)
    return ambiguous


def apply_clarifications(days, clarifications):
    """
    clarifications: dict of food_key -> qty_label string
    Applies labels to all matching ingredients across all days.
    """
    for day in days:
        for meal in day['meals']:
            for ing in meal['ingredients']:
                key = ing['food'].lower().strip()
                if key in clarifications:
                    ing['qty_label'] = clarifications[key]


def has_fruit(days):
    for day in days:
        for meal in day['meals']:
            for ing in meal['ingredients']:
                if is_fruit(ing['food']):
                    return True
    return False


# ── PDF Generation ────────────────────────────────────────────────────────────

PW, PH = A4  # 595.28 x 841.89
LM, RM = 50.0, 545.28
CW = RM - LM

# Brand colours (black & white)
BLACK    = (0.08, 0.08, 0.08)
WHITE    = (1.0,  1.0,  1.0 )
OFFWHITE = (0.96, 0.96, 0.96)
MID_GREY = (0.45, 0.45, 0.45)
LT_GREY  = (0.88, 0.88, 0.88)
DIVIDER  = (0.80, 0.80, 0.80)
DARK_CARD= (0.12, 0.12, 0.12)

HB, H = 'Helvetica-Bold', 'Helvetica'

def pdf_y(top_y): return PH - top_y

def tw(text, font, size): return stringWidth(text, font, size)

def draw_text(c, x, top_y, text, font, size, rgb):
    c.saveState()
    c.setFillColorRGB(*rgb)
    c.setFont(font, size)
    c.drawString(x, pdf_y(top_y + size), text)
    c.restoreState()

def draw_centred(c, cx, top_y, text, font, size, rgb):
    w = tw(text, font, size)
    draw_text(c, cx - w/2, top_y, text, font, size, rgb)

def rounded_rect(c, x, bottom_pdf, w, h, r, fill=None, stroke=None, lw=0):
    c.saveState()
    c.setLineWidth(lw)
    if fill: c.setFillColorRGB(*fill)
    if stroke: c.setStrokeColorRGB(*stroke)
    p = c.beginPath()
    p.moveTo(x+r, bottom_pdf)
    p.lineTo(x+w-r, bottom_pdf)
    p.curveTo(x+w, bottom_pdf, x+w, bottom_pdf, x+w, bottom_pdf+r)
    p.lineTo(x+w, bottom_pdf+h-r)
    p.curveTo(x+w, bottom_pdf+h, x+w, bottom_pdf+h, x+w-r, bottom_pdf+h)
    p.lineTo(x+r, bottom_pdf+h)
    p.curveTo(x, bottom_pdf+h, x, bottom_pdf+h, x, bottom_pdf+h-r)
    p.lineTo(x, bottom_pdf+r)
    p.curveTo(x, bottom_pdf, x, bottom_pdf, x+r, bottom_pdf)
    p.close()
    if fill and stroke: c.drawPath(p, fill=1, stroke=1)
    elif fill: c.drawPath(p, fill=1, stroke=0)
    elif stroke: c.drawPath(p, fill=0, stroke=1)
    c.restoreState()

def wrap_text(text, font, size, max_w):
    words = text.split()
    lines, cur = [], ''
    for word in words:
        test = (cur + ' ' + word).strip()
        if tw(test, font, size) <= max_w:
            cur = test
        else:
            if cur: lines.append(cur)
            cur = word
    if cur: lines.append(cur)
    return lines

def draw_cover(c, client_name, logo_b64):
    """Draw a full cover page."""
    # Full black background
    c.setFillColorRGB(*BLACK)
    c.rect(0, 0, PW, PH, fill=1, stroke=0)

    # Logo centred, top third
    if logo_b64:
        try:
            img_data = base64.b64decode(logo_b64)
            import tempfile, os
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                tmp.write(img_data)
                tmp_path = tmp.name
            logo_size = 130
            logo_x = PW/2 - logo_size/2
            logo_y_pdf = pdf_y(220 + logo_size)
            c.drawImage(tmp_path, logo_x, logo_y_pdf,
                        width=logo_size, height=logo_size, mask='auto')
            os.unlink(tmp_path)
        except Exception as e:
            pass

    # "PROJECT GAIN" wordmark
    draw_centred(c, PW/2, 365, 'PROJECT GAIN', HB, 11, (0.6, 0.6, 0.6))

    # Thin white rule
    c.saveState()
    c.setStrokeColorRGB(1, 1, 1)
    c.setLineWidth(0.5)
    c.line(LM + 80, pdf_y(400), RM - 80, pdf_y(400))
    c.restoreState()

    # Client name — large
    draw_centred(c, PW/2, 420, client_name, HB, 32, WHITE)

    # Subtitle
    draw_centred(c, PW/2, 470, 'Personalised Meal Plan', H, 13, (0.55, 0.55, 0.55))

    # Bottom wordmark
    draw_centred(c, PW/2, PH - 60, 'projectgainofficial.com', H, 8, (0.35, 0.35, 0.35))


FRUIT_NOTE = (
    "Note: Any fruit included in today's meals can be eaten as a snack "
    "at any point throughout the day — it pairs particularly well with "
    "Greek yoghurt. This won't impact your results."
)

def draw_day_header(c, top_y, day_num, total_kcal, total_prot, total_fat, total_carb, include_fruit_note):
    """Draws day header block. Returns y after block."""
    box_h = 100.0
    rounded_rect(c, LM, pdf_y(top_y + box_h), CW, box_h, r=6, fill=BLACK)

    # Day title
    draw_centred(c, PW/2, top_y + 16, f"Day {day_num}", HB, 22, WHITE)

    # Macro pill
    pill_w, pill_h = 340.0, 26.0
    pill_x = PW/2 - pill_w/2
    pill_top = top_y + 46
    rounded_rect(c, pill_x, pdf_y(pill_top + pill_h), pill_w, pill_h, r=5,
                 fill=(0.22, 0.22, 0.22))

    macro_str = f"{total_kcal:,} kcal  ·  {total_prot}g Protein  ·  {total_fat}g Fats  ·  {total_carb}g Carbs"
    draw_centred(c, PW/2, pill_top + 8, macro_str, HB, 8.5, WHITE)

    y = top_y + box_h + 14

    # Fruit note
    if include_fruit_note:
        note_lines = wrap_text(FRUIT_NOTE, H, 8.5, CW - 20)
        note_h = len(note_lines) * 12 + 14
        rounded_rect(c, LM, pdf_y(y + note_h), CW, note_h, r=4, fill=OFFWHITE)
        # Left accent bar
        rounded_rect(c, LM, pdf_y(y + note_h), 3, note_h, r=2, fill=DARK_CARD)
        ny = y + 7
        for line in note_lines:
            draw_text(c, LM + 10, ny, line, H, 8.5, MID_GREY)
            ny += 12
        y += note_h + 10

    return y


def draw_meal_card(c, top_y, meal_label, kcal, prot, fat, carb, ingredients):
    """Draw one meal card. Returns y after card."""
    # Header bar — dark
    hdr_h = 40.0
    rounded_rect(c, LM, pdf_y(top_y + hdr_h), CW, hdr_h, r=6, fill=DARK_CARD)

    # Badge
    badge_w, badge_h = 52.0, 16.0
    badge_x = LM + 10
    badge_top = top_y + 12
    rounded_rect(c, badge_x, pdf_y(badge_top + badge_h), badge_w, badge_h, r=3,
                 fill=(0.28, 0.28, 0.28))
    draw_centred(c, badge_x + badge_w/2, badge_top + 4, meal_label.upper(), HB, 7, WHITE)

    # Dish title placeholder (just meal label for now — recipe names come from Claude)
    # Actually we just show the meal label as-is for simplicity
    # The meal card shows ingredients/method — no dish name since Excel doesn't have one

    y = top_y + hdr_h

    # Macro row — light grey bg
    macro_h = 38.0
    rounded_rect(c, LM, pdf_y(y + macro_h), CW, macro_h, r=0, fill=OFFWHITE)

    divs = [LM + CW*0.25, LM + CW*0.5, LM + CW*0.75]
    col_centres = [
        LM + CW*0.125,
        LM + CW*0.375,
        LM + CW*0.625,
        LM + CW*0.875,
    ]
    vals = [f"{kcal} kcal", f"{prot}g", f"{fat}g", f"{carb}g"]
    labs = ["Calories", "Protein", "Fats", "Carbs"]

    c.saveState()
    c.setStrokeColorRGB(*DIVIDER)
    c.setLineWidth(0.4)
    for dx in divs:
        c.line(dx, pdf_y(y + 8), dx, pdf_y(y + 30))
    c.restoreState()

    for cx, val, lab in zip(col_centres, vals, labs):
        draw_centred(c, cx, y + 7, val, HB, 11, BLACK)
        draw_centred(c, cx, y + 20, lab, H, 7, MID_GREY)

    y += macro_h + 18

    # INGREDIENTS
    draw_text(c, LM, y, 'INGREDIENTS', HB, 8.5, BLACK)
    y += 14

    qty_x = LM + 8
    for ing in ingredients:
        qty_g = ing['qty_g']
        food  = ing['food']
        label = ing.get('qty_label')

        # Format quantity string
        if label:
            # Clarified ingredient: show the label as qty, food name as-is after
            qty_str = label
            draw_text(c, qty_x, y, qty_str, HB, 9.5, BLACK)
            # Don't repeat the food name — the label already contains context
        else:
            # Standard ingredient: bold qty + regular food name
            if qty_g == int(qty_g):
                qty_str = f"{int(qty_g)}g"
            else:
                qty_str = f"{qty_g}g"
            draw_text(c, qty_x, y, qty_str, HB, 9.5, BLACK)
            name_x = qty_x + tw(qty_str, HB, 9.5) + 5.5
            draw_text(c, name_x, y, food, H, 9.5, BLACK)
        y += 14

    # Divider
    y += 6
    c.saveState()
    c.setStrokeColorRGB(*DIVIDER)
    c.setLineWidth(0.4)
    c.line(LM, pdf_y(y), RM, pdf_y(y))
    c.restoreState()
    y += 8

    return y


def generate_pdf(client_name, days, logo_b64):
    """Generate the full PDF. Returns bytes."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"{client_name} — Meal Plan")
    c.setAuthor("Project GAIN")

    has_any_fruit = has_fruit(days)

    # Cover page
    draw_cover(c, client_name, logo_b64)
    c.showPage()

    # Day pages
    for day in days:
        # Check if this day has fruit
        day_has_fruit = has_any_fruit and any(
            is_fruit(ing['food'])
            for meal in day['meals']
            for ing in meal['ingredients']
        )

        y = 48.0
        y = draw_day_header(c, y,
                            day['day_num'],
                            day['total_kcal'], day['total_prot'],
                            day['total_fat'],  day['total_carb'],
                            include_fruit_note=day_has_fruit)

        for meal in day['meals']:
            # Estimate card height
            n_ing = len(meal['ingredients'])
            ing_h = 14 + n_ing * 14 + 14
            macro_h = 38
            hdr_h = 40
            card_h = hdr_h + macro_h + 18 + ing_h + 20

            if y + card_h > PH - 40:
                c.showPage()
                y = 48.0

            y = draw_meal_card(
                c, y,
                meal['meal_label'],
                meal['kcal'], meal['prot'], meal['fat'], meal['carb'],
                meal['ingredients'],
            )
            y += 16

        c.showPage()

    c.save()
    buf.seek(0)
    return buf.read()


# ── API Routes ────────────────────────────────────────────────────────────────

# Store in-memory between requests (simple; stateless via session token)
_sessions = {}

@app.route('/upload', methods=['POST'])
def upload():
    """Step 1: Upload Excel, get back ambiguous ingredients."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx)'}), 400

    file_bytes = f.read()
    filename = f.filename

    try:
        client_name, days = parse_excel(file_bytes, filename)
    except Exception as e:
        return jsonify({'error': f'Could not parse Excel file: {str(e)}'}), 400

    ambiguous = find_ambiguous_ingredients(days)

    # Store session
    import uuid
    session_id = str(uuid.uuid4())
    _sessions[session_id] = {
        'client_name': client_name,
        'days': days,
        'file_bytes': file_bytes,
        'filename': filename,
    }

    return jsonify({
        'session_id': session_id,
        'client_name': client_name,
        'day_count': len(days),
        'ambiguous': ambiguous,
    })


@app.route('/generate', methods=['POST'])
def generate():
    """Step 2: Receive clarifications, generate PDF."""
    data = request.get_json()
    session_id    = data.get('session_id')
    clarifications = data.get('clarifications', {})  # {food_key: qty_label}
    logo_b64      = data.get('logo_b64', '')

    if session_id not in _sessions:
        return jsonify({'error': 'Session expired. Please re-upload the file.'}), 400

    sess = _sessions[session_id]
    client_name = sess['client_name']
    days = sess['days']

    apply_clarifications(days, clarifications)

    try:
        pdf_bytes = generate_pdf(client_name, days, logo_b64)
    except Exception as e:
        import traceback
        return jsonify({'error': f'PDF generation failed: {str(e)}', 'trace': traceback.format_exc()}), 500

    # Clean up session
    del _sessions[session_id]

    pdf_b64 = base64.b64encode(pdf_bytes).decode()
    return jsonify({
        'pdf_b64': pdf_b64,
        'filename': f"{client_name.replace(' ', '_')}_Meal_Plan.pdf",
    })


@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
