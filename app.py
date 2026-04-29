"""
Project GAIN — Meal Plan PDF + Recipe Guide Generator
Backend: Flask API with Anthropic AI for recipe generation
"""

import os, re, io, base64, json, tempfile
from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase.pdfmetrics import stringWidth
import anthropic

# ── Register Inter fonts for PDF ─────────────────────────────────────────────
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def _register_fonts():
    # Try multiple possible font locations
    possible_dirs = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fonts'),
        os.path.join(os.getcwd(), 'fonts'),
        '/opt/render/project/src/fonts',
    ]
    for font_dir in possible_dirs:
        try:
            r = os.path.join(font_dir, 'Inter-Regular.ttf')
            b = os.path.join(font_dir, 'Inter-Bold.ttf')
            xb = os.path.join(font_dir, 'Inter-ExtraBold.ttf')
            if not all(os.path.exists(f) for f in [r, b, xb]):
                continue
            pdfmetrics.registerFont(TTFont('Inter', r))
            pdfmetrics.registerFont(TTFont('Inter-Bold', b))
            pdfmetrics.registerFont(TTFont('Inter-ExtraBold', xb))
            pdfmetrics.registerFontFamily('Inter', normal='Inter', bold='Inter-Bold')
            return True
        except Exception:
            continue
    return False  # Falls back to Helvetica gracefully

INTER_AVAILABLE = _register_fonts()

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})

@app.after_request
def add_cors(response):
    response.headers['Access-Control-Allow-Origin'] = '*'
    response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    response.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
    return response

@app.route('/upload', methods=['OPTIONS'])
@app.route('/clarify', methods=['OPTIONS'])
@app.route('/generate-pdf', methods=['OPTIONS'])
@app.route('/health', methods=['OPTIONS'])
def options():
    return '', 204

# ── Anthropic client ──────────────────────────────────────────────────────────
def get_claude():
    return anthropic.Anthropic(api_key=os.environ.get('ANTHROPIC_API_KEY', ''))

# ── Session store ─────────────────────────────────────────────────────────────
_sessions = {}

# ── Ambiguous ingredient detection ───────────────────────────────────────────
EXCLUSIONS = ['egg white', 'protein water', 'egg powder', 'passionfruit']

# Patterns that flag an ingredient as needing size clarification
AMBIGUOUS_PATTERNS = [
    r'\bboiled eggs?\b',
    r'\bscrambled eggs?\b',
    r'\bfried eggs?\b',
    r'\bpoached eggs?\b',
    r'^eggs?$',
    r'\bapple\b',
    r'\bbanana\b',
    r'\bpeach\b',
    r'\bplum\b',
    r'\borange\b',
    r'\bpear\b',
    r'^mango\b',
    r'\bkiwi\b',
    r'\bgrapes?\b',
    r'\bcarrots?\b',
    r'\bcourgette\b',
    r'\bavocado\b',
    r'\bonion\b',
    r'\bsweet potato\b',
    r'\bpotato\b',
]

def check_ambiguous(qty_g, food_name):
    """Flag ingredient as ambiguous if it matches any pattern. Suggestion is filled in later by AI web search."""
    name_lower = food_name.lower()
    for excl in EXCLUSIONS:
        if excl in name_lower:
            return None
    for pattern in AMBIGUOUS_PATTERNS:
        if re.search(pattern, name_lower, re.IGNORECASE):
            return {
                'food': food_name,
                'qty_g': qty_g,
                'key': name_lower.strip(),
                'suggestion': None,  # filled in by lookup_size_suggestions()
            }
    return None


# Size reference data — weights in grams per unit, indexed by keyword
_SIZE_DATA = {
    'egg':          (60,  'large egg',    'UK large eggs weigh 63-73g each'),
    'apple':        (182, 'large apple',  'large apples weigh around 180-220g'),
    'banana':       (120, 'medium banana','medium bananas weigh around 110-130g'),
    'peach':        (175, 'large peach',  'large peaches weigh around 170-200g'),
    'plum':         (70,  'large plum',   'large plums weigh around 65-85g'),
    'orange':       (180, 'large orange', 'large oranges weigh around 175-210g'),
    'pear':         (170, 'medium pear',  'medium pears weigh around 160-190g'),
    'mango':        (350, 'whole mango',  'a whole large mango weighs around 350g'),
    'kiwi':         (70,  'kiwi',         'a standard kiwi weighs around 70g'),
    'grape':        (5,   'grape',        'individual grapes weigh around 5-8g'),
    'carrot':       (85,  'medium carrot','medium carrots weigh around 75-100g'),
    'courgette':    (200, 'courgette',    'a whole medium courgette weighs around 200g'),
    'avocado':      (200, 'avocado',       'a whole medium avocado weighs around 180-220g'),
    'onion':        (150, 'medium onion', 'a whole medium onion weighs around 140-170g'),
    'sweet potato': (150, 'medium sweet potato', 'a medium sweet potato weighs around 130-180g'),
    'potato':       (175, 'medium potato','a medium potato weighs around 150-200g'),
}

# ── Persistent size suggestion cache ─────────────────────────────────────────
_CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'size_cache.json')

def _load_cache():
    try:
        with open(_CACHE_FILE) as f:
            return json.load(f)
    except Exception:
        return {}

def _save_cache(cache):
    try:
        with open(_CACHE_FILE, 'w') as f:
            json.dump(cache, f, indent=2)
    except Exception:
        pass

_size_cache = _load_cache()


def lookup_size_suggestions(ambiguous_items):
    """
    Instant size suggestions using known food weight data.
    No API call needed — runs in microseconds.
    """
    for item in ambiguous_items:
        name_lower = item['food'].lower()
        qty = item['qty_g']
        qty_int = int(qty) if qty == int(qty) else qty

        matched = None
        # Find matching size data entry
        for keyword, (unit_g, unit_name, reference) in _SIZE_DATA.items():
            if keyword in name_lower:
                matched = (unit_g, unit_name, reference)
                break

        if matched:
            unit_g, unit_name, reference = matched
            count = qty / unit_g
            # Round to nearest sensible fraction
            if count <= 0.4:
                desc = f"a small portion (~{qty_int}g)"
                short = f"~{qty_int}g"
            elif count <= 0.65:
                desc = f"half an {unit_name}" if unit_name[0] in "aeiou" else "half a {unit_name}"
                short = f"half an {unit_name}" if unit_name[0] in "aeiou" else "half a {unit_name}"
            elif count <= 1.3:
                desc = f"1 {unit_name}"
                short = f"1 {unit_name}"
            elif count <= 1.7:
                desc = f"1-2 {unit_name}s"
                short = f"1-2 {unit_name}s"
            else:
                n = round(count)
                desc = f"{n} {unit_name}s"
                short = f"{n} {unit_name}s"

            item['suggestion'] = f"{qty_int}g = {desc} ({reference})"
            item['short'] = short
        else:
            # Mark for AI fallback
            item['suggestion'] = None
            item['short'] = ''

    # For items not in the hardcoded table, check cache first then fall back to Claude
    unknown = [i for i in ambiguous_items if i['suggestion'] is None]

    if unknown:
        # Step 1: fill from cache where possible
        still_unknown = []
        for item in unknown:
            cache_key = f"{item['food'].lower().strip()}:{int(item['qty_g']) if item['qty_g']==int(item['qty_g']) else item['qty_g']}"
            if cache_key in _size_cache:
                cached = _size_cache[cache_key]
                item['suggestion'] = cached['suggestion']
                item['short']      = cached['short']
            else:
                still_unknown.append(item)

        # Step 2: call Claude only for items not in cache
        if still_unknown:
            try:
                items_text = '\n'.join(
                    f"- {i['food']}: {int(i['qty_g']) if i['qty_g']==int(i['qty_g']) else i['qty_g']}g"
                    for i in still_unknown
                )
                prompt = (
                    "You are a nutrition expert. For each ingredient and gram weight below, "
                    "give the most accurate practical size equivalent in one short sentence.\n"
                    "Format: Xg = [description] ([brief weight reference])\n"
                    "Also give a \"short\" field: just the description, no grams.\n\n"
                    "Respond ONLY as valid JSON:\n"
                    '{"suggestions": [{"food": "X", "suggestion": "50g = 1 small egg (large eggs weigh 63-73g)", "short": "1 small egg"}]}\n\n'
                    f"Ingredients:\n{items_text}"
                )
                resp = get_claude().messages.create(
                    model='claude-haiku-4-5-20251001',
                    max_tokens=400,
                    messages=[{'role': 'user', 'content': prompt}]
                )
                raw = resp.content[0].text.strip()
                raw = re.sub(r'^```json\s*', '', raw, flags=re.MULTILINE)
                raw = re.sub(r'^```\s*', '', raw, flags=re.MULTILINE)
                raw = re.sub(r'\s*```$', '', raw)
                data = json.loads(raw)

                sugg_map = {}
                for s in data.get('suggestions', []):
                    k = s['food'].lower().strip()
                    sugg_map[k] = s
                    sugg_map[re.sub(r'\s*\([^)]+\)', '', k).strip()] = s

                cache_updated = False
                for item in still_unknown:
                    k  = item['food'].lower().strip()
                    ck = re.sub(r'\s*\([^)]+\)', '', k).strip()
                    match = sugg_map.get(k) or sugg_map.get(ck)
                    qty_int = int(item['qty_g']) if item['qty_g']==int(item['qty_g']) else item['qty_g']
                    if match:
                        item['suggestion'] = match.get('suggestion', '')
                        item['short']      = match.get('short', '')
                        # Save to cache keyed by food + qty
                        cache_key = f"{k}:{qty_int}"
                        _size_cache[cache_key] = {'suggestion': item['suggestion'], 'short': item['short']}
                        cache_updated = True
                    else:
                        item['suggestion'] = f"{qty_int}g — please confirm the practical size"
                        item['short'] = ''

                if cache_updated:
                    _save_cache(_size_cache)

            except Exception:
                for item in still_unknown:
                    qty_int = int(item['qty_g']) if item['qty_g']==int(item['qty_g']) else item['qty_g']
                    item['suggestion'] = f"{qty_int}g — please confirm the practical size"
                    item['short'] = ''

    return ambiguous_items

def is_fruit(food_name):
    fruits = {'apple','banana','peach','plum','orange','pear','mango','kiwi',
              'grape','blueberr','strawberr','raspberr','blackberr','melon',
              'pineapple','cherry','lemon','lime','grapefruit','fig','date',
              'apricot','nectarine','pomegranate'}
    name = food_name.lower()
    return any(f in name for f in fruits)

# ── Excel parsing ─────────────────────────────────────────────────────────────
def parse_excel(file_bytes, filename):
    stem = os.path.splitext(filename)[0]
    stem = re.sub(r'(?i)meal[\s_-]*plan[\s_-]*[-_]*', '', stem)
    stem = re.sub(r'[\s_-]+\d+[\s_-]*$', '', stem)
    stem = re.sub(r'[_]+', ' ', stem).strip(' -_()[]')
    client_name = ' '.join(w.capitalize() for w in stem.split() if w) or 'Client'

    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    days = []
    day_num = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        nl = sheet_name.lower()
        # Skip non-meal sheets
        if any(skip in nl for skip in ['shopping', 'foodlist', 'food list', 'targets', 'summary', 'notes']):
            continue
        # Accept sheets with day/meal/week indicators, or detect by content below
        has_day_indicator = any(w in nl for w in ['meal', 'day', 'week'])
        # We'll detect by content too — check after header scan

        day_num += 1
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        header_row = None
        for i, row in enumerate(rows):
            row_s = [str(c).lower() if c else '' for c in row]
            if any('meal' in c for c in row_s) and any('food' in c or 'item' in c for c in row_s):
                header_row = i
                break
        if header_row is None:
            continue
        # If sheet has no day indicator in name, only include if it has meal data
        if not has_day_indicator:
            # Check if it looks like a meal plan sheet by scanning for meal-like labels
            has_meal_data = False
            for row in rows[header_row+1:header_row+5]:
                if row and row[0] and str(row[0]).strip().lower() not in ('none', 'nan', ''):
                    has_meal_data = True
                    break
            if not has_meal_data:
                continue

        headers = [str(c).lower().strip() if c else '' for c in rows[header_row]]

        def col(kws):
            for i, h in enumerate(headers):
                if any(k in h for k in kws):
                    return i
            return None

        meal_col = col(['meal'])
        food_col = col(['food', 'item', 'ingredient'])
        qty_col  = col(['quantity', 'qty', 'amount', 'weight'])
        kcal_col = col(['calori', 'kcal', 'energy'])
        prot_col = col(['protein'])
        fat_col  = col(['fat'])
        carb_col = col(['carb'])

        def to_f(v):
            try: return float(v)
            except: return 0.0

        meals_dict = {}
        for row in rows[header_row+1:]:
            if not any(row):
                continue
            mv = str(row[meal_col]).strip() if meal_col is not None and row[meal_col] else ''
            if not mv or mv.lower() in ('none', 'nan', 'meal', 'meal name', 'meal label'):
                continue
            # Accept any non-empty meal label: "Meal 1", "Breakfast", "Snack 1", "Lunch" etc
            # Skip only if it looks like a header or is clearly not a meal
            if mv.lower() in ('total', 'daily total', 'totals'):
                continue
            fv = str(row[food_col]).strip() if food_col is not None and row[food_col] else ''
            if not fv or fv.lower() == 'nan':
                continue

            qty_g  = to_f(row[qty_col]  if qty_col  is not None else 0)
            kcal_v = to_f(row[kcal_col] if kcal_col is not None else 0)
            prot_v = to_f(row[prot_col] if prot_col is not None else 0)
            fat_v  = to_f(row[fat_col]  if fat_col  is not None else 0)
            carb_v = to_f(row[carb_col] if carb_col is not None else 0)

            if mv not in meals_dict:
                meals_dict[mv] = {'meal_label': mv, 'kcal': 0, 'prot': 0, 'fat': 0, 'carb': 0, 'ingredients': []}

            meals_dict[mv]['kcal'] += kcal_v
            meals_dict[mv]['prot'] += prot_v
            meals_dict[mv]['fat']  += fat_v
            meals_dict[mv]['carb'] += carb_v
            meals_dict[mv]['ingredients'].append({
                'food': fv,
                'qty_g': qty_g,
                'qty_label': None,
                'excel_qty_g': qty_g,
            })

        def meal_key(m):
            nums = re.findall(r'\d+', m)
            return int(nums[0]) if nums else 0

        meals = sorted(meals_dict.values(), key=lambda m: meal_key(m['meal_label']))
        for m in meals:
            m['kcal'] = round(m['kcal'])
            m['prot'] = round(m['prot'], 1)
            m['fat']  = round(m['fat'],  1)
            m['carb'] = round(m['carb'], 1)
            m['dish_name'] = ''
            m['recipe_steps'] = []

        days.append({
            'day_num': day_num,
            'sheet': sheet_name,
            'total_kcal': round(sum(m['kcal'] for m in meals)),
            'total_prot': round(sum(m['prot'] for m in meals), 1),
            'total_fat':  round(sum(m['fat']  for m in meals), 1),
            'total_carb': round(sum(m['carb'] for m in meals), 1),
            'meals': meals,
        })

    # Parse shopping lists from same workbook
    shopping_lists = parse_shopping_lists(wb)
    return client_name, days, shopping_lists


def parse_shopping_lists(wb):
    """Parse all shopping list sheets from workbook. Returns list of {name, note, items}."""
    shopping_lists = []
    for sheet_name in wb.sheetnames:
        nl = sheet_name.lower()
        if 'shopping' not in nl:
            continue
        ws = wb[sheet_name]
        rows = [r for r in ws.iter_rows(values_only=True) if any(r)]
        if not rows:
            continue
        header = rows[0]
        col2 = str(header[1]) if len(header) > 1 and header[1] else ''
        note = ''
        if '*' in col2:
            parts = col2.split('*')
            note = parts[1].strip().strip('*').strip() if len(parts) > 1 else ''
        if not note:
            note = 'assumes each meal prepared once'
        items = []
        for row in rows[1:]:
            food = str(row[0]).strip() if row[0] else ''
            qty  = row[1] if len(row) > 1 and row[1] is not None else ''
            if not food or food.lower() in ('none', 'nan', 'food item'):
                continue
            try:
                qty_val = float(qty)
                qty_str = f"{int(qty_val)}g" if qty_val == int(qty_val) else f"{qty_val}g"
            except:
                qty_str = str(qty) if qty else ''
            items.append({'food': food, 'qty': qty_str})
        if items:
            shopping_lists.append({'name': sheet_name, 'note': note, 'items': items})
    return shopping_lists


def find_ambiguous(days):
    seen, result = set(), []
    for day in days:
        for meal in day['meals']:
            for ing in meal['ingredients']:
                a = check_ambiguous(ing['qty_g'], ing['food'])
                if a and a['key'] not in seen:
                    seen.add(a['key'])
                    result.append(a)
    return result


def apply_clarifications(days, clarifications):
    for day in days:
        for meal in day['meals']:
            for ing in meal['ingredients']:
                key = ing['food'].lower().strip()
                if key in clarifications:
                    ing['qty_label'] = clarifications[key]


def has_fruit(days):
    return any(is_fruit(ing['food']) for day in days for meal in day['meals'] for ing in meal['ingredients'])

def day_has_fruit(day):
    return any(is_fruit(ing['food']) for meal in day['meals'] for ing in meal['ingredients'])


# ── AI Recipe Generation ──────────────────────────────────────────────────────
def _clean_json(raw):
    """Clean and parse JSON from Claude, handling common issues."""
    raw = raw.strip()
    # Strip markdown fences
    raw = re.sub(r'^```json\s*', '', raw, flags=re.MULTILINE)
    raw = re.sub(r'^```\s*', '', raw, flags=re.MULTILINE)
    raw = re.sub(r'\s*```$', '', raw)
    # Extract JSON object
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    if m:
        raw = m.group(0)
    # Normalise unicode punctuation
    for old, new in [('‘',"'"),('’',"'"),('“','"'),('”','"'),
                     ('–','-'),('—','-'),('…','...')]:
        raw = raw.replace(old, new)
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # Strip non-ASCII as last resort
        raw = raw.encode('ascii', 'replace').decode('ascii')
        return json.loads(raw)


def _generate_batch(claude, batch_meals, batch_num):
    """Generate recipes for a batch of meals. Returns list of meal dicts."""
    meals_text = []
    for day_num, meal_label, ings_text in batch_meals:
        meals_text.append(f"Day {day_num} — {meal_label}:\n{ings_text}")

    prompt = """You are a nutrition coach writing a recipe guide for a fitness client.

For each meal, generate a dish name and step-by-step cooking instructions.

Respond ONLY with valid JSON — no markdown, no code fences:
{"meals": [{"day": 1, "meal_label": "Breakfast", "dish_name": "Creamy Porridge", "steps": ["Step 1.", "Step 2."], "unclear": false, "unclear_reason": ""}]}

Rules:
- Steps should be practical and motivating, written directly to the client
- Single-item meals (protein water, bar, fruit) need only 1-2 steps
- Set unclear to true only if you genuinely cannot determine the dish
- Never invent ingredients not in the list
- No nutrition advice in steps
- Use only plain ASCII characters — no smart quotes or special dashes

Meals:

""" + '\n\n'.join(meals_text)

    response = claude.messages.create(
        model='claude-sonnet-4-20250514',
        max_tokens=6000,
        messages=[{'role': 'user', 'content': prompt}]
    )
    data = _clean_json(response.content[0].text)
    return data.get('meals', [])


def generate_recipes_ai(days):
    claude = get_claude()

    # Build flat list of all meals with their ingredients text
    all_meals = []
    for day in days:
        for meal in day['meals']:
            ings = []
            for ing in meal['ingredients']:
                label = ing.get('qty_label') or (
                    f"{int(ing['qty_g'])}g" if ing['qty_g'] == int(ing['qty_g'])
                    else f"{ing['qty_g']}g"
                )
                ings.append(f"  - {label} {ing['food']}")
            all_meals.append((day['day_num'], meal['meal_label'], '\n'.join(ings)))

    # Process in batches of 6 meals to avoid token limit truncation
    BATCH_SIZE = 6
    ai_results = []
    for i in range(0, len(all_meals), BATCH_SIZE):
        batch = all_meals[i:i + BATCH_SIZE]
        try:
            results = _generate_batch(claude, batch, i // BATCH_SIZE + 1)
            ai_results.extend(results)
        except Exception as e:
            # If a batch fails, add empty placeholders so other batches still work
            for day_num, meal_label, _ in batch:
                ai_results.append({
                    'day': day_num,
                    'meal_label': meal_label,
                    'dish_name': meal_label,
                    'steps': ['Prepare ingredients as listed and enjoy.'],
                    'unclear': False,
                    'unclear_reason': '',
                })

    ai_map = {(item['day'], item['meal_label']): item for item in ai_results}

    unclear_meals = []
    for day in days:
        for meal in day['meals']:
            key = (day['day_num'], meal['meal_label'])
            if key in ai_map:
                ai = ai_map[key]
                meal['dish_name'] = ai.get('dish_name', meal['meal_label'])
                meal['recipe_steps'] = ai.get('steps', [])
                if ai.get('unclear'):
                    unclear_meals.append({
                        'day': day['day_num'],
                        'meal_label': meal['meal_label'],
                        'unclear_reason': ai.get('unclear_reason', ''),
                        'ingredients': [{'food': i['food'], 'qty_g': i['qty_g']} for i in meal['ingredients']],
                    })

    return days, unclear_meals


# ── PDF Generation ────────────────────────────────────────────────────────────
PW, PH = A4
LM, RM = 50.0, 545.28
CW = RM - LM
BLACK    = (0.08, 0.08, 0.08)
WHITE    = (1.0,  1.0,  1.0)
OFFWHITE = (0.96, 0.96, 0.96)
MID_GREY = (0.45, 0.45, 0.45)
DIVIDER  = (0.80, 0.80, 0.80)
DARK_CARD= (0.12, 0.12, 0.12)
# Use Inter if available, fall back to Helvetica
HB = 'Inter-Bold' if INTER_AVAILABLE else 'Helvetica-Bold'
HXB = 'Inter-ExtraBold' if INTER_AVAILABLE else 'Helvetica-Bold'
H  = 'Inter' if INTER_AVAILABLE else 'Helvetica'

def pdf_y(t): return PH - t
def tw(t, f, s): return stringWidth(t, f, s)

def draw_text(c, x, top_y, text, font, size, rgb):
    c.saveState()
    c.setFillColorRGB(*rgb)
    c.setFont(font, size)
    c.drawString(x, pdf_y(top_y + size), text)
    c.restoreState()

def draw_centred(c, cx, top_y, text, font, size, rgb):
    draw_text(c, cx - tw(text, font, size)/2, top_y, text, font, size, rgb)

def rrect(c, x, bot, w, h, r, fill=None):
    c.saveState()
    c.setLineWidth(0)
    if fill: c.setFillColorRGB(*fill)
    p = c.beginPath()
    p.moveTo(x+r, bot); p.lineTo(x+w-r, bot)
    p.curveTo(x+w, bot, x+w, bot, x+w, bot+r)
    p.lineTo(x+w, bot+h-r)
    p.curveTo(x+w, bot+h, x+w, bot+h, x+w-r, bot+h)
    p.lineTo(x+r, bot+h)
    p.curveTo(x, bot+h, x, bot+h, x, bot+h-r)
    p.lineTo(x, bot+r)
    p.curveTo(x, bot, x, bot, x+r, bot)
    p.close()
    c.drawPath(p, fill=1, stroke=0)
    c.restoreState()

def wrap(text, font, size, max_w):
    words = text.split()
    lines, cur = [], ''
    for word in words:
        test = (cur + ' ' + word).strip()
        if tw(test, font, size) <= max_w: cur = test
        else:
            if cur: lines.append(cur)
            cur = word
    if cur: lines.append(cur)
    return lines

FRUIT_NOTE = ("Note: Any fruit included in today's meals can be eaten as a snack at any point "
              "throughout the day — it pairs particularly well with Greek yoghurt. "
              "This won't impact your results.")

def draw_cover(c, client_name, logo_b64):
    c.setFillColorRGB(*BLACK)
    c.rect(0, 0, PW, PH, fill=1, stroke=0)

    # Layout anchored to vertical centre of page
    # Page height = 841.89pt. Centre = ~421pt from top.
    # Logo sits above centre, name below.
    sz = 100
    logo_top = 280   # top of logo (from top of page)
    logo_bot = logo_top + sz  # 380

    if logo_b64:
        try:
            img_data = base64.b64decode(logo_b64)
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                tmp.write(img_data); tmp_path = tmp.name
            c.drawImage(tmp_path, PW/2 - sz/2, pdf_y(logo_bot), width=sz, height=sz, mask='auto')
            os.unlink(tmp_path)
        except: pass

    # PROJECT GAIN wordmark — 20pt below logo
    draw_centred(c, PW/2, logo_bot + 14, 'PROJECT GAIN', HB, 9, (0.45, 0.45, 0.45))

    # Rule — 16pt below wordmark
    rule_y = logo_bot + 38
    c.saveState()
    c.setStrokeColorRGB(0.25, 0.25, 0.25); c.setLineWidth(0.5)
    c.line(LM + 80, pdf_y(rule_y), RM - 80, pdf_y(rule_y))
    c.restoreState()

    # Client name — 16pt below rule
    draw_centred(c, PW/2, rule_y + 14, client_name, HXB, 30, WHITE)

    # Subtitle — 20pt below name
    draw_centred(c, PW/2, rule_y + 62, 'Personalised Meal Plan + Recipe Guide', H, 11, (0.45, 0.45, 0.45))

    # Footer
    draw_centred(c, PW/2, PH - 50, 'projectgainofficial.com', H, 8, (0.3, 0.3, 0.3))

def draw_day_header(c, top_y, day_num, kcal, prot, fat, carb, fruit_note):
    box_h = 96.0
    rrect(c, LM, pdf_y(top_y+box_h), CW, box_h, r=6, fill=BLACK)
    draw_centred(c, PW/2, top_y+14, f'Day {day_num}', HXB, 22, WHITE)
    pill_w, pill_h, pill_top = 360.0, 26.0, top_y+44
    rrect(c, PW/2-pill_w/2, pdf_y(pill_top+pill_h), pill_w, pill_h, r=5, fill=(0.2, 0.2, 0.2))
    draw_centred(c, PW/2, pill_top+8, f'{kcal:,} kcal  ·  {prot}g Protein  ·  {fat}g Fats  ·  {carb}g Carbs', HB, 8.5, WHITE)
    y = top_y + box_h + 12
    if fruit_note:
        lines = wrap(FRUIT_NOTE, H, 8.5, CW-20)
        note_h = len(lines)*12+14
        rrect(c, LM, pdf_y(y+note_h), CW, note_h, r=4, fill=OFFWHITE)
        rrect(c, LM, pdf_y(y+note_h), 3, note_h, r=2, fill=DARK_CARD)
        ny = y+7
        for line in lines:
            draw_text(c, LM+12, ny, line, H, 8.5, MID_GREY); ny += 12
        y += note_h+10
    return y

def draw_meal_card(c, top_y, meal_label, dish_name, kcal, prot, fat, carb, ingredients, steps):
    hdr_h = 44.0
    rrect(c, LM, pdf_y(top_y+hdr_h), CW, hdr_h, r=6, fill=DARK_CARD)
    bw, bh, bx, bt = 52.0, 16.0, LM+10, top_y+14
    rrect(c, bx, pdf_y(bt+bh), bw, bh, r=3, fill=(0.28, 0.28, 0.28))
    draw_centred(c, bx+bw/2, bt+4, meal_label.upper(), HB, 7, WHITE)
    if dish_name:
        # Truncate dish name if too wide
        max_dish_w = RM - (bx + bw + 20)
        dish_display = dish_name
        while dish_display and tw(dish_display, HB, 13) > max_dish_w:
            dish_display = dish_display[:-1]
        if dish_display != dish_name: dish_display += '...'
        draw_text(c, bx+bw+12, top_y+15, dish_display, HXB, 13, WHITE)

    y = top_y + hdr_h
    macro_h = 38.0
    rrect(c, LM, pdf_y(y+macro_h), CW, macro_h, r=0, fill=OFFWHITE)
    divs = [LM+CW*0.25, LM+CW*0.5, LM+CW*0.75]
    centres = [LM+CW*0.125, LM+CW*0.375, LM+CW*0.625, LM+CW*0.875]
    c.saveState()
    c.setStrokeColorRGB(*DIVIDER); c.setLineWidth(0.4)
    for dx in divs: c.line(dx, pdf_y(y+8), dx, pdf_y(y+30))
    c.restoreState()
    for cx, val, lab in zip(centres, [f'{kcal} kcal', f'{prot}g', f'{fat}g', f'{carb}g'], ['Calories','Protein','Fats','Carbs']):
        draw_centred(c, cx, y+7, val, HB, 11, BLACK)
        draw_centred(c, cx, y+20, lab, H, 7, MID_GREY)
    y += macro_h + 18

    draw_text(c, LM, y, 'INGREDIENTS', HB, 8.5, BLACK)
    y += 14
    for ing in ingredients:
        label = ing.get('qty_label')
        if label:
            draw_text(c, LM+8, y, label, HB, 9.5, BLACK)
        else:
            qty_g = ing['qty_g']
            qs = f"{int(qty_g)}g" if qty_g == int(qty_g) else f"{qty_g}g"
            draw_text(c, LM+8, y, qs, HB, 9.5, BLACK)
            draw_text(c, LM+8+tw(qs, HB, 9.5)+5.5, y, ing['food'], H, 9.5, BLACK)
        y += 14

    y += 6
    c.saveState(); c.setStrokeColorRGB(*DIVIDER); c.setLineWidth(0.4)
    c.line(LM, pdf_y(y), RM, pdf_y(y)); c.restoreState()
    y += 8

    if steps:
        draw_text(c, LM, y, 'METHOD', HB, 8.5, BLACK)
        y += 14
        max_w = RM - (LM+24)
        for i, step in enumerate(steps, 1):
            draw_text(c, LM+8, y, f'{i}.', HB, 9.5, BLACK)
            lines = wrap(step, H, 9.5, max_w)
            for line in lines:
                draw_text(c, LM+24, y, line, H, 9.5, BLACK); y += 13
            y += 2
    y += 12
    return y

def draw_shopping_list(c, shopping_list):
    """Draw a shopping list — rows are 15pt, packed tightly with no gaps."""
    ROW_H = 15

    def draw_header(y):
        draw_text(c, LM, y, shopping_list['name'].upper(), HXB, 18, BLACK)
        y += 28
        note_text = f"* {shopping_list['note'].capitalize()}"
        draw_text(c, LM, y, note_text, H, 9, MID_GREY)
        y += 16
        c.saveState(); c.setStrokeColorRGB(*DIVIDER); c.setLineWidth(0.4)
        c.line(LM, pdf_y(y), RM, pdf_y(y)); c.restoreState()
        y += 12
        draw_text(c, LM, y, 'INGREDIENT', HB, 8, MID_GREY)
        draw_text(c, RM - 60, y, 'QUANTITY', HB, 8, MID_GREY)
        y += 12
        return y

    y = 48.0
    y = draw_header(y)

    for i, item in enumerate(shopping_list['items']):
        # Page break — continue with no header repeat, just items
        if y + ROW_H > PH - 40:
            c.showPage()
            y = 48.0

        if i % 2 == 0:
            rrect(c, LM, pdf_y(y + ROW_H), CW, ROW_H, r=0, fill=OFFWHITE)

        draw_text(c, LM + 8, y + 2, item['food'], H, 9.5, BLACK)
        qty_w = tw(item['qty'], HB, 9.5)
        draw_text(c, RM - qty_w - 8, y + 2, item['qty'], HB, 9.5, BLACK)
        y += ROW_H


def generate_pdf_doc(client_name, days, logo_b64, shopping_lists=None):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle(f"{client_name} — Meal Plan + Recipe Guide")
    c.setAuthor("Project GAIN")
    draw_cover(c, client_name, logo_b64)
    c.showPage()
    any_fruit = has_fruit(days)
    first_day = True
    y = 48.0

    for day in days:
        # Estimate header height
        hdr_h = 96 + (50 if any_fruit and day_has_fruit(day) else 0) + 12

        # If not at top of page and header won't fit, start new page
        if not first_day:
            if y + hdr_h + 150 > PH - 40:
                c.showPage()
                y = 48.0
            else:
                y += 28  # gap between days on same page

        first_day = False

        y = draw_day_header(c, y, day['day_num'],
                            day['total_kcal'], day['total_prot'], day['total_fat'], day['total_carb'],
                            fruit_note=any_fruit and day_has_fruit(day))
        y += 14

        for meal in day['meals']:
            steps = meal.get('recipe_steps', [])
            n_step_lines = sum(len(wrap(s, H, 9.5, RM-(LM+24))) for s in steps)
            # Accurate estimate matching draw_meal_card exactly:
            # header(44) + macro(38) + gap(18) + ING_label(14) + rows + divider(14) + METHOD_label(14) + steps + trailing(12)
            card_h = 44+38+18+14+(len(meal['ingredients'])*14)+14+14+(n_step_lines*13)+(len(steps)*2)+12
            if y + card_h > PH - 40:
                c.showPage(); y = 48.0
            y = draw_meal_card(c, y, meal['meal_label'], meal.get('dish_name',''),
                               meal['kcal'], meal['prot'], meal['fat'], meal['carb'],
                               meal['ingredients'], steps)
            y += 14
    # Shopping list pages — always start on a new page
    if shopping_lists:
        c.showPage()
        for i, sl in enumerate(shopping_lists):
            if i > 0:
                c.showPage()
            draw_shopping_list(c, sl)

    c.save(); buf.seek(0)
    return buf.read()


# ── API Routes ────────────────────────────────────────────────────────────────

@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    f = request.files['file']
    if not f.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'error': 'Please upload an Excel file (.xlsx)'}), 400
    file_bytes = f.read()
    try:
        client_name, days, shopping_lists = parse_excel(file_bytes, f.filename)
    except Exception as e:
        return jsonify({'error': f'Could not parse file: {str(e)}'}), 400

    ambiguous = find_ambiguous(days)

    # Instant size suggestions — no API call, runs in microseconds
    ambiguous = lookup_size_suggestions(ambiguous)

    import uuid
    sid = str(uuid.uuid4())
    _sessions[sid] = {'client_name': client_name, 'days': days, 'shopping_lists': shopping_lists}
    return jsonify({'session_id': sid, 'client_name': client_name, 'day_count': len(days), 'ambiguous': ambiguous})


@app.route('/clarify', methods=['POST'])
def clarify():
    data = request.get_json()
    sid = data.get('session_id')
    if sid not in _sessions:
        return jsonify({'error': 'Session expired. Please re-upload.'}), 400

    sess = _sessions[sid]
    apply_clarifications(sess['days'], data.get('clarifications', {}))

    try:
        days, unclear_meals = generate_recipes_ai(sess['days'])
    except Exception as e:
        import traceback
        return jsonify({'error': f'Recipe generation failed: {str(e)}', 'detail': traceback.format_exc()}), 500

    sess['days'] = days

    review_days = []
    for day in days:
        review_days.append({
            'day_num': day['day_num'],
            'total_kcal': day['total_kcal'], 'total_prot': day['total_prot'],
            'total_fat': day['total_fat'],   'total_carb': day['total_carb'],
            'meals': [{
                'meal_label': m['meal_label'],
                'dish_name':  m.get('dish_name', ''),
                'kcal': m['kcal'], 'prot': m['prot'], 'fat': m['fat'], 'carb': m['carb'],
                'ingredients': [{'food': i['food'], 'qty_g': i['qty_g'],
                                 'qty_label': i.get('qty_label'), 'excel_qty_g': i.get('excel_qty_g', i['qty_g'])} for i in m['ingredients']],
                'recipe_steps': m.get('recipe_steps', []),
                'needs_clarification': m.get('needs_clarification', False),
            } for m in day['meals']],
        })

    return jsonify({'session_id': sid, 'review_days': review_days, 'unclear_meals': unclear_meals})


@app.route('/generate-pdf', methods=['POST'])
def generate_pdf_route():
    data = request.get_json()
    sid = data.get('session_id')
    approved_days = data.get('approved_days', [])
    logo_b64 = data.get('logo_b64', '')

    if sid not in _sessions:
        return jsonify({'error': 'Session expired. Please re-upload.'}), 400

    sess = _sessions[sid]
    days = sess['days']
    client_name = sess['client_name']

    # Merge approved edits + verify quantities
    qty_issues = []
    for i, day in enumerate(days):
        if i >= len(approved_days): continue
        adv = approved_days[i]
        for j, meal in enumerate(day['meals']):
            if j >= len(adv.get('meals', [])): continue
            am = adv['meals'][j]
            meal['dish_name']    = am.get('dish_name', meal.get('dish_name', ''))
            meal['recipe_steps'] = am.get('recipe_steps', meal.get('recipe_steps', []))
            for k, ing in enumerate(meal['ingredients']):
                excel_qty = ing.get('excel_qty_g', ing['qty_g'])
                if k < len(am.get('ingredients', [])):
                    sub_qty = float(am['ingredients'][k].get('qty_g', ing['qty_g']))
                    if abs(excel_qty - sub_qty) > 0.01:
                        qty_issues.append({'day': day['day_num'], 'meal': meal['meal_label'],
                                           'food': ing['food'], 'excel_qty': excel_qty, 'submitted_qty': sub_qty})

    if qty_issues:
        return jsonify({'qty_issues': qty_issues, 'error': 'Quantity mismatch detected — please review.'}), 422

    shopping_lists = sess.get('shopping_lists', [])
    try:
        pdf_bytes = generate_pdf_doc(client_name, days, logo_b64, shopping_lists)
    except Exception as e:
        import traceback
        return jsonify({'error': f'PDF failed: {str(e)}', 'detail': traceback.format_exc()}), 500

    del _sessions[sid]
    return jsonify({'pdf_b64': base64.b64encode(pdf_bytes).decode(),
                    'filename': f"{client_name.replace(' ', '_')}_Meal_Plan.pdf"})


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
