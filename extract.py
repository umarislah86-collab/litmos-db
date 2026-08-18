import sys, io, os, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx
from pypdf import PdfReader

CHECKMARK = '\uf058'
CIRCLE = '\uf111'
BULLET = '\uf0b7'
NULL = '\x00'
ICON_RE = re.compile('[\uf058\uf111\uf0b7\x00]')

NOISE_PATTERNS = [
    r'^Weight: \d+$', r'^\d+/\d+/\d+\,', r'^https?://', r'ASSESSMENT RESULTS',
    r'was completed by', r'^Attempts$', r'^Time Taken$', r'Score \(Passmark',
    r'^\d+% - (Passed|Failed)', r'^Correct$', r'^Incorrect$', r'^\(/home\)$',
    r'/assessment/results', r'/assessment/review', r'/assessment/question',
    r'^\d{2}:\d{2}:\d{2}', r'TRAINING ACHIEVEMENT', r'SECURITY TRAINING',
    r'HCSM Q&A', r'^Review your answers$', r'You have reached the end',
    r'^Time: \d+ minutes', r'HomeContent Library', r'Live Sessions',
    r'Assessment Results\s*\x00', r'^\x00\s*(WD|IT)\s*\x00',
]

Q_KEYWORDS = re.compile(
    r'\b(what|which|how|when|where|who|why|select|choose|please|correct|identify|'
    r'is it|can|should|do we|are we|would you|describe|provide|enter|name|'
    r'define|explain|true or false|yes or no)\b', re.IGNORECASE)


def is_noise(s):
    for p in NOISE_PATTERNS:
        if re.search(p, s, re.IGNORECASE):
            return True
    return False


def normalize(text):
    return re.sub(r'\s+', ' ', text.lower().strip())


def clean_text(text):
    cleaned = ICON_RE.sub('', text)
    return re.sub(r'\s+', ' ', cleaned).strip()


def strip_option_icons(text):
    return re.sub(r'^[\uf058\uf111\uf0b7\x00\s]+', '', text).strip()


def is_real_question(text):
    if text.endswith('?') or text.endswith(':'):
        return True
    return bool(Q_KEYWORDS.search(text))


def is_correct_option(stripped):
    # Icon format: starts with checkmark
    if stripped.startswith(CHECKMARK):
        return True
    # Null-byte format: starts with double null (with or without space between)
    if stripped.startswith(NULL + ' ' + NULL) or stripped.startswith(NULL + NULL):
        return True
    return False


def is_wrong_option(stripped):
    # Icon format: starts with circle only (no checkmark)
    if stripped.startswith(CIRCLE) and not stripped.startswith(CHECKMARK):
        return True
    # Null-byte format: starts with single null only
    if stripped.startswith(NULL) and not is_correct_option(stripped):
        return True
    return False


def looks_like_fragment(text):
    if len(text) > 500:
        return True
    if len(text) < 20:
        return True
    if text[0].islower():
        return True
    return False


def parse_docx(filepath, category):
    doc = docx.Document(filepath)
    questions = []
    cur_q = None
    cur_a_lines = []
    for para in doc.paragraphs:
        text = clean_text(para.text)
        if not text:
            continue
        tl = text.lower()
        if tl.startswith('question:'):
            if cur_q and cur_a_lines:
                questions.append({'question': cur_q, 'answer': ' '.join(cur_a_lines), 'category': category})
            cur_q = text[len('question:'):].strip()
            cur_a_lines = []
        elif tl.startswith('answer:'):
            cur_a_lines = [text[len('answer:'):].strip()]
        elif cur_q and cur_a_lines is not None:
            cur_a_lines.append(text)
    if cur_q and cur_a_lines:
        questions.append({'question': cur_q, 'answer': ' '.join(cur_a_lines), 'category': category})
    return questions


def parse_pdf_assessment(filepath, category):
    reader = PdfReader(filepath)
    questions = []
    all_lines = []
    for page in reader.pages:
        t = page.extract_text(extraction_mode='layout')
        if t:
            all_lines.extend(t.split('\n'))

    cur_q_parts = []
    cur_correct = []
    in_options = False

    def save_q():
        if not cur_q_parts:
            return
        q = clean_text(' '.join(cur_q_parts).strip())
        a_parts = [strip_option_icons(x) for x in cur_correct]
        a_parts = [clean_text(x) for x in a_parts if x.strip()]
        a = ' | '.join(a_parts)
        if len(q) > 20 and a and not looks_like_fragment(q) and is_real_question(q):
            questions.append({'question': q, 'answer': a, 'category': category})

    for line in all_lines:
        stripped = line.strip()
        if not stripped or is_noise(stripped):
            continue
        if is_correct_option(stripped):
            in_options = True
            cur_correct.append(stripped)
        elif is_wrong_option(stripped):
            in_options = True
        else:
            if in_options:
                save_q()
                cur_q_parts = [stripped]
                cur_correct = []
                in_options = False
            else:
                cur_q_parts.append(stripped)
    save_q()
    return questions


def parse_pdf_review(filepath, category):
    reader = PdfReader(filepath)
    all_lines = []
    for page in reader.pages:
        t = page.extract_text(extraction_mode='layout')
        if t:
            all_lines.extend(t.split('\n'))

    questions = []
    cur_q_parts = []
    cur_a_parts = []
    in_answer = False

    def save_q():
        if not cur_q_parts:
            return
        q = clean_text(' '.join(cur_q_parts).strip())
        q = re.sub(r'^\d+\.\s*', '', q)
        a_parts = [clean_text(x) for x in cur_a_parts if clean_text(x)]
        a = ' | '.join(a_parts)
        if len(q) > 20 and a and not looks_like_fragment(q) and is_real_question(q):
            questions.append({'question': q, 'answer': a, 'category': category})

    for line in all_lines:
        stripped = line.strip()
        if not stripped or is_noise(stripped):
            continue
        if stripped.startswith('\u2191') or 'Change this answer' in stripped:
            save_q()
            cur_q_parts = []
            cur_a_parts = []
            in_answer = False
            continue
        if re.match(r'^\d+\.', stripped):
            if cur_q_parts:
                save_q()
                cur_q_parts = []
                cur_a_parts = []
            cur_q_parts = [stripped]
            in_answer = False
        elif cur_q_parts and not in_answer:
            q_so_far = ' '.join(cur_q_parts)
            if q_so_far.rstrip().endswith('?') or q_so_far.rstrip().endswith('.'):
                in_answer = True
                cur_a_parts.append(clean_text(stripped))
            else:
                cur_q_parts.append(stripped)
        elif in_answer:
            cur_a_parts.append(clean_text(stripped))
    save_q()
    return questions


def parse_pdf(filepath, category):
    reader = PdfReader(filepath)
    first_page = (reader.pages[0].extract_text(extraction_mode='layout') or
                  reader.pages[0].extract_text() or '')
    if 'Review your answers' in first_page or 'Change this answer' in first_page:
        return parse_pdf_review(filepath, category)
    else:
        return parse_pdf_assessment(filepath, category)


all_questions = []

for fp, cat in [
    (r'C:\Users\C5407836\litmos\Security\Security.docx', 'Security'),
    (r'C:\Users\C5407836\litmos\HCSM\HCSM.docx', 'HCSM'),
]:
    qs = parse_docx(fp, cat)
    print(f'DOCX {cat}: {len(qs)} Qs')
    all_questions.extend(qs)

for cat, folder in [
    ('General', r'C:\Users\C5407836\litmos\General'),
    ('Security', r'C:\Users\C5407836\litmos\Security'),
    ('HCSM', r'C:\Users\C5407836\litmos\HCSM'),
]:
    for f in sorted(os.listdir(folder)):
        if f.endswith('.pdf'):
            fp = os.path.join(folder, f)
            qs = parse_pdf(fp, cat)
            print(f'PDF {cat}/{f}: {len(qs)} Qs')
            all_questions.extend(qs)

print(f'\nTotal before dedup: {len(all_questions)}')

seen = {}
deduped = []
for q in all_questions:
    key = normalize(q['question'])
    if key not in seen:
        seen[key] = True
        deduped.append(q)

for i, q in enumerate(deduped, 1):
    q['id'] = i

g = sum(1 for q in deduped if q['category'] == 'General')
s = sum(1 for q in deduped if q['category'] == 'Security')
h = sum(1 for q in deduped if q['category'] == 'HCSM')
print(f'After dedup: {len(deduped)} total | General:{g} Security:{s} HCSM:{h}')

long_qs = [q for q in deduped if len(q['question']) > 500]
print(f'Long questions (>500 chars): {len(long_qs)}')
for q in long_qs[:3]:
    print(f'  id={q["id"]} [{q["category"]}] len={len(q["question"])}: {q["question"][:80]}')

os.makedirs(r'C:\Users\C5407836\litmos-app\data', exist_ok=True)
with open(r'C:\Users\C5407836\litmos-app\data\questions.json', 'w', encoding='utf-8') as f:
    json.dump(deduped, f, ensure_ascii=False, indent=2)

js = 'const QUESTIONS = ' + json.dumps(deduped, ensure_ascii=False) + ';'
with open(r'C:\Users\C5407836\litmos-app\questions.js', 'w', encoding='utf-8') as f:
    f.write(js)

print('Saved questions.json and questions.js')
