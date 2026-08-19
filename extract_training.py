import sys, io, os, json, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import docx

TRAINING_FOLDER = r'C:\Users\C5407836\litmos\General\OneDrive_1_8-19-2026'

def clean(text):
    return re.sub(r'\s+', ' ', text.strip())

def extract_docx(filepath):
    doc = docx.Document(filepath)
    source = os.path.splitext(os.path.basename(filepath))[0]
    # Strip "TA - " prefix for cleaner display
    source_label = re.sub(r'^TA\s*-\s*', '', source).strip()

    chunks = []
    cur_h1 = ''
    cur_h2 = ''
    cur_lines = []

    def save_chunk():
        content = ' '.join(l for l in cur_lines if l)
        if content and len(content) > 20:
            chunks.append({
                'topic': cur_h1,
                'section': cur_h2,
                'content': content,
                'source': source_label
            })

    for para in doc.paragraphs:
        text = clean(para.text)
        if not text:
            continue
        style = para.style.name if para.style else 'Normal'

        if 'Heading 1' in style:
            save_chunk()
            cur_h1 = text
            cur_h2 = ''
            cur_lines = []
        elif 'Heading 2' in style:
            save_chunk()
            cur_h2 = text
            cur_lines = []
        elif 'Heading 3' in style:
            save_chunk()
            cur_h2 = text  # treat h3 as subsection
            cur_lines = []
        else:
            cur_lines.append(text)

    save_chunk()
    return chunks

all_chunks = []
for f in sorted(os.listdir(TRAINING_FOLDER)):
    if f.endswith('.docx'):
        fp = os.path.join(TRAINING_FOLDER, f)
        chunks = extract_docx(fp)
        print(f'{f}: {len(chunks)} sections')
        all_chunks.extend(chunks)

for i, c in enumerate(all_chunks, 1):
    c['id'] = i

print(f'\nTotal: {len(all_chunks)} sections from {len(os.listdir(TRAINING_FOLDER))} files')

js = 'const TRAINING = ' + json.dumps(all_chunks, ensure_ascii=False) + ';'
with open(r'C:\Users\C5407836\litmos-app\training.js', 'w', encoding='utf-8') as f:
    f.write(js)

print('Saved training.js')
