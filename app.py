"""
Xelix Brand Aligner — web app
=============================
Upload one Xelix deck, align it to the brand guidelines, download the result.
One comprehensive pass: embeds Barlow so the deck renders correctly everywhere,
sets Barlow typography, ensures one correctly-coloured wordmark per slide
(without doubling an existing one), and repairs any dark-on-dark / light-on-light
text. Backgrounds and layouts are left intact.
"""
import os, uuid, tempfile, traceback
from flask import Flask, render_template, request, jsonify, send_file
from normalizer import normalize_deck

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_BLUE = os.path.join(HERE, 'xelix_blue.png')
LOGO_WHITE = os.path.join(HERE, 'xelix_white.png')
OUT_DIR = os.path.join(tempfile.gettempdir(), 'xelix_aligned')
os.makedirs(OUT_DIR, exist_ok=True)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024   # 300 MB — 100+ slide decks


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/align', methods=['POST'])
def align():
    if 'file' not in request.files:
        return jsonify(success=False, error='No file uploaded.'), 400
    f = request.files['file']
    if not f.filename or not f.filename.lower().endswith('.pptx'):
        return jsonify(success=False, error='Please upload a .pptx file.'), 400

    file_id = uuid.uuid4().hex
    in_path = os.path.join(OUT_DIR, file_id + '_in.pptx')
    out_path = os.path.join(OUT_DIR, file_id + '.pptx')
    f.save(in_path)
    try:
        report = normalize_deck(in_path, out_path,
                                logo_blue=LOGO_BLUE, logo_white=LOGO_WHITE)
    except Exception:
        traceback.print_exc()
        return jsonify(success=False,
                       error='Could not process this deck. It may be corrupt '
                             'or password-protected.'), 500
    finally:
        if os.path.exists(in_path):
            os.remove(in_path)

    return jsonify(
        success=True, file_id=file_id,
        slides=len(report['slides']),
        logos_added=report.get('logos_added', 0),
        logos_kept=report.get('logos_kept', 0),
        fonts_embedded=report.get('fonts_embedded', False),
        warnings=report.get('warnings', []),
        download_name=os.path.splitext(f.filename)[0] + '_on-brand.pptx',
    )


@app.route('/api/download/<file_id>')
def download(file_id):
    if not file_id.isalnum():
        return 'Not found', 404
    path = os.path.join(OUT_DIR, file_id + '.pptx')
    if not os.path.exists(path):
        return 'File not found or expired.', 404
    name = request.args.get('name', 'deck_on-brand.pptx')
    if not name.lower().endswith('.pptx'):
        name = 'deck_on-brand.pptx'
    return send_file(path, as_attachment=True, download_name=name,
                     mimetype='application/vnd.openxmlformats-officedocument'
                              '.presentationml.presentation')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 8080)), debug=False)
