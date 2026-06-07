"""
Kindle SaaS MVP - Flask Web App
Word (.docx) をアップロード → 縦書きEPUBをダウンロード
"""
import os
import uuid
import tempfile
from pathlib import Path
from flask import Flask, request, send_file, jsonify, render_template_string

from docx_to_epub import docx_to_epub

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB上限

ALLOWED_EXT = {'.docx'}

HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>縦書きEPUB変換 - Kindle出版かんたんツール</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: "Hiragino Sans", "Yu Gothic", sans-serif;
      background: #fafafa;
      color: #222;
      min-height: 100vh;
      display: flex;
      flex-direction: column;
      align-items: center;
      padding: 40px 16px;
    }
    .card {
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 2px 16px rgba(0,0,0,0.08);
      padding: 40px;
      max-width: 560px;
      width: 100%;
    }
    h1 {
      font-size: 1.4em;
      font-weight: bold;
      margin-bottom: 8px;
    }
    .subtitle {
      color: #666;
      font-size: 0.9em;
      margin-bottom: 32px;
      line-height: 1.6;
    }
    .drop-zone {
      border: 2px dashed #ccc;
      border-radius: 8px;
      padding: 40px 20px;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
      margin-bottom: 20px;
    }
    .drop-zone.dragover {
      border-color: #4f46e5;
      background: #eef2ff;
    }
    .drop-zone input { display: none; }
    .drop-zone .icon { font-size: 2.5em; margin-bottom: 10px; }
    .drop-zone .label { color: #555; font-size: 0.95em; }
    .drop-zone .file-name {
      margin-top: 10px;
      font-weight: bold;
      color: #4f46e5;
    }
    label.meta { display: block; font-size: 0.85em; margin-bottom: 4px; color: #444; }
    input[type=text] {
      width: 100%;
      padding: 8px 12px;
      border: 1px solid #ddd;
      border-radius: 6px;
      font-size: 0.95em;
      margin-bottom: 16px;
    }
    button[type=submit] {
      width: 100%;
      background: #4f46e5;
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 14px;
      font-size: 1em;
      font-weight: bold;
      cursor: pointer;
      transition: background 0.2s;
    }
    button[type=submit]:hover { background: #4338ca; }
    button[type=submit]:disabled { background: #a5b4fc; cursor: not-allowed; }
    .result {
      margin-top: 24px;
      padding: 16px;
      border-radius: 8px;
      font-size: 0.9em;
      display: none;
    }
    .result.success {
      background: #f0fdf4;
      border: 1px solid #86efac;
      color: #166534;
    }
    .result.error {
      background: #fef2f2;
      border: 1px solid #fca5a5;
      color: #991b1b;
    }
    .result a {
      display: inline-block;
      margin-top: 12px;
      background: #16a34a;
      color: #fff;
      padding: 10px 24px;
      border-radius: 6px;
      text-decoration: none;
      font-weight: bold;
    }
    .warnings { margin-top: 8px; color: #92400e; font-size: 0.85em; }
    .loading { display: none; text-align: center; color: #666; margin-top: 16px; }
    .features {
      margin-top: 32px;
      font-size: 0.82em;
      color: #888;
      line-height: 1.8;
    }
    .features span { display: inline-block; margin-right: 12px; }
  </style>
</head>
<body>
  <div class="card">
    <h1>📖 縦書きEPUB変換</h1>
    <p class="subtitle">
      WordファイルをKindle対応の縦書きEPUBに変換します。<br>
      ルビ・見出し・表に自動対応。変換後すぐダウンロードできます。
    </p>

    <form id="form">
      <div class="drop-zone" id="dropZone">
        <input type="file" id="fileInput" accept=".docx">
        <div class="icon">📄</div>
        <div class="label">Wordファイル (.docx) をドロップ<br>またはクリックして選択</div>
        <div class="file-name" id="fileName"></div>
      </div>

      <label class="meta">タイトル（空欄で自動検出）</label>
      <input type="text" id="title" placeholder="例: 吾輩は猫である">

      <label class="meta">著者名（空欄で「著者不明」）</label>
      <input type="text" id="author" placeholder="例: 夏目漱石">

      <button type="submit" id="btn">EPUBに変換してダウンロード</button>
    </form>

    <div class="loading" id="loading">⏳ 変換中...</div>

    <div class="result" id="result"></div>

    <div class="features">
      対応機能：
      <span>✅ ルビ（振り仮名）</span>
      <span>✅ 縦書きCSS</span>
      <span>✅ 縦中横</span>
      <span>✅ 表</span>
      <span>✅ 見出し自動検出</span>
    </div>
  </div>

<script>
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileName  = document.getElementById('fileName');
const form      = document.getElementById('form');
const btn       = document.getElementById('btn');
const loading   = document.getElementById('loading');
const resultDiv = document.getElementById('result');

dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) fileName.textContent = fileInput.files[0].name;
});
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const f = e.dataTransfer.files[0];
  if (f && f.name.endsWith('.docx')) {
    fileInput.files = e.dataTransfer.files;
    fileName.textContent = f.name;
  }
});

form.addEventListener('submit', async e => {
  e.preventDefault();
  if (!fileInput.files[0]) { alert('ファイルを選択してください'); return; }

  btn.disabled = true;
  loading.style.display = 'block';
  resultDiv.style.display = 'none';

  const fd = new FormData();
  fd.append('file', fileInput.files[0]);
  fd.append('title', document.getElementById('title').value);
  fd.append('author', document.getElementById('author').value);

  try {
    const res = await fetch('/convert', { method: 'POST', body: fd });
    if (res.ok) {
      const blob = await res.blob();
      const epubName = fileInput.files[0].name.replace('.docx', '.epub');
      const url = URL.createObjectURL(blob);
      const warnings = res.headers.get('X-Warnings') || '';
      resultDiv.className = 'result success';
      resultDiv.innerHTML = `
        ✅ 変換完了！
        <a href="${url}" download="${epubName}">⬇ ${epubName} をダウンロード</a>
        ${warnings ? '<div class="warnings">⚠️ ' + decodeURIComponent(warnings) + '</div>' : ''}
      `;
    } else {
      const json = await res.json();
      resultDiv.className = 'result error';
      resultDiv.innerHTML = '❌ ' + (json.error || '変換に失敗しました');
    }
  } catch (err) {
    resultDiv.className = 'result error';
    resultDiv.innerHTML = '❌ 通信エラーが発生しました';
  }

  resultDiv.style.display = 'block';
  loading.style.display = 'none';
  btn.disabled = false;
});
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML)


@app.route('/convert', methods=['POST'])
def convert():
    f = request.files.get('file')
    if not f:
        return jsonify({'error': 'ファイルがありません'}), 400

    if Path(f.filename).suffix.lower() not in ALLOWED_EXT:
        return jsonify({'error': '.docxファイルのみ対応しています'}), 400

    title  = request.form.get('title', '').strip() or None
    author = request.form.get('author', '').strip() or None

    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, 'input.docx')
        epub_path = os.path.join(tmpdir, 'output.epub')
        f.save(docx_path)

        try:
            result = docx_to_epub(docx_path, epub_path, title=title, author=author)
        except Exception as e:
            return jsonify({'error': f'変換エラー: {str(e)}'}), 500

        warnings_str = ' / '.join(result.get('warnings', []))

        response = send_file(
            epub_path,
            mimetype='application/epub+zip',
            as_attachment=True,
            download_name='output.epub'
        )
        if warnings_str:
            response.headers['X-Warnings'] = warnings_str[:500]
        return response


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('FLASK_ENV') != 'production'
    print(f'🚀 起動: http://localhost:{port}')
    app.run(debug=debug, host='0.0.0.0', port=port)
