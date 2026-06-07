"""
Kindle SaaS MVP - Flask Web App
Word (.docx) をアップロード → 縦書きEPUBをダウンロード
"""
import os
import tempfile
from collections import defaultdict
from datetime import date
from pathlib import Path
from flask import Flask, request, send_file, jsonify, render_template_string

from docx_to_epub import docx_to_epub

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 20 * 1024 * 1024  # 20MB上限

ALLOWED_EXT = {'.docx'}

# ── 1日3回制限 ──────────────────────────────────
FREE_DAILY_LIMIT = 3
POLAR_CHECKOUT_URL = 'https://buy.polar.sh/polar_cl_7U8n8aLqQc4gu6JFc01BJaijMgMjNltLDXZjW00RP6Z'

usage_tracker: dict[str, int] = defaultdict(int)

def get_client_ip() -> str:
    """Renderのリバースプロキシ経由でも正しいIPを取得"""
    xff = request.headers.get('X-Forwarded-For', '')
    if xff:
        return xff.split(',')[0].strip()
    return request.remote_addr or '0.0.0.0'

def usage_key(ip: str) -> str:
    return f"{ip}_{date.today().isoformat()}"

def get_remaining(ip: str) -> int:
    return max(0, FREE_DAILY_LIMIT - usage_tracker[usage_key(ip)])

HTML = """
<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>縦書きEPUBコンバーター | Kindle出版のEPUB地獄から解放</title>
  <meta name="description" content="ルビのHTML手打ち、縦書きCSS、Kindleでの崩れ確認……その作業、もう不要です。WordファイルをアップロードするだけでKindle対応EPUBが完成。">
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    :root {
      --indigo: #4f46e5;
      --indigo-dark: #4338ca;
      --indigo-light: #eef2ff;
      --green: #16a34a;
      --text: #111827;
      --muted: #6b7280;
      --border: #e5e7eb;
      --bg: #f9fafb;
    }
    body {
      font-family: "Hiragino Sans", "Yu Gothic UI", "Meiryo", sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.6;
    }

    /* ── NAV ── */
    nav {
      background: #fff;
      border-bottom: 1px solid var(--border);
      padding: 0 24px;
      height: 60px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      position: sticky;
      top: 0;
      z-index: 100;
    }
    .nav-logo { font-weight: 800; font-size: 1.1em; color: var(--indigo); letter-spacing: -0.5px; }
    .nav-badge {
      background: var(--indigo-light);
      color: var(--indigo);
      font-size: 0.72em;
      font-weight: 700;
      padding: 2px 8px;
      border-radius: 99px;
      margin-left: 8px;
    }
    .nav-cta {
      background: var(--indigo);
      color: #fff;
      border: none;
      border-radius: 8px;
      padding: 8px 20px;
      font-size: 0.9em;
      font-weight: 700;
      cursor: pointer;
      text-decoration: none;
    }
    .nav-cta:hover { background: var(--indigo-dark); }

    /* ── HERO ── */
    .hero {
      text-align: center;
      padding: 80px 24px 64px;
      max-width: 720px;
      margin: 0 auto;
    }
    .hero-eyebrow {
      display: inline-block;
      background: var(--indigo-light);
      color: var(--indigo);
      font-size: 0.82em;
      font-weight: 700;
      padding: 4px 14px;
      border-radius: 99px;
      margin-bottom: 20px;
    }
    .hero h1 {
      font-size: clamp(1.8em, 4vw, 2.8em);
      font-weight: 900;
      line-height: 1.25;
      letter-spacing: -1px;
      margin-bottom: 20px;
    }
    .hero h1 em {
      color: var(--indigo);
      font-style: normal;
    }
    .hero p {
      color: var(--muted);
      font-size: 1.05em;
      margin-bottom: 36px;
      max-width: 520px;
      margin-left: auto;
      margin-right: auto;
    }
    .hero-btn {
      display: inline-block;
      background: var(--indigo);
      color: #fff;
      border-radius: 10px;
      padding: 14px 36px;
      font-size: 1.05em;
      font-weight: 800;
      text-decoration: none;
      transition: background 0.2s, transform 0.1s;
    }
    .hero-btn:hover { background: var(--indigo-dark); transform: translateY(-1px); }
    .hero-sub { margin-top: 14px; font-size: 0.82em; color: var(--muted); }

    /* ── PAIN ── */
    .pain-section {
      background: #1e1b4b;
      color: #e0e7ff;
      padding: 64px 24px;
    }
    .pain-section .section-label { color: #a5b4fc; }
    .pain-section .section-title { color: #fff; }
    .pain-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 16px;
      max-width: 860px;
      margin: 40px auto 0;
    }
    .pain-item {
      background: rgba(255,255,255,0.07);
      border: 1px solid rgba(255,255,255,0.12);
      border-radius: 12px;
      padding: 22px 20px;
      font-size: 0.9em;
      line-height: 1.6;
    }
    .pain-item .pain-emoji { font-size: 1.6em; margin-bottom: 10px; display: block; }
    .pain-item strong { color: #fff; display: block; margin-bottom: 4px; }
    .pain-item span { color: #a5b4fc; font-size: 0.88em; }

    /* ── FEATURES ── */
    .features-section {
      background: #fff;
      padding: 64px 24px;
      border-top: 1px solid var(--border);
      border-bottom: 1px solid var(--border);
    }
    .section-label {
      text-align: center;
      font-size: 0.8em;
      font-weight: 700;
      color: var(--indigo);
      letter-spacing: 1px;
      text-transform: uppercase;
      margin-bottom: 12px;
    }
    .section-title {
      text-align: center;
      font-size: clamp(1.3em, 3vw, 1.9em);
      font-weight: 800;
      margin-bottom: 48px;
    }
    .release-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 20px;
      max-width: 900px;
      margin: 0 auto;
    }
    .release-card {
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 24px;
      display: flex;
      gap: 16px;
      align-items: flex-start;
    }
    .release-card .rc-icon {
      width: 40px; height: 40px; min-width: 40px;
      background: var(--indigo-light);
      border-radius: 10px;
      display: flex; align-items: center; justify-content: center;
      font-size: 1.3em;
    }
    .release-card h3 { font-size: 0.95em; font-weight: 700; margin-bottom: 4px; }
    .release-card .before {
      font-size: 0.8em;
      color: #ef4444;
      text-decoration: line-through;
      margin-bottom: 2px;
    }
    .release-card .after {
      font-size: 0.82em;
      color: var(--green);
      font-weight: 600;
    }

    .feature-icon { font-size: 2em; margin-bottom: 12px; }
    .feature-card h3 { font-size: 1em; font-weight: 700; margin-bottom: 6px; }
    .feature-card p { font-size: 0.85em; color: var(--muted); line-height: 1.6; }

    /* ── HOW IT WORKS ── */
    .how-section { padding: 64px 24px; max-width: 700px; margin: 0 auto; text-align: center; }
    .steps { display: flex; flex-direction: column; gap: 0; margin-top: 40px; text-align: left; }
    .step { display: flex; gap: 20px; align-items: flex-start; padding: 24px 0; border-bottom: 1px solid var(--border); }
    .step:last-child { border-bottom: none; }
    .step-num {
      width: 36px; height: 36px; min-width: 36px;
      background: var(--indigo);
      color: #fff;
      border-radius: 50%;
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 0.9em;
    }
    .step h3 { font-size: 1em; font-weight: 700; margin-bottom: 4px; }
    .step p { font-size: 0.88em; color: var(--muted); }

    /* ── CONVERTER ── */
    .converter-section {
      background: #fff;
      border-top: 1px solid var(--border);
      padding: 64px 24px;
    }
    .converter-wrap {
      max-width: 560px;
      margin: 0 auto;
    }
    .drop-zone {
      border: 2px dashed var(--border);
      border-radius: 12px;
      padding: 44px 20px;
      text-align: center;
      cursor: pointer;
      transition: border-color 0.2s, background 0.2s;
      margin-bottom: 20px;
      background: var(--bg);
    }
    .drop-zone.dragover { border-color: var(--indigo); background: var(--indigo-light); }
    .drop-zone input { display: none; }
    .drop-zone .dz-icon { font-size: 2.8em; margin-bottom: 12px; }
    .drop-zone .dz-label { color: var(--muted); font-size: 0.95em; line-height: 1.7; }
    .drop-zone .dz-label strong { color: var(--indigo); }
    .drop-zone .dz-filename { margin-top: 10px; font-weight: 700; color: var(--indigo); font-size: 0.95em; }
    .meta-row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 16px; }
    .meta-group label { display: block; font-size: 0.82em; font-weight: 600; color: var(--muted); margin-bottom: 4px; }
    input[type=text] {
      width: 100%; padding: 9px 12px;
      border: 1px solid var(--border);
      border-radius: 8px; font-size: 0.93em;
      outline: none; transition: border-color 0.2s;
    }
    input[type=text]:focus { border-color: var(--indigo); }
    .convert-btn {
      width: 100%; background: var(--indigo); color: #fff;
      border: none; border-radius: 10px; padding: 15px;
      font-size: 1.05em; font-weight: 800; cursor: pointer;
      transition: background 0.2s;
    }
    .convert-btn:hover { background: var(--indigo-dark); }
    .convert-btn:disabled { background: #a5b4fc; cursor: not-allowed; }
    .loading { display: none; text-align: center; color: var(--muted); margin-top: 20px; font-size: 0.95em; }
    .result { margin-top: 20px; padding: 20px; border-radius: 10px; font-size: 0.9em; display: none; }
    .result.success { background: #f0fdf4; border: 1px solid #86efac; color: #166534; }
    .result.error { background: #fef2f2; border: 1px solid #fca5a5; color: #991b1b; }
    .result a {
      display: inline-block; margin-top: 14px;
      background: var(--green); color: #fff;
      padding: 11px 28px; border-radius: 8px;
      text-decoration: none; font-weight: 700; font-size: 1em;
    }
    .warnings { margin-top: 10px; color: #92400e; font-size: 0.83em; }

    /* ── PRICING ── */
    .pricing-section { padding: 64px 24px; }
    .pricing-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
      gap: 24px;
      max-width: 640px;
      margin: 40px auto 0;
    }
    .pricing-card {
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 16px;
      padding: 32px 28px;
      position: relative;
    }
    .pricing-card.featured {
      border-color: var(--indigo);
      box-shadow: 0 0 0 3px var(--indigo-light);
    }
    .pricing-badge {
      position: absolute; top: -12px; left: 50%; transform: translateX(-50%);
      background: var(--indigo); color: #fff;
      font-size: 0.72em; font-weight: 700;
      padding: 3px 12px; border-radius: 99px;
      white-space: nowrap;
    }
    .pricing-name { font-size: 0.9em; font-weight: 700; color: var(--muted); margin-bottom: 8px; }
    .pricing-price { font-size: 2.2em; font-weight: 900; margin-bottom: 4px; }
    .pricing-price span { font-size: 0.45em; font-weight: 500; color: var(--muted); }
    .pricing-desc { font-size: 0.85em; color: var(--muted); margin-bottom: 20px; }
    .pricing-features { list-style: none; font-size: 0.88em; }
    .pricing-features li { padding: 5px 0; }
    .pricing-features li::before { content: "✓　"; color: var(--green); font-weight: 700; }
    .pricing-cta {
      display: block; width: 100%; margin-top: 24px;
      background: var(--bg); color: var(--text);
      border: 1px solid var(--border); border-radius: 8px;
      padding: 11px; font-size: 0.95em; font-weight: 700;
      text-align: center; cursor: pointer; text-decoration: none;
      transition: background 0.2s;
    }
    .pricing-cta:hover { background: var(--border); }
    .pricing-card.featured .pricing-cta {
      background: var(--indigo); color: #fff; border-color: var(--indigo);
    }
    .pricing-card.featured .pricing-cta:hover { background: var(--indigo-dark); }

    /* ── FOOTER ── */
    footer {
      background: var(--text);
      color: #9ca3af;
      text-align: center;
      padding: 32px 24px;
      font-size: 0.82em;
    }
    footer a { color: #9ca3af; text-decoration: underline; }

    @media (max-width: 480px) {
      .meta-row { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>

<!-- NAV -->
<nav>
  <div>
    <span class="nav-logo">📖 縦書きEPUB</span>
    <span class="nav-badge">β版</span>
  </div>
  <a class="nav-cta" href="#converter">無料で試す</a>
</nav>

<!-- HERO -->
<section class="hero">
  <div class="hero-eyebrow">🛠️ Kindle出版の「手作業を減らす」ためのツール</div>
  <h1>Word原稿を、<br><em>そのままKindleへ。</em></h1>
  <p>すべての原稿を完璧に変換できるわけではありません。<br>でも、縦書きCSS・ルビ保持・表・見出しの書式作業を自動化することで、<strong>あなたの手作業を大幅に減らします。</strong></p>
  <a class="hero-btn" href="#converter">今すぐ無料で変換する →</a>
  <div class="hero-sub">クレジットカード不要・登録なし・今すぐ使える</div>
</section>

<!-- PAIN -->
<section class="pain-section">
  <div class="section-label" style="text-align:center">こんな経験、ありませんか？</div>
  <div class="section-title" style="text-align:center">Kindle出版の「書式地獄」</div>
  <div class="pain-grid">
    <div class="pain-item">
      <span class="pain-emoji">😩</span>
      <strong>変換したらルビが全部消えた</strong>
      <span>Wordで丁寧に振ったルビが、EPUBにしたら跡形もなく消えた……</span>
    </div>
    <div class="pain-item">
      <span class="pain-emoji">🕐</span>
      <strong>縦書きCSSで3時間溶けた</strong>
      <span>writing-modeって何？ vertical-rl？ ネットを調べ続けて日が暮れた。</span>
    </div>
    <div class="pain-item">
      <span class="pain-emoji">💀</span>
      <strong>Kindleビューワーで崩れた</strong>
      <span>やっと変換できたと思ったら、レイアウトが全崩れ。最初からやり直し。</span>
    </div>
    <div class="pain-item">
      <span class="pain-emoji">💸</span>
      <strong>外注に数万円払った</strong>
      <span>「自分でできないならプロに頼むしかない」——もったいなかった。</span>
    </div>
  </div>
</section>

<!-- FEATURES（解放リスト） -->
<section class="features-section">
  <div class="section-label">解決策</div>
  <div class="section-title">もうやらなくていい。<br>全部、自動でやります。</div>
  <div class="release-grid">
    <div class="release-card">
      <div class="rc-icon">🈶</div>
      <div>
        <h3>ルビ（振り仮名）の保持</h3>
        <div class="before">EPUBに変換したらルビが全部消えた</div>
        <div class="after">→ Wordで設定したルビをそのまま保持</div>
        <div style="font-size:0.78em;color:var(--muted);margin-top:6px">※ ルビはWordで事前に設定が必要です。自動付与はしません。</div>
      </div>
    </div>
    <div class="release-card">
      <div class="rc-icon">📜</div>
      <div>
        <h3>縦書きレイアウト</h3>
        <div class="before">縦書きCSSを自分で書く</div>
        <div class="after">→ Kindle対応の縦書きを自動適用</div>
      </div>
    </div>
    <div class="release-card">
      <div class="rc-icon">🔢</div>
      <div>
        <h3>数字の縦中横</h3>
        <div class="before">「第２章」が横倒しになって恥ずかしい</div>
        <div class="after">→ 縦書き中の数字を自動で正立処理</div>
      </div>
    </div>
    <div class="release-card">
      <div class="rc-icon">📊</div>
      <div>
        <h3>表・見出し</h3>
        <div class="before">表とスタイルを手動でHTMLに変換</div>
        <div class="after">→ Word構造をそのまま自動認識</div>
      </div>
    </div>
  </div>
</section>

<!-- HOW IT WORKS -->
<section class="how-section">
  <div class="section-label">使い方</div>
  <div class="section-title">3ステップで完了</div>
  <div class="steps">
    <div class="step">
      <div class="step-num">1</div>
      <div>
        <h3>Wordファイルをアップロード</h3>
        <p>.docx形式のファイルをドラッグ&ドロップ、またはクリックして選択するだけ。</p>
      </div>
    </div>
    <div class="step">
      <div class="step-num">2</div>
      <div>
        <h3>タイトル・著者名を入力（任意）</h3>
        <p>空欄でもOK。ファイル名やWord文書のプロパティから自動取得します。</p>
      </div>
    </div>
    <div class="step">
      <div class="step-num">3</div>
      <div>
        <h3>EPUBをダウンロード</h3>
        <p>数秒で変換完了。KindlePreviewerやKindleアプリですぐに確認できます。</p>
      </div>
    </div>
  </div>
</section>

<!-- CONVERTER -->
<section class="converter-section" id="converter">
  <div class="converter-wrap">
    <div class="section-label" style="text-align:center">今すぐ変換</div>
    <div class="section-title" style="text-align:center;margin-bottom:32px">Wordファイルをアップロード</div>

    <form id="form">
      <div class="drop-zone" id="dropZone">
        <input type="file" id="fileInput" accept=".docx">
        <div class="dz-icon">📄</div>
        <div class="dz-label">
          <strong>クリックしてファイルを選択</strong><br>
          またはここにドラッグ&ドロップ<br>
          <span style="font-size:0.85em">.docx形式・最大20MB</span>
        </div>
        <div class="dz-filename" id="fileName"></div>
      </div>

      <div class="meta-row">
        <div class="meta-group">
          <label>タイトル（省略可）</label>
          <input type="text" id="title" placeholder="例: 吾輩は猫である">
        </div>
        <div class="meta-group">
          <label>著者名（省略可）</label>
          <input type="text" id="author" placeholder="例: 夏目漱石">
        </div>
      </div>

      <button type="submit" class="convert-btn" id="btn">⚡ EPUBに変換してダウンロード</button>
    </form>

    <div class="loading" id="loading">⏳ 変換中です。しばらくお待ちください…</div>
    <div class="result" id="result"></div>

    <!-- 残り回数バッジ -->
    <div id="usageBadge" style="margin-top:20px;text-align:center;font-size:0.85em;color:var(--muted);display:none">
      今日の残り回数：<strong id="remainingCount"></strong> / 3回
    </div>

    <!-- 上限超過バナー -->
    <div id="limitBanner" style="display:none;margin-top:24px;background:#fef3c7;border:1px solid #fcd34d;border-radius:12px;padding:24px;text-align:center">
      <div style="font-size:1.5em;margin-bottom:8px">⚡</div>
      <div style="font-weight:800;font-size:1.05em;margin-bottom:6px">本日の無料枠（3回）を使い切りました</div>
      <div style="font-size:0.88em;color:#92400e;margin-bottom:20px">明日0時にリセットされます。今すぐ続けたい方はProプランへ。</div>
      <a id="polarBtn" href="#" target="_blank"
        style="display:inline-block;background:#4f46e5;color:#fff;border-radius:8px;padding:12px 32px;font-weight:800;text-decoration:none;font-size:0.95em">
        🚀 Proプランにアップグレード（¥980/月）
      </a>
      <div style="margin-top:10px;font-size:0.78em;color:#92400e">変換回数無制限・いつでもキャンセル可</div>
    </div>
  </div>
</section>

<!-- PRICING -->
<section class="pricing-section">
  <div class="section-label" style="text-align:center">料金プラン</div>
  <div class="section-title" style="text-align:center">シンプルな料金体系</div>
  <div class="pricing-grid">
    <div class="pricing-card">
      <div class="pricing-name">FREE</div>
      <div class="pricing-price">¥0<span> / 月</span></div>
      <div class="pricing-desc">まずは試してみたい方に</div>
      <ul class="pricing-features">
        <li>月3回まで変換可能</li>
        <li>ルビ・縦書き・表対応</li>
        <li>最大20MBのファイル</li>
      </ul>
      <a class="pricing-cta" href="#converter">今すぐ試す</a>
    </div>
    <div class="pricing-card featured">
      <div class="pricing-badge">おすすめ</div>
      <div class="pricing-name">PRO</div>
      <div class="pricing-price">¥980<span> / 月</span></div>
      <div class="pricing-desc">本格的に出版したい方に</div>
      <ul class="pricing-features">
        <li>変換回数 無制限</li>
        <li>最大100MBのファイル</li>
        <li>優先サポート</li>
        <li>近日：一括変換・API連携</li>
      </ul>
      <a class="pricing-cta" href="#converter">準備中 — 通知を受け取る</a>
    </div>
  </div>
</section>

<!-- FOOTER -->
<footer>
  <p>© 2026 縦書きEPUBコンバーター &nbsp;·&nbsp; <a href="mailto:support@example.com">お問い合わせ</a></p>
  <p style="margin-top:8px">Kindle・Amazon は Amazon.com, Inc. の商標です。本サービスはAmazonと提携しておりません。</p>
</footer>

<script>
const dropZone      = document.getElementById('dropZone');
const fileInput     = document.getElementById('fileInput');
const fileName      = document.getElementById('fileName');
const form          = document.getElementById('form');
const btn           = document.getElementById('btn');
const loading       = document.getElementById('loading');
const resultDiv     = document.getElementById('result');
const usageBadge    = document.getElementById('usageBadge');
const remainingCount= document.getElementById('remainingCount');
const limitBanner   = document.getElementById('limitBanner');
const polarBtn      = document.getElementById('polarBtn');

let polarUrl = '#';

// ── ページロード時：残り回数を取得 ──
async function loadUsage() {
  try {
    const res = await fetch('/usage');
    const data = await res.json();
    polarUrl = data.polar_url || '#';
    polarBtn.href = polarUrl;
    updateUsageUI(data.remaining);
  } catch (e) { /* サイレントに失敗 */ }
}

function updateUsageUI(remaining) {
  if (remaining <= 0) {
    // 上限超過：フォームを隠してバナー表示
    form.style.display = 'none';
    usageBadge.style.display = 'none';
    limitBanner.style.display = 'block';
    polarBtn.href = polarUrl;
  } else {
    form.style.display = 'block';
    limitBanner.style.display = 'none';
    usageBadge.style.display = 'block';
    remainingCount.textContent = remaining;
    // 残り1回なら警告色
    remainingCount.style.color = remaining === 1 ? '#ef4444' : '#111827';
  }
}

loadUsage();

// ── ファイル選択 ──
dropZone.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', () => {
  if (fileInput.files[0]) {
    fileName.textContent = '📎 ' + fileInput.files[0].name;
    dropZone.style.borderColor = '#4f46e5';
  }
});
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('dragover');
  const f = e.dataTransfer.files[0];
  if (f && f.name.endsWith('.docx')) {
    const dt = new DataTransfer();
    dt.items.add(f);
    fileInput.files = dt.files;
    fileName.textContent = '📎 ' + f.name;
    dropZone.style.borderColor = '#4f46e5';
  }
});

// ── 変換送信 ──
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

    if (res.status === 429) {
      // 上限超過
      const json = await res.json().catch(() => ({}));
      updateUsageUI(0);
      loading.style.display = 'none';
      btn.disabled = false;
      return;
    }

    if (res.ok) {
      const blob = await res.blob();
      const epubName = fileInput.files[0].name.replace('.docx', '.epub');
      const url = URL.createObjectURL(blob);
      const warnings  = res.headers.get('X-Warnings') || '';
      const remaining = parseInt(res.headers.get('X-Remaining') ?? '99', 10);

      resultDiv.className = 'result success';
      resultDiv.innerHTML = `
        <strong>✅ 変換完了！</strong><br>EPUBファイルの準備ができました。
        <br><a href="${url}" download="${epubName}">⬇ ${epubName} をダウンロード</a>
        ${warnings ? '<div class="warnings">⚠️ ' + decodeURIComponent(warnings) + '</div>' : ''}
      `;
      updateUsageUI(remaining);
    } else {
      const json = await res.json().catch(() => ({}));
      resultDiv.className = 'result error';
      resultDiv.innerHTML = '❌ ' + (json.error || '変換に失敗しました。ファイルを確認してください。');
    }
  } catch (err) {
    resultDiv.className = 'result error';
    resultDiv.innerHTML = '❌ 通信エラーが発生しました。時間をおいて再度お試しください。';
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


@app.route('/usage', methods=['GET'])
def usage_status():
    """残り回数をJSONで返す（ページロード時に取得）"""
    ip = get_client_ip()
    remaining = get_remaining(ip)
    return jsonify({
        'remaining': remaining,
        'limit': FREE_DAILY_LIMIT,
        'polar_url': POLAR_CHECKOUT_URL,
    })


@app.route('/convert', methods=['POST'])
def convert():
    ip = get_client_ip()
    key = usage_key(ip)

    # 1日の上限チェック
    if usage_tracker[key] >= FREE_DAILY_LIMIT:
        return jsonify({
            'error': 'limit',
            'message': '本日の無料変換回数（3回）に達しました。明日またご利用いただくか、Proプランにアップグレードしてください。',
            'polar_url': POLAR_CHECKOUT_URL,
        }), 429

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

        # 変換成功 → カウントアップ
        usage_tracker[key] += 1
        remaining = get_remaining(ip)

        warnings_str = ' / '.join(result.get('warnings', []))

        response = send_file(
            epub_path,
            mimetype='application/epub+zip',
            as_attachment=True,
            download_name='output.epub'
        )
        if warnings_str:
            response.headers['X-Warnings'] = warnings_str[:500]
        response.headers['X-Remaining'] = str(remaining)
        response.headers['Access-Control-Expose-Headers'] = 'X-Warnings, X-Remaining'
        return response


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5001))
    debug = os.environ.get('FLASK_ENV') != 'production'
    print(f'🚀 起動: http://localhost:{port}')
    app.run(debug=debug, host='0.0.0.0', port=port)
