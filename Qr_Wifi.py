import csv
import os
from datetime import datetime
from flask import Flask, request, render_template_string, jsonify

app = Flask(__name__)
CSV_FILE = 'maintenance_log.csv'

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>PP2 Car History — Maintenance Log</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@zxing/library@0.21.3/umd/index.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;700&display=swap');
:root{--bg:#060b18;--s1:#090e1f;--s2:#0d1428;--s3:#111c35;--border:#1a2d50;--border2:#243d6a;--blue:#3b82f6;--cyan:#22d3ee;--green:#10b981;--amber:#f59e0b;--violet:#8b5cf6;--rose:#f43f5e;--teal:#14b8a6;--orange:#f97316;--text:#e2e8f0;--text2:#94a3b8;--text3:#475569;--mono:'JetBrains Mono',monospace;--ui:'Inter',sans-serif}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
html{font-size:14px;scroll-behavior:smooth}
body{background:var(--bg);color:var(--text);font-family:var(--ui);min-height:100vh;overflow-x:hidden}
body::before{content:'';position:fixed;inset:0;z-index:0;background-image:linear-gradient(rgba(59,130,246,.03) 1px,transparent 1px),linear-gradient(90deg,rgba(59,130,246,.03) 1px,transparent 1px);background-size:40px 40px;pointer-events:none}

.hero{position:relative;z-index:1;padding:28px 40px 20px;border-bottom:1px solid var(--border);background:linear-gradient(180deg,rgba(59,130,246,.06) 0%,transparent 100%);display:flex;align-items:flex-start;justify-content:space-between;gap:20px;flex-wrap:wrap}
.hero::before{content:'';position:absolute;width:360px;height:200px;background:radial-gradient(ellipse,rgba(59,130,246,.18),transparent 70%);top:-40px;left:-40px;border-radius:50%;filter:blur(80px);pointer-events:none}
.hero-left{position:relative;z-index:2}
.eyebrow{display:inline-flex;align-items:center;gap:8px;background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.25);border-radius:20px;padding:4px 12px;font-size:10.5px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--cyan);margin-bottom:14px}
.eyebrow .live-dot{width:6px;height:6px;background:#10b981;border-radius:50%;animation:blink 2s ease infinite}
@keyframes blink{0%,100%{opacity:1;box-shadow:0 0 0 0 rgba(16,185,129,.4)}50%{opacity:.5;box-shadow:0 0 0 4px rgba(16,185,129,0)}}
.hero h1{font-size:clamp(20px,3vw,30px);font-weight:900;letter-spacing:-.03em;line-height:1.1;background:linear-gradient(135deg,#e2e8f0 20%,#60a5fa 60%,#a78bfa 100%);-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}
.hero-sub{margin-top:6px;font-size:12px;color:var(--text2);font-family:var(--mono);letter-spacing:.04em}
.hero-right{position:relative;z-index:2;display:flex;align-items:center;gap:12px;flex-wrap:wrap}
.stat-chip{background:var(--s3);border:1px solid var(--border2);border-radius:10px;padding:10px 16px;text-align:center;min-width:80px}
.stat-chip .val{font-size:20px;font-weight:800;font-family:var(--mono);color:var(--cyan);line-height:1}
.stat-chip .lbl{font-size:9.5px;color:var(--text3);text-transform:uppercase;letter-spacing:.1em;margin-top:3px}

.car-banner{position:relative;z-index:1;background:var(--s2);border-bottom:2px solid var(--border2);padding:16px 40px}
.car-banner-inner{display:flex;align-items:center;gap:18px;flex-wrap:wrap}
.car-banner-label{font-size:11px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.1em;white-space:nowrap}

.lookup-wrap{position:relative;display:flex;align-items:center;gap:0}
.lookup-wrap input{background:var(--s3);border:1.5px solid var(--border2);border-radius:8px 0 0 8px;color:var(--text);font-family:var(--mono);font-size:13px;padding:10px 14px;outline:none;width:110px;transition:border-color 150ms,box-shadow 150ms}
.lookup-wrap input:focus{border-color:var(--cyan);box-shadow:0 0 0 3px rgba(34,211,238,.15)}
.lookup-wrap input.found{border-color:#10b981;box-shadow:0 0 0 3px rgba(16,185,129,.15)}
.lookup-wrap input.notfound{border-color:#f43f5e;box-shadow:0 0 0 3px rgba(244,63,94,.15)}
.lookup-badge{display:flex;align-items:center;background:var(--s3);border:1.5px solid var(--border2);border-left:none;padding:10px 11px;min-width:36px;justify-content:center;transition:all 150ms}
.lookup-badge.found{border-color:#10b981;color:#10b981}
.lookup-badge.notfound{border-color:#f43f5e;color:#f43f5e}
.lookup-badge svg{width:14px;height:14px}

.btn-scan-qr{display:flex;align-items:center;justify-content:center;gap:6px;background:linear-gradient(135deg,rgba(34,211,238,.15),rgba(99,102,241,.15));border:1.5px solid rgba(34,211,238,.4);border-left:none;border-radius:0 8px 8px 0;padding:10px 13px;cursor:pointer;color:var(--cyan);font-size:11px;font-weight:700;font-family:var(--ui);letter-spacing:.04em;white-space:nowrap;transition:all 150ms;min-width:42px}
.btn-scan-qr:hover{background:linear-gradient(135deg,rgba(34,211,238,.28),rgba(99,102,241,.28));box-shadow:0 0 14px rgba(34,211,238,.25)}
.btn-scan-qr svg{width:15px;height:15px;flex-shrink:0}
.btn-scan-qr.scanning{background:linear-gradient(135deg,rgba(244,63,94,.18),rgba(244,63,94,.1));border-color:rgba(244,63,94,.5);color:#f43f5e;animation:scanpulse 1.2s ease infinite}
@keyframes scanpulse{0%,100%{box-shadow:0 0 0 0 rgba(244,63,94,.3)}50%{box-shadow:0 0 0 6px rgba(244,63,94,0)}}

#scanModal{display:none;position:fixed;inset:0;z-index:1000;background:rgba(4,8,20,.92);backdrop-filter:blur(10px);align-items:center;justify-content:center}
#scanModal.open{display:flex}
.scan-modal-box{background:var(--s1);border:1px solid var(--border2);border-radius:18px;padding:24px;max-width:480px;width:92vw;display:flex;flex-direction:column;gap:16px;box-shadow:0 24px 80px rgba(0,0,0,.8),0 0 0 1px rgba(34,211,238,.1)}
.scan-modal-title{display:flex;align-items:center;justify-content:space-between;gap:12px}
.scan-modal-title h2{font-size:15px;font-weight:700;color:var(--text);display:flex;align-items:center;gap:9px}
.scan-modal-title h2 svg{color:var(--cyan)}
.scan-close{background:rgba(255,255,255,.06);border:1px solid var(--border);border-radius:7px;color:var(--text2);cursor:pointer;padding:6px 10px;font-size:13px;transition:all 140ms}
.scan-close:hover{background:rgba(244,63,94,.15);border-color:rgba(244,63,94,.4);color:#f43f5e}
.scan-cam-wrap{position:relative;width:100%;aspect-ratio:4/3;background:#000;border-radius:12px;overflow:hidden;border:1.5px solid var(--border2)}
#scanVideo{width:100%;height:100%;object-fit:cover;display:block}
.scan-overlay{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;pointer-events:none}
.scan-frame{width:62%;aspect-ratio:1;border:2.5px solid var(--cyan);border-radius:12px;box-shadow:0 0 0 9999px rgba(0,0,0,.45);position:relative}
.scan-frame::before,.scan-frame::after{content:'';position:absolute;width:22px;height:22px;border-color:var(--cyan);border-style:solid}
.scan-frame::before{top:-3px;left:-3px;border-width:3px 0 0 3px;border-radius:6px 0 0 0}
.scan-frame::after{bottom:-3px;right:-3px;border-width:0 3px 3px 0;border-radius:0 0 6px 0}
.scan-line{position:absolute;left:8%;right:8%;height:2px;background:linear-gradient(90deg,transparent,var(--cyan),transparent);animation:scanline 2s ease-in-out infinite;top:10%}
@keyframes scanline{0%,100%{top:10%;opacity:.8}50%{top:88%;opacity:1}}
.scan-status{text-align:center;font-size:12px;font-family:var(--mono);color:var(--text2);display:flex;align-items:center;justify-content:center;gap:8px;min-height:20px}
.scan-status .dot{width:7px;height:7px;border-radius:50%;background:var(--cyan);animation:blink 1.2s ease infinite;flex-shrink:0}
.scan-status.success .dot{background:#10b981}
.scan-status.error .dot{background:#f43f5e}
.scan-cam-sel{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.scan-cam-sel label{font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.08em;white-space:nowrap}
.scan-cam-sel select{flex:1;background:var(--s2);border:1px solid var(--border2);border-radius:7px;color:var(--text);font-family:var(--mono);font-size:11.5px;padding:7px 10px;outline:none;appearance:none;cursor:pointer;min-width:0}
.scan-cam-sel select:focus{border-color:var(--cyan)}
.scan-actions{display:flex;gap:10px}
.scan-actions .btn{flex:1;justify-content:center;padding:10px}

.car-select-wrap{position:relative;flex:0 0 auto}
.car-select-wrap::after{content:'▾';position:absolute;right:12px;top:50%;transform:translateY(-50%);color:var(--cyan);pointer-events:none;font-size:12px}
select#ccSelect{background:var(--s3);border:1.5px solid var(--border2);border-radius:8px;color:var(--text);font-family:var(--mono);font-size:12.5px;padding:10px 32px 10px 14px;outline:none;appearance:none;cursor:pointer;transition:border-color 150ms,box-shadow 150ms;min-width:240px}
select#ccSelect:focus{border-color:var(--cyan);box-shadow:0 0 0 3px rgba(34,211,238,.15)}
select#ccSelect.synced{border-color:#10b981}
.car-info-chip{background:linear-gradient(135deg,rgba(34,211,238,.1),rgba(99,102,241,.1));border:1px solid rgba(34,211,238,.25);border-radius:8px;padding:8px 14px;font-size:12px;font-family:var(--mono);color:var(--cyan);display:flex;align-items:center;gap:8px;white-space:nowrap}
.car-info-chip .sep{color:var(--text3)}
.divider-or{font-size:10px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.1em;padding:0 4px}

.action-bar{position:sticky;top:0;z-index:50;background:rgba(6,11,24,.9);backdrop-filter:blur(16px);border-bottom:1px solid var(--border);padding:10px 40px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.btn{display:inline-flex;align-items:center;gap:7px;padding:8px 18px;border-radius:7px;font-size:12.5px;font-weight:600;font-family:var(--ui);cursor:pointer;border:1px solid transparent;transition:all 140ms ease;letter-spacing:.02em;white-space:nowrap}
.btn-add{background:linear-gradient(135deg,#1d4ed8,#2563eb);color:#fff;border-color:#3b82f6;box-shadow:0 0 16px rgba(59,130,246,.3)}
.btn-add:hover{box-shadow:0 0 28px rgba(59,130,246,.55);transform:translateY(-1px)}
.btn-export{background:rgba(16,185,129,.12);color:#34d399;border-color:rgba(16,185,129,.35)}
.btn-export:hover{background:rgba(16,185,129,.22);box-shadow:0 0 16px rgba(16,185,129,.2);transform:translateY(-1px)}
.btn-reset{background:rgba(245,158,11,.08);color:var(--amber);border-color:rgba(245,158,11,.25)}
.btn-reset:hover{background:rgba(245,158,11,.18);transform:translateY(-1px)}
.bar-right{margin-left:auto;display:flex;align-items:center;gap:14px}
.bar-counter{font-family:var(--mono);font-size:11px;color:var(--text3)}
.bar-counter strong{color:var(--cyan)}

.form-page{position:relative;z-index:1;padding:28px 40px 50px;display:flex;flex-direction:column;gap:24px}
.sec{background:var(--s1);border:1px solid var(--border);border-radius:14px;overflow:hidden;transition:border-color 200ms}
.sec:focus-within{border-color:var(--border2)}
.sec-head{display:flex;align-items:center;gap:12px;padding:12px 18px;border-bottom:1px solid var(--border);position:relative;overflow:hidden}
.sec-head::before{content:'';position:absolute;left:0;top:0;bottom:0;width:3px;background:var(--accent,var(--blue))}
.sec-head::after{content:'';position:absolute;left:0;top:0;right:0;bottom:0;background:linear-gradient(90deg,rgba(var(--accent-rgb,59,130,246),.07),transparent 60%);pointer-events:none}
.sec-icon{width:30px;height:30px;border-radius:7px;flex-shrink:0;display:flex;align-items:center;justify-content:center;background:rgba(var(--accent-rgb,59,130,246),.12);border:1px solid rgba(var(--accent-rgb,59,130,246),.2);position:relative;z-index:1}
.sec-icon svg{width:14px;height:14px}
.sec-title-wrap{position:relative;z-index:1}
.sec-cat{font-size:9px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--accent,var(--blue));margin-bottom:1px}
.sec-title{font-size:13px;font-weight:700;color:var(--text)}
.sec-body{padding:16px 16px 18px}
.fgrid-4{display:grid;grid-template-columns:repeat(4,1fr);gap:10px}
.field{display:flex;flex-direction:column;gap:5px}
.field label{font-size:9.5px;font-weight:600;color:var(--text2);letter-spacing:.06em;text-transform:uppercase;display:flex;align-items:center;gap:5px}
.field label .tag{background:rgba(var(--accent-rgb,59,130,246),.12);color:var(--accent,var(--blue));border-radius:3px;padding:1px 5px;font-size:8.5px;font-weight:700}
.field input{background:var(--s2);border:1px solid var(--border);border-radius:7px;color:var(--text);font-family:var(--mono);font-size:12px;padding:8px 12px;outline:none;transition:all 150ms;width:100%}
.field input::placeholder{color:var(--text3);font-size:11px}
.field input:hover{border-color:var(--border2)}
.field input:focus{border-color:var(--accent,var(--blue));background:var(--s3);box-shadow:0 0 0 3px rgba(var(--accent-rgb,59,130,246),.15);color:#fff}
.field input:not(:placeholder-shown):not(:focus){border-color:rgba(16,185,129,.3)}
.pair-row{display:grid;grid-template-columns:1fr 1fr;gap:10px;padding:10px;background:var(--s2);border-radius:9px;border:1px solid var(--border)}
.pair-label{grid-column:1/-1;font-size:9px;font-weight:700;color:var(--text3);text-transform:uppercase;letter-spacing:.1em;display:flex;align-items:center;gap:7px}
.pair-label .pip{width:7px;height:7px;border-radius:50%;background:var(--accent,var(--blue));box-shadow:0 0 6px var(--accent,var(--blue))}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px}
.sec-grate  {--accent:#3b82f6;--accent-rgb:59,130,246}
.sec-centre {--accent:#10b981;--accent-rgb:16,185,129}
.sec-side   {--accent:#f97316;--accent-rgb:249,115,22}
.sec-ssb    {--accent:#22d3ee;--accent-rgb:34,211,238}
.sec-wheel  {--accent:#f59e0b;--accent-rgb:245,158,11}
.sec-bearing{--accent:#8b5cf6;--accent-rgb:139,92,246}
.sec-pressure{--accent:#14b8a6;--accent-rgb:20,184,166}
.sec-sag    {--accent:#f43f5e;--accent-rgb:244,63,94}
.sec-rev    {--accent:#6366f1;--accent-rgb:99,102,241}
.records-panel{position:relative;z-index:1;margin:0 40px 36px;background:var(--s1);border:1px solid var(--border);border-radius:14px;overflow:hidden}
.records-head{display:flex;align-items:center;justify-content:space-between;padding:12px 18px;border-bottom:1px solid var(--border);background:var(--s2)}
.records-head h3{font-size:12.5px;font-weight:700;display:flex;align-items:center;gap:8px}
.records-head h3 span{background:var(--blue);color:#fff;border-radius:20px;padding:1px 9px;font-size:11px}
.tscroll{overflow-x:auto;max-height:260px;overflow-y:auto;scrollbar-width:thin;scrollbar-color:var(--border) transparent}
.tscroll::-webkit-scrollbar{height:5px;width:5px}
.tscroll::-webkit-scrollbar-thumb{background:var(--border);border-radius:4px}
.rtable{border-collapse:collapse;width:max-content;min-width:100%;font-size:11px}
.rtable th{padding:8px 12px;background:var(--s2);color:var(--text3);font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;white-space:nowrap;border-bottom:1px solid var(--border);position:sticky;top:0}
.rtable td{padding:7px 12px;border-bottom:1px solid #0d1428;white-space:nowrap;font-family:var(--mono);color:var(--text)}
.rtable td.empty{color:var(--text3);font-style:italic;font-size:10px}
.rtable tr:hover td{background:var(--s3)}
.rtable tr:last-child td{border-bottom:none}
.empty-state{text-align:center;padding:32px;color:var(--text3);font-size:12px}
.empty-state svg{display:block;margin:0 auto 12px;opacity:.3}
.filled-badge{margin-left:auto;font-size:10px;color:var(--text3);font-family:var(--mono);background:var(--s3);border:1px solid var(--border);border-radius:20px;padding:2px 8px}
.filled-badge.has-data{color:#10b981;border-color:rgba(16,185,129,.3)}
#toast{position:fixed;bottom:24px;right:24px;z-index:9999;display:none;min-width:220px;background:var(--s3);border-radius:10px;padding:12px 16px;font-size:13px;font-weight:600;box-shadow:0 8px 40px rgba(0,0,0,.6),0 0 0 1px var(--border2);animation:toastin 250ms cubic-bezier(.34,1.56,.64,1)}
@keyframes toastin{from{transform:translateY(20px) scale(.95);opacity:0}to{transform:none;opacity:1}}
@media(max-width:900px){.hero,.form-page,.records-panel,.car-banner{padding-left:20px;padding-right:20px}.action-bar{padding-left:20px;padding-right:20px}.records-panel{margin-left:20px;margin-right:20px}.fgrid-4{grid-template-columns:repeat(2,1fr)}}
@media(max-width:640px){.two-col{grid-template-columns:1fr}.pair-row{grid-template-columns:1fr}.hero h1{font-size:18px}.stat-chip{display:none}}
</style>
</head>
<body>

<!-- QR SCAN MODAL -->
<div id="scanModal">
  <div class="scan-modal-box">
    <div class="scan-modal-title">
      <h2>
        <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <rect x="3" y="3" width="5" height="5" rx="1"/><rect x="16" y="3" width="5" height="5" rx="1"/>
          <rect x="3" y="16" width="5" height="5" rx="1"/>
          <path d="M16 16h5M16 19h2M19 16v5M3 9h3M9 3v3M9 9h3v3H9zM9 16v3M12 16h1M12 12v1"/>
        </svg>
        Scan QR Code
      </h2>
      <button class="scan-close" id="scanCloseBtn">&#x2715; Close</button>
    </div>
    <div class="scan-cam-sel">
      <label>Camera</label>
      <select id="camSelect"><option value="">Detecting cameras…</option></select>
    </div>
    <div class="scan-cam-wrap">
      <video id="scanVideo" playsinline autoplay muted></video>
      <div class="scan-overlay">
        <div class="scan-frame"><div class="scan-line"></div></div>
      </div>
    </div>
    <div class="scan-status" id="scanStatus">
      <span class="dot"></span>
      <span id="scanStatusText">Initialising camera…</span>
    </div>
    <div class="scan-actions">
      <button class="btn btn-reset" id="scanCancelBtn">Cancel</button>
    </div>
  </div>
</div>

<header class="hero">
  <div class="hero-left">
    <div class="eyebrow"><span class="live-dot"></span>PP2 Industrial Maintenance System</div>
    <h1>Car History Maintenance Log</h1>
    <p class="hero-sub">Scan QR or type QR No → CC auto-fills · or select CC No directly · Fill → Add Row → Export</p>
  </div>
  <div class="hero-right">
    <div class="stat-chip"><div class="val" id="filledCount">0</div><div class="lbl">Filled</div></div>
    <div class="stat-chip"><div class="val" id="rowsCount">0</div><div class="lbl">Saved</div></div>
  </div>
</header>

<div class="car-banner">
  <div class="car-banner-inner">
    <span class="car-banner-label">QR No</span>
    <div class="lookup-wrap">
      <input type="text" id="qrInput" placeholder="e.g. 151" autocomplete="off"/>
      <div class="lookup-badge" id="qrBadge">
        <svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="8" r="6"/><path d="M8 5v3l2 2"/></svg>
      </div>
      <button class="btn-scan-qr" id="btnScanQR" title="Scan QR code with camera">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
          <rect x="3" y="3" width="5" height="5" rx="1"/>
          <rect x="16" y="3" width="5" height="5" rx="1"/>
          <rect x="3" y="16" width="5" height="5" rx="1"/>
          <path d="M16 16h5M16 19h2M19 16v5M3 9h3M9 3v3M9 9h3v3H9zM9 16v3M12 16h1M12 12v1"/>
        </svg>
        <span class="scan-label">Scan</span>
      </button>
    </div>
    <span class="divider-or">or</span>
    <span class="car-banner-label">CC No</span>
    <div class="car-select-wrap">
      <select id="ccSelect">
        <option value="">— Choose CC Number —</option>
        <option>CC No-01</option><option>CC No-02</option><option>CC No-SP4</option><option>CC No-05</option><option>CC No-06</option><option>CC No-07</option><option>CC No-08</option><option>CC NO-SP9</option><option>CC No-09</option><option>CC No-AK10</option><option>CC No-12</option><option>CC No-13</option><option>CC No-14</option><option>CC No-15</option><option>CC No-16</option><option>CC No-SP17</option><option>CC No-SP18</option><option>CC No-20</option><option>CC no-21</option><option>CC No-22</option><option>CC no-23</option><option>CC No-25</option><option>CC No-AK26</option><option>CC No-27</option><option>CC No-28</option><option>CC NO-29</option><option>CC No-31</option><option>CC No-32</option><option>CC No-33</option><option>CC No-34</option><option>CC No-35</option><option>CC No-36</option><option>CC No-36-1</option><option>CC No-37</option><option>CC No-38</option><option>CC No-39</option><option>CC No-SP40</option><option>CC No-40 </option><option>CC No-42</option><option>CC No-42|1</option><option>CC No-43</option><option>CC No-44</option><option>CC No-AK45</option><option>CC No-47</option><option>CC No-48</option><option>CC No-49</option><option>CC No-AK51</option><option>CC No-52|1</option><option>CC NO-SP52</option><option>CC No-52</option><option>CC No-AK53</option><option>CC No-AK53-</option><option>CC No-SP54</option><option>CC No-56</option><option>CC No-AN57</option><option>CC No-58</option><option>CC No-60</option><option>CC No-61</option><option>CC No-62</option><option>CC No-63</option><option>CC No-AN64</option><option>CC No-SP66</option><option>CC No-67</option><option>CC No-SP68</option><option>CC No-69</option><option>CC No-71</option><option>CC No-71|1 </option><option>CC No-72</option><option>CC No-72|1 </option><option>CC No-73</option><option>CC No-74</option><option>CC No-75</option><option>CC No-76</option><option>CC No-76|1</option><option>CC No-77</option><option>CC N-78</option><option>CC N0-79</option><option>Cc No-80</option><option>CC No-81</option><option>CC No-AN82</option><option>CC No-83</option><option>CC No-85</option><option>CC NO-86</option><option>CC No-87</option><option>CC No-AK88</option><option>CC No-SP92</option><option>CC No-93</option><option>CC NO-95</option><option>CC No-97</option><option>CCno-98</option><option>CC No -AK99</option><option>CC No-100</option><option>CC No-101</option><option>CC No-105</option><option>CC no-106</option><option>CC No-SP107</option><option>CC No-109</option><option>CC No-110</option><option>CC No-111</option><option>CC No-113</option><option>CC No-114</option><option>CC No-SP117</option><option>CC No-120</option><option>CC No-121</option><option>CC No-SP123</option><option>CC No-124</option><option>CC No-125</option><option>CC No-SP126</option><option>CC No - 128</option><option>CC No-129</option><option>CC No -SP133</option><option>CCNo-134</option><option>CC No -136</option><option>CC No-137</option><option>CC No-138</option><option>CC No-139</option><option>CC No-SP140</option><option>CC No -141 </option><option>CC No-143</option><option>CC No-144</option><option>CC No-145</option><option>CC No-SP146</option><option>CC No - 147</option><option>CC No-148</option><option>CC N0 149</option><option>CC No-152</option><option>CC No-AK154</option><option>CC No-155</option><option>CC No-161</option><option>CC No-SP162</option><option>CC No-163|2</option><option>CC No-163</option><option>CC No-164</option><option>CC No-SP165</option><option>CC No-167</option><option>CC No-167|1</option><option>CC No-168</option><option>CC No-169</option><option>CC No-170|1</option><option>CC No 170|2</option><option>CC No-171</option><option>CC No-SP172</option><option>CC No-AN173</option><option>CC No-174</option><option>CC No-AK175</option><option>CC No-177</option><option>CC No-180</option><option>CC No-181</option><option>CC No-182</option><option>CC No-200</option><option>CC No-AK200</option><option>CC No -SP201</option><option>CC No-001</option><option>CC No-002</option><option>CC No-003</option><option>CC No-004</option><option>CC No-005</option><option>CC No - 006</option><option>CC No-007</option><option>CC No-008</option><option>CC No-009</option><option>CC N0-010</option><option>CC No-011</option><option>CC No-012</option><option>CC No-013</option><option>CC No-014</option><option>CC No-015</option><option>CC No-016</option><option>CC No-017</option><option>CC No-018</option><option>CC No - 019</option><option>CC No-020</option><option>CC N0-021</option><option>CC No-022</option><option>CC No-023</option><option>CC No-024</option><option>CC No-025</option><option>CC No-08(Metso)</option><option>CC No-SP7</option><option>CC No-119</option>
      </select>
    </div>
    <div class="car-info-chip" id="carInfoChip" style="opacity:0;pointer-events:none">
      <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 110 14A7 7 0 018 1zm0 1.5a5.5 5.5 0 100 11 5.5 5.5 0 000-11zM8 4a.75.75 0 01.75.75v3.5h2a.75.75 0 010 1.5h-2.75A.75.75 0 018 9V4.75A.75.75 0 018 4z"/></svg>
      <span id="chipCC">—</span><span class="sep">/</span><span id="chipQR">—</span>
    </div>
  </div>
</div>

<div class="action-bar">
  <button class="btn btn-add" id="btnAddRow">
    <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M8 2a.75.75 0 01.75.75v4.5h4.5a.75.75 0 010 1.5h-4.5v4.5a.75.75 0 01-1.5 0v-4.5h-4.5a.75.75 0 010-1.5h4.5v-4.5A.75.75 0 018 2z"/></svg>
    Add Row to Log
  </button>
  <button class="btn btn-export" id="btnExport">
    <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path d="M2.75 14A1.75 1.75 0 011 12.25v-2.5a.75.75 0 011.5 0v2.5c0 .138.112.25.25.25h10.5a.25.25 0 00.25-.25v-2.5a.75.75 0 011.5 0v2.5A1.75 1.75 0 0113.25 14H2.75zM7.25 7.689L5.03 5.47a.75.75 0 00-1.06 1.06l3.5 3.5a.75.75 0 001.06 0l3.5-3.5a.75.75 0 00-1.06-1.06L8.75 7.689V2a.75.75 0 00-1.5 0v5.689z"/></svg>
    Export Excel
  </button>
  <button class="btn btn-reset" id="btnReset">
    <svg width="13" height="13" viewBox="0 0 16 16" fill="currentColor"><path fill-rule="evenodd" d="M1.705 8.005a.75.75 0 01.834.656 5.5 5.5 0 009.592 2.97l-1.204-1.204a.25.25 0 01.177-.427h3.646a.25.25 0 01.25.25v3.646a.25.25 0 01-.427.177l-1.38-1.38A7.002 7.002 0 011.05 8.84a.75.75 0 01.656-.834zM8 2.5a5.487 5.487 0 00-4.131 1.869l1.204 1.204A.25.25 0 014.896 6H1.25A.25.25 0 011 5.75V2.104a.25.25 0 01.427-.177l1.38 1.38A7.002 7.002 0 0114.95 7.16a.75.75 0 01-1.49.178A5.5 5.5 0 008 2.5z"/></svg>
    Clear Form
  </button>
  <div class="bar-right">
    <div class="bar-counter">Session rows: <strong id="barRows">0</strong></div>
  </div>
</div>

<div class="form-page">
  <div class="two-col">
    <div class="sec sec-grate">
      <div class="sec-head">
        <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><path d="M3 6h18M3 12h18M3 18h18"/></svg></div>
        <div class="sec-title-wrap"><div class="sec-cat">Section 01</div><div class="sec-title">Grate Bars</div></div>
        <div class="filled-badge" id="badge-grate">0 / 2</div>
      </div>
      <div class="sec-body">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
          <div class="field"><label>Replacement Date <span class="tag">DATE</span></label><input id="grate_repl_date" type="text" placeholder="DD/MM/YYYY" data-sec="grate"/></div>
          <div class="field"><label>Make</label><input id="grate_make" type="text" placeholder="e.g. Metso" data-sec="grate"/></div>
        </div>
      </div>
    </div>
    <div class="sec sec-centre">
      <div class="sec-head">
        <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><circle cx="12" cy="12" r="3"/><path d="M12 2v4M12 18v4M4.22 4.22l2.83 2.83M16.95 16.95l2.83 2.83M2 12h4M18 12h4M4.22 19.78l2.83-2.83M16.95 7.05l2.83-2.83"/></svg></div>
        <div class="sec-title-wrap"><div class="sec-cat">Section 02</div><div class="sec-title">Centre Casting</div></div>
        <div class="filled-badge" id="badge-centre">0 / 2</div>
      </div>
      <div class="sec-body">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
          <div class="field"><label>CC Replacement Date <span class="tag">DATE</span></label><input id="cc_repl_date" type="text" placeholder="DD/MM/YYYY" data-sec="centre"/></div>
          <div class="field"><label>CC Make</label><input id="cc_make" type="text" placeholder="Manufacturer" data-sec="centre"/></div>
        </div>
      </div>
    </div>
  </div>

  <div class="sec sec-side">
    <div class="sec-head">
      <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><path d="M2 9h20M2 15h20M9 3v18M15 3v18"/></svg></div>
      <div class="sec-title-wrap"><div class="sec-cat">Section 03</div><div class="sec-title">Side Casting</div></div>
      <div class="filled-badge" id="badge-side">0 / 4</div>
    </div>
    <div class="sec-body">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="pair-row">
          <div class="pair-label"><span class="pip" style="background:#fb923c;box-shadow:0 0 6px #fb923c"></span>East Side</div>
          <div class="field"><label>Replacement Date <span class="tag">DATE</span></label><input id="east_sc_repl" type="text" placeholder="DD/MM/YYYY" data-sec="side"/></div>
          <div class="field"><label>Make</label><input id="east_sc_make" type="text" placeholder="Manufacturer" data-sec="side"/></div>
        </div>
        <div class="pair-row">
          <div class="pair-label"><span class="pip" style="background:#fbbf24;box-shadow:0 0 6px #fbbf24"></span>West Side</div>
          <div class="field"><label>Replacement Date <span class="tag">DATE</span></label><input id="west_sc_repl" type="text" placeholder="DD/MM/YYYY" data-sec="side"/></div>
          <div class="field"><label>Make</label><input id="west_sc_make" type="text" placeholder="Manufacturer" data-sec="side"/></div>
        </div>
      </div>
    </div>
  </div>

  <div class="sec sec-ssb">
    <div class="sec-head">
      <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><rect x="3" y="3" width="7" height="7"/><rect x="14" y="3" width="7" height="7"/><rect x="14" y="14" width="7" height="7"/><rect x="3" y="14" width="7" height="7"/></svg></div>
      <div class="sec-title-wrap"><div class="sec-cat">Section 04</div><div class="sec-title">SSB — Shoe Slide Bearings</div></div>
      <div class="filled-badge" id="badge-ssb">0 / 4</div>
    </div>
    <div class="sec-body">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        <div class="pair-row">
          <div class="pair-label"><span class="pip" style="background:#38bdf8;box-shadow:0 0 6px #38bdf8"></span>East SSB</div>
          <div class="field"><label>Replacement Date <span class="tag">DATE</span></label><input id="east_ssb_repl" type="text" placeholder="DD/MM/YYYY" data-sec="ssb"/></div>
          <div class="field"><label>Make of East SSB</label><input id="east_ssb_make" type="text" placeholder="Manufacturer" data-sec="ssb"/></div>
        </div>
        <div class="pair-row">
          <div class="pair-label"><span class="pip" style="background:#818cf8;box-shadow:0 0 6px #818cf8"></span>West SSB</div>
          <div class="field"><label>Replacement Date <span class="tag">DATE</span></label><input id="west_ssb_repl" type="text" placeholder="DD/MM/YYYY" data-sec="ssb"/></div>
          <div class="field"><label>Make of West SSB</label><input id="west_ssb_make" type="text" placeholder="Manufacturer" data-sec="ssb"/></div>
        </div>
      </div>
    </div>
  </div>

  <div class="sec sec-wheel">
    <div class="sec-head">
      <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="3"/><path d="M12 3v3M12 18v3M3 12h3M18 12h3"/></svg></div>
      <div class="sec-title-wrap"><div class="sec-cat">Section 05</div><div class="sec-title">Wheel Replacement</div></div>
      <div class="filled-badge" id="badge-wheel">0 / 4</div>
    </div>
    <div class="sec-body">
      <div class="fgrid-4">
        <div class="field"><label>EN <span class="tag">DATE</span></label><input id="wr_en" type="text" placeholder="DD/MM/YYYY" data-sec="wheel"/></div>
        <div class="field"><label>ES <span class="tag">DATE</span></label><input id="wr_es" type="text" placeholder="DD/MM/YYYY" data-sec="wheel"/></div>
        <div class="field"><label>WN <span class="tag">DATE</span></label><input id="wr_wn" type="text" placeholder="DD/MM/YYYY" data-sec="wheel"/></div>
        <div class="field"><label>WS <span class="tag">DATE</span></label><input id="wr_ws" type="text" placeholder="DD/MM/YYYY" data-sec="wheel"/></div>
      </div>
    </div>
  </div>

  <div class="sec sec-bearing">
    <div class="sec-head">
      <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="4"/><path d="M12 3v2M12 19v2M3 12h2M19 12h2"/></svg></div>
      <div class="sec-title-wrap"><div class="sec-cat">Section 06</div><div class="sec-title">Wheel Bearing Replacement</div></div>
      <div class="filled-badge" id="badge-bearing">0 / 8</div>
    </div>
    <div class="sec-body">
      <div class="fgrid-4">
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip"></span>EN</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="wb_en" type="text" placeholder="DD/MM/YYYY" data-sec="bearing"/></div>
          <div class="field"><label>Make</label><input id="wb_en_make" type="text" placeholder="Maker" data-sec="bearing"/></div>
        </div>
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip"></span>ES</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="wb_es" type="text" placeholder="DD/MM/YYYY" data-sec="bearing"/></div>
          <div class="field"><label>Make</label><input id="wb_es_make" type="text" placeholder="Maker" data-sec="bearing"/></div>
        </div>
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip"></span>WN</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="wb_wn" type="text" placeholder="DD/MM/YYYY" data-sec="bearing"/></div>
          <div class="field"><label>Make</label><input id="wb_wn_make" type="text" placeholder="Maker" data-sec="bearing"/></div>
        </div>
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip"></span>WS</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="wb_ws" type="text" placeholder="DD/MM/YYYY" data-sec="bearing"/></div>
          <div class="field"><label>Make</label><input id="wb_ws_make" type="text" placeholder="Maker" data-sec="bearing"/></div>
        </div>
      </div>
    </div>
  </div>

  <div class="sec sec-pressure">
    <div class="sec-head">
      <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><rect x="2" y="8" width="20" height="8" rx="4"/><circle cx="8" cy="12" r="2"/><circle cx="16" cy="12" r="2"/></svg></div>
      <div class="sec-title-wrap"><div class="sec-cat">Section 07</div><div class="sec-title">Pressure Roller Replacement</div></div>
      <div class="filled-badge" id="badge-pressure">0 / 8</div>
    </div>
    <div class="sec-body">
      <div class="fgrid-4">
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip" style="background:#14b8a6;box-shadow:0 0 6px #14b8a6"></span>EN</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="pr_en" type="text" placeholder="DD/MM/YYYY" data-sec="pressure"/></div>
          <div class="field"><label>Make</label><input id="pr_en_make" type="text" placeholder="Maker" data-sec="pressure"/></div>
        </div>
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip" style="background:#14b8a6;box-shadow:0 0 6px #14b8a6"></span>ES</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="pr_es" type="text" placeholder="DD/MM/YYYY" data-sec="pressure"/></div>
          <div class="field"><label>Make</label><input id="pr_es_make" type="text" placeholder="Maker" data-sec="pressure"/></div>
        </div>
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip" style="background:#14b8a6;box-shadow:0 0 6px #14b8a6"></span>WN</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="pr_wn" type="text" placeholder="DD/MM/YYYY" data-sec="pressure"/></div>
          <div class="field"><label>Make</label><input id="pr_wn_make" type="text" placeholder="Maker" data-sec="pressure"/></div>
        </div>
        <div class="pair-row" style="grid-template-columns:1fr;gap:8px">
          <div class="pair-label"><span class="pip" style="background:#14b8a6;box-shadow:0 0 6px #14b8a6"></span>WS</div>
          <div class="field"><label>Date <span class="tag">DATE</span></label><input id="pr_ws" type="text" placeholder="DD/MM/YYYY" data-sec="pressure"/></div>
          <div class="field"><label>Make</label><input id="pr_ws_make" type="text" placeholder="Maker" data-sec="pressure"/></div>
        </div>
      </div>
    </div>
  </div>

  <div class="two-col">
    <div class="sec sec-sag">
      <div class="sec-head">
        <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><polyline points="22 7 13.5 15.5 8.5 10.5 2 17"/><polyline points="16 7 22 7 22 13"/></svg></div>
        <div class="sec-title-wrap"><div class="sec-cat">Section 08</div><div class="sec-title">SAG Measurements</div></div>
        <div class="filled-badge" id="badge-sag">0 / 4</div>
      </div>
      <div class="sec-body">
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px">
          <div class="pair-row">
            <div class="pair-label"><span class="pip" style="background:#f43f5e;box-shadow:0 0 6px #f43f5e"></span>N-Sag</div>
            <div class="field"><label>Measured Date <span class="tag">DATE</span></label><input id="n_sag_date" type="text" placeholder="DD/MM/YYYY" data-sec="sag"/></div>
            <div class="field"><label>N Sag Value</label><input id="n_sag_value" type="text" placeholder="0.00 mm" data-sec="sag"/></div>
          </div>
          <div class="pair-row">
            <div class="pair-label"><span class="pip" style="background:#fb923c;box-shadow:0 0 6px #fb923c"></span>S-Sag</div>
            <div class="field"><label>Measured Date <span class="tag">DATE</span></label><input id="s_sag_date" type="text" placeholder="DD/MM/YYYY" data-sec="sag"/></div>
            <div class="field"><label>S Sag Value</label><input id="s_sag_value" type="text" placeholder="0.00 mm" data-sec="sag"/></div>
          </div>
        </div>
      </div>
    </div>
    <div class="sec sec-rev">
      <div class="sec-head">
        <div class="sec-icon"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" style="color:var(--accent)"><path d="M17 1l4 4-4 4M3 11V9a4 4 0 014-4h14M7 23l-4-4 4-4M21 13v2a4 4 0 01-4 4H3"/></svg></div>
        <div class="sec-title-wrap"><div class="sec-cat">Section 09</div><div class="sec-title">Car Reversed</div></div>
        <div class="filled-badge" id="badge-reversed">0 / 1</div>
      </div>
      <div class="sec-body">
        <div class="field"><label>Car Reversed Date <span class="tag">DATE</span></label><input id="car_reversed" type="text" placeholder="DD/MM/YYYY" data-sec="reversed"/></div>
        <p style="margin-top:10px;font-size:10.5px;color:var(--text3);line-height:1.6">Record the date the car was reversed. Leave blank if not applicable.</p>
      </div>
    </div>
  </div>
</div>

<div class="records-panel">
  <div class="records-head">
    <h3>Session Queue <span id="recBadge">0</span></h3>
    <div style="font-size:11px;color:var(--text3);font-family:var(--mono)">Last <strong id="recLast" style="color:var(--text2)">—</strong> shown</div>
  </div>
  <div class="tscroll">
    <table class="rtable" id="recTable">
      <thead id="recHead"></thead>
      <tbody id="recBody">
        <tr><td colspan="99"><div class="empty-state">
          <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"/></svg>
          Scan a QR code or type QR No, fill the form, then click <strong style="color:var(--cyan)">Add Row to Log</strong>.
        </div></td></tr>
      </tbody>
    </table>
  </div>
</div>

<div id="toast"></div>

<script>
/* ══ LOOKUP TABLES ══ */
const CC_TO_QR = {"1":"57","2":"43","SP4":"59","5":"130","6":"46","7":"160","8":"72","SP9":"91","9":"105","AK10":"155","12":"86","13":"4","14":"65","15":"149","16":"111","SP17":"47","SP18":"41","20":"29","21":"97","22":"126","23":"62","25":"87","AK26":"90","27":"73","28":"141","29":"","31":"12","32":"23","33":"6","34":"152","35":"171","36":"161","36-1":"","37":"88","38":"26","39":"82","SP40":"121","40":"55","42":"101","42/1":"22","43":"52","44":"98","AK45":"104","47":"35","48":"129","49":"","AK51":"69","52/1":"107","SP52":"115","52":"","AK53":"78","SP54":"140","56":"","AN57":"124","58":"172","60":"44","61":"95","62":"60","63":"53","AN64":"28","SP66":"159","67":"71","SP68":"110","69":"133","71":"154","71/1":"96","72":"30","72/1":"32","73":"31","74":"74","75":"169","76":"66","76/1":"102","77":"84","78":"37","79":"","80":"13","81":"167","AN82":"93","83":"2","85":"76","86":"","87":"79","AK88":"99","SP92":"135","93":"63","95":"45","97":"139","98":"123","AK99":"56","100":"127","101":"38","105":"100","106":"122","SP107":"125","109":"165","110":"38","111":"61","113":"","114":"157","SP117":"117","120":"25","121":"16","SP123":"108","124":"14","125":"20","SP126":"19","128":"10","129":"81","SP133":"170","134":"131","136":"136","137":"24","138":"1","139":"156","SP140":"42","141":"","143":"","144":"143","145":"50","SP146":"92","147":"33","148":"36","149":"","152":"15","AK154":"153","155":"118","161":"83","SP162":"147","163/2":"64","163":"128","164":"168","SP165":"119","167":"113","167/1":"68","168":"162","169":"173","170/1":"158","170/2":"112","171":"120","SP172":"8","AN173":"146","174":"174","AK175":"17","177":"3","180":"21","181":"103","182":"67","200":"150","AK200":"89","SP201":"48","3":"116","4":"166","10":"132","11":"","17":"75","18":"85","19":"144","24":"148","08(Metso)":"106","SP7":"18","119":"163"};
const QR_TO_CC = {"57":"1","43":"2","11":"2","59":"SP4","130":"5","51":"5","46":"6","109":"6","160":"7","7":"7","72":"8","137":"8","91":"SP9","105":"9","155":"AK10","86":"12","151":"12","4":"13","80":"13","65":"14","27":"14","149":"15","58":"15","111":"16","40":"16","47":"SP17","41":"SP18","29":"20","97":"21","77":"21","126":"22","9":"22","62":"23","164":"23","148":"24","87":"25","94":"25","90":"AK26","73":"27","141":"28","12":"31","23":"32","6":"33","152":"34","171":"35","161":"36","88":"37","26":"38","82":"39","121":"SP40","55":"40","101":"42","22":"42/1","52":"43","98":"44","104":"AK45","35":"47","129":"48","69":"AK51","107":"52/1","115":"SP52","78":"AK53","140":"SP54","124":"AN57","172":"58","44":"60","95":"61","60":"62","53":"63","28":"AN64","159":"SP66","71":"67","110":"SP68","133":"69","154":"71","96":"71/1","30":"72","32":"72/1","31":"73","74":"74","169":"75","66":"76","102":"76/1","84":"77","37":"78","13":"80","167":"81","93":"AN82","2":"83","76":"85","79":"87","99":"AK88","135":"SP92","63":"93","45":"95","139":"97","123":"98","56":"AK99","127":"100","38":"110","100":"105","122":"106","125":"SP107","165":"109","61":"111","157":"114","117":"SP117","25":"120","16":"121","108":"SP123","14":"124","20":"125","19":"SP126","10":"128","81":"129","170":"SP133","131":"134","136":"136","24":"137","1":"138","156":"139","42":"SP140","143":"144","50":"145","92":"SP146","33":"147","36":"148","15":"152","153":"AK154","118":"155","83":"161","147":"SP162","64":"163/2","128":"163","168":"164","119":"SP165","113":"167","68":"167/1","162":"168","173":"169","158":"170/1","112":"170/2","120":"171","8":"SP172","146":"AN173","174":"174","17":"AK175","3":"177","21":"180","103":"181","67":"182","150":"200","89":"AK200","48":"SP201","116":"3","166":"4","132":"10","75":"17","85":"18","144":"19","106":"08(Metso)","18":"SP7","163":"119"};

const CC_SUFFIX_TO_OPTION = {};
document.querySelectorAll('#ccSelect option').forEach(opt => {
  if(!opt.value) return;
  const raw = opt.value.trim();
  const m = raw.match(/(?:CC\s*[Nn][Oo0][-\s]*)([\w\/\|\.]+)/i);
  if(m) CC_SUFFIX_TO_OPTION[m[1].trim()] = raw;
  CC_SUFFIX_TO_OPTION[raw] = raw;
});
function suffixToOption(suffix) {
  if(CC_SUFFIX_TO_OPTION[suffix]) return CC_SUFFIX_TO_OPTION[suffix];
  for(const [k,v] of Object.entries(CC_SUFFIX_TO_OPTION))
    if(k.replace(/[-\s]/g,'').toLowerCase() === suffix.replace(/[-\s]/g,'').toLowerCase()) return v;
  return null;
}

/* ══ QR INPUT LOGIC ══ */
let _suppressCCChange = false;

function applyQRValue(qr) {
  const badge = document.getElementById('qrBadge');
  const sel = document.getElementById('ccSelect');
  const inp = document.getElementById('qrInput');
  inp.value = qr;
  if(!qr) {
    badge.className='lookup-badge';
    badge.innerHTML='<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="8" r="6"/><path d="M8 5v3l2 2"/></svg>';
    inp.className=''; updateChip(); return;
  }
  const ccSuffix = QR_TO_CC[qr];
  if(ccSuffix) {
    const optionText = suffixToOption(ccSuffix);
    if(optionText) { _suppressCCChange=true; sel.value=optionText; sel.classList.add('synced'); _suppressCCChange=false; }
    inp.classList.add('found'); inp.classList.remove('notfound');
    badge.className='lookup-badge found';
    badge.innerHTML='<svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 110 14A7 7 0 018 1zm3.78 4.97a.75.75 0 00-1.06-1.06L7 8.69 5.28 6.97a.75.75 0 00-1.06 1.06l2.25 2.25a.75.75 0 001.06 0l4.25-4.25z"/></svg>';
  } else {
    inp.classList.add('notfound'); inp.classList.remove('found');
    badge.className='lookup-badge notfound';
    badge.innerHTML='<svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 110 14A7 7 0 018 1zm2.78 4.97a.75.75 0 00-1.06 0L8 6.69 6.28 4.97a.75.75 0 00-1.06 1.06L6.94 7.75 5.22 9.47a.75.75 0 001.06 1.06L8 8.81l1.72 1.72a.75.75 0 001.06-1.06L9.06 7.75l1.72-1.72a.75.75 0 000-1.06z"/></svg>';
    sel.value=''; sel.classList.remove('synced');
  }
  updateChip();
}

document.getElementById('qrInput').addEventListener('input', function() {
  applyQRValue(this.value.trim());
});

document.getElementById('ccSelect').addEventListener('change', function() {
  if(_suppressCCChange) return;
  const raw = this.value.trim();
  document.getElementById('qrInput').value=''; document.getElementById('qrInput').className='';
  document.getElementById('qrBadge').className='lookup-badge';
  document.getElementById('qrBadge').innerHTML='<svg viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><circle cx="8" cy="8" r="6"/><path d="M8 5v3l2 2"/></svg>';
  this.classList.remove('synced');
  if(!raw) { updateChip(); return; }
  const m = raw.match(/(?:CC\s*[Nn][Oo0][-\s]*)([\w\/\|\.]+)/i);
  const suffix = m ? m[1].trim() : raw;
  const qr = CC_TO_QR[suffix];
  if(qr) {
    document.getElementById('qrInput').value=qr; document.getElementById('qrInput').classList.add('found');
    document.getElementById('qrBadge').className='lookup-badge found';
    document.getElementById('qrBadge').innerHTML='<svg viewBox="0 0 16 16" fill="currentColor"><path d="M8 1a7 7 0 110 14A7 7 0 018 1zm3.78 4.97a.75.75 0 00-1.06-1.06L7 8.69 5.28 6.97a.75.75 0 00-1.06 1.06l2.25 2.25a.75.75 0 001.06 0l4.25-4.25z"/></svg>';
    this.classList.add('synced');
  }
  updateChip();
});

function updateChip() {
  const cc=document.getElementById('ccSelect').value, qr=document.getElementById('qrInput').value.trim(), chip=document.getElementById('carInfoChip');
  if(cc||qr) {
    document.getElementById('chipCC').textContent=cc||'—';
    document.getElementById('chipQR').textContent=qr?'QR No-'+qr:'—';
    chip.style.opacity='1'; chip.style.pointerEvents='auto';
  } else {
    chip.style.opacity='0'; chip.style.pointerEvents='none';
  }
}

/* ══ SECTION COUNTERS ══ */
const SECS = ['grate','centre','side','ssb','wheel','bearing','pressure','sag','reversed'];
function updateCounters(){
  let total=0,filled=0;
  SECS.forEach(sec=>{
    const inputs=[...document.querySelectorAll('input[data-sec="'+sec+'"]')];
    const f=inputs.filter(i=>i.value.trim()).length;
    const b=document.getElementById('badge-'+sec);
    if(b){b.textContent=f+' / '+inputs.length;b.classList.toggle('has-data',f>0);}
    total+=inputs.length;filled+=f;
  });
  document.getElementById('filledCount').textContent=filled;
}
document.querySelectorAll('input[data-sec]').forEach(i=>i.addEventListener('input',updateCounters));

/* ══ RECORDS ══ */
let records=[];
const COLS=[
  {id:'grate_repl_date',label:'Grate bar replacement Date'},{id:'grate_make',label:'Grate Bar make'},
  {id:'cc_repl_date',label:'CC replacement date'},{id:'cc_make',label:'CC make'},
  {id:'east_sc_repl',label:'East Side casting replacement date'},{id:'east_sc_make',label:'East SC make'},
  {id:'west_sc_repl',label:'West side casting replacement date'},{id:'west_sc_make',label:'West SC make'},
  {id:'east_ssb_repl',label:'East SSB replacement Date'},{id:'east_ssb_make',label:'Make of East SSB'},
  {id:'west_ssb_repl',label:'West SSB relacement Date'},{id:'west_ssb_make',label:'Make of West SSB'},
  {id:'wr_en',label:'Wheel EN'},{id:'wr_es',label:'Wheel ES'},{id:'wr_wn',label:'Wheel WN'},{id:'wr_ws',label:'Wheel WS'},
  {id:'wb_en',label:'WB EN Date'},{id:'wb_en_make',label:'WB EN Make'},{id:'wb_es',label:'WB ES Date'},{id:'wb_es_make',label:'WB ES Make'},
  {id:'wb_wn',label:'WB WN Date'},{id:'wb_wn_make',label:'WB WN Make'},{id:'wb_ws',label:'WB WS Date'},{id:'wb_ws_make',label:'WB WS Make'},
  {id:'pr_en',label:'PR EN Date'},{id:'pr_en_make',label:'PR EN Make'},{id:'pr_es',label:'PR ES Date'},{id:'pr_es_make',label:'PR ES Make'},
  {id:'pr_wn',label:'PR WN Date'},{id:'pr_wn_make',label:'PR WN Make'},{id:'pr_ws',label:'PR WS Date'},{id:'pr_ws_make',label:'PR WS Make'},
  {id:'n_sag_date',label:'N sag measured date'},{id:'n_sag_value',label:'N sag value'},
  {id:'s_sag_date',label:'S sag measured date'},{id:'s_sag_value',label:'S sag value'},
  {id:'car_reversed',label:'Car reversed'},
];

function collectRow(){
  const cc=document.getElementById('ccSelect').value||'', qr=document.getElementById('qrInput').value.trim();
  const row={'SI No':records.length+1,'CC No':cc,'QR No':qr?'QR No-'+qr:''};
  COLS.forEach(c=>{row[c.label]=(document.getElementById(c.id)||{}).value?.trim()||'';});
  return row;
}

document.getElementById('btnAddRow').addEventListener('click',()=>{
  const cc=document.getElementById('ccSelect').value, qr=document.getElementById('qrInput').value.trim();
  if(!cc&&!qr){toast('⚠ Enter QR No or select CC No first!','#f59e0b');return;}
  const row = collectRow();
  records.push(row);
  renderTable();
  syncCounts();
  toast('✔ Row '+records.length+' saved — '+(cc||'?')+' / QR '+(qr||'?'),'#10b981');
  // Also persist to server CSV
  fetch('/log_maintenance', {
    method:'POST',
    headers:{'Content-Type':'application/json'},
    body: JSON.stringify({
      qr: qr ? 'QR No-'+qr : cc,
      cc: cc,
      tech: 'PP2 System',
      dept: '',
      issue: 'Car History Entry',
      priority: '',
      status: 'Completed',
      datetime: new Date().toISOString().replace('T',' ').substring(0,19),
      remarks: JSON.stringify(row)
    })
  }).catch(()=>{});
});

document.getElementById('btnReset').addEventListener('click',()=>{
  document.querySelectorAll('input[data-sec]').forEach(i=>i.value='');
  updateCounters(); toast('Form fields cleared','#f59e0b');
});

function syncCounts(){
  document.getElementById('rowsCount').textContent=records.length;
  document.getElementById('barRows').textContent=records.length;
  document.getElementById('recBadge').textContent=records.length;
}

const PH=['SI No','CC No','QR No','Grate Date','Grate Make','CC Date','CC Make','E-SC Date','E-SC Make','W-SC Date','W-SC Make','E-SSB Date','E-SSB Make','W-SSB Date','W-SSB Make','WR-EN','WR-ES','WR-WN','WR-WS','WB-EN','WB-EN Mk','WB-ES','WB-ES Mk','WB-WN','WB-WN Mk','WB-WS','WB-WS Mk','PR-EN','PR-EN Mk','PR-ES','PR-ES Mk','PR-WN','PR-WN Mk','PR-WS','PR-WS Mk','N-Sag Dt','N-Sag','S-Sag Dt','S-Sag','Car Rev'];
function renderTable(){
  const head=document.getElementById('recHead'), body=document.getElementById('recBody');
  if(!head.childElementCount){const tr=document.createElement('tr');PH.forEach(h=>{const th=document.createElement('th');th.textContent=h;tr.appendChild(th);});head.appendChild(tr);}
  body.innerHTML='';
  const show=records.slice(-8);
  document.getElementById('recLast').textContent=Math.min(8,records.length);
  show.forEach(row=>{
    const tr=document.createElement('tr');
    const vals=[row['SI No'],row['CC No'],row['QR No'],...COLS.map(c=>row[c.label])];
    vals.forEach(v=>{const td=document.createElement('td');td.textContent=v||'';if(!v)td.classList.add('empty');tr.appendChild(td);});
    body.appendChild(tr);
  });
  syncCounts();
}

/* ══ QR CAMERA SCANNER (ZXing) ══ */
let codeReader = null;
let scanActive = false;

const scanModal   = document.getElementById('scanModal');
const scanVideo   = document.getElementById('scanVideo');
const scanStatusEl= document.getElementById('scanStatus');
const scanStatusT = document.getElementById('scanStatusText');
const camSelect   = document.getElementById('camSelect');
const btnScanQR   = document.getElementById('btnScanQR');

function setScanStatus(msg, type) {
  scanStatusEl.className = 'scan-status' + (type ? ' '+type : '');
  scanStatusT.textContent = msg;
}

async function populateCameras() {
  try {
    const devices = await ZXing.BrowserMultiFormatReader.listVideoInputDevices();
    camSelect.innerHTML = '';
    if(!devices.length) { camSelect.innerHTML='<option value="">No camera found</option>'; setScanStatus('No camera detected','error'); return; }
    const sorted = [...devices].sort((a,b)=>{
      const aR=/back|rear|environment/i.test(a.label), bR=/back|rear|environment/i.test(b.label);
      return aR===bR?0:aR?-1:1;
    });
    sorted.forEach((d,i)=>{ const opt=document.createElement('option'); opt.value=d.deviceId; opt.textContent=d.label||'Camera '+(i+1); camSelect.appendChild(opt); });
    setScanStatus('Camera ready — point at a QR code');
  } catch(e) { setScanStatus('Camera permission denied or unavailable','error'); }
}

async function startScan(deviceId) {
  if(codeReader) { try { codeReader.reset(); } catch(e){} }
  codeReader = new ZXing.BrowserMultiFormatReader();
  setScanStatus('Starting camera…');
  try {
    const constraints = deviceId ? {deviceId:{exact:deviceId}} : {facingMode:'environment'};
    await codeReader.decodeFromConstraints({video:constraints}, scanVideo, (result, err) => {
      if(result) {
        const raw = result.getText();
        const numMatch = raw.match(/(\d+)/);
        const qrNum = numMatch ? numMatch[1] : raw;
        onQRScanSuccess(qrNum, raw);
      }
    });
    setScanStatus('Scanning… hold QR code steady in the frame');
  } catch(e) { setScanStatus('Error: '+(e.message||'Could not access camera'),'error'); }
}

function onQRScanSuccess(qrNum, rawText) {
  if(!scanActive) return;
  scanActive = false;
  setScanStatus('✔ QR scanned: '+rawText,'success');
  applyQRValue(qrNum);
  document.querySelector('.scan-frame').style.borderColor='#10b981';
  setTimeout(()=>{ closeScanModal(); toast('✔ QR scanned: '+rawText+(QR_TO_CC[qrNum]?' → CC auto-filled':''),'#10b981'); }, 900);
}

function openScanModal() {
  scanModal.classList.add('open'); scanActive=true;
  document.querySelector('.scan-frame').style.borderColor='var(--cyan)';
  setScanStatus('Initialising camera…');
  btnScanQR.classList.add('scanning');
  populateCameras().then(()=>startScan(camSelect.value||null));
}

function closeScanModal() {
  scanActive=false; scanModal.classList.remove('open'); btnScanQR.classList.remove('scanning');
  if(codeReader){try{codeReader.reset();}catch(e){}codeReader=null;}
}

btnScanQR.addEventListener('click',openScanModal);
document.getElementById('scanCloseBtn').addEventListener('click',closeScanModal);
document.getElementById('scanCancelBtn').addEventListener('click',closeScanModal);
scanModal.addEventListener('click',e=>{if(e.target===scanModal)closeScanModal();});
document.addEventListener('keydown',e=>{if(e.key==='Escape'&&scanModal.classList.contains('open'))closeScanModal();});
camSelect.addEventListener('change',function(){if(scanActive)startScan(this.value);});

/* ══ EXPORT — ExcelJS ══ */
document.getElementById('btnExport').addEventListener('click', async () => {
  if(!records.length){toast('⚠ No records to export!','#f59e0b');return;}
  toast('Building styled Excel…','#22d3ee');
  const wb = new ExcelJS.Workbook();
  wb.creator='PP2 Maintenance System'; wb.created=new Date();
  const CLR={indexFill:'FFC000',indexFont:'002060',grateBg:'A9CD90',centreBg:'92D050',sideBg:'8E98A5',ssbBg:'8EAADB',wheelBg:'BF9000',wbBg:'FFD965',prBg:'1FD3E1',sagBg:'0070C0'};
  const thinBorder={top:{style:'thin'},left:{style:'thin'},bottom:{style:'thin'},right:{style:'thin'}};
  const topBorder={top:{style:'thin'},left:{style:'thin'},right:{style:'thin'}};
  function hdrStyle(bgHex){return{fill:{type:'pattern',pattern:'solid',fgColor:{argb:'FF'+bgHex}},font:{bold:true,name:'Calibri',size:11},alignment:{horizontal:'center',vertical:'middle',wrapText:true},border:thinBorder};}
  function subHdrStyle(){return{font:{bold:false,name:'Calibri',size:10},alignment:{horizontal:'center',vertical:'middle',wrapText:true},border:topBorder};}
  const SHEET_COLS=[['SI No',6],['Grate bar replacement Date',13],['Grate Bar make',11],['CC replacement date',16],['CC make',9],['East Side casting replacement date',17],['East SC make',13],['West side casting replacement date',17],['West SC make',13],['East SSB replacement Date',13],['Make of East SSB',13],['West SSB relacement Date',18],['Make of West SSB',12],['EN',13],['ES',13],['WN',13],['WS',13],['EN',13],['EN Make',13],['ES',13],['ES Make',13],['WN',13],['WN Make',13],['WS',13],['WS Make',13],['EN',13],['EN Make',13],['ES',13],['ES Make',13],['WN',13],['WN Make',13],['WS',13],['WS Make',13],['N sag measured date',11],['N sag value',8],['S sag measured date',11],['S sag value',8],['Car reversed',13]];
  const SECTIONS=[{label:'Grate Bars',start:2,end:3,color:CLR.grateBg},{label:'Centre Casting',start:4,end:5,color:CLR.centreBg},{label:'Side Casting',start:6,end:9,color:CLR.sideBg},{label:'SSB',start:10,end:13,color:CLR.ssbBg},{label:'Wheel Replacement Date',start:14,end:17,color:CLR.wheelBg},{label:'Wheel bearing replacement Date',start:18,end:25,color:CLR.wbBg},{label:'Pressure roller Replacement Date',start:26,end:33,color:CLR.prBg},{label:'SAG',start:34,end:37,color:CLR.sagBg},{label:'Car Reversed',start:38,end:38,color:'D9D9D9'}];
  const byCC={};
  records.forEach(r=>{const k=r['CC No']||'Unknown';if(!byCC[k])byCC[k]=[];byCC[k].push(r);});
  for(const [ccName,rows] of Object.entries(byCC)){
    const safe=ccName.replace(/[:\\/\?\*\[\]]/g,'').substring(0,31)||'Sheet';
    const ws=wb.addWorksheet(safe,{views:[{showGridLines:true}]});
    ws.columns=SHEET_COLS.map(([,w])=>({width:w}));
    ws.getRow(1).height=15;
    const idxCell=ws.getRow(1).getCell(2);
    idxCell.value='Index';idxCell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF'+CLR.indexFill}};
    idxCell.font={bold:true,color:{argb:'FF'+CLR.indexFont},name:'Calibri',size:11};
    idxCell.alignment={horizontal:'center',vertical:'middle'};ws.mergeCells(1,2,1,38);
    ws.getRow(2).height=18;
    const infoCell=ws.getRow(2).getCell(2);
    const qrVal=rows[0]?.['QR No']||'';
    infoCell.value=ccName+(qrVal?',  '+qrVal:'');
    infoCell.font={name:'Calibri',size:11};infoCell.alignment={horizontal:'left',vertical:'middle'};ws.mergeCells(2,2,2,38);
    ws.getRow(3).height=18;ws.getRow(3).getCell(1).border=thinBorder;
    SECTIONS.forEach(sec=>{
      const cell=ws.getRow(3).getCell(sec.start);
      cell.value=sec.label;Object.assign(cell,hdrStyle(sec.color));
      if(sec.start!==sec.end)ws.mergeCells(3,sec.start,3,sec.end);
    });
    ws.getRow(4).height=43;
    const siCell=ws.getRow(4).getCell(1);siCell.value='SI No';siCell.font={bold:true,name:'Calibri',size:10};siCell.alignment={horizontal:'center',vertical:'middle',wrapText:true};siCell.border=topBorder;
    SHEET_COLS.slice(1).forEach(([label],idx)=>{const cell=ws.getRow(4).getCell(idx+2);cell.value=label;Object.assign(cell,subHdrStyle());});
    const dataKeys=['SI No','Grate bar replacement Date','Grate Bar make','CC replacement date','CC make','East Side casting replacement date','East SC make','West side casting replacement date','West SC make','East SSB replacement Date','Make of East SSB','West SSB relacement Date','Make of West SSB','Wheel EN','Wheel ES','Wheel WN','Wheel WS','WB EN Date','WB EN Make','WB ES Date','WB ES Make','WB WN Date','WB WN Make','WB WS Date','WB WS Make','PR EN Date','PR EN Make','PR ES Date','PR ES Make','PR WN Date','PR WN Make','PR WS Date','PR WS Make','N sag measured date','N sag value','S sag measured date','S sag value','Car reversed'];
    rows.forEach((row,ri)=>{const exRow=ws.getRow(5+ri);exRow.height=15;dataKeys.forEach((key,ci)=>{const cell=exRow.getCell(ci+1);cell.value=row[key]||'';cell.font={name:'Calibri',size:10};cell.alignment={horizontal:'center',vertical:'middle'};cell.border={left:{style:'thin'},right:{style:'thin'},bottom:{style:'thin'}};});});
  }
  const buf=await wb.xlsx.writeBuffer();
  const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url=URL.createObjectURL(blob);const a=document.createElement('a');
  const now=new Date();const stamp=now.getFullYear()+String(now.getMonth()+1).padStart(2,'0')+String(now.getDate()).padStart(2,'0');
  a.href=url;a.download='PP2_MaintenanceLog_'+stamp+'.xlsx';document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  toast('✔ '+records.length+' rows exported','#22d3ee');
});

/* ══ Toast ══ */
function toast(msg,color){
  color=color||'#34d399';
  const el=document.getElementById('toast');
  el.textContent=msg;el.style.color=color;el.style.borderColor=color+'55';
  el.style.boxShadow='0 8px 40px rgba(0,0,0,.6),0 0 0 1px '+color+'44';
  el.style.display='block';clearTimeout(el._t);
  el._t=setTimeout(()=>el.style.display='none',3400);
}

/* ══ Enter key nav ══ */
document.addEventListener('keydown',e=>{
  if(e.key!=='Enter')return;
  const all=[document.getElementById('qrInput'),document.getElementById('ccSelect'),...document.querySelectorAll('input[data-sec]')];
  const idx=all.indexOf(document.activeElement);
  if(idx>=0&&idx<all.length-1){all[idx+1].focus();e.preventDefault();}
});

function checkLayout(){document.querySelectorAll('.two-col').forEach(tc=>{tc.style.gridTemplateColumns=window.innerWidth<680?'1fr':'1fr 1fr';});}
window.addEventListener('resize',checkLayout);checkLayout();
updateCounters();
</script>
</body>
</html>
"""


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/log_maintenance', methods=['POST'])
def log_maintenance():
    content = request.json
    if not isinstance(content, dict):
        return jsonify({"status": "error", "message": "Invalid request"}), 400

    qr        = content.get('qr', '').strip()
    tech      = content.get('tech', '').strip()
    dept      = content.get('dept', '').strip()
    issue     = content.get('issue', '').strip()
    priority  = content.get('priority', '').strip()
    status    = content.get('status', '').strip()
    datetime_ = content.get('datetime', '').strip()
    remarks   = content.get('remarks', '').strip()

    if not qr or not tech or not issue or not status:
        return jsonify({"status": "error", "message": "Missing required fields"}), 400

    try:
        received_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        already_exists = False
        if os.path.exists(CSV_FILE):
            with open(CSV_FILE, mode='r', newline='', encoding='utf-8') as f:
                reader = csv.reader(f)
                next(reader, None)
                for row in reader:
                    if row and row[0] == qr:
                        already_exists = True
                        break

        if already_exists:
            print(f"Warning: '{qr}' already logged.")
            return jsonify({"status": "error", "message": "Already exists"}), 200

        file_is_empty = not os.path.exists(CSV_FILE) or os.stat(CSV_FILE).st_size == 0
        with open(CSV_FILE, mode='a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            if file_is_empty:
                writer.writerow([
                    "QR_DATA", "TECHNICIAN", "DEPARTMENT",
                    "ISSUE_TYPE", "PRIORITY", "STATUS",
                    "ENTRY_DATETIME", "REMARKS", "SAVED_AT"
                ])
            writer.writerow([qr, tech, dept, issue, priority, status, datetime_, remarks, received_at])

        print(f"Maintenance logged: '{qr}' by {tech} at {received_at}")
        return jsonify({"status": "success", "saved_at": received_at}), 200

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"status": "error", "message": str(e)}), 400


if __name__ == '__main__':
    # Access from your phone via https://[Your-PC-IP]:5000
    app.run(host='0.0.0.0', port=5000)