#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
app.py — Bermudan Swaption Pricer Web UI v2
Launch:  python app.py → Opens http://localhost:5000
"""
import os, sys, json, webbrowser, threading, tempfile, io
from datetime import datetime
from contextlib import redirect_stdout

src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src")
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

from flask import Flask, request, jsonify, send_file
import yaml, numpy as np, openpyxl

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Bermudan Swaption Pricer</title>
<link href="https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600&family=DM+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{--bg:#0a0e17;--bg2:#111827;--bg3:#1a2235;--card:#151d2e;--border:#2a3650;--border-hi:#3b82f6;--text:#e2e8f0;--text2:#94a3b8;--text3:#64748b;--accent:#3b82f6;--accent2:#60a5fa;--green:#22c55e;--green-bg:rgba(34,197,94,0.1);--red:#ef4444;--red-bg:rgba(239,68,68,0.1);--amber:#f59e0b;--amber-bg:rgba(245,158,11,0.1)}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--text);min-height:100vh}
.header{background:linear-gradient(135deg,var(--bg2),var(--bg3));border-bottom:1px solid var(--border);padding:20px 32px;display:flex;align-items:center;justify-content:space-between}
.header h1{font-size:20px;font-weight:700;letter-spacing:-0.5px}.header h1 span{color:var(--accent2)}
.header .subtitle{font-size:12px;color:var(--text3);font-family:'JetBrains Mono',monospace;margin-top:2px}
.status-badge{font-size:11px;font-family:'JetBrains Mono',monospace;padding:4px 12px;border-radius:20px;background:var(--green-bg);color:var(--green);border:1px solid rgba(34,197,94,0.2)}
.container{max-width:1440px;margin:0 auto;padding:24px;display:grid;grid-template-columns:420px 1fr;gap:24px}
.panel{background:var(--card);border:1px solid var(--border);border-radius:12px;overflow:hidden}
.panel-header{padding:14px 20px;border-bottom:1px solid var(--border);font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:var(--text2);display:flex;align-items:center;gap:8px}
.panel-header .dot{width:6px;height:6px;border-radius:50%;background:var(--accent)}
.panel-body{padding:16px 20px}
.section-label{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:1.5px;color:var(--text3);margin:16px 0 8px}.section-label:first-child{margin-top:0}
.field{margin-bottom:8px;display:grid;grid-template-columns:130px 1fr;align-items:center;gap:8px}
.field label{font-size:12px;color:var(--text2);font-weight:500}
.field input,.field select{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:7px 10px;color:var(--text);font-family:'JetBrains Mono',monospace;font-size:12px;outline:none}
.field input:focus,.field select:focus{border-color:var(--accent)}.field select{cursor:pointer}
.field-check{margin-bottom:8px;display:flex;align-items:center;gap:8px}
.field-check label{font-size:12px;color:var(--text2);font-weight:500;cursor:pointer}
.field-check input[type=checkbox]{width:16px;height:16px;cursor:pointer;accent-color:var(--accent)}
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:18px;text-align:center;cursor:pointer;transition:all 0.3s;margin:10px 0}
.upload-zone:hover{border-color:var(--accent);background:rgba(59,130,246,0.03)}
.upload-zone.loaded{border-color:var(--green);border-style:solid;background:var(--green-bg)}
.upload-zone .icon{font-size:26px;margin-bottom:4px}.upload-zone .label{font-size:13px;color:var(--text2)}
.upload-zone .sublabel{font-size:11px;color:var(--text3);margin-top:3px}
.upload-zone.loaded .label{color:var(--green)}.upload-zone input[type=file]{display:none}
.file-info{font-family:'JetBrains Mono',monospace;font-size:11px;color:var(--green);padding:8px 12px;background:var(--green-bg);border-radius:6px;margin-top:6px;display:none}
.file-info.show{display:block}
.btn-price{width:100%;margin-top:14px;padding:14px;background:linear-gradient(135deg,var(--accent),#2563eb);color:white;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;transition:all 0.3s;font-family:'DM Sans',sans-serif}
.btn-price:hover{transform:translateY(-1px);box-shadow:0 8px 24px rgba(59,130,246,0.3)}
.btn-price:disabled{opacity:0.5;cursor:not-allowed;transform:none;box-shadow:none}
.btn-price.running{background:linear-gradient(135deg,#475569,#334155);animation:pulse 1.5s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:0.7}}
.data-section{margin-top:10px}.data-toggle{font-size:12px;color:var(--accent2);cursor:pointer;padding:5px 0}
.data-toggle:hover{color:var(--accent)}.data-area{display:none;margin-top:4px}.data-area.open{display:block}
.data-area textarea{width:100%;height:160px;background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:10px;color:var(--text);font-family:'JetBrains Mono',monospace;font-size:10px;line-height:1.5;resize:vertical;outline:none}
.data-area textarea:focus{border-color:var(--accent)}.data-area label{display:block;font-size:11px;color:var(--text3);margin-bottom:3px}
.results-area{display:flex;flex-direction:column;gap:14px}
.result-cards{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
.rcard{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:14px;text-align:center;transition:border-color 0.3s}
.rcard:hover{border-color:var(--border-hi)}.rcard .label{font-size:10px;text-transform:uppercase;letter-spacing:1.5px;color:var(--text3);margin-bottom:5px}
.rcard .value{font-family:'JetBrains Mono',monospace;font-size:20px;font-weight:600}
.rcard .value.match{color:var(--green)}.rcard .sub{font-size:11px;color:var(--text3);margin-top:2px;font-family:'JetBrains Mono',monospace}
.cmp-table{width:100%;border-collapse:collapse;font-size:13px}
.cmp-table th{text-align:left;padding:9px 14px;font-size:11px;text-transform:uppercase;letter-spacing:1px;color:var(--text3);border-bottom:1px solid var(--border);font-weight:600}
.cmp-table td{padding:10px 14px;border-bottom:1px solid rgba(42,54,80,0.5);font-family:'JetBrains Mono',monospace}
.cmp-table tr:hover{background:rgba(59,130,246,0.03)}
.cmp-table .name{color:var(--text2);font-family:'DM Sans',sans-serif;font-weight:500}
.cmp-table .val{color:var(--text);text-align:right}.cmp-table .bbg{color:var(--text3);text-align:right}.cmp-table .diff{text-align:right}
.diff-good{color:var(--green)}.diff-ok{color:var(--amber)}.diff-bad{color:var(--red)}
.model-bar{display:flex;gap:20px;padding:12px 18px;font-family:'JetBrains Mono',monospace;font-size:12px;color:var(--text2);flex-wrap:wrap}
.model-bar span{color:var(--accent2)}
.log-area{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:14px;font-family:'JetBrains Mono',monospace;font-size:11px;line-height:1.7;color:var(--text3);max-height:220px;overflow-y:auto;white-space:pre-wrap}
.placeholder{display:flex;flex-direction:column;align-items:center;justify-content:center;min-height:400px;color:var(--text3);gap:12px}
.placeholder svg{opacity:0.3}
.btn-export{padding:10px 20px;background:var(--bg3);border:1px solid var(--border);color:var(--text2);border-radius:6px;font-size:12px;cursor:pointer;font-family:'DM Sans',sans-serif;font-weight:600}
.btn-export:hover{border-color:var(--accent);color:var(--text)}
.scroll-left{max-height:calc(100vh - 100px);overflow-y:auto}
@media(max-width:900px){.container{grid-template-columns:1fr}.result-cards{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="header">
<div><h1>Bermudan <span>Swaption</span> Pricer</h1><div class="subtitle">CAD CORRA OIS · Hull-White 1F · v12 hybrid</div></div>
<div class="status-badge" id="statusBadge">READY</div>
</div>
<div class="container">
<!-- LEFT -->
<div class="scroll-left">
<div class="panel">
<div class="panel-header"><div class="dot"></div> Deal Parameters</div>
<div class="panel-body">
<div class="section-label">Deal</div>
<div class="field"><label>Valuation Date</label><input type="date" id="val_date" value="2026-02-11"></div>
<div class="field"><label>Notional</label><input type="number" id="notional" value="10000000" step="1000000"></div>
<div class="field"><label>Strike (%)</label><input type="number" id="strike" value="3.245112" step="0.000001"></div>
<div class="field"><label>Direction</label><select id="direction"><option value="Receiver">Receiver</option><option value="Payer">Payer</option></select></div>
<div class="field"><label>Swap Start</label><input type="date" id="swap_start" value="2027-02-12"></div>
<div class="field"><label>Swap End</label><input type="date" id="swap_end" value="2032-02-12"></div>
<div class="field"><label>Frequency</label><select id="frequency"><option value="SemiAnnual" selected>SemiAnnual</option><option value="Quarterly">Quarterly</option><option value="Annual">Annual</option></select></div>
<div class="field"><label>Day Count</label><select id="day_count"><option value="ACT/365" selected>ACT/365</option><option value="ACT/360">ACT/360</option><option value="30/360">30/360</option></select></div>
<div class="field"><label>Payment Lag</label><input type="number" id="payment_lag" value="2"></div>
<div class="field"><label>Currency</label><input type="text" id="currency" value="CAD"></div>

<div class="section-label">Model</div>
<div class="field"><label>Mean Reversion</label><input type="number" id="mean_rev" value="0.03" step="0.001"></div>
<div class="field-check"><input type="checkbox" id="calib_a"><label for="calib_a">Calibrate a (mean reversion) — if unchecked, a is fixed</label></div>
<div class="field"><label>FDM Grid</label><input type="number" id="fdm_grid" value="300"></div>

<div class="section-label">Calibration Mode</div>
<div class="field-check"><input type="checkbox" id="standalone_mode" onchange="toggleBBG()"><label for="standalone_mode">Standalone (no BBG) — pure ATM calibration only</label></div>

<div id="bbgSection">
<div class="section-label">BBG Valuation Results</div>
<div class="field"><label>NPV</label><input type="number" id="bbg_npv" value="255683.06" step="0.01"></div>
<div class="field"><label>ATM Strike (%)</label><input type="number" id="bbg_atm" value="2.922733" step="0.000001"></div>
<div class="field"><label>Yield Value (bp)</label><input type="number" id="bbg_yv" value="56.389" step="0.001"></div>
<div class="field"><label>Und. Premium (%)</label><input type="number" id="bbg_uprem" value="1.46175" step="0.00001"></div>
<div class="field"><label>Premium (%)</label><input type="number" id="bbg_prem" value="2.55683" step="0.00001"></div>

<div class="section-label">BBG Greeks</div>
<div class="field"><label>DV01</label><input type="number" id="bbg_dv01" value="2832.42" step="0.01"></div>
<div class="field"><label>Gamma (1bp)</label><input type="number" id="bbg_gamma" value="22.06" step="0.01"></div>
<div class="field"><label>Vega (1bp)</label><input type="number" id="bbg_vega" value="2542.10" step="0.01"></div>
<div class="field"><label>Theta (1 day)</label><input type="number" id="bbg_theta" value="-109.14" step="0.01"></div>
</div><!-- end bbgSection -->
</div>
</div>

<!-- MARKET DATA -->
<div class="panel" style="margin-top:14px">
<div class="panel-header"><div class="dot" style="background:var(--amber)"></div> Market Data</div>
<div class="panel-body">
<div class="upload-zone" id="uploadZone" onclick="document.getElementById('fileInput').click()">
<div class="icon">📁</div><div class="label">Click to load market data (.xlsx)</div>
<div class="sublabel">Excel with sheets: Curve_CAD_OIS + BVOL_CAD_RFR_Normal</div>
<input type="file" id="fileInput" accept=".xlsx,.xls" onchange="uploadFile(this)">
</div>
<div class="file-info" id="fileInfo"></div>
<div class="data-section"><div class="data-toggle" onclick="toggleData('curve')">▸ Manual: Curve Data</div>
<div class="data-area" id="curveData"><label>date,discount_factor (one per line)</label>
<textarea id="curveText"></textarea></div></div>
<div class="data-section"><div class="data-toggle" onclick="toggleData('vol')">▸ Manual: Vol Surface</div>
<div class="data-area" id="volData"><label>BPx10 matrix (comma or tab separated)</label>
<textarea id="volText"></textarea></div></div>
<button class="btn-price" id="btnPrice" onclick="runPricer()">▶ PRICE</button>
</div>
</div>
</div>

<!-- RIGHT: RESULTS -->
<div class="results-area" id="resultsArea">
<div class="placeholder"><svg width="64" height="64" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M3 3v18h18"/><path d="M7 16l4-8 4 4 4-6"/></svg>
<p>Load market data, set deal parameters, click <strong>PRICE</strong></p></div>
</div>
</div>

<script>
const EXPIRY_LABELS=["1Mo","3Mo","6Mo","9Mo","1Yr","2Yr","3Yr","4Yr","5Yr","6Yr","7Yr","8Yr","9Yr","10Yr","12Yr","15Yr","20Yr","25Yr","30Yr"];
const TENOR_LABELS=["1Y","2Y","3Y","4Y","5Y","6Y","7Y","8Y","9Y","10Y","12Y","15Y","20Y","25Y","30Y"];
let loadedExpLabels=null;

function toggleBBG(){
    const sa=document.getElementById('standalone_mode').checked;
    document.getElementById('bbgSection').style.display=sa?'none':'block';
}

function toggleData(id){const el=document.getElementById(id+'Data');el.classList.toggle('open');const t=el.previousElementSibling;t.textContent=(el.classList.contains('open')?'▾':'▸')+t.textContent.slice(1)}
function fmt(n,dec=2){if(n===null||n===undefined)return'N/A';return parseFloat(n).toLocaleString('en-US',{minimumFractionDigits:dec,maximumFractionDigits:dec})}
function diffClass(pct){const a=Math.abs(pct);if(a<3)return'diff-good';if(a<10)return'diff-ok';return'diff-bad'}
function diffBpClass(diffBp, refBp){const pct=refBp?Math.abs(diffBp/refBp*100):0;if(pct<3)return'diff-good';if(pct<10)return'diff-ok';return'diff-bad'}

function uploadFile(input){
    const file=input.files[0];if(!file)return;
    const formData=new FormData();formData.append('file',file);
    const info=document.getElementById('fileInfo'),zone=document.getElementById('uploadZone');
    info.className='file-info show';info.textContent='⟳ Reading '+file.name+'...';info.style.color='var(--amber)';info.style.background='var(--amber-bg)';
    fetch('/api/upload_excel',{method:'POST',body:formData}).then(r=>r.json()).then(data=>{
        if(data.error){info.textContent='✗ '+data.error;info.style.color='var(--red)';info.style.background='var(--red-bg)';return}
        loadedExpLabels=data.expiry_labels;
        document.getElementById('curveText').value=data.curve.map(r=>r[0]+','+r[1]).join('\n');
        document.getElementById('volText').value=data.vol_values.map(r=>r.join(',')).join('\n');
        zone.classList.add('loaded');zone.querySelector('.icon').textContent='✓';
        zone.querySelector('.label').textContent=file.name;zone.querySelector('.sublabel').textContent='Click to load a different file';
        info.textContent='✓ Loaded '+data.curve.length+' curve nodes + '+data.vol_values.length+'×'+data.vol_values[0].length+' vol surface';
        info.style.color='var(--green)';info.style.background='var(--green-bg)';
    }).catch(err=>{info.textContent='✗ '+err;info.style.color='var(--red)';info.style.background='var(--red-bg)';});
}

function runPricer(){
    const btn=document.getElementById('btnPrice'),badge=document.getElementById('statusBadge');
    btn.disabled=true;btn.classList.add('running');btn.textContent='⟳ PRICING...';
    badge.textContent='RUNNING';badge.style.background='rgba(245,158,11,0.1)';badge.style.color='#f59e0b';

    const volLines=document.getElementById('volText').value.trim().split('\n');
    const volValues=volLines.filter(l=>l.trim()).map(l=>l.split(/[,\t]+/).map(Number));
    const curveLines=document.getElementById('curveText').value.trim().split('\n');
    const curveData=curveLines.filter(l=>l.trim()).map(l=>{const p=l.split(/[,\t]+/);return[p[0].trim(),parseFloat(p[1])]});
    let expLabels=loadedExpLabels||EXPIRY_LABELS.slice(0,volValues.length);
    let tnrLabels=TENOR_LABELS.slice(0,volValues[0]?volValues[0].length:15);

    const standalone=document.getElementById('standalone_mode').checked;

    const bbg=standalone?{npv:0,atm_strike:0,yield_value_bp:0,underlying_premium:0,premium:0,dv01:0,gamma_1bp:0,vega_1bp:0,theta_1d:0}:{
        npv:parseFloat(document.getElementById('bbg_npv').value)||0,
        atm_strike:parseFloat(document.getElementById('bbg_atm').value)||0,
        yield_value_bp:parseFloat(document.getElementById('bbg_yv').value)||0,
        underlying_premium:parseFloat(document.getElementById('bbg_uprem').value)||0,
        premium:parseFloat(document.getElementById('bbg_prem').value)||0,
        dv01:parseFloat(document.getElementById('bbg_dv01').value)||0,
        gamma_1bp:parseFloat(document.getElementById('bbg_gamma').value)||0,
        vega_1bp:parseFloat(document.getElementById('bbg_vega').value)||0,
        theta_1d:parseFloat(document.getElementById('bbg_theta').value)||0,
    };

    const payload={
        deal:{valuation_date:document.getElementById('val_date').value,notional:parseFloat(document.getElementById('notional').value),
            strike:parseFloat(document.getElementById('strike').value),direction:document.getElementById('direction').value,
            swap_start:document.getElementById('swap_start').value,swap_end:document.getElementById('swap_end').value,
            fixed_frequency:document.getElementById('frequency').value,day_count:document.getElementById('day_count').value,
            payment_lag:parseInt(document.getElementById('payment_lag').value),currency:document.getElementById('currency').value},
        model:{mean_reversion:parseFloat(document.getElementById('mean_rev').value),
            calibrate_a:document.getElementById('calib_a').checked,
            fdm_time_grid:parseInt(document.getElementById('fdm_grid').value),fdm_space_grid:parseInt(document.getElementById('fdm_grid').value)},
        benchmark:bbg,curve_data:curveData,
        vol_surface_data:{expiry_labels:expLabels,tenor_labels:tnrLabels,values:volValues},
        exercise:{mode:"auto"},data_source:{mode:"manual"},
        greeks:{dv01_bump_bp:1,gamma_bump_bp:1,vega_bump_bp:1,compute_theta:true,theta_annualization:"none"},
        output:{print_console:false,export_excel:false}
    };

    fetch('/api/price',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)})
    .then(r=>r.json()).then(data=>{
        btn.disabled=false;btn.classList.remove('running');btn.textContent='▶ PRICE';
        if(data.error){badge.textContent='ERROR';badge.style.background='rgba(239,68,68,0.1)';badge.style.color='#ef4444';
            document.getElementById('resultsArea').innerHTML='<div class="panel"><div class="panel-header"><div class="dot" style="background:var(--red)"></div> Error</div><div class="panel-body"><div class="log-area" style="color:var(--red)">'+data.error+'</div></div></div>';return}
        badge.textContent='PRICED';badge.style.background='rgba(34,197,94,0.1)';badge.style.color='#22c55e';
        renderResults(data,bbg);
    }).catch(err=>{btn.disabled=false;btn.classList.remove('running');btn.textContent='▶ PRICE';
        badge.textContent='ERROR';badge.style.background='rgba(239,68,68,0.1)';badge.style.color='#ef4444';
        document.getElementById('resultsArea').innerHTML='<div class="panel"><div class="panel-body"><div class="log-area" style="color:var(--red)">'+err+'</div></div></div>';
    });
}

function renderResults(d,bbg){
    const g=d.greeks, moneyBp=d.moneyness_bp;
    const sa=document.getElementById('standalone_mode').checked;
    const hasBBG=bbg.npv>0;

    // Top cards — always shown
    let npvSub=sa?'standalone calibration':'';
    if(hasBBG){const npvPct=((d.npv-bbg.npv)/bbg.npv*100);npvSub=`${npvPct>=0?'+':''}${fmt(npvPct,4)}% vs BBG`;}

    let html=`<div class="result-cards">
        <div class="rcard"><div class="label">NPV</div><div class="value match">${fmt(d.npv)}</div><div class="sub">${npvSub}</div></div>
        <div class="rcard"><div class="label">σ total</div><div class="value">${fmt(d.sigma_total*10000,2)}</div><div class="sub">bp short-rate vol</div></div>
        <div class="rcard"><div class="label">Yield Value</div><div class="value">${fmt(d.yield_value,3)}</div><div class="sub">bps</div></div>
        <div class="rcard"><div class="label">ATM Rate</div><div class="value">${fmt(d.fair_rate*100,4)}%</div><div class="sub">Moneyness: ${moneyBp>=0?'+':''}${fmt(moneyBp,1)} bp</div></div>
        <div class="rcard"><div class="label">Premium</div><div class="value">${fmt(d.premium_pct,4)}%</div><div class="sub">of notional</div></div>
        <div class="rcard"><div class="label">Und. NPV</div><div class="value">${fmt(d.underlying_npv)}</div><div class="sub">${fmt(d.underlying_prem_pct,4)}%</div></div>
    </div>`;

    // Model decomposition
    let modelInfo=`a = <span>${d.a_used}</span> ${d.a_calibrated?'(calibrated)':'(fixed)'} &nbsp;|&nbsp; σ_ATM = <span>${fmt(d.sigma_atm*10000,2)} bp</span>`;
    if(hasBBG) modelInfo+=` + Δσ = <span>${fmt(d.delta_spread*10000,2)} bp</span> → σ_total = <span>${fmt(d.sigma_total*10000,2)} bp</span>`;
    else modelInfo+=` &nbsp;(standalone — no Δσ spread)`;
    html+=`<div class="panel"><div class="panel-header"><div class="dot"></div> Model</div><div class="model-bar">${modelInfo}</div></div>`;

    // BBG comparison — only if not standalone
    if(hasBBG){
        const npvPct=((d.npv-bbg.npv)/bbg.npv*100);
        const atmDiffBp=(d.fair_rate-bbg.atm_strike/100)*10000;
        const yvDiff=d.yield_value-bbg.yield_value_bp;
        const upDiff=d.underlying_prem_pct-bbg.underlying_premium;
        const prDiff=d.premium_pct-bbg.premium;
        const valRows=[
            ['NPV (CAD)',fmt(bbg.npv),fmt(d.npv),fmt(npvPct,4)+'%',diffClass(npvPct)],
            ['ATM Strike (%)',fmt(bbg.atm_strike,6),fmt(d.fair_rate*100,6),fmt(atmDiffBp,2)+' bp',diffBpClass(atmDiffBp,bbg.atm_strike*100)],
            ['Yield Value (bp)',fmt(bbg.yield_value_bp,3),fmt(d.yield_value,3),fmt(yvDiff,3)+' bp',diffBpClass(yvDiff,bbg.yield_value_bp)],
            ['Und. Premium (%)',fmt(bbg.underlying_premium,5),fmt(d.underlying_prem_pct,5),fmt(upDiff,5)+'%',diffBpClass(upDiff*100,bbg.underlying_premium)],
            ['Premium (%)',fmt(bbg.premium,5),fmt(d.premium_pct,5),fmt(prDiff,5)+'%',diffClass(prDiff/bbg.premium*100)],
        ].map(r=>`<tr><td class="name">${r[0]}</td><td class="bbg">${r[1]}</td><td class="val">${r[2]}</td><td class="diff ${r[4]}">${r[3]}</td></tr>`).join('');
        html+=`<div class="panel"><div class="panel-header"><div class="dot"></div> Valuation — BBG Comparison</div>
        <div class="panel-body" style="padding:0"><table class="cmp-table"><thead><tr><th>Metric</th><th style="text-align:right">Bloomberg</th><th style="text-align:right">QuantLib</th><th style="text-align:right">Diff</th></tr></thead><tbody>${valRows}</tbody></table></div></div>`;
    }

    // Greeks — always shown, BBG columns only if available
    const greeks=[
        {name:'DV01',ql:g.dv01,bbg:hasBBG?bbg.dv01:null},{name:'Gamma (1bp)',ql:g.gamma_1bp,bbg:hasBBG?bbg.gamma_1bp:null},
        {name:'Vega (1bp)',ql:g.vega_1bp,bbg:hasBBG?bbg.vega_1bp:null},{name:'Theta (1d)',ql:g.theta_1d,bbg:hasBBG?bbg.theta_1d:null},
        {name:'Delta',ql:g.delta_hedge,bbg:null},{name:'Und. DV01',ql:g.underlying_dv01,bbg:null}
    ];
    if(hasBBG){
        const greekRows=greeks.map(gr=>{
            const diff=gr.bbg!=null?gr.ql-gr.bbg:null;const pct=(gr.bbg&&gr.bbg!==0)?(diff/Math.abs(gr.bbg)*100):null;
            const dc=pct!==null?diffClass(pct):'';
            return`<tr><td class="name">${gr.name}</td><td class="bbg">${gr.bbg!=null?fmt(gr.bbg):'—'}</td><td class="val">${fmt(gr.ql)}</td><td class="diff ${dc}">${diff!=null?(diff>=0?'+':'')+fmt(diff):'—'}</td><td class="diff ${dc}">${pct!=null?(pct>=0?'+':'')+fmt(pct,1)+'%':'—'}</td></tr>`;
        }).join('');
        html+=`<div class="panel"><div class="panel-header"><div class="dot"></div> Greeks — BBG Comparison</div>
        <div class="panel-body" style="padding:0"><table class="cmp-table"><thead><tr><th>Greek</th><th style="text-align:right">Bloomberg</th><th style="text-align:right">QuantLib</th><th style="text-align:right">Diff</th><th style="text-align:right">%</th></tr></thead><tbody>${greekRows}</tbody></table></div></div>`;
    } else {
        const greekRows=greeks.map(gr=>`<tr><td class="name">${gr.name}</td><td class="val" style="text-align:right">${fmt(gr.ql)}</td></tr>`).join('');
        html+=`<div class="panel"><div class="panel-header"><div class="dot"></div> Greeks</div>
        <div class="panel-body" style="padding:0"><table class="cmp-table"><thead><tr><th>Greek</th><th style="text-align:right">Value</th></tr></thead><tbody>${greekRows}</tbody></table></div></div>`;
    }

    html+=`<div class="panel"><div class="panel-header"><div class="dot"></div> Execution Log</div><div class="panel-body"><div class="log-area">${d.log||''}</div></div></div>`;
    html+=`<button class="btn-export" onclick="window.location.href='/api/export'">⬇ Export to Excel</button>`;
    html+=`<button class="btn-export" style="margin-left:10px;border-color:var(--amber)" onclick="window.location.href='/api/export_pbi'">📊 Export for Power BI</button>`;

    document.getElementById('resultsArea').innerHTML=html;
}
</script>
</body></html>"""

# ═══════════════════════════════════════════════════════════════════════════
@app.route("/")
def index(): return HTML

@app.route("/api/upload_excel", methods=["POST"])
def api_upload_excel():
    try:
        f=request.files.get("file")
        if not f: return jsonify({"error":"No file uploaded"})
        tmp=os.path.join(tempfile.gettempdir(),"mkt_data.xlsx"); f.save(tmp)
        wb=openpyxl.load_workbook(tmp,data_only=True); result={}
        # Curve sheet
        curve_sheet=None
        for name in wb.sheetnames:
            if "curve" in name.lower() or "ois" in name.lower(): curve_sheet=name; break
        if not curve_sheet: curve_sheet=wb.sheetnames[0]
        ws=wb[curve_sheet]; curve_data=[]
        # Auto-detect columns: find "date" col and "discount" col from header
        header = [str(c.value or "").strip().lower() for c in ws[1]]
        date_col = 0  # default: first column
        df_col = 1    # default: second column
        for i, h in enumerate(header):
            if "date" in h:
                date_col = i
            if "discount" in h or h == "df":
                df_col = i
        # If no "discount" found but values in col B are > 1, try col D
        all_rows = list(ws.iter_rows(min_row=2, values_only=True))
        if all_rows and df_col == 1:
            try:
                test_val = float(all_rows[0][1])
                if test_val > 1.0:  # looks like a rate, not a DF
                    # Try to find a column with values < 1
                    for ci in range(len(all_rows[0])):
                        try:
                            tv = float(all_rows[0][ci])
                            if 0 < tv < 1.0:
                                df_col = ci
                                break
                        except: pass
            except: pass

        for row in all_rows:
            if row[date_col] is None: continue
            d = row[date_col]
            d = d.strftime("%Y-%m-%d") if isinstance(d, datetime) else str(d).strip().split()[0]
            try: curve_data.append([d, float(row[df_col])])
            except: continue
        result["curve"] = curve_data
        # Vol sheet
        vol_sheet=None
        for name in wb.sheetnames:
            if "vol" in name.lower() or "bvol" in name.lower(): vol_sheet=name; break
        if not vol_sheet and len(wb.sheetnames)>1: vol_sheet=wb.sheetnames[1]
        if vol_sheet:
            ws=wb[vol_sheet]; rows=list(ws.iter_rows(values_only=True))
            tenor_labels=[str(c).strip() for c in rows[0][1:] if c is not None]
            expiry_labels=[]; vol_values=[]
            for row in rows[1:]:
                if row[0] is None: continue
                expiry_labels.append(str(row[0]).strip())
                vol_values.append([float(c) if c else 0.0 for c in row[1:1+len(tenor_labels)]])
            result["vol_values"]=vol_values; result["expiry_labels"]=expiry_labels; result["tenor_labels"]=tenor_labels
        wb.close(); return jsonify(result)
    except Exception as e:
        import traceback; return jsonify({"error":f"{e}\n{traceback.format_exc()}"})

@app.route("/api/price", methods=["POST"])
def api_price():
    try:
        cfg=request.json
        vol_values=np.array(cfg.get("vol_surface_data",{}).get("values",[]),dtype=float)
        from bbg_fetcher import labels_to_years,EXPIRY_LABEL_TO_YEARS,TENOR_LABEL_TO_YEARS
        exp_labels=cfg.get("vol_surface_data",{}).get("expiry_labels",[])
        tnr_labels=cfg.get("vol_surface_data",{}).get("tenor_labels",[])
        market_data={"curve":cfg.get("curve_data",[]),"vol_surface":vol_values,
            "expiry_grid":labels_to_years(exp_labels,EXPIRY_LABEL_TO_YEARS),
            "tenor_grid":labels_to_years(tnr_labels,TENOR_LABEL_TO_YEARS),
            "bbg_npv":float(cfg.get("benchmark",{}).get("npv",0))}
        log_buf=io.StringIO()
        with redirect_stdout(log_buf):
            from pricer import BermudanPricer
            pricer=BermudanPricer(cfg,market_data); pricer.setup(); pricer.calibrate(); pricer.compute_greeks()
        bps_leg=abs(float(pricer.swap.fixedLegBPS())); yv=pricer.npv/bps_leg if bps_leg else 0
        app.config["LAST_PRICER"]=pricer
        app.config["LAST_CFG"]=cfg
        return jsonify({"npv":pricer.npv,"sigma_atm":pricer.sigma_atm,"sigma_total":pricer.sigma_total,
            "delta_spread":pricer.delta_spread,"fair_rate":pricer.fair_rate,"underlying_npv":pricer.underlying_npv,
            "yield_value":yv,"premium_pct":pricer.npv/pricer.notional*100,
            "underlying_prem_pct":pricer.underlying_npv/pricer.notional*100,
            "moneyness_bp":(pricer.strike-pricer.fair_rate)*10000,"greeks":pricer.greeks,
            "a_used":pricer.a,"a_calibrated":pricer.calib_a,"log":log_buf.getvalue()})
    except Exception as e:
        import traceback; return jsonify({"error":f"{e}\n\n{traceback.format_exc()}"})

@app.route("/api/export")
def api_export():
    pricer=app.config.get("LAST_PRICER")
    if not pricer: return "No results. Run pricer first.",400
    xlsx=os.path.join(tempfile.gettempdir(),"bermudan_results.xlsx"); pricer.export_excel(xlsx)
    return send_file(xlsx,as_attachment=True,download_name="bermudan_results.xlsx")

@app.route("/api/export_pbi")
def api_export_pbi():
    """Export structured Excel optimized for Power BI."""
    pricer=app.config.get("LAST_PRICER")
    cfg=app.config.get("LAST_CFG")
    if not pricer: return "No results. Run pricer first.",400
    try:
        from run_and_export import export_pbi_excel
        xlsx=os.path.join(tempfile.gettempdir(),"pbi_data.xlsx")
        export_pbi_excel(pricer, cfg, xlsx)
        return send_file(xlsx,as_attachment=True,download_name="pbi_data.xlsx")
    except Exception as e:
        import traceback; return str(e)+"\n"+traceback.format_exc(), 500

def open_browser(): webbrowser.open("http://localhost:5000")
if __name__=="__main__":
    print("="*60);print("  Bermudan Swaption Pricer — Web UI");print("  http://localhost:5000");print("="*60);print("  Press Ctrl+C to stop\n")
    threading.Timer(1.5,open_browser).start(); app.run(host="127.0.0.1",port=5000,debug=False)
