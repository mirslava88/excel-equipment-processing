# Основной модуль - Множественная обработка файлов

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from starlette.responses import FileResponse
import uvicorn
import os
import json
import tempfile
import zipfile
from typing import List, Optional

from .excel_logic import (
    save_temp_file,
    get_engine,
    get_sheet_names,
    get_columns,
    auto_detect_columns,
    process_excels,
)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Хранилище сессии
session_data: dict = {
    "base_file": None,  # База данных {path, engine, filename}
    "process_files": [],  # Массив файлов для обработки
    "results": []  # Массив результатов {filename, path, stats}
}


# ─── HTML ────────────────────────────────────────────────────────────────────

HTML_PAGE = """
<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Множественная обработка Excel</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
           background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
           color: #333; min-height: 100vh; padding: 20px; }
    .container { max-width: 900px; margin: 0 auto; }
    h1 { text-align: center; margin-bottom: 8px; font-size: 2rem; color: #fff; text-shadow: 0 2px 4px rgba(0,0,0,0.2); }
    .subtitle { text-align: center; color: #f0f0f0; margin-bottom: 32px; font-size: 1rem; }
    .card { background: #fff; border-radius: 16px; padding: 32px;
            box-shadow: 0 8px 32px rgba(0,0,0,0.15); margin-bottom: 24px; }
    .step-title { font-size: 1.2rem; font-weight: 700; margin-bottom: 20px;
                  display: flex; align-items: center; gap: 12px; color: #1a1a2e; }
    .step-num { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: #fff; width: 36px; height: 36px; border-radius: 50%;
                display: flex; align-items: center; justify-content: center;
                font-size: 1rem; flex-shrink: 0; box-shadow: 0 4px 8px rgba(102, 126, 234, 0.3); }
    
    .base-file-section { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
                         padding: 20px; border-radius: 12px; margin-bottom: 24px; color: #fff; }
    .base-file-section label { color: #fff; font-weight: 600; margin-bottom: 8px; display: block; }
    
    .process-files-section { border: 2px dashed #e0e0e0; border-radius: 12px; padding: 20px; margin-bottom: 20px; }
    .file-item { background: #f8f9ff; border-radius: 10px; padding: 16px; margin-bottom: 12px;
                 border-left: 4px solid #667eea; position: relative; }
    .file-item-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
    .file-num { font-weight: 700; color: #667eea; font-size: 1.1rem; }
    .btn-remove { background: #ff4757; color: #fff; border: none; padding: 6px 12px;
                  border-radius: 6px; cursor: pointer; font-size: 0.85rem; }
    .btn-remove:hover { background: #ee5a6f; }
    
    label { display: block; font-weight: 600; margin-bottom: 6px; margin-top: 12px;
            font-size: 0.9rem; color: #555; }
    input[type="file"] { width: 100%; padding: 12px; border: 2px solid #e0e0e0;
                         border-radius: 10px; background: #fff; cursor: pointer;
                         font-size: 0.95rem; }
    input[type="file"]:hover { border-color: #667eea; }
    
    select { width: 100%; padding: 10px 12px; border: 2px solid #e0e0e0;
             border-radius: 10px; font-size: 0.95rem; background: #fff; }
    select:disabled { background: #f5f5f5; color: #999; }
    
    .checkbox-group { display: flex; gap: 24px; margin-top: 12px; }
    .checkbox-label { display: flex; align-items: center; gap: 8px; font-size: 0.95rem;
                      cursor: pointer; user-select: none; }
    input[type="checkbox"] { width: 18px; height: 18px; cursor: pointer; }
    
    button { padding: 14px 32px; border: none; border-radius: 10px; font-size: 1rem;
             cursor: pointer; font-weight: 600; transition: all 0.2s; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .btn-primary { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: #fff; }
    .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 6px 16px rgba(102, 126, 234, 0.4); }
    .btn-primary:disabled { background: #ccc; cursor: not-allowed; transform: none; }
    .btn-success { background: linear-gradient(135deg, #84fab0 0%, #8fd3f4 100%); color: #333; }
    .btn-success:hover { transform: translateY(-2px); }
    .btn-add { background: #4cd137; color: #fff; width: 100%; margin-top: 12px; }
    .btn-add:hover { background: #44bd32; }
    
    .actions { margin-top: 24px; display: flex; gap: 16px; justify-content: center; flex-wrap: wrap; }
    .hidden { display: none; }
    .status { padding: 14px 18px; border-radius: 10px; margin-top: 16px; font-size: 0.95rem; }
    .status-info { background: #e8f4fd; color: #1565c0; border-left: 4px solid #1565c0; }
    .status-ok { background: #e8f5e9; color: #2e7d32; border-left: 4px solid #2e7d32; }
    .status-err { background: #fdecea; color: #c62828; border-left: 4px solid #c62828; }
    
    .spinner { display: inline-block; width: 16px; height: 16px; border: 2px solid #ddd;
               border-top: 2px solid #667eea; border-radius: 50%;
               animation: spin 0.8s linear infinite; vertical-align: middle; margin-right: 8px; }
    @keyframes spin { to { transform: rotate(360deg); } }
    
    .result-item { background: #f8f9ff; border-radius: 10px; padding: 16px; margin-bottom: 16px;
                   border-left: 4px solid #2e7d32; }
    .result-stats { display: flex; gap: 20px; margin: 10px 0; font-size: 0.9rem; }
    .stat { display: flex; align-items: center; gap: 6px; }
    .stat-label { color: #666; }
    .stat-value { font-weight: 700; color: #667eea; }
    
    .file-name { font-size: 0.85rem; color: #666; margin-top: 4px; font-style: italic; }
    .auto-hint { font-size: 0.8rem; color: #667eea; margin-top: 4px; }
    .auto-hint.empty { color: #ff6b6b; }
    
    /* Табы навигации */
    .tabs { display: flex; gap: 8px; margin-bottom: 24px; background: rgba(255,255,255,0.2); 
            padding: 8px; border-radius: 12px; }
    .tab { background: transparent; color: #fff; padding: 12px 24px; border-radius: 8px; 
           border: 2px solid transparent; cursor: pointer; transition: all 0.3s; 
           font-size: 1rem; font-weight: 600; }
    .tab:hover { background: rgba(255,255,255,0.1); }
    .tab.active { background: #fff; color: #667eea; border-color: #fff; 
                  box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .tab-content { display: none; }
    .tab-content.active { display: block; }
    
    /* Склад - фильтры */
    .warehouse-filters { display: grid; grid-template-columns: 1fr 1fr auto; gap: 16px; 
                         margin-bottom: 24px; align-items: end; }
    
    /* Склад - таблица */
    .warehouse-table { width: 100%; border-collapse: collapse; margin-top: 20px; }
    .warehouse-table th { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                          color: #fff; padding: 14px; text-align: left; font-weight: 600; 
                          font-size: 0.9rem; }
    .warehouse-table td { padding: 12px; border-bottom: 1px solid #e0e0e0; font-size: 0.9rem; }
    .warehouse-table tr:hover { background: #f8f9ff; }
    .warehouse-table tr:last-child td { border-bottom: none; }
    .warehouse-empty { text-align: center; padding: 40px; color: #666; font-size: 1rem; }
    .warehouse-count { color: #667eea; font-weight: 700; margin-bottom: 12px; font-size: 1.1rem; }
  </style>
</head>
<body>
<div class="container">
  <h1>📊 Множественная обработка Excel</h1>
  <p class="subtitle">База данных + неограниченное количество файлов для обработки</p>

  <!-- Табы навигации -->
  <div class="tabs">
    <button class="tab active" onclick="switchTab('processing')">📄 Обработка файлов</button>
    <button class="tab" onclick="switchTab('warehouse')">📦 Склад</button>
  </div>

  <!-- Вкладка: Обработка файлов -->
  <div id="tabProcessing" class="tab-content active">
  <!-- STEP 1: Загрузка файлов -->
  <div class="card" id="step1">
    <div class="step-title"><span class="step-num">1</span> Загрузка файлов</div>
    
    <!-- База данных -->
    <div class="base-file-section">
      <label>📁 База данных (склад для сверки)</label>
      <input type="file" id="baseFile" accept=".xlsx,.xlsb">
      <div class="file-name" id="baseName"></div>
    </div>
    
    <!-- Файлы для обработки -->
    <div class="process-files-section">
      <label style="color: #667eea; font-size: 1rem; margin-bottom: 12px;">📄 Файлы для обработки</label>
      <div id="processFilesList"></div>
      <button class="btn-add" id="btnAddFile">+ Добавить файл</button>
    </div>
    
    <div class="actions">
      <button class="btn-primary" id="btnUpload" disabled>Загрузить все файлы</button>
    </div>
    <div id="uploadStatus" class="hidden"></div>
  </div>

  <!-- STEP 2: Настройка -->
  <div class="card hidden" id="step2">
    <div class="step-title"><span class="step-num">2</span> Настройка обработки</div>
    
    <!-- База данных -->
    <div style="background: #f0f2f5; padding: 16px; border-radius: 10px; margin-bottom: 20px;">
      <h3 style="font-size: 1rem; margin-bottom: 12px; color: #667eea;">База данных</h3>
      <div style="display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px;">
        <div>
          <label>Лист</label>
          <select id="baseSheet"></select>
        </div>
        <div>
          <label>Столбец с серийными номерами</label>
          <select id="baseSerial" disabled></select>
          <div class="auto-hint" id="hintBaseSerial"></div>
        </div>
        <div>
          <label>Столбец с датой</label>
          <select id="baseDate" disabled></select>
          <div class="auto-hint" id="hintBaseDate"></div>
        </div>
      </div>
    </div>
    
    <!-- Файлы для обработки -->
    <div id="configFilesList"></div>
    
    <div class="actions">
      <button class="btn-success" id="btnProcess" disabled>🚀 Обработать все файлы</button>
    </div>
    <div id="processStatus" class="hidden"></div>
  </div>

  <!-- STEP 3: Результаты -->
  <div class="card hidden" id="step3">
    <div class="step-title"><span class="step-num">3</span> Результаты обработки</div>
    <div id="resultsList"></div>
    <div class="actions">
      <a id="downloadAllLink" href="#"><button class="btn-primary">📦 Скачать все файлы (ZIP)</button></a>
      <button class="btn-success" onclick="location.reload()">🔄 Начать заново</button>
    </div>
  </div>
  </div> <!-- /tabProcessing -->

  <!-- Вкладка: Склад -->
  <div id="tabWarehouse" class="tab-content">
    <div class="card">
      <div class="step-title">📦 Поиск оборудования на складе</div>
      
      <!-- Загрузка базы данных для склада -->
      <div class="base-file-section" style="margin-bottom: 24px;">
        <label>📁 База данных (файл с листом "Возврат")</label>
        <input type="file" id="warehouseFile" accept=".xlsx,.xlsb">
        <div class="file-name" id="warehouseFileName"></div>
        <div style="margin-top: 8px;">
          <button class="btn-primary" id="btnLoadWarehouse" disabled>Загрузить базу</button>
        </div>
      </div>
      
      <div class="warehouse-filters hidden" id="warehouseFiltersSection">
        <div>
          <label>Тип оборудования</label>
          <select id="warehouseType" disabled>
            <option value="">— Выберите тип —</option>
          </select>
        </div>
        <div>
          <label>Модель</label>
          <select id="warehouseModel" disabled>
            <option value="">— Все модели —</option>
          </select>
        </div>
        <div>
          <button class="btn-primary" id="btnSearchWarehouse" disabled>🔍 Найти</button>
        </div>
      </div>
      
      <div id="warehouseStatus" class="hidden"></div>
      <div id="warehouseResults"></div>
    </div>
  </div>

</div> <!-- /container -->

<script>
const $ = id => document.getElementById(id);
const API = '';

// State
let baseFile = null;
let processFiles = [];
let fileCounter = 0;

// --- Step 1: Управление файлами ---
$('baseFile').onchange = e => {
  baseFile = e.target.files[0];
  $('baseName').textContent = baseFile?.name || '';
  checkUploadReady();
};

$('btnAddFile').onclick = () => addProcessFileInput();

function addProcessFileInput() {
  fileCounter++;
  const id = fileCounter;
  const div = document.createElement('div');
  div.className = 'file-item';
  div.id = `fileItem${id}`;
  div.innerHTML = `
    <div class="file-item-header">
      <span class="file-num">Файл #${id}</span>
      <button class="btn-remove" onclick="removeFile(${id})">✕ Удалить</button>
    </div>
    <input type="file" id="processFile${id}" accept=".xlsx,.xlsb">
    <div class="file-name" id="fileName${id}"></div>
  `;
  $('processFilesList').appendChild(div);
  
  $(`processFile${id}`).onchange = e => {
    const file = e.target.files[0];
    $(`fileName${id}`).textContent = file?.name || '';
    processFiles[id] = file;
    checkUploadReady();
  };
}

function removeFile(id) {
  $(`fileItem${id}`).remove();
  delete processFiles[id];
  checkUploadReady();
}

function checkUploadReady() {
  const hasBase = !!baseFile;
  const hasProcess = Object.values(processFiles).some(f => f);
  $('btnUpload').disabled = !(hasBase && hasProcess);
}

// Добавляем первый файл по умолчанию
addProcessFileInput();

$('btnUpload').onclick = async () => {
  $('btnUpload').disabled = true;
  showStatus('uploadStatus', 'info', '<span class="spinner"></span> Загрузка файлов...');
  
  const fd = new FormData();
  fd.append('base_file', baseFile);
  
  Object.entries(processFiles).forEach(([id, file]) => {
    if (file) fd.append('process_files', file);
  });
  
  try {
    const r = await fetch(API + '/upload_multiple', { method: 'POST', body: fd });
    const d = await r.json();
    if (!r.ok) throw new Error(d.detail || 'Ошибка загрузки');
    
    showStatus('uploadStatus', 'ok', `✓ Загружено: база данных + ${d.files_count} файлов`);
    
    // Заполняем настройки базы
    fillSelect('baseSheet', d.base_sheets);
    
    // Создаем конфигурацию для каждого файла
    d.process_files_info.forEach((info, idx) => {
      createFileConfig(idx, info);
    });
    
    $('step2').classList.remove('hidden');
    $('baseSheet').dispatchEvent(new Event('change'));
  } catch(e) {
    showStatus('uploadStatus', 'err', '✗ ' + e.message);
    $('btnUpload').disabled = false;
  }
};

// --- Step 2: Настройка ---
$('baseSheet').onchange = async () => {
  const sheet = $('baseSheet').value;
  if (!sheet) return;
  const r = await fetch(API + `/columns?file_type=base&sheet=${encodeURIComponent(sheet)}`);
  const d = await r.json();
  fillSelect('baseSerial', d.columns);
  fillSelect('baseDate', d.columns);
  $('baseSerial').disabled = false;
  $('baseDate').disabled = false;
  if (d.detected_serial) {
    $('baseSerial').value = d.detected_serial;
    $('hintBaseSerial').textContent = '↑ Автоопределён: ' + d.detected_serial;
  }
  if (d.detected_date) {
    $('baseDate').value = d.detected_date;
    $('hintBaseDate').textContent = '↑ Автоопределён: ' + d.detected_date;
  }
  checkProcessReady();
};

$('baseSerial').onchange = checkProcessReady;
$('baseDate').onchange = checkProcessReady;

function createFileConfig(idx, info) {
  const div = document.createElement('div');
  div.className = 'file-item';
  div.innerHTML = `
    <div class="file-item-header">
      <span class="file-num">📄 ${info.filename}</span>
    </div>
    <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 12px;">
      <div>
        <label>Лист</label>
        <select id="sheet${idx}"></select>
      </div>
      <div>
        <label>Серийный номер</label>
        <select id="serial${idx}" disabled></select>
        <div class="auto-hint" id="hintSerial${idx}"></div>
      </div>
    </div>
    <div>
      <label>Дата отражения проводки (для техрефреша)</label>
      <select id="date${idx}" disabled>
        <option value="">— не выбран (пропустить техрефреш) —</option>
      </select>
      <div class="auto-hint" id="hintDate${idx}"></div>
    </div>
    <div class="checkbox-group">
      <label class="checkbox-label">
        <input type="checkbox" id="opCompare${idx}" checked>
        <span>Сверка с базой данных</span>
      </label>
      <label class="checkbox-label">
        <input type="checkbox" id="opTechRefresh${idx}" checked>
        <span>Анализ устаревшего оборудования</span>
      </label>
    </div>
  `;
  $('configFilesList').appendChild(div);
  
  fillSelect(`sheet${idx}`, info.sheets);
  
  $(`sheet${idx}`).onchange = async () => {
    const sheet = $(`sheet${idx}`).value;
    if (!sheet) return;
    const r = await fetch(API + `/columns?file_type=process&file_idx=${idx}&sheet=${encodeURIComponent(sheet)}`);
    const d = await r.json();
    
    fillSelect(`serial${idx}`, d.columns);
    fillSelect(`date${idx}`, d.columns, true);
    $(`serial${idx}`).disabled = false;
    $(`date${idx}`).disabled = false;
    
    if (d.detected_serial) {
      $(`serial${idx}`).value = d.detected_serial;
      $(`hintSerial${idx}`).textContent = '↑ Автоопределён: ' + d.detected_serial;
    }
    if (d.detected_date) {
      $(`date${idx}`).value = d.detected_date;
      $(`hintDate${idx}`).textContent = '↑ Автоопределён: ' + d.detected_date;
    }
    
    checkProcessReady();
  };
  
  $(`serial${idx}`).onchange = checkProcessReady;
  $(`sheet${idx}`).dispatchEvent(new Event('change'));
}

function checkProcessReady() {
  const baseReady = $('baseSerial').value && $('baseDate').value;
  $('btnProcess').disabled = !baseReady;
}

$('btnProcess').onclick = async () => {
  $('btnProcess').disabled = true;
  showStatus('processStatus', 'info', '<span class="spinner"></span> Обработка файлов...');
  
  // Собираем конфигурацию
  const config = {
    base_sheet: $('baseSheet').value,
    base_serial: $('baseSerial').value,
    base_date: $('baseDate').value,
    files_config: []
  };
  
  const fileConfigs = document.querySelectorAll('#configFilesList .file-item');
  fileConfigs.forEach((item, idx) => {
    config.files_config.push({
      sheet: $(`sheet${idx}`).value,
      serial_col: $(`serial${idx}`).value,
      date_col: $(`date${idx}`).value || null,
      compare: $(`opCompare${idx}`).checked,
      tech_refresh: $(`opTechRefresh${idx}`).checked
    });
  });
  
  try {
    const r = await fetch(API + '/process_multiple', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(config)
    });
    const d = await r.json();
    if (!r.ok) throw new Error(d.detail || 'Ошибка обработки');
    
    showStatus('processStatus', 'ok', `✓ Обработано файлов: ${d.results.length}`);
    
    // Отображаем результаты
    d.results.forEach((res, idx) => {
      createResultItem(idx, res);
    });
    
    $('downloadAllLink').href = API + '/download_all';
    $('step3').classList.remove('hidden');
  } catch(e) {
    showStatus('processStatus', 'err', '✗ ' + e.message);
    $('btnProcess').disabled = false;
  }
};

// --- Step 3: Результаты ---
function createResultItem(idx, result) {
  const div = document.createElement('div');
  div.className = 'result-item';
  div.innerHTML = `
    <h3 style="font-size: 1rem; margin-bottom: 8px; color: #333;">
      ${result.source_filename}
    </h3>
    <div class="result-stats">
      <div class="stat">
        <span class="stat-label">Строк:</span>
        <span class="stat-value">${result.total_rows}</span>
      </div>
      ${result.matched !== null ? `
        <div class="stat">
          <span class="stat-label">На складе:</span>
          <span class="stat-value">${result.matched}</span>
        </div>
      ` : ''}
      ${result.outdated !== null ? `
        <div class="stat">
          <span class="stat-label">Устарело:</span>
          <span class="stat-value">${result.outdated}</span>
        </div>
      ` : ''}
    </div>
    <a href="${API}/download_single?idx=${idx}" style="text-decoration: none;">
      <button class="btn-primary" style="padding: 8px 20px; font-size: 0.9rem; margin-top: 8px;">
        📥 Скачать ${result.result_filename}
      </button>
    </a>
  `;
  $('resultsList').appendChild(div);
}

// --- Helpers ---
function showStatus(id, type, html) {
  const el = $(id);
  el.className = 'status status-' + type;
  el.innerHTML = html;
  el.classList.remove('hidden');
}

function fillSelect(id, items, addEmpty) {
  const sel = $(id);
  sel.innerHTML = '';
  if (addEmpty) {
    const o = document.createElement('option');
    o.value = '';
    o.textContent = '— не выбран (пропустить техрефреш) —';
    sel.appendChild(o);
  }
  items.forEach(item => {
    const o = document.createElement('option');
    o.value = item;
    o.textContent = item;
    sel.appendChild(o);
  });
}

// Делаем функцию removeFile глобальной
window.removeFile = removeFile;

// ─── Переключение табов ───
function switchTab(tabName) {
  // Скрыть все вкладки
  document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  
  // Показать выбранную вкладку
  if (tabName === 'processing') {
    $('tabProcessing').classList.add('active');
    event.target.classList.add('active');
  } else if (tabName === 'warehouse') {
    $('tabWarehouse').classList.add('active');
    event.target.classList.add('active');
  }
}

window.switchTab = switchTab;

// ─── Склад: Управление файлом ───
let warehouseFileSelected = null;

$('warehouseFile').onchange = e => {
  warehouseFileSelected = e.target.files[0];
  $('warehouseFileName').textContent = warehouseFileSelected?.name || '';
  $('btnLoadWarehouse').disabled = !warehouseFileSelected;
};

$('btnLoadWarehouse').onclick = async () => {
  if (!warehouseFileSelected) return;
  
  $('btnLoadWarehouse').disabled = true;
  showStatus('warehouseStatus', 'info', '<span class="spinner"></span> Загрузка базы данных...');
  
  try {
    const formData = new FormData();
    formData.append('file', warehouseFileSelected);
    
    const r = await fetch(API + '/warehouse/upload', {
      method: 'POST',
      body: formData
    });
    
    if (!r.ok) {
      const errText = await r.text();
      throw new Error(errText);
    }
    
    showStatus('warehouseStatus', 'ok', '✓ База данных загружена');
    
    // Загружаем данные для фильтров
    await loadWarehouseData();
    
    // Показываем фильтры
    $('warehouseFiltersSection').classList.remove('hidden');
    
    setTimeout(() => $('warehouseStatus').classList.add('hidden'), 2000);
    
  } catch (e) {
    showStatus('warehouseStatus', 'err', '❌ Ошибка загрузки: ' + e.message);
    $('btnLoadWarehouse').disabled = false;
  }
};

// ─── Склад: Загрузка данных ───
async function loadWarehouseData() {
  try {
    // Загружаем типы оборудования
    const r = await fetch(API + '/warehouse/types');
    if (!r.ok) {
      const errText = await r.text();
      throw new Error(errText);
    }
    const data = await r.json();
    
    fillSelect('warehouseType', data.types);
    $('warehouseType').disabled = false;
    $('btnSearchWarehouse').disabled = false;
    
    // Устанавливаем обработчики
    $('warehouseType').onchange = async () => {
      const type = $('warehouseType').value;
      if (!type) {
        $('warehouseModel').disabled = true;
        fillSelect('warehouseModel', []);
        return;
      }
      
      // Загружаем модели для выбранного типа
      const r = await fetch(API + `/warehouse/models?type=${encodeURIComponent(type)}`);
      const d = await r.json();
      fillSelect('warehouseModel', d.models);
      $('warehouseModel').disabled = false;
    };
    
    $('btnSearchWarehouse').onclick = searchWarehouse;
    
  } catch (e) {
    showStatus('warehouseStatus', 'err', '❌ Ошибка загрузки данных склада: ' + e.message);
  }
}

// ─── Склад: Поиск ───
async function searchWarehouse() {
  const type = $('warehouseType').value;
  if (!type) {
    showStatus('warehouseStatus', 'err', '❌ Выберите тип оборудования');
    return;
  }
  
  showStatus('warehouseStatus', 'info', '<span class="spinner"></span> Поиск...');
  
  try {
    const model = $('warehouseModel').value;
    let url = API + `/warehouse/search?type=${encodeURIComponent(type)}`;
    if (model) url += `&model=${encodeURIComponent(model)}`;
    
    const r = await fetch(url);
    if (!r.ok) throw new Error(await r.text());
    const data = await r.json();
    
    displayWarehouseResults(data.items, data.total);
    $('warehouseStatus').classList.add('hidden');
    
  } catch (e) {
    showStatus('warehouseStatus', 'err', '❌ Ошибка поиска: ' + e.message);
  }
}

// ─── Склад: Отображение результатов ───
function displayWarehouseResults(items, total) {
  const container = $('warehouseResults');
  
  if (items.length === 0) {
    container.innerHTML = '<div class="warehouse-empty">🔍 Оборудование не найдено</div>';
    return;
  }
  
  let html = `<div class="warehouse-count">📦 Найдено: ${total} шт.</div>`;
  html += '<table class="warehouse-table">';
  html += '<thead><tr>';
  html += '<th>Адрес</th>';
  html += '<th>Корпус/Этаж</th>';
  html += '<th>Местоположение</th>';
  html += '<th>Тип оборудования</th>';
  html += '<th>Марка</th>';
  html += '<th>Модель</th>';
  html += '<th>Серийный номер</th>';
  html += '<th>Инвентарный номер</th>';
  html += '</tr></thead><tbody>';
  
  items.forEach(item => {
    html += '<tr>';
    html += `<td>${item['Адрес'] || '-'}</td>`;
    html += `<td>${item['корпус/этаж'] || '-'}</td>`;
    html += `<td>${item['Местоположение'] || '-'}</td>`;
    html += `<td>${item['Тип оборудования'] || '-'}</td>`;
    html += `<td>${item['Марка'] || '-'}</td>`;
    html += `<td>${item['Модель'] || '-'}</td>`;
    html += `<td>${item['Серийный номер'] || '-'}</td>`;
    html += `<td>${item['Инвентарный номер'] || '-'}</td>`;
    html += '</tr>';
  });
  
  html += '</tbody></table>';
  container.innerHTML = html;
}

</script>
</body>
</html>
"""


# ─── API Routes ──────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse)
def main_form():
    return HTML_PAGE


@app.post("/upload_multiple")
async def upload_multiple(
    base_file: UploadFile = File(...),
    process_files: List[UploadFile] = File(...)
):
    """Загрузка базового файла + массива файлов для обработки"""
    allowed_ext = (".xlsx", ".xlsb")
    
    # Проверка базового файла
    if not base_file.filename.lower().endswith(allowed_ext):
        raise HTTPException(400, f"Базовый файл {base_file.filename} — неподдерживаемый формат")
    
    # Сохраняем базовый файл
    base_path = save_temp_file(base_file)
    base_engine = get_engine(base_file.filename)
    
    try:
        base_sheets = get_sheet_names(base_path, base_engine)
    except Exception as e:
        raise HTTPException(500, f"Не удалось прочитать листы базового файла: {e}")
    
    session_data["base_file"] = {
        "path": base_path,
        "engine": base_engine,
        "filename": base_file.filename,
        "sheets": base_sheets
    }
    
    # Обрабатываем файлы для обработки
    session_data["process_files"] = []
    process_files_info = []
    
    for pf in process_files:
        if not pf.filename.lower().endswith(allowed_ext):
            raise HTTPException(400, f"Файл {pf.filename} — неподдерживаемый формат")
        
        path = save_temp_file(pf)
        engine = get_engine(pf.filename)
        
        try:
            sheets = get_sheet_names(path, engine)
        except Exception as e:
            raise HTTPException(500, f"Не удалось прочитать листы файла {pf.filename}: {e}")
        
        session_data["process_files"].append({
            "path": path,
            "engine": engine,
            "filename": pf.filename,
            "sheets": sheets
        })
        
        process_files_info.append({
            "filename": pf.filename,
            "sheets": sheets
        })
    
    return {
        "base_sheets": base_sheets,
        "files_count": len(process_files),
        "process_files_info": process_files_info
    }


@app.get("/columns")
def get_cols(file_type: str, sheet: str, file_idx: Optional[int] = None):
    """Возвращает столбцы указанного листа"""
    if file_type == "base":
        if not session_data.get("base_file"):
            raise HTTPException(400, "Базовый файл не загружен")
        file_info = session_data["base_file"]
    elif file_type == "process":
        if file_idx is None:
            raise HTTPException(400, "Не указан индекс файла")
        if file_idx >= len(session_data["process_files"]):
            raise HTTPException(400, "Некорректный индекс файла")
        file_info = session_data["process_files"][file_idx]
    else:
        raise HTTPException(400, "Некорректный тип файла")
    
    try:
        cols = get_columns(file_info["path"], file_info["engine"], sheet)
        detected = auto_detect_columns(cols)
    except Exception as e:
        raise HTTPException(500, f"Ошибка чтения столбцов: {e}")
    
    return {
        "columns": cols,
        "detected_serial": detected["serial"],
        "detected_date": detected["date"]
    }


@app.post("/process_multiple")
async def process_multiple(config: dict):
    """Обработка всех файлов"""
    if not session_data.get("base_file") or not session_data.get("process_files"):
        raise HTTPException(400, "Файлы не загружены")
    
    base = session_data["base_file"]
    results = []
    session_data["results"] = []
    
    for idx, file_info in enumerate(session_data["process_files"]):
        file_config = config["files_config"][idx]
        
        try:
            result_path = process_excels(
                path1=file_info["path"],
                path2=base["path"],
                engine1=file_info["engine"],
                engine2=base["engine"],
                sheet1=file_config["sheet"],
                sheet2=config["base_sheet"],
                serial_col1=file_config["serial_col"],
                serial_col2=config["base_serial"],
                date_col1=file_config["date_col"],
                date_col2=config["base_date"],
                compare=file_config["compare"],
                tech_refresh=file_config["tech_refresh"]
            )
        except Exception as e:
            raise HTTPException(500, f"Ошибка обработки файла {file_info['filename']}: {e}")
        
        # Статистика
        import pandas as pd
        df = pd.read_excel(result_path, engine="calamine")
        
        matched = None
        outdated = None
        
        if file_config["compare"] and "Передано на склад" in df.columns:
            matched = int((df["Передано на склад"] == "Да").sum())
        
        if file_config["tech_refresh"] and "Оборудование устарело" in df.columns:
            outdated = int(df["Оборудование устарело"].str.startswith("Да", na=False).sum())
        
        result_filename = f"result_{idx + 1}_{file_info['filename']}"
        
        session_data["results"].append({
            "path": result_path,
            "filename": result_filename
        })
        
        results.append({
            "source_filename": file_info["filename"],
            "result_filename": result_filename,
            "total_rows": len(df),
            "matched": matched,
            "outdated": outdated
        })
    
    return {"results": results}


@app.get("/download_single")
def download_single(idx: int):
    """Скачать отдельный результат"""
    if idx >= len(session_data["results"]):
        raise HTTPException(400, "Некорректный индекс файла")
    
    result = session_data["results"][idx]
    return FileResponse(
        result["path"],
        filename=result["filename"],
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.get("/download_all")
def download_all():
    """Скачать все результаты в ZIP"""
    if not session_data.get("results"):
        raise HTTPException(400, "Нет результатов для скачивания")
    
    # Создаем ZIP архив
    zip_path = os.path.join(tempfile.gettempdir(), "results_all.zip")
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for result in session_data["results"]:
            zipf.write(result["path"], result["filename"])
    
    return FileResponse(
        zip_path,
        filename="results_all.zip",
        media_type="application/zip"
    )


# ─── Склад ───────────────────────────────────────────────────────────────────

@app.post("/warehouse/upload")
async def warehouse_upload(file: UploadFile = File(...)):
    """Загрузить файл базы данных для склада"""
    try:
        # Сохраняем файл
        file_path = save_temp_file(file.file, file.filename)
        engine = get_engine(file.filename)
        
        # Проверяем наличие листа "Возврат"
        sheets = get_sheet_names(file_path, engine)
        if "Возврат" not in sheets:
            raise HTTPException(400, f"Лист 'Возврат' не найден. Доступные листы: {', '.join(sheets)}")
        
        # Сохраняем в session_data
        session_data["base_file"] = {
            "path": file_path,
            "engine": engine,
            "filename": file.filename,
            "sheets": sheets
        }
        
        return {"status": "ok", "filename": file.filename, "sheets": sheets}
    
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, f"Ошибка загрузки файла: {str(e)}")


@app.get("/warehouse/types")
def warehouse_types():
    """Получить уникальные типы оборудования из листа Возврат"""
    if not session_data.get("base_file"):
        raise HTTPException(400, "База данных не загружена")
    
    try:
        import pandas as pd
        from .excel_logic import _read_sheet_safe
        
        base = session_data["base_file"]
        df = _read_sheet_safe(base["path"], base["engine"], "Возврат")
        
        if "Тип оборудования" not in df.columns:
            raise HTTPException(400, "Столбец 'Тип оборудования' не найден на листе 'Возврат'")
        
        # Получаем уникальные типы, исключая пустые значения
        types = df["Тип оборудования"].dropna().unique().tolist()
        types = sorted([str(t).strip() for t in types if str(t).strip()])
        
        return {"types": types}
    
    except Exception as e:
        raise HTTPException(500, f"Ошибка чтения данных: {str(e)}")


@app.get("/warehouse/models")
def warehouse_models(type: str):
    """Получить модели по типу оборудования"""
    if not session_data.get("base_file"):
        raise HTTPException(400, "База данных не загружена")
    
    try:
        import pandas as pd
        from .excel_logic import _read_sheet_safe
        
        base = session_data["base_file"]
        df = _read_sheet_safe(base["path"], base["engine"], "Возврат")
        
        if "Тип оборудования" not in df.columns or "Модель" not in df.columns:
            raise HTTPException(400, "Необходимые столбцы не найдены")
        
        # Фильтруем по типу
        filtered = df[df["Тип оборудования"] == type]
        
        # Получаем уникальные модели
        models = filtered["Модель"].dropna().unique().tolist()
        models = sorted([str(m).strip() for m in models if str(m).strip()])
        
        return {"models": models}
    
    except Exception as e:
        raise HTTPException(500, f"Ошибка чтения данных: {str(e)}")


@app.get("/warehouse/search")
def warehouse_search(type: str, model: Optional[str] = None):
    """Поиск оборудования на складе"""
    if not session_data.get("base_file"):
        raise HTTPException(400, "База данных не загружена")
    
    try:
        import pandas as pd
        from .excel_logic import _read_sheet_safe
        
        base = session_data["base_file"]
        df = _read_sheet_safe(base["path"], base["engine"], "Возврат")
        
        # Проверяем наличие всех необходимых столбцов
        required_cols = ["Адрес", "корпус/этаж", "Местоположение", "Тип оборудования", 
                        "Марка", "Модель", "Серийный номер", "Инвентарный номер"]
        
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise HTTPException(400, f"Отсутствуют столбцы: {', '.join(missing)}")
        
        # Фильтруем по типу
        filtered = df[df["Тип оборудования"] == type]
        
        # Фильтруем по модели, если указана
        if model:
            filtered = filtered[filtered["Модель"] == model]
        
        # Преобразуем в список словарей
        items = filtered[required_cols].fillna("").to_dict('records')
        
        return {
            "items": items,
            "total": len(items)
        }
    
    except Exception as e:
        raise HTTPException(500, f"Ошибка поиска: {str(e)}")


if __name__ == "__main__":
    uvicorn.run("app.main:app", host="127.0.0.1", port=8001, reload=True)
