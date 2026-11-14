/**
 * Simulación de Interfaz Microsoft Excel 365 — Modelos EOQ / CEP Clásico y con Faltantes
 * Controla la barra de fórmulas interactiva (fx), gráficos dinámicos y navegación de hojas.
 */

document.addEventListener('DOMContentLoaded', () => {
  initFormulaBarInteraction();
  initSheetTabs();
  drawExcelEOQCurves();

  // Estado inicial: Mostrar métricas de la Hoja 1 (Nestlé)
  updateStatusMetrics('sheet-cep');
});

/**
 * Permite hacer clic en cualquier celda de datos/fórmulas para actualizar
 * el Cuadro de Nombres (ej. G8) y la Barra de Fórmulas (fx) igual que en Excel.
 */
function initFormulaBarInteraction() {
  const cells = document.querySelectorAll('.selectable-cell');
  const addressBox = document.getElementById('cellAddressBox');
  const formulaInput = document.getElementById('formulaInput');
  const formulaDesc = document.getElementById('formulaDesc');

  cells.forEach(cell => {
    cell.addEventListener('click', () => {
      cells.forEach(c => c.classList.remove('selected-cell'));
      cell.classList.add('selected-cell');

      const addr = cell.getAttribute('data-cell') || 'A1';
      const form = cell.getAttribute('data-formula') || cell.innerText.trim();
      const desc = cell.getAttribute('data-desc') || '';

      addressBox.textContent = addr;
      formulaInput.value = form;
      formulaDesc.textContent = desc;
    });
  });
}

/**
 * Dibuja el Gráfico 1 (CEP Clásico - Nestlé Nido)
 */
function drawExcelEOQCurves() {
  const canvas = document.getElementById('excelEoqCanvas');
  if (!canvas) return;

  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  if (rect.width === 0) return;

  canvas.width = rect.width * dpr;
  canvas.height = 175 * dpr;
  ctx.scale(dpr, dpr);

  const width = rect.width;
  const height = 175;
  ctx.clearRect(0, 0, width, height);

  const D = 3600;
  const Co = 27;
  const Cc = 6;
  const minQ = 50;
  const maxQ = 320;
  const maxCost = 2000;

  const padL = 44;
  const padR = 20;
  const padT = 16;
  const padB = 26;
  const plotW = width - padL - padR;
  const plotH = height - padT - padB;

  const toX = q => padL + ((q - minQ) / (maxQ - minQ)) * plotW;
  const toY = c => padT + plotH - (c / maxCost) * plotH;

  ctx.strokeStyle = '#E1DFDD';
  ctx.fillStyle = '#605E5C';
  ctx.font = '600 9.5px "Plus Jakarta Sans"';
  ctx.textAlign = 'right';

  [0, 500, 1000, 1500, 2000].forEach(cost => {
    const y = toY(cost);
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(width - padR, y);
    ctx.stroke();
    ctx.fillText(`$${cost}`, padL - 5, y + 3);
  });

  ctx.textAlign = 'center';
  [60, 120, 180, 240, 300].forEach(q => {
    const x = toX(q);
    ctx.fillText(`Q=${q}`, x, height - 7);
  });

  function plotCurve(fn, color, lineWidth, dashed = false) {
    ctx.beginPath();
    ctx.strokeStyle = color;
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(dashed ? [4, 3] : []);
    for (let q = minQ; q <= maxQ; q += 4) {
      const cost = fn(q);
      const x = toX(q);
      const y = toY(Math.min(cost, maxCost * 1.05));
      if (q === minQ) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    }
    ctx.stroke();
    ctx.setLineDash([]);
  }

  plotCurve(q => Cc * (q / 2), '#21A366', 2.2);
  plotCurve(q => Co * (D / q), '#2B579A', 2.2, true);
  plotCurve(q => Co * (D / q) + Cc * (q / 2), '#185C37', 2.8);

  const optX = toX(180);
  const optYTotal = toY(1080);
  const optYCross = toY(540);

  ctx.strokeStyle = '#107C41';
  ctx.lineWidth = 1.5;
  ctx.setLineDash([3, 3]);
  ctx.beginPath();
  ctx.moveTo(optX, optYTotal);
  ctx.lineTo(optX, padT + plotH);
  ctx.stroke();
  ctx.setLineDash([]);

  ctx.fillStyle = '#2B579A';
  ctx.beginPath();
  ctx.arc(optX, optYCross, 4.5, 0, Math.PI * 2);
  ctx.fill();

  ctx.fillStyle = '#107C41';
  ctx.beginPath();
  ctx.arc(optX, optYTotal, 5, 0, Math.PI * 2);
  ctx.fill();

  ctx.font = '800 9.5px "Plus Jakarta Sans"';
  ctx.textAlign = 'left';
  ctx.fillStyle = '#185C37';
  ctx.fillText('CT* Mín = $1,080 (Q*=180)', optX + 7, optYTotal - 4);
  ctx.fillStyle = '#2B579A';
  ctx.fillText('Equilibrio Ordenar/Mtto = $540', optX + 7, optYCross + 3);
}

/**
 * Dibuja el Gráfico 2 (CEP con Faltantes - Grupo Q / Vehículos)
 */
function drawExcelFaltantesCurves() {
  const canvas = document.getElementById('excelFaltantesCanvas');
  if (!canvas) return;

  const ctx = canvas.getContext('2d');
  const dpr = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  if (rect.width === 0) return;

  canvas.width = rect.width * dpr;
  canvas.height = 175 * dpr;
  ctx.scale(dpr, dpr);

  const width = rect.width;
  const height = 175;
  ctx.clearRect(0, 0, width, height);

  const D = 1200;
  const Co = 10000;
  const Cc = 6250;
  const Cs = 15000;

  const minQ = 30;
  const maxQ = 140;
  const maxCost = 600000;

  const padL = 52;
  const padR = 20;
  const padT = 16;
  const padB = 26;
  const plotW = width - padL - padR;
  const plotH = height - padT - padB;

  const toX = q => padL + ((q - minQ) / (maxQ - minQ)) * plotW;
  const toY = c => padT + plotH - (c / maxCost) * plotH;

  ctx.strokeStyle = '#E1DFDD';
  ctx.fillStyle = '#605E5C';
  ctx.font = '600 9.5px "Plus Jakarta Sans"';
  ctx.textAlign = 'right';

  [0, 150000, 300000, 450000, 600000].forEach(cost => {
    const y = toY(cost);
    ctx.beginPath();
    ctx.moveTo(padL, y);
    ctx.lineTo(width - padR, y);
    ctx.stroke();
    ctx.fillText(`$${cost/1000}k`, padL - 5, y + 3);
  });

  ctx.textAlign = 'center';
  [40, 60, 76, 100, 120].forEach(q => {
    const x = toX(q);
    ctx.fillText(`Q=${q}`, x, height - 7);
  });

  function plotCurve(fn, color, lineWidth, dashed = false) {
    ctx.beginPath();
    ctx.strokeStyle = color;
    ctx.lineWidth = lineWidth;
    ctx.setLineDash(dashed ? [4, 3] : []);
    for (let q = minQ; q <= maxQ; q += 2) {
      const cost = fn(q);
      const x = toX(q);
      const y = toY(Math.min(cost, maxCost * 1.05));
      if (q === minQ) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
    }
    ctx.stroke();
    ctx.setLineDash([]);
  }

  plotCurve(q => Co * (D / q), '#2B579A', 2.0, true);
  plotCurve(q => Cc * Math.pow(q - ((q*Cc)/(Cs+Cc)), 2) / (2*q), '#21A366', 2.0);
  plotCurve(q => Math.sqrt(2 * Co * Cc * D) * Math.sqrt(Cs / (Cs + Cc)), '#185C37', 2.8);

  const optX = toX(76);
  const optYTotal = toY(316227);

  ctx.strokeStyle = '#107C41';
  ctx.lineWidth = 1.5;
  ctx.setLineDash([3, 3]);
  ctx.beginPath();
  ctx.moveTo(optX, optYTotal);
  ctx.lineTo(optX, padT + plotH);
  ctx.stroke();
  ctx.setLineDash([]);

  ctx.fillStyle = '#107C41';
  ctx.beginPath();
  ctx.arc(optX, optYTotal, 5, 0, Math.PI * 2);
  ctx.fill();

  ctx.font = '800 9.5px "Plus Jakarta Sans"';
  ctx.textAlign = 'left';
  ctx.fillStyle = '#185C37';
  ctx.fillText('Ct* = $316.2k (Q*=76)', optX + 7, optYTotal - 4);
}

/**
 * Actualiza los valores de la barra de estado según la hoja activa
 */
function updateStatusMetrics(sheetId) {
  const avgEl = document.getElementById('statusAverage');
  const sumEl = document.getElementById('statusSum');

  if (!avgEl || !sumEl) return;

  if (sheetId === 'sheet-tema2') {
    avgEl.textContent = '$158,113.88';
    sumEl.textContent = '$316,227.77 USD';
  } else {
    avgEl.textContent = '$540.00';
    sumEl.textContent = '$1,080.00 USD';
  }
}

/**
 * Control robusto para cambiar entre la Hoja 1 y la Hoja 2
 */
function initSheetTabs() {
  const tabs = document.querySelectorAll('.sheet-tab');
  const sheetCep = document.getElementById('sheet-cep');
  const sheetTema2 = document.getElementById('sheet-tema2');

  if (!tabs.length || !sheetCep || !sheetTema2) return;

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const sheetId = tab.getAttribute('data-sheet');

      // Actualizar clases activas en los botones
      tabs.forEach(t => t.classList.remove('active'));
      tab.classList.add('active');

      // Mostrar u ocultar contenedores correspondientes y actualizar barra inferior
      updateStatusMetrics(sheetId);

      if (sheetId === 'sheet-tema2') {
        sheetCep.classList.add('hidden-sheet');
        sheetTema2.classList.remove('hidden-sheet');
        drawExcelFaltantesCurves();
        // Fuerza el renderizado de KaTeX al mostrar la hoja
        if (typeof renderMathInElement === 'function') {
          renderMathInElement(sheetTema2);
        }
      } else {
        sheetTema2.classList.add('hidden-sheet');
        sheetCep.classList.remove('hidden-sheet');
        drawExcelEOQCurves();
      }
    });
  });
}
