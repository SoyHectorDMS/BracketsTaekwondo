/* ============================================================
   BRACKET TAEKWONDO — app.js
   ============================================================ */

'use strict';

/* ----------------------------------------------------------
   ESTADO GLOBAL
   ---------------------------------------------------------- */

/**
 * Estructura plana de todos los partidos, indexados por id.
 * Cada partido tiene:
 *   id, slotA, slotB, winner,
 *   isBye,          — si uno de los dos slots es null desde el inicio
 *   roundIdx,       — índice de ronda (0 = primera)
 *   matchIdx,       — posición dentro de la ronda
 *   parentId,       — id del partido padre (null si es la final)
 *   parentSlot,     — 'A' o 'B': qué slot del padre alimenta este partido
 *   childAId,       — id del partido hijo que alimenta slotA
 *   childBId        — id del partido hijo que alimenta slotB
 * @type {Object.<string, Object>}
 */
let matches = {};

/** Orden de rondas: rounds[ri] = [id, id, ...] */
let rounds = [];

/** Partido(s) por 3.° lugar */
let tercerMatch = null;

/** Combates preliminares para evitar un cuadro lleno de pases directos. */
let prelimMatches = [];
let prelimPlacements = [];


/* ----------------------------------------------------------
   OPCIONES TOMADAS DE LA PLANTILLA EXCEL
   ---------------------------------------------------------- */

const EXCEL_SELECT_CONFIG = {
  'KYORUGUI': {
    categoryOptions: [
      'PRE INFANTIL (4-6 años)',
      'INFANTIL (7-9 años)',
      'PRE CADETE (10-11 años)',
      'CADETE (12-14 años)',
      'JUVENIL (15-17 años)',
      'MAYORES (+18 años)',
    ],
    genderOptions: ['Masculino', 'Femenino'],
    divisionLabel: 'División',
    divisionPlaceholder: 'Selecciona edad y género',
    categoryCodes: {
      'PRE INFANTIL (4-6 años)': 'PI',
      'INFANTIL (7-9 años)': 'I',
      'PRE CADETE (10-11 años)': 'PC',
      'CADETE (12-14 años)': 'C',
      'JUVENIL (15-17 años)': 'J',
      'MAYORES (+18 años)': 'M',
    },
    weightsByCategoryGender: {
      'PI_Femenino': ['-19 kg', '-22 kg', '-25 kg', '-27 kg', '-29 kg', '-31 kg', '+34 kg'],
      'PI_Masculino': ['-19 kg', '-22 kg', '-25 kg', '-27 kg', '-29 kg', '-31 kg', '+34 kg'],
      'I_Femenino': ['-21 kg', '-24 kg', '-27 kg', '-30 kg', '-33 kg', '-36 kg', '-39 kg', '+39 kg'],
      'I_Masculino': ['-24 kg', '-27 kg', '-30 kg', '-33 kg', '-36 kg', '-39 kg', '-42 kg', '+42 kg'],
      'PC_Femenino': ['-27 kg', '-30 kg', '-33 kg', '-36 kg', '-39 kg', '-42 kg', '-45 kg', '-48 kg', '-51 kg', '+51 kg'],
      'PC_Masculino': ['-30 kg', '-33 kg', '-36 kg', '-39 kg', '-42 kg', '-45 kg', '-48 kg', '-51 kg', '-54 kg', '+54 kg'],
      'C_Femenino': ['-29 kg', '-33 kg', '-37 kg', '-41 kg', '-44 kg', '-47 kg', '-51 kg', '-55 kg', '-59 kg', '+59 kg'],
      'C_Masculino': ['-33 kg', '-37 kg', '-41 kg', '-45 kg', '-49 kg', '-53 kg', '-57 kg', '-61 kg', '-65 kg', '+65 kg'],
      'J_Femenino': ['-42 kg', '-44 kg', '-46 kg', '-49 kg', '-52 kg', '-55 kg', '-59 kg', '-63 kg', '-68 kg', '+68 kg'],
      'J_Masculino': ['-45 kg', '-48 kg', '-51 kg', '-55 kg', '-59 kg', '-63 kg', '-68 kg', '-73 kg', '-78 kg', '+78 kg'],
      'M_Femenino': ['-46 kg', '-49 kg', '-53 kg', '-57 kg', '-62 kg', '-67 kg', '-73 kg', '+73 kg'],
      'M_Masculino': ['-54 kg', '-58 kg', '-63 kg', '-68 kg', '-74 kg', '-80 kg', '-87 kg', '+87 kg'],
    },
  },
  'CIRCUITO MOTOR': {
    categoryOptions: ['PRE INFANTIL (4-6 años)', 'INFANTIL (7-9 años)'],
    genderOptions: ['Masculino', 'Femenino'],
    divisionLabel: 'Kup',
    divisionPlaceholder: 'Seleccionar Kup',
    divisionOptions: ['10 Kup', '9 Kup', '8 Kup', '7 Kup', '6 Kup', '5 Kup', '4 Kup', '3 Kup', '2 Kup', '1 Kup'],
  },
  'POOMSAE': {
    categoryOptions: [
      'PRE INFANTIL (4-6 años)',
      'INFANTIL (7-9 años)',
      'PRE CADETE (10-11 años)',
      'CADETE (12-14 años)',
      'JUVENIL (15-17 años)',
      'MAYORES (+18 años)',
    ],
    genderOptions: ['Masculino', 'Femenino'],
    divisionLabel: 'Nombre del Poomsae',
    divisionPlaceholder: 'Seleccionar Poomsae',
    divisionOptions: [
      'Kibom (9no Kup)',
      'Il Jang (8vo Kup)',
      'I Jang (7mo Kup)',
      'Sam Jang (6to Kup)',
      'Sa Jang (5to Kup)',
      'Oh Jang (4to Kup)',
      'Yuk Jang (3er Kup)',
      'Chil Jang (2do Kup)',
      'Pal Jang (1er Kup)',
      'Koryo (1er Poom o Dan)',
      'Keumgang (2do Poom o Dan)',
      'Taebaeck (3er Poom o Dan)',
      'Pyongwon (4to Dan)',
      'Sipjing (5to Dan)',
      'Ji Tar (6to Dan)',
      'Chongkwon (7mo Dan)',
      'Hansu (8vo Dan)',
    ],
  },
  'FESTIVAL': {
    categoryOptions: [
      'PRE INFANTIL (4-6 años)',
      'INFANTIL (7-9 años)',
      'PRE CADETE (10-11 años)',
      'CADETE (12-14 años)',
      'JUVENIL (15-17 años)',
      'MAYORES (+18 años)',
    ],
    genderOptions: ['Masculino', 'Femenino'],
    divisionLabel: 'Grado',
    divisionPlaceholder: 'Seleccionar grado',
    divisionOptions: ['Blanco', 'Blanco punta amarilla', 'Amarillo'],
  },
};

function getSelectedSheetName() {
  return document.getElementById('excelHojaDestino')?.value || 'KYORUGUI';
}

function getSelectedExcelConfig() {
  return EXCEL_SELECT_CONFIG[getSelectedSheetName()] || EXCEL_SELECT_CONFIG.KYORUGUI;
}

function setSelectOptions(selectEl, options, placeholder, keepValue = true) {
  if (!selectEl) return;
  const prev = keepValue ? selectEl.value : '';
  const opts = Array.isArray(options) ? options : [];
  selectEl.innerHTML = `<option value="">${esc(placeholder || 'Seleccionar opción')}</option>` +
    opts.map(value => `<option value="${esc(value)}">${esc(value)}</option>`).join('');

  if (prev && opts.includes(prev)) {
    selectEl.value = prev;
  } else {
    selectEl.value = '';
  }
}

function getDivisionOptionsForCurrentSelection() {
  const config = getSelectedExcelConfig();
  if (getSelectedSheetName() !== 'KYORUGUI') {
    return config.divisionOptions || [];
  }

  const categoria = getCategoriaEdad();
  const genero = getGenero();
  const code = config.categoryCodes?.[categoria];
  const key = code && genero ? `${code}_${genero}` : '';
  return config.weightsByCategoryGender?.[key] || [];
}

function refreshExcelDrivenSelects(options = {}) {
  const { keepValues = true, updateDivisionOnly = false } = options;
  const config = getSelectedExcelConfig();

  const categoriaEl = document.getElementById('categoria');
  const generoEl = document.getElementById('genero');
  const divisionEl = document.getElementById('division');
  const divisionLabelEl = document.getElementById('divisionLabel');

  if (!updateDivisionOnly) {
    setSelectOptions(categoriaEl, config.categoryOptions, 'Seleccionar categoría', keepValues);
    setSelectOptions(generoEl, config.genderOptions, 'Seleccionar género', keepValues);
  }

  if (divisionLabelEl) {
    divisionLabelEl.textContent = config.divisionLabel || 'División / Peso';
  }

  const divisionOptions = getDivisionOptionsForCurrentSelection();
  const placeholder = getSelectedSheetName() === 'KYORUGUI' && divisionOptions.length === 0
    ? 'Selecciona edad y género'
    : (config.divisionPlaceholder || 'Seleccionar opción');

  setSelectOptions(divisionEl, divisionOptions, placeholder, keepValues);
  if (divisionEl) divisionEl.disabled = divisionOptions.length === 0;
}

/* ----------------------------------------------------------
   UTILIDADES
   ---------------------------------------------------------- */

function shuffle(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

function nextPow2(n) {
  let p = 1;
  while (p < n) p *= 2;
  return p;
}

function esc(s) {
  if (!s) return '';
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function getCompetidores() {
  return document.getElementById('competidores')
    .value.split('\n').map(s => s.trim()).filter(Boolean);
}

function distribuirCompetidoresConByesBalanceados(competidores) {
  const n = competidores.length;
  const size = nextPow2(n);

  if (n === size) {
    return [...competidores];
  }

  const slots = Array(size).fill(null);
  const firstRoundMatches = size / 2;

  /*
    Distribución balanceada: primero coloca un competidor en cada combate
    de primera ronda y luego completa los combates necesarios. Así los BYE
    quedan repartidos y nadie avanza varias rondas seguidas sin pelear.
  */
  competidores.forEach((competidor, index) => {
    if (index < firstRoundMatches) {
      slots[index * 2] = competidor;
    } else {
      const extraIndex = index - firstRoundMatches;
      slots[extraIndex * 2 + 1] = competidor;
    }
  });

  return slots;
}

const PRELIM_PLACEHOLDER_PREFIX = '__PRELIM_WINNER_';

function highestPow2LessOrEqual(n) {
  let p = 1;
  while (p * 2 <= n) p *= 2;
  return p;
}

function makePrelimPlaceholder(index) {
  return `${PRELIM_PLACEHOLDER_PREFIX}${index}__`;
}

function isPrelimPlaceholder(value) {
  return typeof value === 'string' && value.startsWith(PRELIM_PLACEHOLDER_PREFIX);
}

function getPrelimPlaceholderLabel(value) {
  const match = String(value || '').match(/^__PRELIM_WINNER_(\d+)__$/);
  const num = match ? Number(match[1]) + 1 : '';
  return `Ganador preliminar ${num}`;
}

/**
 * Modelos tomados de las llaves de referencia enviadas por el usuario.
 * D = competidor directo al cuadro principal.
 * P = ganador de un combate preliminar de 2 competidores.
 */
function generarOrdenSiembra(size) {
  if (size === 2) return [1, 2];

  let order = [1, 2];
  let currentSize = 2;

  while (currentSize < size) {
    const nextSize = currentSize * 2;
    order = order.flatMap(seed => [seed, nextSize + 1 - seed]);
    currentSize = nextSize;
  }

  return order;
}

function compactarPrimeraRondaDesdeSiembra(n, drawSize) {
  const orden = generarOrdenSiembra(drawSize);
  const slots = [];
  let prelimCount = 0;

  for (let i = 0; i < orden.length; i += 2) {
    const seedA = orden[i];
    const seedB = orden[i + 1];
    const hasA = seedA <= n;
    const hasB = seedB <= n;

    if (hasA && hasB) {
      slots.push([seedA, seedB]);
      prelimCount++;
    } else if (hasA) {
      slots.push(seedA);
    } else if (hasB) {
      slots.push(seedB);
    } else {
      slots.push(null);
    }
  }

  return { baseSize: drawSize / 2, slots, prelimCount };
}

function getDrawSizeForCompetitors(n) {
  if (n <= 2) return 2;
  if (n <= 4) return 4;
  if (n <= 8) return 8;
  if (n <= 16) return 16;
  if (n <= 32) return 32;
  return nextPow2(n);
}

function getCompactPrelimTemplate(n) {
  /*
    Nuevo criterio:
    Se muestra el cuadro oficial completo desde la primera columna.
    Si un competidor pasa por BYE, igual aparece desde el inicio en su cruce,
    y luego se replica automáticamente en la siguiente ronda.
  */
  const baseSize = getDrawSizeForCompetitors(n);
  return {
    baseSize,
    slots: generarOrdenSiembra(baseSize),
  };
}

function prepararBracketSegunModelos(competidores) {
  const template = getCompactPrelimTemplate(competidores.length);
  if (!template) return null;

  prelimMatches = [];
  prelimPlacements = [];

  const slots = Array(template.baseSize).fill(null);

  template.slots.forEach((entry, slotIndex) => {
    if (entry === null || entry === undefined) {
      slots[slotIndex] = null;
      return;
    }

    if (Array.isArray(entry)) {
      const prelimIdx = prelimMatches.length;
      const seedA = entry[0];
      const seedB = entry[1];
      const slotA = competidores[seedA - 1] || null;
      const slotB = competidores[seedB - 1] || null;
      const placeholder = makePrelimPlaceholder(prelimIdx);

      prelimMatches.push({
        id: `prelim-m${prelimIdx}`,
        slotA,
        slotB,
        winner: null,
        isBye: false,
        roundIdx: -1,
        matchIdx: prelimIdx,
        placeholder,
        targetMatchId: `r0m${Math.floor(slotIndex / 2)}`,
        targetSlot: slotIndex % 2 === 0 ? 'A' : 'B',
      });

      prelimPlacements[prelimIdx] = slotIndex;
      slots[slotIndex] = placeholder;
      return;
    }

    slots[slotIndex] = competidores[entry - 1] || null;
  });

  return {
    baseSize: template.baseSize,
    slots,
    prelimCount: prelimMatches.length,
  };
}

function conectarPreliminaresAlCuadroPrincipal() {
  prelimMatches.forEach((prelim) => {
    if (!prelim.targetMatchId || !prelim.targetSlot) return;
  });
}

function aplicarGanadorPreliminar(prelim) {
  if (!prelim || !prelim.targetMatchId || !prelim.targetSlot) return;

  const target = matches[prelim.targetMatchId];
  if (!target) return;

  const propName = prelim.targetSlot === 'A' ? 'slotA' : 'slotB';
  const oldValue = target[propName];
  const newValue = prelim.winner || prelim.placeholder;

  target[propName] = newValue;

  if (target.winner === oldValue && oldValue !== newValue) target.winner = null;

  propagate(target);
}

function getCategoriaEdad() {
  return document.getElementById('categoria')?.value.trim() || '';
}

function getCategoriaCodigoExcel(categoriaEdad) {
  const config = EXCEL_SELECT_CONFIG.KYORUGUI;
  return config.categoryCodes?.[categoriaEdad] || '';
}

function getGenero() {
  return document.getElementById('genero')?.value.trim() || '';
}

function getDivision() {
  return document.getElementById('division')?.value.trim() || '';
}

function getCategoriaDisplay() {
  const categoria = getCategoriaEdad();
  const genero    = getGenero();
  return [categoria, genero].filter(Boolean).join(' — ');
}

function buildChampionCardHtml(nombre) {
  return `<div class="champion-card">
    <div class="champion-label">
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 9H4.5a2 2 0 0 1 0-4H6"/><path d="M18 9h1.5a2 2 0 0 0 0-4H18"/><path d="M4 22h16"/><path d="M10 14.66V17c0 .55-.47.98-.97 1.21C7.85 18.75 7 20.24 7 22"/><path d="M14 14.66V17c0 .55.47.98.97 1.21C16.15 18.75 17 20.24 17 22"/><path d="M18 2H6v7a6 6 0 0 0 12 0V2Z"/></svg>
      Campeón
    </div>
    <div class="champion-name" id="championName">${esc(nombre || '—')}</div>
  </div>`;
}

function nombreRonda(idxDesdeElFinal) {
  switch (idxDesdeElFinal) {
    case 0: return 'Final';
    case 1: return 'Semifinal';
    case 2: return 'Cuartos de final';
    case 3: return 'Octavos de final';
    default: return `Ronda ${idxDesdeElFinal + 1}`;
  }
}

/* ----------------------------------------------------------
   CONSTRUCCIÓN DEL ÁRBOL — referencia directa padre/hijo
   ---------------------------------------------------------- */

/**
 * Construye todos los partidos con referencias cruzadas.
 * @param {(string|null)[]} slots  Arreglo de tamaño potencia-de-2
 */
function buildMatches(slots) {
  matches = {};
  rounds  = [];

  /* Ronda 0 */
  const r0 = [];
  for (let i = 0; i < slots.length; i += 2) {
    const sA  = slots[i]     || null;
    const sB  = slots[i + 1] || null;
    const id  = `r0m${i / 2}`;
    const isBye = (!!sA) !== (!!sB);   // exactamente uno es null
    matches[id] = {
      id,
      slotA:    sA,
      slotB:    sB,
      winner:   isBye ? (sA || sB) : null,
      isBye,
      roundIdx: 0,
      matchIdx: i / 2,
      parentId:    null,
      parentSlot:  null,
      childAId:    null,
      childBId:    null,
    };
    r0.push(id);
  }
  rounds.push(r0);

  /* Rondas siguientes */
  let prevRound = r0;
  let ri = 1;
  while (prevRound.length > 1) {
    const curRound = [];
    for (let i = 0; i < prevRound.length; i += 2) {
      const childAId = prevRound[i];
      const childBId = prevRound[i + 1] || null;
      const id       = `r${ri}m${i / 2}`;

      const winA = childAId ? matches[childAId].winner : null;
      const winB = childBId ? matches[childBId].winner : null;

      /* Un partido de ronda superior solo es BYE automático si uno
         de los dos partidos hijos no existe (bracket incompleto).
         Si ambos hijos existen pero uno ya tiene ganador (por BYE
         de ronda 0), el partido padre igual requiere que el usuario
         elija ganador — NO se auto-resuelve. */
      const isBye = !childBId;   // solo si no hay segundo hijo

      matches[id] = {
        id,
        slotA:    winA,
        slotB:    winB,
        winner:   isBye ? (winA || winB) : null,
        isBye,
        roundIdx: ri,
        matchIdx: i / 2,
        parentId:    null,
        parentSlot:  null,
        childAId,
        childBId,
      };

      /* Apuntar hijos hacia este padre */
      if (childAId) {
        matches[childAId].parentId   = id;
        matches[childAId].parentSlot = 'A';
      }
      if (childBId) {
        matches[childBId].parentId   = id;
        matches[childBId].parentSlot = 'B';
      }

      curRound.push(id);
    }
    rounds.push(curRound);
    prevRound = curRound;
    ri++;
  }
}

/* ----------------------------------------------------------
   PROPAGACIÓN — sube en cascada usando referencias parentId
   ---------------------------------------------------------- */

/**
 * Después de cambiar match.winner, propaga hacia arriba.
 * @param {Object} changedMatch
 */
function propagate(changedMatch) {
  let current = changedMatch;

  while (current.parentId) {
    const parent    = matches[current.parentId];
    const slot      = current.parentSlot;   // 'A' o 'B'
    const newValue  = current.winner;       // ganador del partido hijo (puede ser null)
    const oldValue  = slot === 'A' ? parent.slotA : parent.slotB;

    if (slot === 'A') parent.slotA = newValue;
    else              parent.slotB = newValue;

    /* Si el ganador del padre era el que cambió, resetear */
    if (parent.winner === oldValue && oldValue !== newValue) {
      parent.winner = null;
    }

    /* Recalcular isBye del padre: solo si no tiene hijo en uno de los lados */
    parent.isBye = !parent.childBId;
    if (parent.isBye) {
      parent.winner = parent.slotA || parent.slotB;
    }

    current = parent;
  }

  /* Actualizar terceros lugares de Taekwondo */
  actualizarTercerosTaekwondo();
}

function getLoser(match) {
  if (!match || !match.winner) return null;
  return match.winner === match.slotA ? match.slotB : match.slotA;
}

/**
 * En Taekwondo no se disputa combate por 3.° lugar.
 * Los dos perdedores de semifinal ocupan ambos terceros lugares.
 */
function actualizarTercerosTaekwondo() {
  if (!tercerMatch) return;

  tercerMatch.thirds = [null, null];

  const numRounds = rounds.length;
  if (numRounds < 2) return;

  const semiIds = rounds[numRounds - 2] || [];
  tercerMatch.thirds = [
    semiIds[0] ? getLoser(matches[semiIds[0]]) : null,
    semiIds[1] ? getLoser(matches[semiIds[1]]) : null,
  ];
}

/* ----------------------------------------------------------
   RENDERIZADO
   ---------------------------------------------------------- */

function renderSlot(nombre, seedNum, lado, match) {
  const esPrelimPlaceholder = isPrelimPlaceholder(nombre);
  const esBye  = match.roundIdx === 0 && !nombre && (lado === 'A' ? !!match.slotB : !!match.slotA);
  const esTbd  = (!nombre && !esBye) || esPrelimPlaceholder;
  const winner = match.winner;

  let clases = 'match-slot';
  let content = '';

  if (esBye) {
    clases  += ' slot-bye';
    content  = `<span class="slot-name">Pase automático</span>
                <span class="bye-badge">BYE</span>`;
  } else if (esTbd) {
    clases  += ' slot-tbd';
    content  = `<span class="slot-name">${esPrelimPlaceholder ? getPrelimPlaceholderLabel(nombre) : 'Por definir'}</span>`;
  } else {
    const esWinner = !!winner && winner === nombre;
    const esLoser  = !!winner && winner !== nombre;
    if (esWinner) clases += ' slot-winner';
    if (esLoser)  clases += ' slot-loser';
    content = `<span class="slot-seed">${seedNum}</span>
               <span class="slot-name">${esc(nombre)}</span>
               <span class="slot-win-icon" aria-hidden="true">✓</span>`;
  }

  const dataAttrs = (!esBye && !esTbd)
    ? `data-match="${match.id}" data-lado="${lado}"`
    : '';

  return `<div class="${clases}" ${dataAttrs}>
    ${content}
  </div>`;
}

function renderMatchCard(match) {
  return `<div class="match ${match.isBye ? 'match-with-bye' : ''}" id="match-${match.id}">
    ${renderSlot(match.slotA, match.matchIdx * 2 + 1, 'A', match)}
    ${renderSlot(match.slotB, match.matchIdx * 2 + 2, 'B', match)}
  </div>`;
}

function getRoundOffset(ri, matchH, matchGap) {
  if (ri === 0) return 0;
  return ((Math.pow(2, ri) - 1) * (matchH + matchGap)) / 2;
}

function renderConnectors(numMatches, matchH, gapBetween) {
  const w      = 100;
  const slotH  = matchH / 2;
  const totalH = numMatches * matchH + (numMatches - 1) * gapBetween;
  let paths    = '';

  for (let i = 0; i < numMatches; i += 2) {
    const y1   = i * (matchH + gapBetween) + slotH;
    const y2   = (i + 1) * (matchH + gapBetween) + slotH;
    const yMid = (y1 + y2) / 2;

    paths += `<line x1="0" y1="${y1}" x2="${w/2}" y2="${y1}" stroke="var(--line-color)" stroke-width="1"/>`;
    paths += `<line x1="0" y1="${y2}" x2="${w/2}" y2="${y2}" stroke="var(--line-color)" stroke-width="1"/>`;
    paths += `<line x1="${w/2}" y1="${y1}" x2="${w/2}" y2="${y2}" stroke="var(--line-color)" stroke-width="1"/>`;
    paths += `<line x1="${w/2}" y1="${yMid}" x2="${w}" y2="${yMid}" stroke="var(--line-color)" stroke-width="1"/>`;
  }

  return `<svg class="connector-svg"
    width="${w}" height="${totalH}"
    viewBox="0 0 ${w} ${totalH}"
    fill="none" aria-hidden="true">${paths}</svg>`;
}

function renderPrelimConnectors(items, height, width) {
  const startX = 0;
  const midX   = Math.round(width * 0.56);
  const endX   = width;

  const paths = items.map(item => {
    const y1 = item.sourceCenter;
    const y2 = item.targetCenter;
    return `
      <path d="M ${startX} ${y1} H ${midX} V ${y2} H ${endX}"
        stroke="var(--line-color)" stroke-width="1" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`;
  }).join('');

  return `<svg class="connector-svg" width="${width}" height="${height}"
    viewBox="0 0 ${width} ${height}" fill="none" aria-hidden="true">${paths}</svg>`;
}

function renderBracket() {
  const bracketEl = document.getElementById('bracketEl');
  const matchH    = 72;
  const matchGap  = 28;
  const headerH   = 32;

  bracketEl.innerHTML = '';

  if (prelimMatches.length > 0) {
    const prelimCol = document.createElement('div');
    prelimCol.className = 'round prelim-round';
    prelimCol.innerHTML = `<div class="round-header">Preliminar</div>`;

    const prelimContainer = document.createElement('div');
    prelimContainer.className = 'round-matches prelim-matches';

    /*
      Rediseño: los preliminares se apilan con separación vertical uniforme.
      Luego un SVG dibuja las conexiones reales hacia el cuadro principal.
    */
    const prelimGap = 18;
    const prelimLayout = [];

    prelimMatches.forEach((match, mi) => {
      const wrapper = document.createElement('div');
      wrapper.className = 'match-wrapper';
      wrapper.style.marginTop = mi === 0 ? '0' : `${prelimGap}px`;
      wrapper.style.marginBottom = '0';
      wrapper.innerHTML = renderMatchCard(match);
      prelimContainer.appendChild(wrapper);

      const top = mi * (matchH + prelimGap);
      const pos = prelimPlacements[mi] ?? mi;
      const targetMatchIdx = Math.floor(pos / 2);
      const targetCenter = targetMatchIdx * (matchH + matchGap) +
        (pos % 2 === 0 ? matchH / 4 : (matchH * 3) / 4);

      prelimLayout.push({
        sourceCenter: top + matchH / 2,
        targetCenter,
      });
    });

    prelimCol.appendChild(prelimContainer);
    bracketEl.appendChild(prelimCol);

    const prelimConn = document.createElement('div');
    prelimConn.className = 'prelim-connector-col';
    prelimConn.style.marginTop = `${headerH + 12}px`;
    const connectorHeight = Math.max(
      prelimLayout.length ? prelimLayout[prelimLayout.length - 1].sourceCenter + matchH / 2 : 0,
      ...prelimLayout.map(item => item.targetCenter + matchH / 2)
    );
    prelimConn.innerHTML = renderPrelimConnectors(prelimLayout, connectorHeight, 118);
    bracketEl.appendChild(prelimConn);
  }

  rounds.forEach((roundIds, ri) => {
    const numMatches  = roundIds.length;
    const isLast      = ri === rounds.length - 1;
    const spacingFact = Math.pow(2, ri);
    const gapBetween  = matchGap + (spacingFact - 1) * (matchH + matchGap);
    const roundOffset = getRoundOffset(ri, matchH, matchGap);

    /* Columna de ronda */
    const col = document.createElement('div');
    col.className = 'round';

    const idxDesdeElFinal = rounds.length - 1 - ri;
    col.innerHTML = `<div class="round-header">${nombreRonda(idxDesdeElFinal)}</div>`;

    const matchesContainer = document.createElement('div');
    matchesContainer.className = 'round-matches';
    matchesContainer.style.marginTop = `${roundOffset}px`;

    roundIds.forEach((mid, mi) => {
      const match   = matches[mid];
      const wrapper = document.createElement('div');
      wrapper.className  = 'match-wrapper';
      wrapper.style.marginBottom = mi < numMatches - 1 ? `${gapBetween}px` : '0';
      wrapper.innerHTML  = renderMatchCard(match);
      matchesContainer.appendChild(wrapper);
    });

    col.appendChild(matchesContainer);
    bracketEl.appendChild(col);

    /* Conectores */
    if (!isLast) {
      const connDiv = document.createElement('div');
      connDiv.className = 'connector-col';
      connDiv.style.marginTop = `${headerH + 12 + roundOffset}px`;
      connDiv.innerHTML = renderConnectors(numMatches, matchH, gapBetween);
      bracketEl.appendChild(connDiv);
    }
  });

  /* Resultado final + línea + campeón, todo centrado sobre el mismo eje */
  const finalOffset = getRoundOffset(rounds.length - 1, matchH, matchGap);
  const finalMatch = matches[rounds[rounds.length - 1][0]];
  const finalOutcome = document.createElement('div');
  finalOutcome.className = 'final-outcome-group';
  finalOutcome.style.marginTop = `${headerH + 12 + finalOffset + (matchH / 2)}px`;
  finalOutcome.innerHTML = `
    <svg class="final-outcome-line" width="96" height="2" aria-hidden="true">
      <line x1="0" y1="1" x2="96" y2="1" stroke="var(--line-color)" stroke-width="1"/>
    </svg>
    ${buildChampionCardHtml(finalMatch.winner || '—')}
  `;
  bracketEl.appendChild(finalOutcome);

  const championArea = document.getElementById('championArea');
  if (championArea) championArea.classList.add('hidden');

  /* Tercer lugar */
  renderTercerLugar();

  /* Eventos */
  attachClickHandlers();
}


function renderTercerLugar() {
  const area = document.getElementById('tercerLugarArea');
  const el   = document.getElementById('tercerBracketEl');

  if (!tercerMatch) {
    area.classList.add('hidden');
    return;
  }

  actualizarTercerosTaekwondo();

  area.classList.remove('hidden');
  el.innerHTML = '';

  const terceros = tercerMatch.thirds || [null, null];

  const podio = document.createElement('div');
  podio.className = 'tercer-podium';
  podio.innerHTML = [0, 1].map(index => {
    const nombre = terceros[index] || 'Por definir';
    const isPending = !terceros[index];
    return `<div class="tercer-card ${isPending ? 'tercer-card-pending' : ''}">
      <div class="tercer-card-label">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
          <circle cx="12" cy="8" r="6"/><path d="M15.477 12.89 17 22l-5-3-5 3 1.523-9.11"/>
        </svg>
        3.° lugar
      </div>
      <div class="tercer-card-name">${esc(nombre)}</div>
      <div class="tercer-card-note">Perdedor de semifinal ${index + 1}</div>
    </div>`;
  }).join('');

  el.appendChild(podio);
}

/* ----------------------------------------------------------
   EVENTOS DE CLIC
   ---------------------------------------------------------- */

function attachClickHandlers() {
  document.querySelectorAll('.match-slot[data-match]').forEach(slot => {
    slot.addEventListener('click', onSlotClick);
  });
}

function onSlotClick(e) {
  const slot    = e.currentTarget;
  const matchId = slot.dataset.match;
  const lado    = slot.dataset.lado;

  /* Buscar en bracket principal o preliminar */
  let match    = matches[matchId] || null;
  let isPrelim = false;

  if (!match && prelimMatches.length > 0) {
    match = prelimMatches.find(item => item.id === matchId) || null;
    isPrelim = !!match;
  }

  if (!match) return;

  const winner = lado === 'A' ? match.slotA : match.slotB;
  if (!winner) return;

  /* Toggle */
  match.winner = (match.winner === winner) ? null : winner;

  if (isPrelim) {
    aplicarGanadorPreliminar(match);
  } else {
    propagate(match);
  }

  renderBracket();
}

/* ----------------------------------------------------------
   GENERACIÓN PRINCIPAL
   ---------------------------------------------------------- */

function generarBracket() {
  const competidores = getCompetidores();
  const n = competidores.length;

  if (n < 2) {
    alert('Ingresa al menos 2 competidores.');
    return;
  }

  const nombre    = document.getElementById('torneoNombre').value.trim();
  const categoria = getCategoriaDisplay();
  const conTercer = document.getElementById('tercerLugar').checked;

  let size = 0;
  let prelimCount = 0;
  let byes = 0;
  let slots = [];

  const modelo = prepararBracketSegunModelos(competidores);

  if (modelo) {
    size = modelo.baseSize;
    prelimCount = modelo.prelimCount;
    byes = Math.max(0, size - n);
    slots = modelo.slots;
  } else {
    prelimMatches = [];
    prelimPlacements = [];
    size = nextPow2(n);
    byes = size - n;
    slots = distribuirCompetidoresConByesBalanceados(competidores);
  }

  /* Construir cuadro principal */
  buildMatches(slots);
  if (prelimCount > 0) {
    conectarPreliminaresAlCuadroPrincipal();
  }

  /* Terceros lugares de Taekwondo: los dos perdedores de semifinal */
  tercerMatch = null;
  if (conTercer && rounds.length >= 2) {
    tercerMatch = {
      id:      'terceros-taekwondo',
      thirds:  [null, null],
    };
    actualizarTercerosTaekwondo();
  }

  /* Mostrar sección */
  document.getElementById('bracketSection').classList.remove('hidden');

  /* Título */
  document.getElementById('bracketTitulo').textContent   = nombre || 'Torneo de Taekwondo';
  document.getElementById('bracketCategoria').textContent = categoria;

  /* Badges */
  document.getElementById('bracketBadges').innerHTML = `
    <span class="badge badge-players">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
      ${n} competidor${n !== 1 ? 'es' : ''}
    </span>
    <span class="badge badge-rounds">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"/></svg>
      ${rounds.length} ronda${rounds.length !== 1 ? 's' : ''}
    </span>
    ${byes > 0 ? `<span class="badge badge-bye">
      <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polygon points="5 3 19 12 5 21 5 3"/></svg>
      ${byes} BYE
    </span>` : ''}`;

  /* Aviso */
  const byeInfoEl = document.getElementById('byeInfo');
  if (byes > 0) {
    byeInfoEl.classList.remove('hidden');
    byeInfoEl.innerHTML = `
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0" aria-hidden="true"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
      Con <strong>${n}</strong> competidores se usa el cuadro oficial de <strong>${size}</strong>.
      Los competidores con <strong>BYE</strong> se muestran desde la primera columna y aparecen automáticamente en la siguiente ronda.`;
  } else {
    byeInfoEl.classList.add('hidden');
  }

  renderBracket();
  document.getElementById('bracketSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ----------------------------------------------------------
   SORTEO ALEATORIO
   ---------------------------------------------------------- */

function sortearYGenerar() {
  const c = getCompetidores();
  if (c.length < 2) {
    alert('Ingresa al menos 2 competidores antes de sortear.');
    return;
  }
  document.getElementById('competidores').value = shuffle(c).join('\n');
  actualizarContador();
  generarBracket();
}

/* ----------------------------------------------------------
   RESET
   ---------------------------------------------------------- */

function resetear() {
  document.getElementById('competidores').value  = '';
  document.getElementById('torneoNombre').value  = '';
  document.getElementById('categoria').value     = '';
  const generoEl = document.getElementById('genero');
  if (generoEl) generoEl.value = '';
  const divisionEl = document.getElementById('division');
  if (divisionEl) divisionEl.value = '';
  refreshExcelDrivenSelects({ keepValues: false });
  document.getElementById('tercerLugar').checked  = false;
  actualizarContador();

  matches     = {};
  rounds      = [];
  prelimMatches = [];
  prelimPlacements = [];
  tercerMatch = null;

  document.getElementById('bracketSection').classList.add('hidden');
  document.getElementById('bracketEl').innerHTML       = '';
  document.getElementById('tercerBracketEl').innerHTML = '';
  document.getElementById('championName').textContent  = '—';
  const championArea = document.getElementById('championArea');
  if (championArea) championArea.classList.add('hidden');
}

/* ----------------------------------------------------------
   EXPORTACIÓN A EXCEL — plantilla .xlsm
   ---------------------------------------------------------- */

const EXCEL_TEMPLATE_URL = 'plantilla_taekwondo.xlsm';
const EXCEL_START_ROW = 4;          // Fila 4: primera fila debajo de encabezados
const EXCEL_MAX_ROWS  = 20;         // La plantilla visible tiene 20 filas útiles (A4:A23)
const EXCEL_NAME_COL     = 'B';     // Columna B: Nombres y apellidos
const EXCEL_NUM_COL      = 'A';     // Columna A: Nº
const EXCEL_AGE_COL      = 'C';     // Columna C: Edad / categoría por edad
const EXCEL_GENDER_COL   = 'D';     // Columna D: Género
const EXCEL_DIVISION_COL = 'E';     // Columna E: Peso / Kup / Poomsae / Grado
const EXCEL_CATEGORY_CODE_COL = 'H'; // Columna H: código interno usado por la lista de Kg

const XLSX_MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const XLSX_REL_NS  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XLSM_MIME    = 'application/vnd.ms-excel.sheet.macroEnabled.12';

function setExcelStatus(msg, type = '') {
  const el = document.getElementById('excelStatus');
  if (!el) return;
  el.textContent = msg || '';
  el.className = `excel-status ${type}`.trim();
}

async function readExcelTemplateBuffer() {
  const input = document.getElementById('excelTemplateFile');

  if (input && input.files && input.files[0]) {
    return input.files[0].arrayBuffer();
  }

  const templateUrls = [
    EXCEL_TEMPLATE_URL,
    'TAEKWONDO - TOPE UNIÓN V2.xlsm',
    'TAEKWONDO - TOPE UNION V2.xlsm',
  ];

  for (const url of templateUrls) {
    try {
      const response = await fetch(encodeURI(url), { cache: 'no-store' });
      if (response.ok) return response.arrayBuffer();
    } catch (err) {
      /* Probar el siguiente nombre de plantilla. */
    }
  }

  throw new Error('No se pudo cargar la plantilla automática. Verifica que el archivo plantilla_taekwondo.xlsm esté en la misma carpeta o selecciónalo manualmente desde el campo de plantilla.');
}

function parseXml(xmlText) {
  const doc = new DOMParser().parseFromString(xmlText, 'application/xml');
  const parserError = doc.getElementsByTagName('parsererror')[0];
  if (parserError) throw new Error('No se pudo leer la estructura XML del Excel.');
  return doc;
}

function getElements(doc, tagName) {
  return Array.from(doc.getElementsByTagNameNS(XLSX_MAIN_NS, tagName));
}

function findWorkbookSheetPath(zip, workbookXml, relsXml, sheetName) {
  const wbDoc = parseXml(workbookXml);
  const relsDoc = parseXml(relsXml);
  const sheets = getElements(wbDoc, 'sheet');
  const sheet = sheets.find(node => node.getAttribute('name') === sheetName);

  if (!sheet) {
    throw new Error(`La hoja "${sheetName}" no existe en la plantilla.`);
  }

  const relId = sheet.getAttribute('r:id') || sheet.getAttributeNS(XLSX_REL_NS, 'id');
  const relationships = Array.from(relsDoc.getElementsByTagName('Relationship'));
  const rel = relationships.find(node => node.getAttribute('Id') === relId);

  if (!rel) {
    throw new Error(`No se encontró la relación interna de la hoja "${sheetName}".`);
  }

  let target = rel.getAttribute('Target') || '';
  target = target.replace(/^\//, '');
  const sheetPath = target.startsWith('xl/') ? target : `xl/${target}`;

  if (!zip.file(sheetPath)) {
    throw new Error(`No se encontró el archivo interno ${sheetPath} dentro de la plantilla.`);
  }

  return sheetPath;
}

function colToNumber(col) {
  return col.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);
}

function getCellRef(col, row) {
  return `${col}${row}`;
}

function escapeXml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function cellRegex(cellRef) {
  return new RegExp(`<c\\b[^>]*\\br="${cellRef}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`);
}

function rowRegex(rowNumber) {
  return new RegExp(`<row\\b[^>]*\\br="${rowNumber}"[^>]*(?:\\/>|>[\\s\\S]*?<\\/row>)`);
}

function getCellXmlFromSheet(sheetXml, cellRef) {
  const match = sheetXml.match(cellRegex(cellRef));
  return match ? match[0] : null;
}

function getStyleIdFromCellXml(cellXml) {
  const match = cellXml ? cellXml.match(/\bs="([^"]+)"/) : null;
  return match ? match[1] : null;
}

function getStyleAttr(styleId) {
  return styleId !== null && styleId !== undefined ? ` s="${styleId}"` : '';
}

function makeBlankCellXml(cellRef, styleId) {
  return `<c r="${cellRef}"${getStyleAttr(styleId)}/>`;
}

function makeNumberCellXml(cellRef, styleId, value) {
  return `<c r="${cellRef}"${getStyleAttr(styleId)}><v>${Number(value)}</v></c>`;
}

function makeInlineStringCellXml(cellRef, styleId, value) {
  const text = escapeXml(value);
  const preserve = /^\s|\s$/.test(String(value)) ? ' xml:space="preserve"' : '';
  return `<c r="${cellRef}"${getStyleAttr(styleId)} t="inlineStr"><is><t${preserve}>${text}</t></is></c>`;
}

function makeFormulaStringCellXml(cellRef, styleId, formula, cachedValue = '') {
  const safeFormula = escapeXml(formula);
  const safeValue = escapeXml(cachedValue || '');
  return `<c r="${cellRef}"${getStyleAttr(styleId)} t="str"><f>${safeFormula}</f><v>${safeValue}</v></c>`;
}

function getRowOpenTagTemplate(sheetXml, sourceRowNumber) {
  const match = sheetXml.match(rowRegex(sourceRowNumber));
  if (!match) return '<row r="__ROW__">';
  const rowXml = match[0];
  const openTagMatch = rowXml.match(/^<row\b[^>]*>/);
  if (!openTagMatch) return '<row r="__ROW__">';

  let openTag = openTagMatch[0]
    .replace(/\br="\d+"/, 'r="__ROW__"')
    .replace(/\bspans="[^"]*"/, 'spans="1:26"');

  if (!/\br="__ROW__"/.test(openTag)) {
    openTag = openTag.replace('<row', '<row r="__ROW__"');
  }

  return openTag;
}

function parseRowCells(rowXml) {
  const openTagMatch = rowXml.match(/^<row\b[^>]*>/);
  const closeTagMatch = rowXml.match(/<\/row>$/);

  if (!openTagMatch || !closeTagMatch) {
    return { openTag: '<row>', closeTag: '</row>', cells: [] };
  }

  const openTag = openTagMatch[0];
  const closeTag = closeTagMatch[0];
  const inner = rowXml.slice(openTag.length, rowXml.length - closeTag.length);
  const cellMatches = inner.match(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g) || [];
  const cells = cellMatches.map(cellXml => {
    const refMatch = cellXml.match(/\br="([A-Z]+\d+)"/);
    return {
      ref: refMatch ? refMatch[1] : '',
      col: refMatch ? refMatch[1].replace(/[0-9]/g, '') : '',
      xml: cellXml,
    };
  }).filter(cell => cell.ref);

  return { openTag, closeTag, cells };
}

function buildRowXml(openTag, closeTag, cells) {
  const sorted = [...cells].sort((a, b) => colToNumber(a.col) - colToNumber(b.col));
  return `${openTag}${sorted.map(cell => cell.xml).join('')}${closeTag}`;
}

function replaceOrInsertCell(rowXml, newCellXml) {
  const cellRefMatch = newCellXml.match(/\br="([A-Z]+\d+)"/);
  if (!cellRefMatch) return rowXml;

  const cellRef = cellRefMatch[1];
  const col = cellRef.replace(/[0-9]/g, '');
  const parsed = parseRowCells(rowXml);
  const idx = parsed.cells.findIndex(cell => cell.ref === cellRef);

  if (idx >= 0) {
    parsed.cells[idx].xml = newCellXml;
    parsed.cells[idx].col = col;
  } else {
    parsed.cells.push({ ref: cellRef, col, xml: newCellXml });
  }

  return buildRowXml(parsed.openTag, parsed.closeTag, parsed.cells);
}

function replaceOrInsertRow(sheetXml, rowNumber, newRowXml) {
  const existingRe = rowRegex(rowNumber);
  if (existingRe.test(sheetXml)) {
    return sheetXml.replace(existingRe, newRowXml);
  }

  const rowMatches = Array.from(sheetXml.matchAll(/<row\b[^>]*\br="(\d+)"[^>]*(?:\/>|>[\s\S]*?<\/row>)/g));
  for (const match of rowMatches) {
    const currentRow = Number(match[1]);
    if (currentRow > rowNumber) {
      return sheetXml.slice(0, match.index) + newRowXml + sheetXml.slice(match.index);
    }
  }

  return sheetXml.replace('</sheetData>', `${newRowXml}</sheetData>`);
}

function clearCellValueInRow(rowXml, cellRef, fallbackStyleId = null) {
  const parsed = parseRowCells(rowXml);
  const idx = parsed.cells.findIndex(cell => cell.ref === cellRef);
  if (idx < 0) return rowXml;

  const existingCellXml = parsed.cells[idx].xml;
  const styleId = getStyleIdFromCellXml(existingCellXml) ?? fallbackStyleId;
  parsed.cells[idx].xml = makeBlankCellXml(cellRef, styleId);
  return buildRowXml(parsed.openTag, parsed.closeTag, parsed.cells);
}

function replaceFormulaCellCachedValue(rowXml, cellRef, value) {
  const parsed = parseRowCells(rowXml);
  const idx = parsed.cells.findIndex(cell => cell.ref === cellRef);
  if (idx < 0) return rowXml;

  let cellXml = parsed.cells[idx].xml;
  const escapedValue = escapeXml(value || 'EMPTY');
  if (/<v>[\s\S]*?<\/v>/.test(cellXml)) {
    cellXml = cellXml.replace(/<v>[\s\S]*?<\/v>/, `<v>${escapedValue}</v>`);
  } else if (/<\/c>$/.test(cellXml)) {
    cellXml = cellXml.replace('</c>', `<v>${escapedValue}</v></c>`);
  }

  parsed.cells[idx].xml = cellXml;
  return buildRowXml(parsed.openTag, parsed.closeTag, parsed.cells);
}

function patchCompetitorRows(sheetXml, competidores, categoriaEdad, genero, division) {
  const endRow = EXCEL_START_ROW + EXCEL_MAX_ROWS - 1;

  const nameStyleId     = getStyleIdFromCellXml(getCellXmlFromSheet(sheetXml, getCellRef(EXCEL_NAME_COL, EXCEL_START_ROW)));
  const ageStyleId      = getStyleIdFromCellXml(getCellXmlFromSheet(sheetXml, getCellRef(EXCEL_AGE_COL, EXCEL_START_ROW)));
  const genderStyleId   = getStyleIdFromCellXml(getCellXmlFromSheet(sheetXml, getCellRef(EXCEL_GENDER_COL, EXCEL_START_ROW)));
  const divisionStyleId = getStyleIdFromCellXml(getCellXmlFromSheet(sheetXml, getCellRef(EXCEL_DIVISION_COL, EXCEL_START_ROW)));
  const codeStyleId     = getStyleIdFromCellXml(getCellXmlFromSheet(sheetXml, getCellRef(EXCEL_CATEGORY_CODE_COL, EXCEL_START_ROW)));
  const keyStyleId      = getStyleIdFromCellXml(getCellXmlFromSheet(sheetXml, getCellRef('I', EXCEL_START_ROW)));
  const isKyorugui      = getSelectedSheetName() === 'KYORUGUI';
  const categoryCode    = isKyorugui ? getCategoriaCodigoExcel(categoriaEdad) : '';

  for (let rowNumber = EXCEL_START_ROW; rowNumber <= endRow; rowNumber++) {
    const rowMatch = sheetXml.match(rowRegex(rowNumber));
    if (!rowMatch) continue;

    let rowXml = rowMatch[0];
    const idx = rowNumber - EXCEL_START_ROW;

    const nameRef     = getCellRef(EXCEL_NAME_COL, rowNumber);
    const ageRef      = getCellRef(EXCEL_AGE_COL, rowNumber);
    const genderRef   = getCellRef(EXCEL_GENDER_COL, rowNumber);
    const divisionRef = getCellRef(EXCEL_DIVISION_COL, rowNumber);
    const codeRef     = getCellRef(EXCEL_CATEGORY_CODE_COL, rowNumber);
    const keyRef      = getCellRef('I', rowNumber);

    /* No se toca la columna A: usa fórmulas compartidas de la plantilla. */
    rowXml = clearCellValueInRow(rowXml, nameRef, nameStyleId);
    rowXml = clearCellValueInRow(rowXml, ageRef, ageStyleId);
    rowXml = clearCellValueInRow(rowXml, genderRef, genderStyleId);
    rowXml = clearCellValueInRow(rowXml, divisionRef, divisionStyleId);

    if (idx < competidores.length) {
      rowXml = replaceOrInsertCell(rowXml, makeInlineStringCellXml(nameRef, nameStyleId, competidores[idx]));

      if (categoriaEdad) {
        rowXml = replaceOrInsertCell(rowXml, makeInlineStringCellXml(ageRef, ageStyleId, categoriaEdad));
      }
      if (genero) {
        rowXml = replaceOrInsertCell(rowXml, makeInlineStringCellXml(genderRef, genderStyleId, genero));
      }
      if (division) {
        rowXml = replaceOrInsertCell(rowXml, makeInlineStringCellXml(divisionRef, divisionStyleId, division));
      }
    }

    if (isKyorugui) {
      const codeFormula = `IFERROR(VLOOKUP(C${rowNumber},$W$2:$X$7,2,FALSE),"")`;
      const keyFormula = `IF(OR($H${rowNumber}="",$D${rowNumber}=""),"EMPTY",$H${rowNumber}&"_"&$D${rowNumber})`;
      const cachedCode = idx < competidores.length ? categoryCode : '';
      const cachedKey = idx < competidores.length && categoryCode && genero ? `${categoryCode}_${genero}` : 'EMPTY';

      rowXml = replaceOrInsertCell(rowXml, makeFormulaStringCellXml(codeRef, codeStyleId, codeFormula, cachedCode));
      rowXml = replaceOrInsertCell(rowXml, makeFormulaStringCellXml(keyRef, keyStyleId, keyFormula, cachedKey));
    }

    sheetXml = replaceOrInsertRow(sheetXml, rowNumber, rowXml);
  }

  return sheetXml;
}

function extendSqrefValue(sqref, endRow) {
  return sqref.split(/\s+/).map(part => {
    const match = part.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) return part;

    const startRow = Number(match[2]);
    const currentEndRow = Number(match[4]);

    if (startRow >= EXCEL_START_ROW && currentEndRow >= 20) {
      return `${match[1]}${startRow}:${match[3]}${Math.max(currentEndRow, endRow)}`;
    }

    return part;
  }).join(' ');
}

function extendDataValidationRanges(sheetXml) {
  const endRow = EXCEL_START_ROW + EXCEL_MAX_ROWS - 1;
  return sheetXml.replace(/(<dataValidation\b[^>]*\bsqref=")([^"]+)(")/g,
    (_, before, sqref, after) => `${before}${extendSqrefValue(sqref, endRow)}${after}`);
}

function updateDimensionXml(sheetXml, endRow) {
  return sheetXml.replace(/<dimension\b[^>]*\bref="([^"]+)"[^>]*\/>/, (full, ref) => {
    const currentEnd = ref.includes(':') ? ref.split(':')[1] : ref;
    const currentEndCol = currentEnd.replace(/[0-9]/g, '') || 'J';
    const currentEndRow = Number(currentEnd.replace(/[A-Z]/gi, '')) || 1;
    const nextRef = `A1:${currentEndCol}${Math.max(currentEndRow, endRow)}`;
    return full.replace(/ref="[^"]+"/, `ref="${nextRef}"`);
  });
}


function setDataValidationsCount(sheetXml) {
  const matches = sheetXml.match(/<dataValidation\b/g);
  if (!matches) return sheetXml;
  return sheetXml.replace(/<dataValidations\b([^>]*)\bcount="\d+"/, `<dataValidations$1count="${matches.length}"`);
}

function ensureKyoruguiWeightValidation(sheetXml) {
  if (getSelectedSheetName() !== 'KYORUGUI') return sheetXml;

  const endRow = EXCEL_START_ROW + EXCEL_MAX_ROWS - 1;
  const formula = 'OFFSET($J$2,0,MATCH($I4,$J$1:$V$1,0)-1,COUNTA(OFFSET($J$2,0,MATCH($I4,$J$1:$V$1,0)-1,30,1)),1)';
  const escapedFormula = formula.replace(/&/g, '&amp;');
  const validationXml = `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E${EXCEL_START_ROW}:E${endRow}"><formula1>${escapedFormula}</formula1></dataValidation>`;

  let nextXml = sheetXml.replace(/<dataValidation\b[^>]*\bsqref="E\d+:E\d+"[^>]*(?:>[\s\S]*?<\/dataValidation>|\/>)?/g, '');

  if (/<dataValidations\b[^>]*>/.test(nextXml)) {
    nextXml = nextXml.replace(/<dataValidations\b[^>]*>/, match => `${match}${validationXml}`);
  } else {
    nextXml = nextXml.replace('<pageMargins', `<dataValidations count="1">${validationXml}</dataValidations><pageMargins`);
  }

  return setDataValidationsCount(nextXml);
}

function ensureKyoruguiKgDropdown(sheetXml) {
  if (getSelectedSheetName() !== 'KYORUGUI') return sheetXml;

  const endRow = EXCEL_START_ROW + EXCEL_MAX_ROWS - 1;
  const formula = 'OFFSET($J$2,0,MATCH($I4,$J$1:$V$1,0)-1,COUNTA(OFFSET($J$2,0,MATCH($I4,$J$1:$V$1,0)-1,30,1)),1)';
  const validationXml = `<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E${EXCEL_START_ROW}:E${endRow}"><formula1>${formula}</formula1></dataValidation>`;

  let nextXml = sheetXml;
  const kgValidationRe = /<dataValidation\b[^>]*\bsqref="E\d+:E\d+"[^>]*(?:>[\s\S]*?<\/dataValidation>|\/>)?/;

  if (kgValidationRe.test(nextXml)) {
    nextXml = nextXml.replace(kgValidationRe, validationXml);
  } else if (/<dataValidations\b[^>]*>/.test(nextXml)) {
    nextXml = nextXml.replace(/<dataValidations\b[^>]*>/, match => `${match}${validationXml}`);
  } else {
    nextXml = nextXml.replace('<pageMargins', `<dataValidations count="1">${validationXml}</dataValidations><pageMargins`);
  }

  return setDataValidationsCount(nextXml);
}

async function removeCalcChain(zip) {
  if (zip.file('xl/calcChain.xml')) {
    zip.remove('xl/calcChain.xml');
  }

  const contentTypesFile = zip.file('[Content_Types].xml');
  if (contentTypesFile) {
    const contentTypesXml = await contentTypesFile.async('text');
    const nextContentTypesXml = contentTypesXml.replace(/<Override\b[^>]*PartName="\/xl\/calcChain\.xml"[^>]*\/>/g, '');
    zip.file('[Content_Types].xml', nextContentTypesXml);
  }

  const relsFile = zip.file('xl/_rels/workbook.xml.rels');
  if (relsFile) {
    const relsXml = await relsFile.async('text');
    const nextRelsXml = relsXml.replace(/<Relationship\b[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/calcChain"[^>]*\/>/g, '');
    zip.file('xl/_rels/workbook.xml.rels', nextRelsXml);
  }
}

function markWorkbookForRecalculation(workbookXml) {
  if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
    return workbookXml.replace(/<calcPr\b([^>]*)\/>/, (full, attrs) => {
      let next = attrs
        .replace(/\bcalcMode="[^"]*"/g, '')
        .replace(/\bfullCalcOnLoad="[^"]*"/g, '')
        .replace(/\bforceFullCalc="[^"]*"/g, '')
        .replace(/\s+/g, ' ')
        .trim();
      next = next ? ` ${next}` : '';
      return `<calcPr${next} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>`;
    });
  }

  return workbookXml.replace('</workbook>', '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/></workbook>');
}

function patchSheetXmlForExport(sheetXml, competidores, categoriaEdad, genero) {
  let nextXml = patchCompetitorRows(sheetXml, competidores.slice(0, EXCEL_MAX_ROWS), categoriaEdad, genero, '');
  nextXml = ensureKyoruguiKgDropdown(nextXml);
  return nextXml;
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

async function exportarCompetidoresExcel() {
  try {
    const competidores = getCompetidores();
    if (competidores.length === 0) {
      alert('Ingresa al menos un competidor para exportar.');
      return;
    }

    if (typeof JSZip === 'undefined') {
      alert('No se cargó la librería JSZip. Revisa tu conexión o descarga la librería de forma local.');
      return;
    }

    const sheetName = document.getElementById('excelHojaDestino').value;
    const categoriaEdad = getCategoriaEdad();
    const genero = getGenero();
    setExcelStatus('Preparando Excel...', 'loading');

    const buffer = await readExcelTemplateBuffer();
    const zip = await JSZip.loadAsync(buffer);
    const workbookXml = await zip.file('xl/workbook.xml').async('text');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('text');
    const sheetPath = findWorkbookSheetPath(zip, workbookXml, relsXml, sheetName);

    const sheetXml = await zip.file(sheetPath).async('text');
    const patchedSheetXml = patchSheetXmlForExport(sheetXml, competidores, categoriaEdad, genero);
    zip.file(sheetPath, patchedSheetXml);
    zip.file('xl/workbook.xml', markWorkbookForRecalculation(workbookXml));
    await removeCalcChain(zip);

    const torneo = document.getElementById('torneoNombre').value.trim() || 'torneo-taekwondo';
    const categoria = getCategoriaDisplay();
    const fileBase = [torneo, categoria, sheetName]
      .filter(Boolean)
      .join(' - ')
      .replace(/[\/:*?"<>|]+/g, '-')
      .replace(/\s+/g, ' ')
      .trim();

    const blob = await zip.generateAsync({
      type: 'blob',
      mimeType: XLSM_MIME,
      compression: 'DEFLATE'
    });
    downloadBlob(blob, `${fileBase}.xlsm`);

    const omitidos = competidores.length > EXCEL_MAX_ROWS
      ? ` Se exportaron solo los primeros ${EXCEL_MAX_ROWS}.`
      : '';
    const detalles = [];
    if (categoriaEdad) detalles.push(`Edad: ${categoriaEdad}`);
    if (genero) detalles.push(`Género: ${genero}`);
    const detailMsg = detalles.length ? ` Se llenó ${detalles.join(' / ')}.` : '';
    setExcelStatus(`Excel generado en la hoja ${sheetName}. Se conservaron las listas desplegables.${detailMsg}${omitidos}`, 'ok');
  } catch (err) {
    console.error(err);
    setExcelStatus(err.message || 'No se pudo exportar el Excel.', 'error');
    alert(err.message || 'No se pudo exportar el Excel.');
  }
}

/* ----------------------------------------------------------
   CONTADOR EN VIVO
   ---------------------------------------------------------- */

function actualizarContador() {
  document.getElementById('contadorNum').textContent = getCompetidores().length;
}

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('competidores').addEventListener('input', actualizarContador);

  const hojaDestinoEl = document.getElementById('excelHojaDestino');
  if (hojaDestinoEl) {
    hojaDestinoEl.addEventListener('change', () => refreshExcelDrivenSelects({ keepValues: false }));
  }

  const categoriaEl = document.getElementById('categoria');
  if (categoriaEl) {
    categoriaEl.addEventListener('change', () => refreshExcelDrivenSelects({ keepValues: true, updateDivisionOnly: true }));
  }

  const generoEl = document.getElementById('genero');
  if (generoEl) {
    generoEl.addEventListener('change', () => refreshExcelDrivenSelects({ keepValues: true, updateDivisionOnly: true }));
  }

  refreshExcelDrivenSelects({ keepValues: false });
  actualizarContador();
});