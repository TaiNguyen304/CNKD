import express from 'express';
import http from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { Server } from 'socket.io';
import xlsx from 'xlsx';
import fs from 'fs';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const server = http.createServer(app);
const io = new Server(server, {
  cors: {
    origin: '*',
    methods: ['GET', 'POST']
  }
});

const PORT = process.env.PORT || 3000;

app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Initial default round questions & multi-topic data (Hoàn toàn không có câu hỏi ô chữ cứng trong code)
const initialRoundsData = {
  'Đoán nhanh 1': [],
  'Đoán nhanh 2': [],
  'Vòng 1': [],
  'Vòng 2': [],
  'Vòng 3': [],
  'Đoán nhanh 3': [],
  'ĐNLH 1': [],
  'ĐNLH 2': [],
  'ĐNLH 3': [],
  'Vòng 4': [],
  'Vòng khán giả': [],
  'Vòng đặc biệt': []
};

// Available wheels
const AVAILABLE_WHEELS = [
  { id: 'vong7.png', label: 'Nón 1 (vong7.png)' },
  { id: 'vong13.png', label: 'Nón 2 (vong13.png)' },
  { id: 'vong14.png', label: 'Nón 3 (vong14.png)' },
  { id: 'vong18.png', label: 'Nón 4 (vong18.png)' },
  { id: 'vong19.png', label: 'Nón 5 (vong19.png)' },
  { id: 'vong23.png', label: 'Nón 6 (vong23.png)' }
];

// Centralized Game State
const gameState = {
  players: [
    { id: 1, name: 'Người chơi 1', roundScore: 0, totalScore: 0, canSpin: false, buzzed: false, buzzTime: null },
    { id: 2, name: 'Người chơi 2', roundScore: 0, totalScore: 0, canSpin: false, buzzed: false, buzzTime: null },
    { id: 3, name: 'Người chơi 3', roundScore: 0, totalScore: 0, canSpin: false, buzzed: false, buzzTime: null }
  ],
  currentRound: 'Vòng 1',
  puzzle: {
    category: '',
    clue: '',
    answer: '',
    gridMatrix: null, // 4x16 array from Excel coordinates
    revealedIndices: [], // array of cell indices 0..63 (r*16 + c)
    markedIndices: [],
    boardState: 'hidden', // 'hidden', 'visible', 'revealed', 'cleared'
    revealedLetters: [],
    showCategory: true,
    showClue: true,
    borderColor: '#800080'
  },
  // Selection mode for Vòng 1-4 and Vòng đặc biệt
  topicSelection: {
    active: false,
    round: null,
    mode: null, // 'choose_topic' or 'fixed'
    topics: []
  },
  revealProgress: {
    active: false,
    intervalMs: 1500
  },
  wheel: {
    selectedWheel: 'vong7.png', // Default wheel as requested
    spinning: false,
    currentAngle: 0,
    startAngle: 0,
    finalAngle: 0,
    startTime: 0,
    duration: 22000, // Exactly 22 seconds as requested
    result: null,
    activeSpinner: null,
    lockedAll: true,
    pointers: {
      p1: { light: true, visible: true }, // Kim 1 (Đỏ - Trái, -30 deg)
      p2: { light: true, visible: true }, // Kim 2 (Vàng - Giữa, 0 deg)
      p3: { light: true, visible: true }  // Kim 3 (Xanh - Phải, +30 deg)
    }
  },
  buzzer: {
    active: false,
    enabled: false,
    winner: null,
    timestamp: null
  },
  roundsData: JSON.parse(JSON.stringify(initialRoundsData))
};

// Try loading Puzzle.xlsx if present to initialize roundsData
function loadPuzzlesFromExcelFile(filePath) {
  try {
    if (!fs.existsSync(filePath)) return false;
    const wb = xlsx.readFile(filePath);
    const parsed = parseWorkbookData(wb);
    if (Object.keys(parsed).length > 0) {
      gameState.roundsData = { ...gameState.roundsData, ...parsed };
      console.log('Successfully pre-loaded questions from', filePath);
      return true;
    }
  } catch (err) {
    console.error('Error reading Puzzle.xlsx:', err.message);
  }
  return false;
}

loadPuzzlesFromExcelFile(path.join(__dirname, 'Puzzle.xlsx'));

function parseWorkbookData(workbook) {
  const result = {};
  if (!workbook || !workbook.SheetNames) return result;

  workbook.SheetNames.forEach(sheetName => {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws['!ref']) return;

    const range = xlsx.utils.decode_range(ws['!ref']);
    
    // Tìm tất cả các dòng header có chữ 'Chủ đề' ở Cột A (hoặc bất kỳ cột nào)
    const headerRows = [];
    for (let r = range.s.r; r <= range.e.r; r++) {
      const cellA = ws[xlsx.utils.encode_cell({ r, c: 0 })];
      const valA = cellA && cellA.v !== undefined ? String(cellA.v).trim().toLowerCase() : '';
      if (valA === 'chủ đề' || valA === 'chu de' || valA.startsWith('chủ đề') || valA.startsWith('chu de')) {
        headerRows.push(r);
      }
    }

    if (headerRows.length === 0) {
      for (let r = 0; r <= range.e.r; r += 6) {
        headerRows.push(r);
      }
    }

    const blocks = [];

    headerRows.forEach((headerR, bIdx) => {
      // Dòng headerR: Cột A = "Chủ đề", Cột B = "Ô chữ" (chỉ là tên cột)
      
      // Dòng headerR + 1: Cột A = Tên chủ đề
      const catCell = ws[xlsx.utils.encode_cell({ r: headerR + 1, c: 0 })];
      let category = catCell && catCell.v !== undefined ? String(catCell.v).trim() : '';
      
      // Dòng headerR + 2: Cột A = Gợi ý / Câu hỏi
      const clueCell = ws[xlsx.utils.encode_cell({ r: headerR + 2, c: 0 })];
      let clue = clueCell && clueCell.v !== undefined ? String(clueCell.v).trim() : '';

      // Trích xuất chính xác ma trận 4 hàng x 16 cột (B..Q tương ứng c = 1..16)
      // Dòng 1 trên sân khấu = headerR + 1 trong Excel (ngang hàng với tên Chủ đề)
      // Dòng 2 trên sân khấu = headerR + 2 trong Excel (ngang hàng với Gợi ý/Câu hỏi)
      // Dòng 3 trên sân khấu = headerR + 3 trong Excel
      // Dòng 4 trên sân khấu = headerR + 4 trong Excel
      let gridMatrix = Array.from({ length: 4 }, () => Array(16).fill(''));
      let fullWords = [];
      let cellFoundCount = 0;

      for (let rowIdx = 0; rowIdx < 4; rowIdx++) {
        const r = headerR + 1 + rowIdx;
        let curWord = '';
        for (let colIdx = 0; colIdx < 16; colIdx++) {
          const c = colIdx + 1; // Col B (1) -> Col Q (16)
          const cell = ws[xlsx.utils.encode_cell({ r, c })];
          const val = cell && cell.v !== undefined ? String(cell.v).trim().toUpperCase() : '';
          gridMatrix[rowIdx][colIdx] = val;
          if (val) {
            cellFoundCount++;
            curWord += val;
          } else {
            if (curWord) {
              fullWords.push(curWord);
              curWord = '';
            }
          }
        }
        if (curWord) {
          fullWords.push(curWord);
        }
      }

      const answer = fullWords.join(' ').replace(/\s+/g, ' ').trim().toUpperCase();

      if (!category) {
        category = `Chủ đề ${bIdx + 1}`;
      }

      if (answer || category || cellFoundCount > 0) {
        const gridLines = gridMatrix.map(row => row.join('').trim()).filter(l => l.length > 0);
        blocks.push({
          blockIndex: bIdx + 1,
          category: category.toUpperCase(),
          clue: clue || `Gợi ý cho chủ đề ${category}`,
          answer: answer || 'CHIẾC NÓN KỲ DIỆU',
          gridMatrix,
          gridLines
        });
      }
    });

    if (blocks.length > 0) {
      result[sheetName] = blocks;
      // Tự động alias các tên vòng chơi
      if (sheetName === 'Khán giả') result['Vòng khán giả'] = blocks;
      if (sheetName === 'Vòng khán giả') result['Khán giả'] = blocks;
      if (sheetName === 'Đặc biệt') result['Vòng đặc biệt'] = blocks;
      if (sheetName === 'Vòng đặc biệt') result['Đặc biệt'] = blocks;
      if (sheetName === 'ĐNLH') {
        result['Đoán nhanh 3'] = blocks;
        if (blocks[0]) result['ĐNLH 1'] = [blocks[0]];
        if (blocks[1]) result['ĐNLH 2'] = [blocks[1]];
        if (blocks[2]) result['ĐNLH 3'] = [blocks[2]];
      }
    }
  });

  return result;
}

// Timer for automatic letter revealing
let revealIntervalTimer = null;

function stopAutoReveal() {
  if (revealIntervalTimer) {
    clearInterval(revealIntervalTimer);
    revealIntervalTimer = null;
  }
  gameState.revealProgress.active = false;
}

function removeToneMarks(str) {
  if (!str) return '';
  const toneMap = {
    'Á':'A','À':'A','Ả':'A','Ã':'A','Ạ':'A',
    'Ắ':'Ă','Ằ':'Ă','Ẳ':'Ă','Ẵ':'Ă','Ặ':'Ă',
    'Ấ':'Â','Ầ':'Â','Ẩ':'Â','Ẫ':'Â','Ậ':'Â',
    'É':'E','È':'E','Ẻ':'E','Ẽ':'E','Ẹ':'E',
    'Ế':'Ê','Ề':'Ê','Ể':'Ê','Ễ':'Ê','Ệ':'Ê',
    'Í':'I','Ì':'I','Ỉ':'I','Ĩ':'I','Ị':'I',
    'Ó':'O','Ò':'O','Ỏ':'O','Õ':'O','Ọ':'O',
    'Ố':'Ô','Ồ':'Ô','Ổ':'Ô','Ỗ':'Ô','Ộ':'Ô',
    'Ớ':'Ơ','Ờ':'Ơ','Ở':'Ơ','Ỡ':'Ơ','Ợ':'Ơ',
    'Ú':'U','Ù':'U','Ủ':'U','Ũ':'U','Ụ':'U',
    'Ứ':'Ư','Ừ':'Ư','Ử':'Ư','Ữ':'Ư','Ự':'Ư',
    'Ý':'Y','Ỳ':'Y','Ỷ':'Y','Ỹ':'Y','Ỵ':'Y'
  };
  return str.split('').map(c => toneMap[c.toUpperCase()] || c.toUpperCase()).join('');
}

function buildGridMatrixFromAnswer(answer) {
  const grid = Array.from({ length: 4 }, () => Array(16).fill(''));
  if (!answer) return grid;

  const words = answer.trim().toUpperCase().split(/\s+/).filter(w => w.length > 0);
  const rows = [[], [], [], []];
  let curRow = 0;
  let curLen = 0;

  for (let w = 0; w < words.length; w++) {
    const word = words[w];
    const needed = curLen === 0 ? word.length : word.length + 1;
    if (curLen + needed <= 16) {
      rows[curRow].push(word);
      curLen += needed;
    } else {
      curRow++;
      if (curRow >= 4) {
        curRow = 3;
        rows[curRow].push(word);
      } else {
        rows[curRow].push(word);
        curLen = word.length;
      }
    }
  }

  const activeRows = rows.filter(r => r.length > 0);
  const usedCount = activeRows.length;
  let startRow = 0;
  if (usedCount === 1) startRow = 1;
  else if (usedCount === 2) startRow = 1;
  else if (usedCount === 3) startRow = 0;

  for (let r = 0; r < usedCount; r++) {
    const targetRowIdx = startRow + r;
    if (targetRowIdx >= 4) break;
    const textInRow = rows[r].join(' ');
    const leftPad = Math.max(0, Math.floor((16 - textInRow.length) / 2));
    let charIdx = 0;
    for (let c = 0; c < 16; c++) {
      if (c >= leftPad && charIdx < textInRow.length) {
        const ch = textInRow[charIdx];
        charIdx++;
        if (ch !== ' ') {
          grid[targetRowIdx][c] = ch;
        }
      }
    }
  }
  return grid;
}

function getActivePuzzleMatrix() {
  if (gameState.puzzle.gridMatrix && Array.isArray(gameState.puzzle.gridMatrix) && gameState.puzzle.gridMatrix.length === 4) {
    return gameState.puzzle.gridMatrix;
  }
  return buildGridMatrixFromAnswer(gameState.puzzle.answer);
}

function getUnrevealedCharIndices() {
  const matrix = getActivePuzzleMatrix();
  const unrevealed = [];
  for (let r = 0; r < 4; r++) {
    for (let c = 0; c < 16; c++) {
      const ch = (matrix[r] && matrix[r][c]) ? matrix[r][c] : '';
      if (ch && ch.trim()) {
        const cellIdx = r * 16 + c;
        if (!gameState.puzzle.revealedIndices.includes(cellIdx)) {
          unrevealed.push(cellIdx);
        }
      }
    }
  }
  return unrevealed;
}

function ensureWheelStateCurrent() {
  if (gameState.wheel.spinning && gameState.wheel.startTime) {
    const elapsed = Date.now() - gameState.wheel.startTime;
    if (elapsed >= (gameState.wheel.duration || 22000)) {
      gameState.wheel.spinning = false;
      gameState.wheel.lockedAll = true;
      gameState.wheel.activeSpinner = null;
    }
  }
}

function broadcastState(eventName = 'game:state') {
  ensureWheelStateCurrent();
  io.emit(eventName, gameState);
}

// Socket.IO logic
io.on('connection', (socket) => {
  ensureWheelStateCurrent();
  socket.emit('game:state', gameState);

  socket.on('state:request', () => {
    ensureWheelStateCurrent();
    socket.emit('game:state', gameState);
  });

  // Player controls (3 rows)
  socket.on('player:set', (data) => {
    const { playerId, name, roundScore } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      if (typeof name === 'string' && name.trim()) {
        player.name = name.trim();
      }
      if (!isNaN(roundScore) && roundScore !== null && roundScore !== '') {
        player.roundScore = Number(roundScore);
      }
      broadcastState();
    }
  });

  socket.on('player:add_round', (data) => {
    const { playerId, amount } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      const val = Number(amount) || 0;
      player.roundScore += val;
      broadcastState();
    }
  });

  socket.on('player:subtract_round', (data) => {
    const { playerId, amount } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      const val = Number(amount) || 0;
      player.roundScore = Math.max(0, player.roundScore - val);
      broadcastState();
    }
  });

  socket.on('player:reset_round', (data) => {
    const { playerId } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      player.roundScore = 0;
      broadcastState();
    }
  });

  socket.on('player:multiply_round', (data) => {
    const { playerId, factor } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      if (factor === 2) {
        player.roundScore = player.roundScore * 2;
      } else if (factor === 0.5) {
        player.roundScore = Math.floor(player.roundScore / 2);
      }
      broadcastState();
    }
  });

  socket.on('player:set_total', (data) => {
    const { playerId, totalScore } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      if (!isNaN(totalScore) && totalScore !== null && totalScore !== '') {
        player.totalScore = Number(totalScore);
      }
      broadcastState();
    }
  });

  socket.on('player:add_to_total', (data) => {
    const { playerId, resetRound } = data;
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player) {
      player.totalScore += player.roundScore;
      if (resetRound !== false) {
        player.roundScore = 0;
      }
      broadcastState();
    }
  });

  // Direct Round selection (for quick guess and non-modal rounds)
  socket.on('round:set', (roundName) => {
    if (typeof roundName === 'string') {
      gameState.currentRound = roundName;
      gameState.buzzer.active = true;
      gameState.buzzer.winner = null;
      gameState.players.forEach(p => {
        p.buzzed = false;
        p.buzzTime = null;
      });

      // Load matching round default puzzle if exists
      const roundTopics = gameState.roundsData[roundName];
      if (roundTopics && roundTopics.length > 0) {
        const first = roundTopics[0];
        gameState.puzzle.category = first.category;
        gameState.puzzle.clue = first.clue || '';
        gameState.puzzle.answer = first.answer;
        gameState.puzzle.gridMatrix = first.gridMatrix || buildGridMatrixFromAnswer(first.answer);
        gameState.puzzle.revealedIndices = [];
        gameState.puzzle.markedIndices = [];
        gameState.puzzle.revealedLetters = [];
        gameState.puzzle.boardState = 'visible';
      }

      broadcastState();
      io.emit('sound:play', { type: 'round_change', round: roundName });
    }
  });

  // Open Topic Selection Popup state (for Vòng 1-4 and Vòng đặc biệt)
  socket.on('round:open_selection_flow', (roundName) => {
    gameState.currentRound = roundName;
    const roundTopics = gameState.roundsData[roundName] || [];
    gameState.topicSelection = {
      active: true,
      round: roundName,
      mode: null,
      topics: roundTopics.slice(0, 3)
    };
    // Board is hidden until user chooses fixed or a specific topic
    gameState.puzzle.boardState = 'hidden';
    gameState.buzzer.winner = null;
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastState();
    io.emit('sound:play', { type: 'round_change', round: roundName });
  });

  // Choice 1: "Ô chữ cố định" -> "ô chữ trống hiện ra"
  socket.on('round:choose_fixed', (roundName) => {
    const targetRound = roundName || gameState.currentRound;
    gameState.currentRound = targetRound;
    const roundTopics = gameState.roundsData[targetRound] || [];
    const chosen = roundTopics[0] || {
      category: 'CHỦ ĐỀ CỐ ĐỊNH',
      clue: '',
      answer: 'CHIẾC NÓN KỲ DIỆU'
    };

    gameState.puzzle.category = chosen.category;
    gameState.puzzle.clue = chosen.clue || '';
    gameState.puzzle.answer = (chosen.answer || '').toUpperCase().trim();
    gameState.puzzle.gridMatrix = chosen.gridMatrix || buildGridMatrixFromAnswer(gameState.puzzle.answer);
    gameState.puzzle.revealedIndices = [];
    gameState.puzzle.markedIndices = [];
    gameState.puzzle.revealedLetters = [];
    // Blank boxes immediately appear!
    gameState.puzzle.boardState = 'visible';
    gameState.topicSelection.active = false;
    stopAutoReveal();

    broadcastState();
    io.emit('sound:play', { type: 'puzzle_show' });
  });

  // Choice 2: "Chọn 1 trong 3 chủ đề" -> shows the 3 topics, boxes do NOT appear yet
  socket.on('round:start_topic_selection', (roundName) => {
    const targetRound = roundName || gameState.currentRound;
    gameState.currentRound = targetRound;
    const roundTopics = gameState.roundsData[targetRound] || [];

    gameState.topicSelection = {
      active: true,
      round: targetRound,
      mode: 'choose_topic',
      topics: roundTopics.slice(0, 3)
    };
    // Must choose topic before puzzle boxes appear
    gameState.puzzle.boardState = 'hidden';
    stopAutoReveal();

    broadcastState();
  });

  // When a topic among the 3 is selected: "phải chọn chủ đề thì ô chữ mới hiện ra"
  socket.on('round:pick_topic', (data) => {
    const { round, topicIndex } = data;
    const targetRound = round || gameState.currentRound;
    const roundTopics = gameState.roundsData[targetRound] || [];
    const chosen = roundTopics[topicIndex] || roundTopics[0];

    if (chosen) {
      gameState.puzzle.category = chosen.category;
      gameState.puzzle.clue = chosen.clue || '';
      gameState.puzzle.answer = (chosen.answer || '').toUpperCase().trim();
      gameState.puzzle.gridMatrix = chosen.gridMatrix || buildGridMatrixFromAnswer(gameState.puzzle.answer);
      gameState.puzzle.revealedIndices = [];
      gameState.puzzle.markedIndices = [];
      gameState.puzzle.revealedLetters = [];
      // Blank boxes now appear!
      gameState.puzzle.boardState = 'visible';
      gameState.topicSelection.active = false;
      stopAutoReveal();

      broadcastState();
      io.emit('sound:play', { type: 'puzzle_show' });
    }
  });

  // Wheel selection (Nón 1 to Nón 6)
  socket.on('wheel:select', (wheelId) => {
    const valid = AVAILABLE_WHEELS.find(w => w.id === wheelId);
    if (valid) {
      gameState.wheel.selectedWheel = wheelId;
      broadcastState();
      io.emit('sound:play', { type: 'wheel_changed', wheel: wheelId });
    }
  });

  let wheelSpinTimeout = null;

  // Wheel Spin (22 seconds per turn! Đồng bộ góc dừng chính xác trên tất cả các máy)
  socket.on('wheel:spin', (data = {}) => {
    const currentAngle = gameState.wheel.currentAngle || 0;
    // Góc quay: 18 vòng đầy đủ (6480 độ) + góc ngẫu nhiên [0..359]
    const randomWedgeDeg = Math.floor(Math.random() * 360);
    const deltaAngle = (data.targetAngle !== undefined) ? data.targetAngle : (360 * 18 + randomWedgeDeg);
    const finalAngle = currentAngle + deltaAngle;
    const duration = 22000; // Strictly 22 seconds
    const startTime = Date.now();

    gameState.wheel.spinning = true;
    gameState.wheel.startAngle = currentAngle;
    gameState.wheel.finalAngle = finalAngle;
    gameState.wheel.currentAngle = finalAngle; // Authoritative destination
    gameState.wheel.duration = duration;
    gameState.wheel.startTime = startTime;
    gameState.wheel.activeSpinner = data.spinner || 'controller';

    if (wheelSpinTimeout) clearTimeout(wheelSpinTimeout);
    wheelSpinTimeout = setTimeout(() => {
      if (gameState.wheel.spinning) {
        gameState.wheel.spinning = false;
        gameState.players.forEach(p => p.canSpin = false);
        gameState.wheel.lockedAll = true;
        gameState.wheel.activeSpinner = null;
        broadcastState();
      }
    }, duration + 300);

    io.emit('wheel:start_spin', {
      startAngle: gameState.wheel.startAngle,
      finalAngle: gameState.wheel.finalAngle,
      duration,
      startTime,
      spinner: gameState.wheel.activeSpinner,
      selectedWheel: gameState.wheel.selectedWheel
    });
    broadcastState();
    io.emit('sound:play', { type: 'spin_start', duration });
  });

  socket.on('wheel:landed', (result) => {
    if (wheelSpinTimeout) {
      clearTimeout(wheelSpinTimeout);
      wheelSpinTimeout = null;
    }
    if (gameState.wheel.spinning) {
      gameState.wheel.spinning = false;
      gameState.wheel.result = result || {};
      gameState.players.forEach(p => p.canSpin = false);
      gameState.wheel.lockedAll = true;
      gameState.wheel.activeSpinner = null;
      broadcastState();
      io.emit('sound:play', { type: 'wheel_result', result });
    }
  });

  socket.on('wheel:allow_player', (playerId) => {
    if (wheelSpinTimeout) {
      clearTimeout(wheelSpinTimeout);
      wheelSpinTimeout = null;
    }
    gameState.wheel.spinning = false;
    gameState.wheel.lockedAll = false;
    const targetId = Number(playerId);
    gameState.players.forEach(p => {
      p.canSpin = (p.id === targetId);
    });
    broadcastState();
    io.emit('player:spin_enabled', { playerId: targetId });
    io.emit('sound:play', { type: 'turn_granted', playerId: targetId });
  });

  socket.on('wheel:lock_all', () => {
    if (wheelSpinTimeout) {
      clearTimeout(wheelSpinTimeout);
      wheelSpinTimeout = null;
    }
    gameState.wheel.lockedAll = true;
    gameState.players.forEach(p => {
      p.canSpin = false;
    });
    broadcastState();
    io.emit('player:spin_disabled', {});
  });

  // Pointer Lights & Visibility Controls
  socket.on('wheel:toggle_pointer_light', (pointer) => {
    if (gameState.wheel.pointers && gameState.wheel.pointers[pointer]) {
      gameState.wheel.pointers[pointer].light = !gameState.wheel.pointers[pointer].light;
      broadcastState();
    }
  });

  socket.on('wheel:set_pointer_light', ({ pointer, light }) => {
    if (gameState.wheel.pointers && gameState.wheel.pointers[pointer]) {
      gameState.wheel.pointers[pointer].light = !!light;
      broadcastState();
    }
  });

  socket.on('wheel:set_all_pointer_lights', ({ light }) => {
    if (gameState.wheel.pointers) {
      ['p1', 'p2', 'p3'].forEach(k => {
        if (gameState.wheel.pointers[k]) gameState.wheel.pointers[k].light = !!light;
      });
      broadcastState();
    }
  });

  socket.on('wheel:toggle_pointer_visibility', (pointer) => {
    if (gameState.wheel.pointers && gameState.wheel.pointers[pointer]) {
      gameState.wheel.pointers[pointer].visible = !gameState.wheel.pointers[pointer].visible;
      broadcastState();
    }
  });

  socket.on('wheel:set_pointer_visibility', ({ pointer, visible }) => {
    if (gameState.wheel.pointers && gameState.wheel.pointers[pointer]) {
      gameState.wheel.pointers[pointer].visible = !!visible;
      broadcastState();
    }
  });

  // Puzzle Actions
  socket.on('puzzle:set', (data) => {
    const { category, clue, answer, gridMatrix } = data;
    if (answer || gridMatrix) {
      gameState.puzzle.category = category || 'CHỦ ĐỀ';
      gameState.puzzle.clue = clue || '';
      gameState.puzzle.answer = (answer || '').toUpperCase().trim();
      gameState.puzzle.gridMatrix = gridMatrix || buildGridMatrixFromAnswer(gameState.puzzle.answer);
      gameState.puzzle.revealedIndices = [];
      gameState.puzzle.markedIndices = [];
      gameState.puzzle.revealedLetters = [];
      gameState.puzzle.boardState = 'hidden';
      gameState.topicSelection.active = false;
      stopAutoReveal();
      broadcastState();
    }
  });

  socket.on('puzzle:show', () => {
    gameState.puzzle.boardState = 'visible';
    gameState.puzzle.revealedIndices = [];
    gameState.puzzle.markedIndices = [];
    gameState.puzzle.revealedLetters = [];
    gameState.topicSelection.active = false;
    stopAutoReveal();
    broadcastState();
    io.emit('sound:play', { type: 'puzzle_show' });
  });

  socket.on('puzzle:solve', () => {
    const matrix = getActivePuzzleMatrix();
    const allIndices = [];
    for (let r = 0; r < 4; r++) {
      for (let c = 0; c < 16; c++) {
        const ch = (matrix[r] && matrix[r][c]) ? matrix[r][c] : '';
        if (ch && ch.trim()) {
          allIndices.push(r * 16 + c);
        }
      }
    }
    gameState.puzzle.revealedIndices = allIndices;
    gameState.puzzle.markedIndices = [];
    gameState.puzzle.boardState = 'revealed';
    stopAutoReveal();
    broadcastState();
    io.emit('sound:play', { type: 'puzzle_solve' });
  });

  socket.on('puzzle:reset', () => {
    gameState.puzzle.boardState = 'cleared';
    gameState.puzzle.revealedIndices = [];
    gameState.puzzle.markedIndices = [];
    gameState.puzzle.revealedLetters = [];
    gameState.puzzle.borderColor = '#800080';
    gameState.topicSelection.active = false;
    stopAutoReveal();
    broadcastState();
    io.emit('sound:play', { type: 'puzzle_reset' });
  });

  // Ẩn/Hiện thanh chủ đề & Ẩn/Hiện câu hỏi
  socket.on('puzzle:toggle_category', () => {
    gameState.puzzle.showCategory = !(gameState.puzzle.showCategory !== false);
    broadcastState();
  });

  socket.on('puzzle:toggle_clue', () => {
    gameState.puzzle.showClue = !(gameState.puzzle.showClue !== false);
    broadcastState();
  });

  socket.on('puzzle:set_category_visibility', (show) => {
    gameState.puzzle.showCategory = !!show;
    broadcastState();
  });

  socket.on('puzzle:set_clue_visibility', (show) => {
    gameState.puzzle.showClue = !!show;
    broadcastState();
  });

  // Đổi màu viền ô chữ theo yêu cầu Controller (Tím #800080, Đỏ #ff0000, Vàng #ffff00, Xanh #0000ff, Xanh lá #00ff00, Trắng #ffffff)
  socket.on('puzzle:set_border_color', (color) => {
    if (color) {
      gameState.puzzle.borderColor = color;
      broadcastState();
    }
  });

  // 1. Check letter occurrences (Dò xem có bao nhiêu ô chứa chữ cái tương ứng theo toạ độ 4x16)
  socket.on('puzzle:check_letter', (letter) => {
    if (!letter) return;
    const targetChar = removeToneMarks(letter.toUpperCase());
    const matrix = getActivePuzzleMatrix();
    const matchingIndices = [];
    for (let r = 0; r < 4; r++) {
      for (let c = 0; c < 16; c++) {
        const ch = (matrix[r] && matrix[r][c]) ? matrix[r][c] : '';
        if (ch && ch.trim()) {
          const stripped = removeToneMarks(ch);
          const cellIdx = r * 16 + c;
          if (stripped === targetChar && !gameState.puzzle.revealedIndices.includes(cellIdx)) {
            matchingIndices.push(cellIdx);
          }
        }
      }
    }
    socket.emit('puzzle:check_result', {
      letter: targetChar,
      count: matchingIndices.length,
      indices: matchingIndices
    });
    io.emit('sound:play', {
      type: matchingIndices.length > 0 ? 'letter_correct' : 'letter_wrong',
      count: matchingIndices.length,
      letter: targetChar
    });
  });

  // 2. Mở lần 1 (Đánh dấu): Các ô trống sẽ được đánh dấu (màu blue)
  socket.on('puzzle:mark_letter', (letter) => {
    if (!letter) return;
    const targetChar = removeToneMarks(letter.toUpperCase());
    const matrix = getActivePuzzleMatrix();
    let count = 0;
    for (let r = 0; r < 4; r++) {
      for (let c = 0; c < 16; c++) {
        const ch = (matrix[r] && matrix[r][c]) ? matrix[r][c] : '';
        if (ch && ch.trim()) {
          const stripped = removeToneMarks(ch);
          const cellIdx = r * 16 + c;
          if (stripped === targetChar && !gameState.puzzle.revealedIndices.includes(cellIdx)) {
            if (!gameState.puzzle.markedIndices.includes(cellIdx)) {
              gameState.puzzle.markedIndices.push(cellIdx);
              count++;
            }
          }
        }
      }
    }
    broadcastState();
    io.emit('sound:play', { type: count > 0 ? 'letter_correct' : 'letter_wrong', count, letter: targetChar });
  });

  // 3. Mở lần 2 (Mở chữ): Các ô đánh dấu được mở (không hiện dấu thanh hỏi, sắc, huyền, ngã, nặng)
  socket.on('puzzle:reveal_marked', (letter) => {
    if (gameState.puzzle.markedIndices && gameState.puzzle.markedIndices.length > 0) {
      gameState.puzzle.markedIndices.forEach(idx => {
        if (!gameState.puzzle.revealedIndices.includes(idx)) {
          gameState.puzzle.revealedIndices.push(idx);
        }
      });
      gameState.puzzle.markedIndices = [];
    }
    if (letter) {
      const char = removeToneMarks(letter.toUpperCase());
      if (!gameState.puzzle.revealedLetters.includes(char)) {
        gameState.puzzle.revealedLetters.push(char);
      }
    }
    broadcastState();
    io.emit('sound:play', { type: 'letter_correct' });
  });

  socket.on('puzzle:clear_marked', () => {
    gameState.puzzle.markedIndices = [];
    broadcastState();
  });

  socket.on('puzzle:reveal_single_letter', (letter) => {
    if (!letter) return;
    const targetChar = removeToneMarks(letter.toUpperCase());
    const matrix = getActivePuzzleMatrix();
    let count = 0;

    for (let r = 0; r < 4; r++) {
      for (let c = 0; c < 16; c++) {
        const ch = (matrix[r] && matrix[r][c]) ? matrix[r][c] : '';
        if (ch && ch.trim()) {
          const stripped = removeToneMarks(ch);
          const cellIdx = r * 16 + c;
          if (stripped === targetChar) {
            if (!gameState.puzzle.revealedIndices.includes(cellIdx)) {
              gameState.puzzle.revealedIndices.push(cellIdx);
              count++;
            }
            const mIdx = gameState.puzzle.markedIndices.indexOf(cellIdx);
            if (mIdx !== -1) {
              gameState.puzzle.markedIndices.splice(mIdx, 1);
            }
          }
        }
      }
    }
    if (!gameState.puzzle.revealedLetters.includes(targetChar)) {
      gameState.puzzle.revealedLetters.push(targetChar);
    }
    if (gameState.puzzle.boardState !== 'visible' && gameState.puzzle.boardState !== 'revealed') {
      gameState.puzzle.boardState = 'visible';
    }
    broadcastState();
    io.emit('sound:play', { type: count > 0 ? 'letter_correct' : 'letter_wrong', count, letter: targetChar });
  });

  // Nút bắt đầu hiện chữ: Các ô trống lần lượt hiện chữ, đồng thời chuông của cả 3 người chơi được mở.
  socket.on('puzzle:start_reveal', (options = {}) => {
    if (gameState.puzzle.boardState !== 'visible') {
      gameState.puzzle.boardState = 'visible';
    }
    stopAutoReveal();
    gameState.revealProgress.active = true;
    gameState.revealProgress.intervalMs = options.intervalMs || 1500;

    // Mở chuông của cả 3 người chơi
    gameState.buzzer.enabled = true;
    gameState.buzzer.winner = null;
    gameState.buzzer.timestamp = null;
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });

    revealIntervalTimer = setInterval(() => {
      const unrevealed = getUnrevealedCharIndices();
      if (unrevealed.length === 0) {
        stopAutoReveal();
        gameState.buzzer.enabled = false;
        gameState.puzzle.boardState = 'revealed';
        broadcastState();
        io.emit('sound:play', { type: 'puzzle_solve' });
        return;
      }
      const nextIndex = unrevealed[0];
      gameState.puzzle.revealedIndices.push(nextIndex);
      broadcastState();
      io.emit('sound:play', { type: 'letter_flip' });
    }, gameState.revealProgress.intervalMs);

    broadcastState();
  });

  // Dừng hiện chữ: Ngưng việc lật chữ lại.
  socket.on('puzzle:stop_reveal', () => {
    stopAutoReveal();
    gameState.buzzer.enabled = false;
    broadcastState();
  });

  // Tiếp tục hiện chữ: Các ô trống còn lại tiếp tục lần lượt hiện chữ, đồng thời chuông của cả 3 người chơi được mở, khung bảng trở về màu tím (#800080).
  socket.on('puzzle:resume_reveal', () => {
    if (gameState.puzzle.boardState !== 'visible') {
      gameState.puzzle.boardState = 'visible';
    }
    stopAutoReveal();
    gameState.revealProgress.active = true;

    // Khung bảng ô chữ trở về màu tím (#800080)
    gameState.puzzle.borderColor = '#800080';

    // Chuông của cả 3 người chơi được mở
    gameState.buzzer.enabled = true;
    gameState.buzzer.winner = null;
    gameState.buzzer.timestamp = null;
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });

    revealIntervalTimer = setInterval(() => {
      const unrevealed = getUnrevealedCharIndices();
      if (unrevealed.length === 0) {
        stopAutoReveal();
        gameState.buzzer.enabled = false;
        gameState.puzzle.boardState = 'revealed';
        broadcastState();
        io.emit('sound:play', { type: 'puzzle_solve' });
        return;
      }
      const nextIndex = unrevealed[0];
      gameState.puzzle.revealedIndices.push(nextIndex);
      broadcastState();
      io.emit('sound:play', { type: 'letter_flip' });
    }, gameState.revealProgress.intervalMs || 1500);

    broadcastState();
  });

  // Buzzer Controls
  // Nếu có 1 người bấm: tương đương với việc bấm nút Dừng hiện chữ và chuông của 2 người còn lại không bấm được. Khung bảng ô chữ đổi sang màu của người bấm chuông (P1: Đỏ #ff0000, P2: Vàng #ffff00, P3: Xanh #0000ff).
  socket.on('buzzer:press', (playerId) => {
    const player = gameState.players.find(p => p.id === Number(playerId));
    if (player && gameState.buzzer.enabled && !gameState.buzzer.winner) {
      gameState.buzzer.winner = player.id;
      gameState.buzzer.timestamp = Date.now();
      gameState.buzzer.enabled = false; // Chuông của 2 người còn lại KHÔNG bấm được!
      player.buzzed = true;
      player.buzzTime = new Date().toLocaleTimeString();

      // Khi Player bấm chuông, khung bảng ô chữ cũng đổi sang màu tương ứng
      if (player.id === 1) {
        gameState.puzzle.borderColor = '#ff0000'; // Đỏ
      } else if (player.id === 2) {
        gameState.puzzle.borderColor = '#ffff00'; // Vàng
      } else if (player.id === 3) {
        gameState.puzzle.borderColor = '#0000ff'; // Xanh
      }

      // Tương đương với việc bấm nút Dừng hiện chữ
      if (gameState.revealProgress.active) {
        stopAutoReveal();
      }

      broadcastState();
      io.emit('sound:play', { type: 'buzzer_hit', player });
    }
  });

  // 4 hành động chuông theo yêu cầu: Mở chuông, Khóa chuông, Reset và mở, Reset và khóa
  // 1. Mở chuông
  socket.on('buzzer:enable', () => {
    gameState.buzzer.enabled = true;
    gameState.buzzer.winner = null;
    gameState.buzzer.timestamp = null;
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastState();
  });

  // 2. Khóa chuông
  socket.on('buzzer:disable', () => {
    gameState.buzzer.enabled = false;
    broadcastState();
  });

  // 3. Reset và mở
  socket.on('buzzer:reset_and_enable', () => {
    gameState.buzzer.winner = null;
    gameState.buzzer.timestamp = null;
    gameState.buzzer.enabled = true;
    gameState.puzzle.borderColor = '#800080';
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastState();
  });

  // 4. Reset và khóa
  socket.on('buzzer:reset_and_disable', () => {
    gameState.buzzer.winner = null;
    gameState.buzzer.timestamp = null;
    gameState.buzzer.enabled = false;
    gameState.puzzle.borderColor = '#800080';
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastState();
  });

  // Backward compatibility
  socket.on('buzzer:reset', () => {
    gameState.buzzer.winner = null;
    gameState.buzzer.timestamp = null;
    gameState.buzzer.enabled = false;
    gameState.puzzle.borderColor = '#800080';
    gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastState();
  });

  // Import parsed rounds data from Excel
  socket.on('rounds:import_data', (importedRounds) => {
    if (importedRounds && typeof importedRounds === 'object') {
      gameState.roundsData = { ...gameState.roundsData, ...importedRounds };
      broadcastState();
      socket.emit('rounds:imported_success', {
        sheetCount: Object.keys(importedRounds).length
      });
    }
  });

  // Full reset
  socket.on('game:reset_all', () => {
    gameState.players.forEach((p, idx) => {
      p.name = `Người chơi ${idx + 1}`;
      p.roundScore = 0;
      p.totalScore = 0;
      p.canSpin = false;
      p.buzzed = false;
      p.buzzTime = null;
    });
    gameState.currentRound = 'Vòng 1';
    gameState.puzzle.boardState = 'hidden';
    gameState.puzzle.revealedIndices = [];
    gameState.puzzle.revealedLetters = [];
    gameState.puzzle.borderColor = '#800080';
    gameState.topicSelection.active = false;
    stopAutoReveal();
    gameState.wheel.spinning = false;
    gameState.wheel.result = null;
    gameState.wheel.lockedAll = true;
    gameState.wheel.selectedWheel = 'vong7.png';
    gameState.buzzer.winner = null;
    broadcastState();
  });
});

// REST Endpoints
app.get('/api/state', (req, res) => {
  ensureWheelStateCurrent();
  res.json(gameState);
});

app.get('/api/wheels', (req, res) => {
  res.json(AVAILABLE_WHEELS);
});

app.get('/api/rounds', (req, res) => {
  res.json(gameState.roundsData);
});

// Serve static assets from public and root folder
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// Route handlers for the 7 requested pages
const pages = [
  'Controller',
  'Host',
  'Puzzleboard',
  'Wheel',
  'Player1',
  'Player2',
  'Player3'
];

pages.forEach((page) => {
  app.get(`/${page}.html`, (req, res) => {
    res.sendFile(path.join(__dirname, `${page}.html`));
  });
  app.get(`/${page.toLowerCase()}.html`, (req, res) => {
    res.sendFile(path.join(__dirname, `${page}.html`));
  });
  app.get(`/${page.toLowerCase()}`, (req, res) => {
    res.sendFile(path.join(__dirname, `${page}.html`));
  });
  app.get(`/${page}`, (req, res) => {
    res.sendFile(path.join(__dirname, `${page}.html`));
  });
});

// Root route
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Chiếc Nón Kỳ Diệu Game Server running on port ${PORT}`);
  console.log(`Available screens:`);
  console.log(` - http://localhost:${PORT}/Controller.html`);
  console.log(` - http://localhost:${PORT}/Host.html`);
  console.log(` - http://localhost:${PORT}/Puzzleboard.html`);
  console.log(` - http://localhost:${PORT}/Wheel.html`);
  console.log(` - http://localhost:${PORT}/Player1.html`);
  console.log(` - http://localhost:${PORT}/Player2.html`);
  console.log(` - http://localhost:${PORT}/Player3.html`);
});
