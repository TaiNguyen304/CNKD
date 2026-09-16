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

// Initial default round questions template
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

let globalLoadedRoundsData = JSON.parse(JSON.stringify(initialRoundsData));

// Parser for Puzzle.xlsx
function parseWorkbookData(workbook) {
  const result = {};
  if (!workbook || !workbook.SheetNames) return result;

  workbook.SheetNames.forEach(sheetName => {
    const ws = workbook.Sheets[sheetName];
    if (!ws || !ws['!ref']) return;

    const range = xlsx.utils.decode_range(ws['!ref']);
    
    // Tìm các dòng header có chữ 'Chủ đề' ở Cột A
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
      const catCell = ws[xlsx.utils.encode_cell({ r: headerR + 1, c: 0 })];
      let category = catCell && catCell.v !== undefined ? String(catCell.v).trim() : '';
      
      const clueCell = ws[xlsx.utils.encode_cell({ r: headerR + 2, c: 0 })];
      let clue = clueCell && clueCell.v !== undefined ? String(clueCell.v).trim() : '';

      let gridMatrix = Array.from({ length: 4 }, () => Array(16).fill(''));
      let fullWords = [];
      let cellFoundCount = 0;

      for (let rowIdx = 0; rowIdx < 4; rowIdx++) {
        const r = headerR + 1 + rowIdx;
        let curWord = '';
        for (let colIdx = 0; colIdx < 16; colIdx++) {
          const c = colIdx + 1;
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

function loadInitialPuzzles() {
  try {
    const filePath = path.join(__dirname, 'Puzzle.xlsx');
    if (fs.existsSync(filePath)) {
      const wb = xlsx.readFile(filePath);
      const parsed = parseWorkbookData(wb);
      if (Object.keys(parsed).length > 0) {
        globalLoadedRoundsData = { ...globalLoadedRoundsData, ...parsed };
        console.log('Successfully pre-loaded questions from Puzzle.xlsx');
      }
    }
  } catch (err) {
    console.error('Error pre-loading Puzzle.xlsx:', err.message);
  }
}
loadInitialPuzzles();

// Factory for a fresh Game State
function createNewGameState() {
  return {
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
      gridMatrix: null,
      revealedIndices: [],
      markedIndices: [],
      boardState: 'hidden',
      revealedLetters: [],
      showCategory: true,
      showClue: true,
      borderColor: '#800080'
    },
    topicSelection: {
      active: false,
      round: null,
      mode: null,
      topics: []
    },
    revealProgress: {
      active: false,
      intervalMs: 800
    },
    wheel: {
      selectedWheel: 'vong7.png',
      spinning: false,
      currentAngle: 0,
      startAngle: 0,
      finalAngle: 0,
      startTime: 0,
      duration: 22000,
      result: null,
      activeSpinner: null,
      lockedAll: true,
      pointers: {
        p1: { light: true, visible: true },
        p2: { light: true, visible: true },
        p3: { light: true, visible: true }
      }
    },
    wheelOverlays: {},
    overlayAngles: {},
    spinMusic: 'Nhạc quay Nón 1.mp3',
    puzzleSounds: {
      showBlank: 'reveal.mp3',
      markLetter: 'ding.wav',
      openLetter: '2nd_ding.wav',
      solvePuzzle: 'ClearPuzzle.mp3'
    },
    buzzer: {
      active: false,
      enabled: false,
      winner: null,
      timestamp: null
    },
    roundsData: JSON.parse(JSON.stringify(globalLoadedRoundsData))
  };
}

// Room Management (Multi-room architecture)
const rooms = new Map();

function generateRandomRoomId() {
  let id = '';
  do {
    id = Math.floor(100000 + Math.random() * 900000).toString(); // 6 digits
  } while (rooms.has(id));
  return id;
}

function generateRandomPass() {
  return Math.floor(1000 + Math.random() * 9000).toString(); // 4 digits
}

function createRoom(customRoomId = null, customPasses = null) {
  const roomId = (customRoomId && String(customRoomId).trim()) || generateRandomRoomId();
  const p1 = (customPasses && (customPasses[1] || customPasses['1'] || customPasses.p1)) || generateRandomPass();
  const p2 = (customPasses && (customPasses[2] || customPasses['2'] || customPasses.p2)) || generateRandomPass();
  const p3 = (customPasses && (customPasses[3] || customPasses['3'] || customPasses.p3)) || generateRandomPass();
  const passwords = {
    1: p1,
    2: p2,
    3: p3,
    '1': p1,
    '2': p2,
    '3': p3,
    p1: p1,
    p2: p2,
    p3: p3
  };

  const room = {
    id: roomId,
    passwords,
    createdAt: Date.now(),
    gameState: createNewGameState(),
    wheelSpinTimeout: null,
    revealIntervalTimer: null
  };

  rooms.set(roomId, room);
  return room;
}

// Ensure default room 100000 exists
createRoom('100000', { 1: '1111', 2: '2222', 3: '3333', p1: '1111', p2: '2222', p3: '3333' });

function getRoom(roomId) {
  if (!roomId) return null;
  const cleanId = String(roomId).trim();
  return rooms.get(cleanId) || null;
}

function getOrCreateRoom(roomId) {
  if (!roomId) return createRoom();
  const cleanId = String(roomId).trim();
  let r = rooms.get(cleanId);
  if (!r) {
    r = createRoom(cleanId);
  }
  return r;
}

// SANITIZE GAME STATE FOR PUBLIC CLIENTS (PLAYER, PUZZLEBOARD, WHEEL)
// Strict answer protection: Never expose secret letters before reveal
function sanitizeGameStateForPublic(state) {
  if (!state) return null;
  const clone = JSON.parse(JSON.stringify(state));

  if (clone.puzzle) {
    const isRevealed = clone.puzzle.boardState === 'revealed';
    if (!isRevealed) {
      clone.puzzle.answer = ''; // Mask the raw answer text

      if (clone.puzzle.gridMatrix && Array.isArray(clone.puzzle.gridMatrix)) {
        const revealedSet = new Set(clone.puzzle.revealedIndices || []);
        clone.puzzle.gridMatrix = clone.puzzle.gridMatrix.map((row, r) => {
          return row.map((char, c) => {
            const cellIdx = r * 16 + c;
            if (!char || !char.trim()) return '';
            // If already revealed, send actual letter. Otherwise send mask indicator
            if (revealedSet.has(cellIdx)) {
              return char;
            } else {
              return '#'; // Masked character indicating box contains a letter
            }
          });
        });
      }
    }
  }

  // Sanitize topic options to prevent reading future answers
  if (clone.topicSelection && Array.isArray(clone.topicSelection.topics)) {
    clone.topicSelection.topics = clone.topicSelection.topics.map(t => ({
      category: t.category,
      clue: t.clue
    }));
  }

  // Remove full answer bank from public clients
  if (clone.roundsData) {
    clone.roundsData = {};
  }

  return clone;
}

// Broadcast room state to all connected sockets in that room
function broadcastRoomState(roomId) {
  const room = getRoom(roomId);
  if (!room) return;

  const socketsInRoom = io.sockets.adapter.rooms.get(roomId);
  if (!socketsInRoom) return;

  const fullState = room.gameState;
  const publicState = sanitizeGameStateForPublic(fullState);

  for (const socketId of socketsInRoom) {
    const s = io.sockets.sockets.get(socketId);
    if (s) {
      const role = s.data.role;
      if (role === 'controller' || role === 'host') {
        s.emit('game:state', fullState);
      } else {
        s.emit('game:state', publicState);
      }
    }
  }
}

// Stop auto letter reveal timer for a room
function stopRoomAutoReveal(room) {
  if (room && room.revealIntervalTimer) {
    clearInterval(room.revealIntervalTimer);
    room.revealIntervalTimer = null;
  }
  if (room && room.gameState && room.gameState.revealProgress) {
    room.gameState.revealProgress.active = false;
  }
}

// Vietnamese tone removal for letter comparison
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

// SOCKET.IO EVENT HANDLING
io.on('connection', (socket) => {
  // Client joins a specific room
  socket.on('room:join', (data = {}) => {
    const { roomId, role, auth, slot } = data;
    if (!roomId) {
      socket.emit('room:join_error', { message: 'Vui lòng cung cấp mã phòng (roomid)!' });
      return;
    }

    const room = getOrCreateRoom(roomId);
    socket.data.roomId = room.id;
    socket.data.role = role || 'guest';
    socket.data.slot = Number(slot) || null;

    // Validate player authentication
    if (role === 'player') {
      const playerSlot = Number(slot);
      const expectedPass = room.passwords[playerSlot] || room.passwords['p' + playerSlot] || room.passwords[String(playerSlot)];
      if (!auth || String(auth).trim() !== String(expectedPass).trim()) {
        socket.emit('room:auth_failed', { message: 'Mật khẩu người chơi không chính xác hoặc phòng không tồn tại!' });
        return;
      }
      socket.data.authenticated = true;
    }

    socket.join(room.id);
    socket.emit('room:joined', {
      roomId: room.id,
      role: socket.data.role,
      passwords: (role === 'controller' || role === 'host') ? room.passwords : undefined
    });

    // Send initial state
    if (role === 'controller' || role === 'host') {
      socket.emit('game:state', room.gameState);
    } else {
      socket.emit('game:state', sanitizeGameStateForPublic(room.gameState));
    }
  });

  // State request
  socket.on('state:request', () => {
    const roomId = socket.data.roomId;
    const room = getRoom(roomId);
    if (!room) return;

    if (socket.data.role === 'controller' || socket.data.role === 'host') {
      socket.emit('game:state', room.gameState);
    } else {
      socket.emit('game:state', sanitizeGameStateForPublic(room.gameState));
    }
  });

  // NTP Time Synchronization for zero-latency cross-client audio sync
  socket.on('timesync:ping', (data = {}) => {
    socket.emit('timesync:pong', {
      t0: data.t0 || 0,
      serverTime: Date.now()
    });
  });

  // Helper to get socket's room
  function currentRoom() {
    const roomId = socket.data.roomId;
    return getRoom(roomId);
  }

  // Helper to broadcast synchronized sound event with timestamp scheduling
  function broadcastSound(room, eventName, payload) {
    if (!room || !room.id) return;
    const serverTime = Date.now();
    const playAt = serverTime + 60; // 60ms sync horizon for zero-lag alignment
    const fullPayload = (typeof payload === 'object' && payload !== null)
      ? { ...payload, serverTime, playAt }
      : { file: payload, serverTime, playAt };
    io.to(room.id).emit(eventName, fullPayload);
  }

  // Room creation / update from Controller
  socket.on('room:create_new', (customData = {}) => {
    const newRoomId = customData.roomId || generateRandomRoomId();
    const newPasses = customData.passwords || {
      p1: generateRandomPass(),
      p2: generateRandomPass(),
      p3: generateRandomPass()
    };
    const room = createRoom(newRoomId, newPasses);
    socket.data.roomId = room.id;
    socket.data.role = 'controller';
    socket.join(room.id);

    socket.emit('room:created', {
      roomId: room.id,
      passwords: room.passwords
    });
    broadcastRoomState(room.id);
  });

  // Sound selection config
  socket.on('puzzle:set_sound', ({ key, value }) => {
    const room = currentRoom();
    if (!room || !key || !value) return;
    if (!room.gameState.puzzleSounds) room.gameState.puzzleSounds = {};
    room.gameState.puzzleSounds[key] = value;
    broadcastRoomState(room.id);
  });

  // Player details & scores
  socket.on('player:set', (data) => {
    const room = currentRoom();
    if (!room) return;
    const { playerId, name, roundScore, totalScore } = data;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      if (name !== undefined) p.name = name;
      if (roundScore !== undefined) p.roundScore = Number(roundScore);
      if (totalScore !== undefined) p.totalScore = Number(totalScore);
      broadcastRoomState(room.id);
    }
  });

  socket.on('player:add_round', ({ playerId, amount }) => {
    const room = currentRoom();
    if (!room) return;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      p.roundScore = (p.roundScore || 0) + Number(amount);
      broadcastRoomState(room.id);
    }
  });

  socket.on('player:subtract_round', ({ playerId, amount }) => {
    const room = currentRoom();
    if (!room) return;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      p.roundScore = (p.roundScore || 0) - Number(amount);
      broadcastRoomState(room.id);
    }
  });

  socket.on('player:multiply_round', ({ playerId, factor }) => {
    const room = currentRoom();
    if (!room) return;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      p.roundScore = Math.floor((p.roundScore || 0) * Number(factor));
      broadcastRoomState(room.id);
    }
  });

  socket.on('player:reset_round', ({ playerId }) => {
    const room = currentRoom();
    if (!room) return;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      p.roundScore = 0;
      broadcastRoomState(room.id);
    }
  });

  socket.on('player:set_total', ({ playerId, totalScore }) => {
    const room = currentRoom();
    if (!room) return;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      p.totalScore = Number(totalScore);
      broadcastRoomState(room.id);
    }
  });

  socket.on('player:add_to_total', ({ playerId, resetRound }) => {
    const room = currentRoom();
    if (!room) return;
    const p = room.gameState.players.find(x => x.id === Number(playerId));
    if (p) {
      p.totalScore = (p.totalScore || 0) + (p.roundScore || 0);
      if (resetRound) p.roundScore = 0;
      broadcastRoomState(room.id);
    }
  });

  // Round Change
  socket.on('round:set', (roundName) => {
    const room = currentRoom();
    if (!room || !roundName) return;

    room.gameState.currentRound = roundName;
    room.gameState.players.forEach(p => p.roundScore = 0);
    room.gameState.topicSelection.active = false;
    room.gameState.buzzer.winner = null;
    room.gameState.buzzer.timestamp = null;
    room.gameState.puzzle.borderColor = '#800080';

    stopRoomAutoReveal(room);

    // Auto-load question from roundsData
    const questions = room.gameState.roundsData[roundName];
    if (questions && questions.length > 0) {
      const q = questions[0];
      room.gameState.puzzle.category = q.category || roundName.toUpperCase();
      room.gameState.puzzle.clue = q.clue || '';
      room.gameState.puzzle.answer = q.answer || '';
      room.gameState.puzzle.gridMatrix = q.gridMatrix || null;
      room.gameState.puzzle.revealedIndices = [];
      room.gameState.puzzle.markedIndices = [];
      room.gameState.puzzle.revealedLetters = [];
      room.gameState.puzzle.boardState = 'visible';

      const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.showBlank) || 'reveal.mp3';
      if (soundFile && soundFile !== 'None') {
        broadcastSound(room, 'sound:play', { type: 'puzzle_show', sound: soundFile });
      }
    }

    broadcastRoomState(room.id);
    broadcastSound(room, 'sound:play', { type: 'round_change', round: roundName });
  });

  // Choice 1: "Ô chữ cố định"
  socket.on('round:choose_fixed', (roundName) => {
    const room = currentRoom();
    if (!room || !roundName) return;

    room.gameState.currentRound = roundName;
    room.gameState.players.forEach(p => p.roundScore = 0);
    room.gameState.buzzer.winner = null;
    room.gameState.puzzle.borderColor = '#800080';

    room.gameState.topicSelection = {
      active: false,
      round: roundName,
      mode: 'fixed',
      topics: []
    };

    const questions = room.gameState.roundsData[roundName] || [];
    if (questions.length > 0) {
      const q = questions[0];
      room.gameState.puzzle.category = q.category || roundName.toUpperCase();
      room.gameState.puzzle.clue = q.clue || '';
      room.gameState.puzzle.answer = q.answer || '';
      room.gameState.puzzle.gridMatrix = q.gridMatrix || null;
    }

    room.gameState.puzzle.boardState = 'visible';
    room.gameState.puzzle.revealedIndices = [];
    room.gameState.puzzle.markedIndices = [];
    room.gameState.puzzle.revealedLetters = [];

    stopRoomAutoReveal(room);
    broadcastRoomState(room.id);

    const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.showBlank) || 'reveal.mp3';
    if (soundFile && soundFile !== 'None') {
      io.to(room.id).emit('sound:play', { type: 'puzzle_show', sound: soundFile });
    }
  });

  // Choice 2: "Chọn 1 trong 3 chủ đề"
  socket.on('round:start_topic_selection', (roundName) => {
    const room = currentRoom();
    if (!room || !roundName) return;

    room.gameState.currentRound = roundName;
    room.gameState.players.forEach(p => p.roundScore = 0);
    room.gameState.buzzer.winner = null;
    room.gameState.puzzle.borderColor = '#800080';

    stopRoomAutoReveal(room);

    let topics = (room.gameState.roundsData[roundName] || []).slice(0, 3);
    if (topics.length === 0) {
      topics = [
        { category: 'CHỦ ĐỀ 1', answer: 'Ô CHỮ MỘT', clue: 'Gợi ý 1' },
        { category: 'CHỦ ĐỀ 2', answer: 'Ô CHỮ HAI', clue: 'Gợi ý 2' },
        { category: 'CHỦ ĐỀ 3', answer: 'Ô CHỮ BA', clue: 'Gợi ý 3' }
      ];
    }

    room.gameState.topicSelection = {
      active: true,
      round: roundName,
      mode: 'choose_topic',
      topics
    };

    room.gameState.puzzle.boardState = 'hidden';
    room.gameState.puzzle.revealedIndices = [];
    room.gameState.puzzle.markedIndices = [];
    room.gameState.puzzle.revealedLetters = [];

    broadcastRoomState(room.id);
  });

  // Pick one of 3 topics
  socket.on('round:pick_topic', ({ round, topicIndex }) => {
    const room = currentRoom();
    if (!room) return;

    const roundName = round || room.gameState.currentRound;
    const questions = room.gameState.roundsData[roundName] || [];
    const chosen = questions[topicIndex] || (room.gameState.topicSelection.topics && room.gameState.topicSelection.topics[topicIndex]);

    if (chosen) {
      room.gameState.puzzle.category = chosen.category || `CHỦ ĐỀ ${topicIndex + 1}`;
      room.gameState.puzzle.clue = chosen.clue || '';
      room.gameState.puzzle.answer = chosen.answer || '';
      room.gameState.puzzle.gridMatrix = chosen.gridMatrix || null;

      room.gameState.topicSelection.active = false;
      room.gameState.puzzle.boardState = 'visible';
      room.gameState.puzzle.revealedIndices = [];
      room.gameState.puzzle.markedIndices = [];
      room.gameState.puzzle.revealedLetters = [];

      stopRoomAutoReveal(room);
      broadcastRoomState(room.id);

      const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.showBlank) || 'reveal.mp3';
      if (soundFile && soundFile !== 'None') {
        io.to(room.id).emit('sound:play', { type: 'puzzle_show', sound: soundFile });
      }
    }
  });

  // Wheel Spin (22s duration)
  socket.on('wheel:spin', (data = {}) => {
    const room = currentRoom();
    if (!room) return;

    const currentAngle = room.gameState.wheel.currentAngle || 0;
    const randomWedgeDeg = Math.floor(Math.random() * 360);
    const deltaAngle = (data.targetAngle !== undefined) ? data.targetAngle : (360 * 18 + randomWedgeDeg);
    const finalAngle = currentAngle + deltaAngle;
    const duration = 22000;
    const startTime = Date.now();

    room.gameState.wheel.spinning = true;
    room.gameState.wheel.startAngle = currentAngle;
    room.gameState.wheel.finalAngle = finalAngle;
    room.gameState.wheel.currentAngle = finalAngle;
    room.gameState.wheel.duration = duration;
    room.gameState.wheel.startTime = startTime;
    room.gameState.wheel.activeSpinner = data.spinner || 'controller';

    if (room.wheelSpinTimeout) clearTimeout(room.wheelSpinTimeout);
    room.wheelSpinTimeout = setTimeout(() => {
      if (room.gameState.wheel.spinning) {
        room.gameState.wheel.spinning = false;
        room.gameState.players.forEach(p => p.canSpin = false);
        room.gameState.wheel.lockedAll = true;
        room.gameState.wheel.activeSpinner = null;
        broadcastRoomState(room.id);
      }
    }, duration + 300);

    const spinMusic = room.gameState.spinMusic || 'Nhạc quay Nón 1.mp3';
    broadcastSound(room, 'wheel:start_spin', {
      startAngle: room.gameState.wheel.startAngle,
      finalAngle: room.gameState.wheel.finalAngle,
      duration,
      startTime,
      spinner: room.gameState.wheel.activeSpinner,
      selectedWheel: room.gameState.wheel.selectedWheel,
      music: spinMusic
    });
    broadcastRoomState(room.id);
  });

  // Soundboard broadcast handlers (scoped to room)
  socket.on('soundboard:play', (file) => {
    const room = currentRoom();
    if (room) broadcastSound(room, 'soundboard:play_file', { file });
  });
  socket.on('soundboard:stop', () => {
    const room = currentRoom();
    if (room) io.to(room.id).emit('soundboard:stop_all');
  });
  socket.on('PLAY_SOUNDBOARD', (file) => {
    const room = currentRoom();
    if (room) broadcastSound(room, 'soundboard:play_file', { file });
  });
  socket.on('STOP_SOUNDBOARD', () => {
    const room = currentRoom();
    if (room) io.to(room.id).emit('soundboard:stop_all');
  });

  // Wheel configuration
  socket.on('wheel:select', (wheelId) => {
    const room = currentRoom();
    if (!room || !wheelId) return;
    room.gameState.wheel.selectedWheel = wheelId;
    broadcastRoomState(room.id);
  });

  socket.on('wheel:set_spin_music', (musicFile) => {
    const room = currentRoom();
    if (!room || !musicFile) return;
    room.gameState.spinMusic = musicFile.trim();
    broadcastRoomState(room.id);
  });

  // Wheel Overlays
  socket.on('wheel:toggle_overlay', (itemId) => {
    const room = currentRoom();
    if (!room || !itemId) return;
    room.gameState.wheelOverlays[itemId] = !room.gameState.wheelOverlays[itemId];
    broadcastRoomState(room.id);
  });

  socket.on('wheel:set_overlay', ({ itemId, visible }) => {
    const room = currentRoom();
    if (!room || !itemId) return;
    room.gameState.wheelOverlays[itemId] = !!visible;
    broadcastRoomState(room.id);
  });

  socket.on('wheel:set_all_overlays', (visible) => {
    const room = currentRoom();
    if (!room) return;
    const allIds = [
      'pn1_a', 'pn1_b', 'bm1_1', 'bm2_1', 'st1', 'dd1', 'o1000', 'o2000', 'o100k_1',
      'pn2', 'bm1_2', 'bm2_2', 'st2', 'dd2', 'o3000', 'o4000', 'oqd', 'o100k_2',
      't_ch', 't_qn'
    ];
    allIds.forEach(id => {
      room.gameState.wheelOverlays[id] = !!visible;
    });
    broadcastRoomState(room.id);
  });

  socket.on('wheel:set_overlay_angle', ({ itemId, angle }) => {
    const room = currentRoom();
    if (!room || !itemId || angle === undefined) return;
    if (!room.gameState.overlayAngles) room.gameState.overlayAngles = {};
    room.gameState.overlayAngles[itemId] = Number(angle);
    broadcastRoomState(room.id);
  });

  socket.on('wheel:landed', (result) => {
    const room = currentRoom();
    if (!room) return;
    if (room.wheelSpinTimeout) {
      clearTimeout(room.wheelSpinTimeout);
      room.wheelSpinTimeout = null;
    }
    if (room.gameState.wheel.spinning) {
      room.gameState.wheel.spinning = false;
      room.gameState.wheel.result = result || {};
      room.gameState.players.forEach(p => p.canSpin = false);
      room.gameState.wheel.lockedAll = true;
      room.gameState.wheel.activeSpinner = null;
      broadcastRoomState(room.id);
      broadcastSound(room, 'sound:play', { type: 'wheel_result', result });
    }
  });

  socket.on('wheel:allow_player', (playerId) => {
    const room = currentRoom();
    if (!room) return;
    if (room.wheelSpinTimeout) {
      clearTimeout(room.wheelSpinTimeout);
      room.wheelSpinTimeout = null;
    }
    room.gameState.wheel.spinning = false;
    room.gameState.wheel.lockedAll = false;
    const targetId = Number(playerId);
    room.gameState.players.forEach(p => {
      p.canSpin = (p.id === targetId);
    });
    broadcastRoomState(room.id);
    io.to(room.id).emit('player:spin_enabled', { playerId: targetId });
  });

  socket.on('wheel:lock_all', () => {
    const room = currentRoom();
    if (!room) return;
    if (room.wheelSpinTimeout) {
      clearTimeout(room.wheelSpinTimeout);
      room.wheelSpinTimeout = null;
    }
    room.gameState.wheel.lockedAll = true;
    room.gameState.players.forEach(p => p.canSpin = false);
    broadcastRoomState(room.id);
    io.to(room.id).emit('player:spin_disabled', {});
  });

  // Pointer Lights & Visibility
  socket.on('wheel:toggle_pointer_light', (pointer) => {
    const room = currentRoom();
    if (room && room.gameState.wheel.pointers && room.gameState.wheel.pointers[pointer]) {
      room.gameState.wheel.pointers[pointer].light = !room.gameState.wheel.pointers[pointer].light;
      broadcastRoomState(room.id);
    }
  });

  socket.on('wheel:set_all_pointer_lights', ({ light }) => {
    const room = currentRoom();
    if (room && room.gameState.wheel.pointers) {
      ['p1', 'p2', 'p3'].forEach(k => {
        if (room.gameState.wheel.pointers[k]) room.gameState.wheel.pointers[k].light = !!light;
      });
      broadcastRoomState(room.id);
    }
  });

  socket.on('wheel:toggle_pointer_visibility', (pointer) => {
    const room = currentRoom();
    if (room && room.gameState.wheel.pointers && room.gameState.wheel.pointers[pointer]) {
      room.gameState.wheel.pointers[pointer].visible = !room.gameState.wheel.pointers[pointer].visible;
      broadcastRoomState(room.id);
    }
  });

  // Puzzle Actions
  socket.on('puzzle:set', (data) => {
    const room = currentRoom();
    if (!room || !data) return;
    const { category, clue, answer, gridMatrix } = data;
    if (answer || gridMatrix) {
      room.gameState.puzzle.category = category || 'CHỦ ĐỀ';
      room.gameState.puzzle.clue = clue || '';
      room.gameState.puzzle.answer = answer || '';
      room.gameState.puzzle.gridMatrix = gridMatrix || null;
      room.gameState.puzzle.revealedIndices = [];
      room.gameState.puzzle.markedIndices = [];
      room.gameState.puzzle.revealedLetters = [];
      room.gameState.puzzle.boardState = 'visible';
      broadcastRoomState(room.id);
    }
  });

  socket.on('puzzle:show', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.puzzle.boardState = 'visible';
    room.gameState.topicSelection.active = false;
    stopRoomAutoReveal(room);
    broadcastRoomState(room.id);
    const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.showBlank) || 'reveal.mp3';
    if (soundFile && soundFile !== 'None') {
      io.to(room.id).emit('sound:play', { type: 'puzzle_show', sound: soundFile });
    }
  });

  socket.on('puzzle:solve', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.puzzle.boardState = 'revealed';
    stopRoomAutoReveal(room);
    broadcastRoomState(room.id);

    const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.solvePuzzle) || 'ClearPuzzle.mp3';
    if (soundFile && soundFile !== 'None') {
      broadcastSound(room, 'sound:play', { type: 'puzzle_solve', sound: soundFile });
    }
  });

  socket.on('puzzle:reset', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.puzzle.boardState = 'cleared';
    room.gameState.puzzle.revealedIndices = [];
    room.gameState.puzzle.markedIndices = [];
    room.gameState.puzzle.revealedLetters = [];
    room.gameState.puzzle.borderColor = '#800080';
    room.gameState.topicSelection.active = false;
    stopRoomAutoReveal(room);
    broadcastRoomState(room.id);
  });

  socket.on('puzzle:set_border_color', (color) => {
    const room = currentRoom();
    if (!room || !color) return;
    room.gameState.puzzle.borderColor = color;
    broadcastRoomState(room.id);
  });

  socket.on('puzzle:toggle_category', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.puzzle.showCategory = !room.gameState.puzzle.showCategory;
    broadcastRoomState(room.id);
  });

  socket.on('puzzle:toggle_clue', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.puzzle.showClue = !room.gameState.puzzle.showClue;
    broadcastRoomState(room.id);
  });

  // Letter checking and marking
  socket.on('puzzle:check_letter', (letter) => {
    const room = currentRoom();
    if (!room || !letter) return;
    const cleanLetter = removeToneMarks(letter.trim().toUpperCase());
    let count = 0;
    const matchingIndices = [];

    if (room.gameState.puzzle.gridMatrix && Array.isArray(room.gameState.puzzle.gridMatrix)) {
      for (let r = 0; r < 4; r++) {
        for (let c = 0; c < 16; c++) {
          const char = (room.gameState.puzzle.gridMatrix[r] && room.gameState.puzzle.gridMatrix[r][c]) ? room.gameState.puzzle.gridMatrix[r][c] : '';
          if (char && char.trim()) {
            const cellIdx = r * 16 + c;
            if (removeToneMarks(char) === cleanLetter && !room.gameState.puzzle.revealedIndices.includes(cellIdx)) {
              count++;
              matchingIndices.push(cellIdx);
            }
          }
        }
      }
    } else if (room.gameState.puzzle.answer) {
      const chars = room.gameState.puzzle.answer.toUpperCase().split('');
      for (let i = 0; i < chars.length; i++) {
        const c = chars[i];
        if (c !== ' ' && removeToneMarks(c) === cleanLetter && !room.gameState.puzzle.revealedIndices.includes(i)) {
          count++;
          matchingIndices.push(i);
        }
      }
    }

    socket.emit('puzzle:check_result', {
      letter: cleanLetter,
      count,
      indices: matchingIndices
    });
  });

  // Step 1: Mark matching letters (Blue)
  socket.on('puzzle:mark_letter', (letter) => {
    const room = currentRoom();
    if (!room || !letter) return;
    const cleanLetter = removeToneMarks(letter.trim().toUpperCase());
    const newlyMarked = [];

    if (room.gameState.puzzle.gridMatrix && Array.isArray(room.gameState.puzzle.gridMatrix)) {
      for (let r = 0; r < 4; r++) {
        for (let c = 0; c < 16; c++) {
          const char = (room.gameState.puzzle.gridMatrix[r] && room.gameState.puzzle.gridMatrix[r][c]) ? room.gameState.puzzle.gridMatrix[r][c] : '';
          if (char && char.trim()) {
            const cellIdx = r * 16 + c;
            if (removeToneMarks(char) === cleanLetter && !room.gameState.puzzle.revealedIndices.includes(cellIdx)) {
              newlyMarked.push(cellIdx);
            }
          }
        }
      }
    } else if (room.gameState.puzzle.answer) {
      const chars = room.gameState.puzzle.answer.toUpperCase().split('');
      for (let i = 0; i < chars.length; i++) {
        const c = chars[i];
        if (c !== ' ' && removeToneMarks(c) === cleanLetter && !room.gameState.puzzle.revealedIndices.includes(i)) {
          newlyMarked.push(i);
        }
      }
    }

    room.gameState.puzzle.markedIndices = newlyMarked;
    broadcastRoomState(room.id);

    const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.markLetter) || 'ding.wav';
    if (soundFile && soundFile !== 'None') {
      broadcastSound(room, 'sound:play', { type: 'letter_mark', sound: soundFile });
    }
  });

  // Step 2: Reveal marked letters
  socket.on('puzzle:reveal_marked', (letter) => {
    const room = currentRoom();
    if (!room) return;
    const marked = room.gameState.puzzle.markedIndices || [];
    if (marked.length > 0) {
      room.gameState.puzzle.revealedIndices = Array.from(new Set([...room.gameState.puzzle.revealedIndices, ...marked]));
      room.gameState.puzzle.markedIndices = [];
      if (letter) {
        const cleanLetter = removeToneMarks(letter.trim().toUpperCase());
        if (!room.gameState.puzzle.revealedLetters.includes(cleanLetter)) {
          room.gameState.puzzle.revealedLetters.push(cleanLetter);
        }
      }

      broadcastRoomState(room.id);

      const soundFile = (room.gameState.puzzleSounds && room.gameState.puzzleSounds.openLetter) || '2nd_ding.wav';
      if (soundFile && soundFile !== 'None') {
        broadcastSound(room, 'sound:play', { type: 'letter_open', sound: soundFile });
      }
    }
  });

  socket.on('puzzle:clear_marked', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.puzzle.markedIndices = [];
    broadcastRoomState(room.id);
  });

  // Auto Letter Reveal
  socket.on('puzzle:start_reveal', (options = {}) => {
    const room = currentRoom();
    if (!room) return;
    const intervalMs = options.intervalMs || 1500;
    stopRoomAutoReveal(room);

    room.gameState.revealProgress.active = true;
    room.gameState.revealProgress.intervalMs = intervalMs;
    room.gameState.puzzle.borderColor = '#800080';

    // Bổ sung: Mở chuông cho cả 3 người chơi, xóa trạng thái đã bấm chuông trước đó
    room.gameState.buzzer.enabled = true;
    room.gameState.buzzer.winner = null;
    room.gameState.buzzer.timestamp = null;
    room.gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });

    room.revealIntervalTimer = setInterval(() => {
      let unrevealed = [];
      if (room.gameState.puzzle.gridMatrix && Array.isArray(room.gameState.puzzle.gridMatrix)) {
        for (let r = 0; r < 4; r++) {
          for (let c = 0; c < 16; c++) {
            const char = (room.gameState.puzzle.gridMatrix[r] && room.gameState.puzzle.gridMatrix[r][c]) ? room.gameState.puzzle.gridMatrix[r][c] : '';
            if (char && char.trim()) {
              const cellIdx = r * 16 + c;
              if (!room.gameState.puzzle.revealedIndices.includes(cellIdx)) {
                unrevealed.push(cellIdx);
              }
            }
          }
        }
      } else if (room.gameState.puzzle.answer) {
        const chars = room.gameState.puzzle.answer.toUpperCase().split('');
        for (let i = 0; i < chars.length; i++) {
          if (chars[i] !== ' ' && !room.gameState.puzzle.revealedIndices.includes(i)) {
            unrevealed.push(i);
          }
        }
      }

      if (unrevealed.length === 0) {
        stopRoomAutoReveal(room);
        broadcastRoomState(room.id);
        return;
      }

      const randomIndex = Math.floor(Math.random() * unrevealed.length);
      const nextIndex = unrevealed[randomIndex];
      room.gameState.puzzle.revealedIndices.push(nextIndex);
      broadcastRoomState(room.id);
      broadcastSound(room, 'sound:play', { type: 'letter_flip' });
    }, intervalMs);

    broadcastRoomState(room.id);
  });

  socket.on('puzzle:stop_reveal', () => {
    const room = currentRoom();
    if (room) {
      stopRoomAutoReveal(room);
      broadcastRoomState(room.id);
    }
  });

  socket.on('puzzle:resume_reveal', () => {
    const room = currentRoom();
    if (!room) return;
    stopRoomAutoReveal(room);

    room.gameState.revealProgress.active = true;
    room.gameState.puzzle.borderColor = '#800080';

    // Nút tiếp tục hiện: Khung bảng và 3 ô điểm trên Puzzleboard trở về màu tím, chuông của cả 3 người chơi được mở lại
    room.gameState.buzzer.enabled = true;
    room.gameState.buzzer.winner = null;
    room.gameState.buzzer.timestamp = null;
    room.gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });

    room.revealIntervalTimer = setInterval(() => {
      let unrevealed = [];
      if (room.gameState.puzzle.gridMatrix && Array.isArray(room.gameState.puzzle.gridMatrix)) {
        for (let r = 0; r < 4; r++) {
          for (let c = 0; c < 16; c++) {
            const char = (room.gameState.puzzle.gridMatrix[r] && room.gameState.puzzle.gridMatrix[r][c]) ? room.gameState.puzzle.gridMatrix[r][c] : '';
            if (char && char.trim()) {
              const cellIdx = r * 16 + c;
              if (!room.gameState.puzzle.revealedIndices.includes(cellIdx)) {
                unrevealed.push(cellIdx);
              }
            }
          }
        }
      } else if (room.gameState.puzzle.answer) {
        const chars = room.gameState.puzzle.answer.toUpperCase().split('');
        for (let i = 0; i < chars.length; i++) {
          if (chars[i] !== ' ' && !room.gameState.puzzle.revealedIndices.includes(i)) {
            unrevealed.push(i);
          }
        }
      }

      if (unrevealed.length === 0) {
        stopRoomAutoReveal(room);
        broadcastRoomState(room.id);
        return;
      }
      const randomIndex = Math.floor(Math.random() * unrevealed.length);
      const nextIndex = unrevealed[randomIndex];
      room.gameState.puzzle.revealedIndices.push(nextIndex);
      broadcastRoomState(room.id);
      broadcastSound(room, 'sound:play', { type: 'letter_flip' });
    }, room.gameState.revealProgress.intervalMs || 1500);

    broadcastRoomState(room.id);
  });

  // Buzzer Controls
  socket.on('buzzer:press', (playerId) => {
    const room = currentRoom();
    if (!room) return;

    const player = room.gameState.players.find(p => p.id === Number(playerId));
    if (player && room.gameState.buzzer.enabled && !room.gameState.buzzer.winner) {
      room.gameState.buzzer.winner = player.id;
      room.gameState.buzzer.timestamp = Date.now();
      room.gameState.buzzer.enabled = false;
      player.buzzed = true;
      player.buzzTime = new Date().toLocaleTimeString();

      if (player.id === 1) room.gameState.puzzle.borderColor = '#ff0000';
      else if (player.id === 2) room.gameState.puzzle.borderColor = '#ffff00';
      else if (player.id === 3) room.gameState.puzzle.borderColor = '#0000ff';

      if (room.gameState.revealProgress.active) {
        stopRoomAutoReveal(room);
      }

      broadcastRoomState(room.id);
      broadcastSound(room, 'sound:play', { type: 'buzzer_hit', player });
    }
  });

  socket.on('buzzer:enable', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.buzzer.enabled = true;
    room.gameState.buzzer.winner = null;
    room.gameState.buzzer.timestamp = null;
    room.gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastRoomState(room.id);
  });

  socket.on('buzzer:disable', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.buzzer.enabled = false;
    broadcastRoomState(room.id);
  });

  socket.on('buzzer:reset_and_enable', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.buzzer.winner = null;
    room.gameState.buzzer.timestamp = null;
    room.gameState.buzzer.enabled = true;
    room.gameState.puzzle.borderColor = '#800080';
    room.gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastRoomState(room.id);
  });

  socket.on('buzzer:reset_and_disable', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.buzzer.winner = null;
    room.gameState.buzzer.timestamp = null;
    room.gameState.buzzer.enabled = false;
    room.gameState.puzzle.borderColor = '#800080';
    room.gameState.players.forEach(p => {
      p.buzzed = false;
      p.buzzTime = null;
    });
    broadcastRoomState(room.id);
  });

  socket.on('rounds:import_data', (importedRounds) => {
    const room = currentRoom();
    if (!room || !importedRounds) return;
    room.gameState.roundsData = { ...room.gameState.roundsData, ...importedRounds };
    broadcastRoomState(room.id);
    socket.emit('rounds:imported_success', {
      sheetCount: Object.keys(importedRounds).length
    });
  });

  // Full reset for room
  socket.on('game:reset_all', () => {
    const room = currentRoom();
    if (!room) return;
    room.gameState.players.forEach((p, idx) => {
      p.name = `Người chơi ${idx + 1}`;
      p.roundScore = 0;
      p.totalScore = 0;
      p.canSpin = false;
      p.buzzed = false;
      p.buzzTime = null;
    });
    room.gameState.currentRound = 'Vòng 1';
    room.gameState.puzzle.boardState = 'hidden';
    room.gameState.puzzle.revealedIndices = [];
    room.gameState.puzzle.revealedLetters = [];
    room.gameState.puzzle.borderColor = '#800080';
    room.gameState.topicSelection.active = false;
    stopRoomAutoReveal(room);
    room.gameState.wheel.spinning = false;
    room.gameState.wheel.result = null;
    room.gameState.wheel.lockedAll = true;
    room.gameState.wheel.selectedWheel = 'vong7.png';
    room.gameState.buzzer.winner = null;
    broadcastRoomState(room.id);
  });
});

// REST ENDPOINTS

// Room management endpoints
app.post('/api/room/create', (req, res) => {
  const { roomId, passwords } = req.body || {};
  const room = createRoom(roomId, passwords);
  res.json({
    success: true,
    roomId: room.id,
    passwords: room.passwords
  });
});

app.post('/api/room/verify', (req, res) => {
  const { roomId, slot, auth } = req.body || {};
  const room = getRoom(roomId);
  if (!room) {
    return res.status(404).json({ success: false, message: 'Mã phòng không tồn tại!' });
  }

  if (slot) {
    const pSlot = Number(slot);
    const expectedPass = room.passwords[pSlot] || room.passwords['p' + pSlot] || room.passwords[String(pSlot)];
    if (!auth || String(auth).trim() !== String(expectedPass).trim()) {
      return res.status(401).json({ success: false, message: 'Mật khẩu vị trí này không chính xác!' });
    }
  }

  res.json({
    success: true,
    roomId: room.id,
    slot: Number(slot) || null
  });
});

app.get('/api/room/info', (req, res) => {
  const roomId = req.query.roomid;
  const room = getRoom(roomId);
  if (!room) {
    return res.status(404).json({ exists: false });
  }
  res.json({
    exists: true,
    roomId: room.id,
    passwords: room.passwords
  });
});

app.get('/api/state', (req, res) => {
  const roomId = req.query.roomid;
  const room = getRoom(roomId) || getRoom('100000');
  if (!room) {
    return res.json(createNewGameState());
  }

  const role = req.query.role;
  if (role === 'controller' || role === 'host') {
    res.json(room.gameState);
  } else {
    res.json(sanitizeGameStateForPublic(room.gameState));
  }
});

app.get('/api/wheels', (req, res) => {
  res.json(AVAILABLE_WHEELS);
});

// Static assets
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// HTML Page Route aliases
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

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`Chiếc Nón Kỳ Diệu Game Server running on port ${PORT}`);
  console.log(`Ready for production & Render deployments`);
});
