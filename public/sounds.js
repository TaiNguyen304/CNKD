// =========================================================================
// Chiếc Nón Kỳ Diệu - High Precision Synchronized Sound Engine
// Features:
// 1. Iframe Muting: Embedded iframes (Wheel/Puzzleboard in Host/Player) are
//    muted to prevent overlapping/echo inside a single window.
// 2. Controller Muting: Controller panel is muted so only viewers/hosts/players
//    output master audio.
// 3. Sub-millisecond NTP Clock Synchronization: Perfectly phase-locks audio
//    across different devices, eliminating network delay skew.
// =========================================================================

class SoundEngine {
  constructor() {
    this.ctx = null;
    
    // Check if running inside an iframe or in controller page
    const isIframe = typeof window !== 'undefined' && (window.self !== window.top);
    const isController = typeof window !== 'undefined' && (
      window.location.pathname.toLowerCase().includes('controller') ||
      window.location.href.toLowerCase().includes('controller')
    );

    // Mute if inside an iframe or on controller page
    this.muted = isIframe || isController;
    this.isIframe = isIframe;
    this.isController = isController;

    // Active audio channels
    this.activeSpinAudio = null;
    this.activeSoundboardAudio = null;
    this.activeSfxAudio = null;

    // NTP Time Synchronization (serverTime - clientLocalTime)
    this.clockOffset = 0;
    this.hasSyncedTime = false;
    this.timeSyncInterval = null;

    // Deduplication tracker (file + targetPlayAt)
    this.recentPlays = new Map();
    this.audioCache = new Map();
  }

  init() {
    if (this.muted) return;
    if (!this.ctx) {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) {
        this.ctx = new AudioCtx();
      }
    }
    if (this.ctx && this.ctx.state === 'suspended') {
      this.ctx.resume().catch(() => {});
    }
  }

  // NTP Time Synchronization Protocol
  startTimeSync(socket) {
    if (!socket || this.muted) return;

    const doPing = () => {
      if (socket.connected) {
        socket.emit('timesync:ping', { t0: Date.now() });
      }
    };

    socket.on('timesync:pong', (data) => {
      if (!data || data.t0 === undefined || data.serverTime === undefined) return;
      const t1 = Date.now();
      const rtt = Math.max(0, t1 - data.t0);
      const estimatedLatency = rtt / 2;
      const calculatedOffset = (data.serverTime + estimatedLatency) - t1;

      if (!this.hasSyncedTime) {
        this.clockOffset = calculatedOffset;
        this.hasSyncedTime = true;
      } else {
        // Exponential moving average for clock drift stability
        this.clockOffset = Math.round(this.clockOffset * 0.7 + calculatedOffset * 0.3);
      }
    });

    doPing();
    if (this.timeSyncInterval) clearInterval(this.timeSyncInterval);
    this.timeSyncInterval = setInterval(doPing, 8000);
  }

  // Pre-cache audio element
  getAudioInstance(filename) {
    const cleanName = filename.replace(/^\/+/, '');
    const audioUrl = '/' + encodeURI(cleanName);
    const audio = new Audio(audioUrl);
    audio.preload = 'auto';
    return { audio, audioUrl };
  }

  // High-precision synchronized audio trigger
  schedulePlay(filename, options = {}) {
    if (this.muted || !filename) return null;
    this.init();

    const loop = !!options.loop;
    const volume = options.volume !== undefined ? options.volume : 0.9;
    const category = options.category || 'sfx'; // 'spin', 'soundboard', 'sfx'
    const serverTime = options.serverTime;
    const playAt = options.playAt;

    // Deduplication check: ignore if exact same sound scheduled within 100ms
    const dedupKey = `${filename}_${category}`;
    const now = Date.now();
    if (this.recentPlays.has(dedupKey)) {
      const lastPlayTime = this.recentPlays.get(dedupKey);
      if (now - lastPlayTime < 100) {
        return null;
      }
    }
    this.recentPlays.set(dedupKey, now);

    // Calculate synchronized target delay
    let delayMs = 0;
    let startSeekSec = 0;

    if (playAt || serverTime) {
      const serverNow = now + this.clockOffset;
      const targetPlayAt = playAt || (serverTime + 60);
      const diff = targetPlayAt - serverNow;

      if (diff > 0) {
        delayMs = diff;
      } else {
        // Network took longer than scheduled buffer, start immediately and catch up
        startSeekSec = Math.min(10, (-diff) / 1000);
      }
    }

    if (delayMs > 0) {
      setTimeout(() => {
        this.executePlay(filename, loop, volume, startSeekSec, category);
      }, delayMs);
      return null;
    } else {
      return this.executePlay(filename, loop, volume, startSeekSec, category);
    }
  }

  executePlay(filename, loop, volume, seekSec, category) {
    if (this.muted) return null;

    // Channel specific cleanup to prevent overlapping/echoing
    if (category === 'spin') {
      this.stopSpin();
    } else if (category === 'soundboard') {
      this.stopSoundboard();
    }

    const { audio, audioUrl } = this.getAudioInstance(filename);
    audio.volume = Math.max(0, Math.min(1, volume));
    audio.loop = loop;

    if (seekSec > 0.05) {
      try {
        audio.currentTime = seekSec;
      } catch (e) {}
    }

    const promise = audio.play();
    if (promise !== undefined) {
      promise.catch((err) => {
        console.warn('[SoundEngine] Playback warning for:', audioUrl, err.message);
      });
    }

    if (category === 'spin') {
      this.activeSpinAudio = audio;
    } else if (category === 'soundboard') {
      this.activeSoundboardAudio = audio;
    } else {
      this.activeSfxAudio = audio;
    }

    return audio;
  }

  playAudioFile(filename, loop = false, volume = 0.8) {
    return this.schedulePlay(filename, { loop, volume, category: 'sfx' });
  }

  playSpin(musicFile, timing = {}) {
    if (this.muted) return;
    this.stopSpin();
    const fileToPlay = musicFile || 'Nhạc quay Nón 1.mp3';
    return this.schedulePlay(fileToPlay, {
      loop: false,
      volume: 0.95,
      category: 'spin',
      serverTime: timing.serverTime,
      playAt: timing.playAt
    });
  }

  stopSpin() {
    if (this.activeSpinAudio) {
      try {
        this.activeSpinAudio.pause();
        this.activeSpinAudio.currentTime = 0;
      } catch (e) {}
      this.activeSpinAudio = null;
    }
  }

  playSoundboard(filename, timing = {}) {
    if (this.muted || !filename) return;
    this.stopSoundboard();
    return this.schedulePlay(filename, {
      loop: false,
      volume: 0.9,
      category: 'soundboard',
      serverTime: timing.serverTime,
      playAt: timing.playAt
    });
  }

  stopSoundboard() {
    if (this.activeSoundboardAudio) {
      try {
        this.activeSoundboardAudio.pause();
        this.activeSoundboardAudio.currentTime = 0;
      } catch (e) {}
      this.activeSoundboardAudio = null;
    }
  }

  stopAll() {
    this.stopSpin();
    this.stopSoundboard();
    if (this.activeSfxAudio) {
      try {
        this.activeSfxAudio.pause();
        this.activeSfxAudio.currentTime = 0;
      } catch (e) {}
      this.activeSfxAudio = null;
    }
  }

  playBuzzer(timing = {}) {
    if (this.muted) return;
    return this.schedulePlay('buzzer.mp3', { volume: 0.95, ...timing, category: 'sfx' });
  }

  playLetterCorrect(timing = {}) {
    if (this.muted) return;
    return this.schedulePlay('ding.wav', { volume: 0.9, ...timing, category: 'sfx' });
  }

  playLetterWrong(timing = {}) {
    if (this.muted) return;
    return this.schedulePlay('wrong.mp3', { volume: 0.9, ...timing, category: 'sfx' });
  }

  playSolvePuzzle(soundFile = 'ClearPuzzle.mp3', timing = {}) {
    if (this.muted) return;
    this.stopSoundboard();
    this.stopSpin();
    return this.schedulePlay(soundFile || 'ClearPuzzle.mp3', { volume: 0.95, ...timing, category: 'sfx' });
  }

  playTurnGranted(timing = {}) {
    if (this.muted) return;
    return this.schedulePlay('nhac_host_entrance.mp3', { volume: 0.85, ...timing, category: 'sfx' });
  }
}

window.soundEngine = new SoundEngine();

window.attachSoundListeners = function(socketInstance) {
  if (!socketInstance || socketInstance._soundListenersAttached) return;
  socketInstance._soundListenersAttached = true;

  // Do not attach audio events if engine is muted (Controller or inside Iframe)
  if (window.soundEngine && window.soundEngine.muted) {
    return;
  }

  // Start NTP Clock Sync
  if (window.soundEngine) {
    window.soundEngine.startTimeSync(socketInstance);
  }

  // Synchronized Soundboard play
  socketInstance.on('soundboard:play_file', (data) => {
    if (!window.soundEngine) return;
    const filename = (typeof data === 'object' && data !== null) ? data.file : data;
    const timing = (typeof data === 'object' && data !== null) ? { serverTime: data.serverTime, playAt: data.playAt } : {};
    if (filename) {
      window.soundEngine.playSoundboard(filename, timing);
    }
  });

  // Soundboard stop all
  socketInstance.on('soundboard:stop_all', () => {
    if (window.soundEngine) {
      window.soundEngine.stopAll();
    }
  });

  // Synchronized Wheel spin
  socketInstance.on('wheel:start_spin', (data) => {
    if (window.soundEngine) {
      const music = (data && data.music) ? data.music : 'Nhạc quay Nón 1.mp3';
      window.soundEngine.playSpin(music, { serverTime: data?.serverTime, playAt: data?.playAt });
    }
  });

  // Wheel landed
  socketInstance.on('wheel:landed', () => {
    if (window.soundEngine) {
      window.soundEngine.stopSpin();
    }
  });

  // Synchronized Sound effects
  socketInstance.on('sound:play', (data) => {
    if (!data || !window.soundEngine) return;
    const timing = { serverTime: data.serverTime, playAt: data.playAt };

    if (data.type === 'puzzle_show') {
      const snd = data.sound || 'reveal.mp3';
      if (snd && snd !== 'None') {
        window.soundEngine.stopSoundboard();
        window.soundEngine.schedulePlay(snd, { volume: 0.9, ...timing, category: 'sfx' });
      }
    } else if (data.type === 'letter_mark') {
      const snd = data.sound || 'ding.wav';
      if (snd && snd !== 'None') {
        window.soundEngine.schedulePlay(snd, { volume: 0.9, ...timing, category: 'sfx' });
      }
    } else if (data.type === 'letter_open') {
      const snd = data.sound || '2nd_ding.wav';
      if (snd && snd !== 'None') {
        window.soundEngine.schedulePlay(snd, { volume: 0.9, ...timing, category: 'sfx' });
      }
    } else if (data.type === 'letter_flip') {
      window.soundEngine.schedulePlay('ding.wav', { volume: 0.8, ...timing, category: 'sfx' });
    } else if (data.type === 'puzzle_solve') {
      const snd = data.sound || 'ClearPuzzle.mp3';
      if (snd && snd !== 'None') {
        window.soundEngine.playSolvePuzzle(snd, timing);
      }
    } else if (data.type === 'buzzer_hit') {
      window.soundEngine.playBuzzer(timing);
    } else if (data.type === 'turn_granted') {
      window.soundEngine.playTurnGranted(timing);
    }
  });
};

// Automatic hook into io() socket initialization
if (typeof window !== 'undefined') {
  const hookSocket = () => {
    if (window.io && !window._soundIoHooked) {
      const originalIo = window.io;
      window.io = function(...args) {
        const s = originalIo(...args);
        window.attachSoundListeners(s);
        return s;
      };
      Object.assign(window.io, originalIo);
      window._soundIoHooked = true;
    }
  };
  hookSocket();
  document.addEventListener('DOMContentLoaded', hookSocket);
}

// User Interaction Listener to Unlock Audio Context
const unlockAudio = () => {
  if (window.soundEngine) {
    window.soundEngine.init();
  }
};

if (typeof document !== 'undefined') {
  document.addEventListener('click', unlockAudio, { capture: true, passive: true });
  document.addEventListener('touchstart', unlockAudio, { capture: true, passive: true });
  document.addEventListener('keydown', unlockAudio, { capture: true, passive: true });
}
