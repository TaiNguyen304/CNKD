// Sound Asset Player
class SoundEngine {
  constructor() {
    this.ctx = null;
    const isController = typeof window !== 'undefined' && (
      window.location.pathname.toLowerCase().includes('controller') ||
      window.location.href.toLowerCase().includes('controller')
    );
    this.muted = isController;
    this.audioElements = {};
    this.activeSpinAudio = null;
    this.activeSoundboardAudio = null;
  }

  init() {
    if (!this.ctx) {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (AudioContext) {
        this.ctx = new AudioContext();
      }
    }
    if (this.ctx && this.ctx.state === 'suspended') {
      this.ctx.resume().catch(() => {});
    }
  }

  playAudioFile(filename, loop = false, volume = 0.8) {
    if (this.muted || !filename) return null;
    this.init();

    // Clean filename
    const cleanName = filename.replace(/^\/+/, '');
    
    // Construct standard URL to file in root / public directory
    const audioUrl = '/' + encodeURI(cleanName);

    const audio = new Audio(audioUrl);
    audio.volume = volume;
    audio.loop = loop;

    const promise = audio.play();
    if (promise !== undefined) {
      promise.catch(err => {
        console.warn('[SoundEngine] Playback issue for:', audioUrl, err);
      });
    }

    return audio;
  }

  playSpin(musicFile) {
    if (this.muted) return;
    this.stopSpin();
    const fileToPlay = musicFile || 'Nhạc quay Nón 1.mp3';
    this.activeSpinAudio = this.playAudioFile(fileToPlay, false, 0.9);
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

  playSoundboard(filename) {
    if (this.muted || !filename) return;
    if (this.lastPlayedFile === filename && (Date.now() - (this.lastPlayedTime || 0) < 300)) {
      return;
    }
    this.lastPlayedFile = filename;
    this.lastPlayedTime = Date.now();
    this.stopSoundboard();
    this.activeSoundboardAudio = this.playAudioFile(filename, false, 0.9);
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

  playClick() {
    // Silent
  }

  playWheelTick() {
    // Silent
  }

  playBuzzer() {
    if (this.muted) return;
    return this.playAudioFile('buzzer.mp3', false, 0.9);
  }

  playLetterCorrect() {
    if (this.muted) return;
    return this.playAudioFile('ding.wav', false, 0.9) || this.playAudioFile('Open_2.mp3', false, 0.9);
  }

  playLetterWrong() {
    if (this.muted) return;
    return this.playAudioFile('wrong.mp3', false, 0.9);
  }

  playSolvePuzzle() {
    if (this.muted) return;
    return this.playAudioFile('solved.mp3', false, 0.9) || this.playAudioFile('solved 2001.mp3', false, 0.9);
  }

  playRoundChange() {
    // Disabled automatic player music on round switch
    return;
  }

  playTurnGranted() {
    if (this.muted) return;
    return this.playAudioFile('nhac_host_entrance.mp3', false, 0.8) ||
           this.playAudioFile('HostEntrance.mp3', false, 0.8);
  }

  playResultEffect(result) {
    return;
  }
}

window.soundEngine = new SoundEngine();

window.attachSoundListeners = function(socketInstance) {
  if (!socketInstance || socketInstance._soundListenersAttached) return;
  socketInstance._soundListenersAttached = true;

  const isController = typeof window !== 'undefined' && (
    window.location.pathname.toLowerCase().includes('controller') ||
    window.location.href.toLowerCase().includes('controller')
  );
  if (isController) {
    if (window.soundEngine) {
      window.soundEngine.muted = true;
    }
    return;
  }

  socketInstance.on('soundboard:play_file', (filename) => {
    if (window.soundEngine && filename) {
      window.soundEngine.playSoundboard(filename);
    }
  });

  socketInstance.on('soundboard:stop_all', () => {
    if (window.soundEngine) {
      window.soundEngine.stopSoundboard();
      window.soundEngine.stopSpin();
    }
  });

  socketInstance.on('wheel:start_spin', (data) => {
    if (window.soundEngine) {
      const music = (data && data.music) ? data.music : 'Nhạc quay Nón 1.mp3';
      window.soundEngine.playSpin(music);
    }
  });

  socketInstance.on('wheel:landed', () => {
    if (window.soundEngine) {
      window.soundEngine.stopSpin();
    }
  });

  socketInstance.on('sound:play', (data) => {
    if (!data || !window.soundEngine) return;
    if (data.type === 'spin_start') {
      window.soundEngine.playSpin(data.music || 'Nhạc quay Nón 1.mp3');
    } else if (data.type === 'round_change') {
      window.soundEngine.playRoundChange();
    } else if (data.type === 'puzzle_show') {
      const snd = data.sound || 'reveal.mp3';
      if (snd && snd !== 'None') {
        window.soundEngine.stopSoundboard();
        window.soundEngine.playAudioFile(snd);
      }
    } else if (data.type === 'letter_mark') {
      const snd = data.sound || 'ding.wav';
      if (snd && snd !== 'None') window.soundEngine.playAudioFile(snd);
    } else if (data.type === 'letter_open') {
      const snd = data.sound || '2nd_ding.wav';
      if (snd && snd !== 'None') window.soundEngine.playAudioFile(snd);
    } else if (data.type === 'puzzle_solve') {
      const snd = data.sound || 'ClearPuzzle.mp3';
      if (snd && snd !== 'None') {
        window.soundEngine.stopSoundboard();
        window.soundEngine.stopSpin();
        window.soundEngine.playAudioFile(snd);
      }
    } else if (data.type === 'buzzer_hit') {
      window.soundEngine.playBuzzer();
    } else if (data.type === 'letter_correct') {
      window.soundEngine.playLetterCorrect();
    } else if (data.type === 'letter_wrong') {
      window.soundEngine.playLetterWrong();
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

// Invisible User Interaction Listener to Unlock Audio Context
const unlockAudio = () => {
  if (window.soundEngine) {
    window.soundEngine.init();
  }
};

document.addEventListener('click', unlockAudio, { capture: true });
document.addEventListener('touchstart', unlockAudio, { capture: true });
document.addEventListener('keydown', unlockAudio, { capture: true });
