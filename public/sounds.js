// Enhanced Sound Synthesizer & Audio Asset Player
class SoundEngine {
  constructor() {
    this.ctx = null;
    this.muted = false;
    this.audioElements = {};
    this.activeSpinAudio = null;
  }

  init() {
    if (!this.ctx) {
      const AudioContext = window.AudioContext || window.webkitAudioContext;
      if (AudioContext) {
        this.ctx = new AudioContext();
      }
    }
    if (this.ctx && this.ctx.state === 'suspended') {
      this.ctx.resume();
    }
  }

  playAudioFile(filename, loop = false, volume = 0.8) {
    if (this.muted) return null;
    try {
      const audio = new Audio(`/${encodeURIComponent(filename)}`);
      audio.volume = volume;
      audio.loop = loop;
      audio.play().catch(() => {});
      return audio;
    } catch (e) {
      return null;
    }
  }

  playSpin() {
    if (this.muted) return;
    this.stopSpin();
    // spin.mp3 plays for the duration of the wheel spin
    this.activeSpinAudio = this.playAudioFile('spin.mp3', false, 0.9);
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

  playClick() {
    if (this.muted) return;
    this.init();
    if (!this.ctx) return;
    const osc = this.ctx.createOscillator();
    const gain = this.ctx.createGain();
    osc.type = 'triangle';
    osc.frequency.setValueAtTime(800, this.ctx.currentTime);
    osc.frequency.exponentialRampToValueAtTime(120, this.ctx.currentTime + 0.04);
    gain.gain.setValueAtTime(0.25, this.ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.01, this.ctx.currentTime + 0.04);
    osc.connect(gain);
    gain.connect(this.ctx.destination);
    osc.start();
    osc.stop(this.ctx.currentTime + 0.04);
  }

  playWheelTick(pitch = 900) {
    if (this.muted) return;
    this.init();
    if (!this.ctx) return;
    const osc = this.ctx.createOscillator();
    const gain = this.ctx.createGain();
    osc.type = 'sawtooth';
    osc.frequency.setValueAtTime(pitch, this.ctx.currentTime);
    osc.frequency.exponentialRampToValueAtTime(150, this.ctx.currentTime + 0.035);
    gain.gain.setValueAtTime(0.3, this.ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.001, this.ctx.currentTime + 0.035);
    osc.connect(gain);
    gain.connect(this.ctx.destination);
    osc.start();
    osc.stop(this.ctx.currentTime + 0.035);
  }

  playBuzzer() {
    if (this.muted) return;
    this.init();
    if (!this.ctx) return;
    const now = this.ctx.currentTime;
    [260, 310].forEach(freq => {
      const osc = this.ctx.createOscillator();
      const gain = this.ctx.createGain();
      osc.type = 'sawtooth';
      osc.frequency.setValueAtTime(freq, now);
      gain.gain.setValueAtTime(0.35, now);
      gain.gain.exponentialRampToValueAtTime(0.01, now + 0.7);
      osc.connect(gain);
      gain.connect(this.ctx.destination);
      osc.start(now);
      osc.stop(now + 0.7);
    });
  }

  playLetterCorrect() {
    if (this.muted) return;
    this.init();
    if (!this.ctx) return;
    const now = this.ctx.currentTime;
    const notes = [1046.50, 1567.98];
    notes.forEach((freq, idx) => {
      const osc = this.ctx.createOscillator();
      const gain = this.ctx.createGain();
      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, now + idx * 0.12);
      gain.gain.setValueAtTime(0.28, now + idx * 0.12);
      gain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.12 + 0.6);
      osc.connect(gain);
      gain.connect(this.ctx.destination);
      osc.start(now + idx * 0.12);
      osc.stop(now + idx * 0.12 + 0.6);
    });
  }

  playLetterWrong() {
    if (this.muted) return;
    this.init();
    if (!this.ctx) return;
    const now = this.ctx.currentTime;
    const osc = this.ctx.createOscillator();
    const gain = this.ctx.createGain();
    osc.type = 'sawtooth';
    osc.frequency.setValueAtTime(140, now);
    osc.frequency.linearRampToValueAtTime(90, now + 0.4);
    gain.gain.setValueAtTime(0.25, now);
    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.4);
    osc.connect(gain);
    gain.connect(this.ctx.destination);
    osc.start(now);
    osc.stop(now + 0.4);
  }

  playSolvePuzzle() {
    if (this.muted) return;
    this.playAudioFile('Intro.mp3', false, 0.7);
  }

  playRoundChange() {
    if (this.muted) return;
    this.playAudioFile('Người chơi.mp3', false, 0.7);
  }

  playTurnGranted() {
    if (this.muted) return;
    this.playAudioFile('HostEntrance.mp3', false, 0.6);
  }

  playResultEffect(result) {
    // Bỏ luôn các âm thanh tự sinh dựa trên cái Nón chứa text cũ (Coi như nón text cũ không tồn tại)
    return;
  }
}

window.soundEngine = new SoundEngine();
document.addEventListener('click', () => {
  if (window.soundEngine) window.soundEngine.init();
}, { once: true });
