(() => {
  const button = document.createElement('button');
  button.className = 'sound-toggle';
  button.type = 'button';
  button.setAttribute('aria-pressed', 'false');
  button.setAttribute('aria-label', '배경 음악 재생');
  button.innerHTML = '<i></i><span>MUSIC ON</span>';
  document.body.append(button);

  let context, timer, playing = false;
  const melody = [523.25, 659.25, 783.99, 659.25, 587.33, 698.46, 880, 783.99];

  function note(frequency, when, duration, volume, type = 'triangle') {
    const osc = context.createOscillator();
    const gain = context.createGain();
    osc.type = type;
    osc.frequency.setValueAtTime(frequency, when);
    gain.gain.setValueAtTime(0.0001, when);
    gain.gain.exponentialRampToValueAtTime(volume, when + 0.018);
    gain.gain.exponentialRampToValueAtTime(0.0001, when + duration);
    osc.connect(gain).connect(context.destination);
    osc.start(when);
    osc.stop(when + duration + 0.03);
  }

  function playPhrase() {
    if (!playing) return;
    const start = context.currentTime + 0.04;
    melody.forEach((frequency, index) => {
      const time = start + index * 0.31;
      note(frequency, time, 0.24, 0.075);
      if (index % 2 === 0) note(frequency * 2, time + 0.055, 0.12, 0.018, 'sine');
    });
    timer = window.setTimeout(playPhrase, 3000);
  }

  button.addEventListener('click', async () => {
    if (!context) context = new (window.AudioContext || window.webkitAudioContext)();
    if (context.state === 'suspended') await context.resume();
    playing = !playing;
    button.setAttribute('aria-pressed', String(playing));
    button.setAttribute('aria-label', playing ? '배경 음악 정지' : '배경 음악 재생');
    button.querySelector('span').textContent = playing ? 'MUSIC OFF' : 'MUSIC ON';
    if (playing) playPhrase(); else window.clearTimeout(timer);
  });
})();
