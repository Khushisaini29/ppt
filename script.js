(() => {
  const slides = Array.from(document.querySelectorAll('.slide'));
  const counter = document.querySelector('#counter');
  const progress = document.querySelector('#progress');
  const prev = document.querySelector('#prev');
  const next = document.querySelector('#next');
  const focusToggle = document.querySelector('#focusToggle');
  let current = 0;
  let touchStartX = 0;

  function goTo(index) {
    const target = Math.max(0, Math.min(index, slides.length - 1));
    if (target === current && slides[current].classList.contains('active')) return;
    slides[current].classList.remove('active');
    current = target;
    slides[current].classList.add('active');
    slides[current].scrollTop = 0;
    counter.textContent = `${String(current + 1).padStart(2, '0')} / ${String(slides.length).padStart(2, '0')}`;
    progress.style.transform = `scaleX(${(current + 1) / slides.length})`;
    prev.disabled = current === 0;
    next.disabled = current === slides.length - 1;
    document.title = `${slides[current].dataset.title} | Introduction to Web Development`;
  }

  prev.addEventListener('click', () => goTo(current - 1));
  next.addEventListener('click', () => goTo(current + 1));
  focusToggle.addEventListener('click', () => document.body.classList.toggle('presentation-mode'));
  document.addEventListener('keydown', event => {
    if (event.target.matches('button, a, input, textarea, select')) return;
    if (event.key === 'ArrowRight' || event.key === 'PageDown' || event.key === ' ') { event.preventDefault(); goTo(current + 1); }
    if (event.key === 'ArrowLeft' || event.key === 'PageUp') { event.preventDefault(); goTo(current - 1); }
    if (event.key.toLowerCase() === 'p' || event.key.toLowerCase() === 'f') document.body.classList.toggle('presentation-mode');
    if (event.key === 'Home') goTo(0);
    if (event.key === 'End') goTo(slides.length - 1);
  });
  document.addEventListener('touchstart', event => { touchStartX = event.changedTouches[0].screenX; }, { passive: true });
  document.addEventListener('touchend', event => {
    const distance = event.changedTouches[0].screenX - touchStartX;
    if (Math.abs(distance) > 60) goTo(current + (distance < 0 ? 1 : -1));
  }, { passive: true });
  goTo(0);
})();
