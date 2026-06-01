/* ===================================================
   PILATESRAUM BY MALU — Präsentation Navigation
=================================================== */

(function () {
  'use strict';

  const slides = document.querySelectorAll('.slide');
  const counter = document.getElementById('slide-counter');
  const btnPrev = document.getElementById('btn-prev');
  const btnNext = document.getElementById('btn-next');
  const btnScroll = document.getElementById('btn-scroll');

  let current = 0;
  let scrollMode = false;

  const total = slides.length;

  function goTo(index) {
    if (scrollMode) return;
    slides[current].classList.remove('active');
    current = Math.max(0, Math.min(total - 1, index));
    slides[current].classList.add('active');
    updateCounter();
    updateButtons();
  }

  function updateCounter() {
    counter.textContent = (current + 1) + ' / ' + total;
  }

  function updateButtons() {
    btnPrev.disabled = current === 0;
    btnNext.disabled = current === total - 1;
  }

  function toggleScrollMode() {
    scrollMode = !scrollMode;
    document.body.classList.toggle('scroll-mode', scrollMode);
    if (scrollMode) {
      btnScroll.title = 'Folien-Modus';
      btnScroll.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="3" width="20" height="14" rx="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>';
    } else {
      btnScroll.title = 'Scroll-Modus';
      btnScroll.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><line x1="12" y1="5" x2="12" y2="19"/><polyline points="19 12 12 19 5 12"/></svg>';
      goTo(current);
    }
  }

  // Init
  goTo(0);

  // Buttons
  btnPrev.addEventListener('click', () => goTo(current - 1));
  btnNext.addEventListener('click', () => goTo(current + 1));
  btnScroll.addEventListener('click', toggleScrollMode);

  // Keyboard
  document.addEventListener('keydown', function (e) {
    if (scrollMode) return;
    if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === 'PageDown' || e.key === ' ') {
      e.preventDefault();
      goTo(current + 1);
    } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
      e.preventDefault();
      goTo(current - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      goTo(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      goTo(total - 1);
    }
  });

  // Swipe support
  let touchStartX = 0;
  document.addEventListener('touchstart', function (e) {
    touchStartX = e.changedTouches[0].screenX;
  }, { passive: true });

  document.addEventListener('touchend', function (e) {
    if (scrollMode) return;
    const delta = e.changedTouches[0].screenX - touchStartX;
    if (Math.abs(delta) > 50) {
      if (delta < 0) goTo(current + 1);
      else goTo(current - 1);
    }
  });

  // Auto-hide placeholder if image loads
  document.querySelectorAll('.screenshot-box img').forEach(function (img) {
    function onLoad() {
      img.classList.add('loaded');
      const ph = img.nextElementSibling;
      if (ph && ph.classList.contains('placeholder-label')) {
        ph.style.display = 'none';
      }
    }
    if (img.complete && img.naturalWidth > 0) {
      onLoad();
    } else {
      img.addEventListener('load', onLoad);
    }
  });

  // Mockup card selection
  document.querySelectorAll('.mockup-card').forEach(function (card) {
    card.addEventListener('click', function () {
      const siblings = card.parentElement.querySelectorAll('.mockup-card');
      siblings.forEach(function (s) { s.classList.remove('selected'); });
      card.classList.add('selected');
    });
  });

})();
