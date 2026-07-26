/* ============================================
   MARIAGE SISSI & GUILLAUME
   Interactions : countdown, menu, FAQ, smooth scroll
   ============================================ */

(() => {
  'use strict';

  // ===== 1. COMPTE À REBOURS =====
  const WEDDING_DATE = new Date('2026-12-05T15:30:00+01:00');

  function updateCountdown() {
    const now = new Date();
    const diff = WEDDING_DATE - now;

    if (diff <= 0) {
      document.getElementById('countdown')?.style.setProperty('display', 'none');
      return;
    }

    const days = Math.floor(diff / (1000 * 60 * 60 * 24));
    const hours = Math.floor((diff / (1000 * 60 * 60)) % 24);
    const minutes = Math.floor((diff / (1000 * 60)) % 60);
    const seconds = Math.floor((diff / 1000) % 60);

    set('days',    String(days).padStart(3, '0'));
    set('hours',   String(hours).padStart(2, '0'));
    set('minutes', String(minutes).padStart(2, '0'));
    set('seconds', String(seconds).padStart(2, '0'));
  }

  function set(unit, value) {
    const el = document.querySelector(`[data-unit="${unit}"]`);
    if (el) el.textContent = value;
  }

  updateCountdown();
  setInterval(updateCountdown, 1000);

  // ===== 2. MENU MOBILE =====
  const burger = document.getElementById('burger');
  const mobileMenu = document.getElementById('mobile-menu');
  const closeBtn = document.getElementById('mobile-menu-close');

  function openMenu() {
    mobileMenu?.classList.add('is-open');
    document.body.style.overflow = 'hidden';
  }
  function closeMenu() {
    mobileMenu?.classList.remove('is-open');
    document.body.style.overflow = '';
  }

  burger?.addEventListener('click', openMenu);
  closeBtn?.addEventListener('click', closeMenu);
  mobileMenu?.querySelectorAll('a').forEach(a => {
    a.addEventListener('click', closeMenu);
  });

  // ===== 3. SMOOTH SCROLL AVEC OFFSET POUR LA NAV FIXE =====
  document.querySelectorAll('a[href^="#"]').forEach(link => {
    link.addEventListener('click', (e) => {
      const targetId = link.getAttribute('href');
      if (targetId === '#') return;
      const target = document.querySelector(targetId);
      if (!target) return;
      e.preventDefault();
      const offset = 80; // hauteur de la nav
      const top = target.getBoundingClientRect().top + window.pageYOffset - offset;
      window.scrollTo({ top, behavior: 'smooth' });
    });
  });

  // ===== 4. LIGHTBOX GALERIE (minimaliste) =====
  const galleryItems = document.querySelectorAll('[data-lightbox]');
  galleryItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const imgSrc = item.getAttribute('href');
      const overlay = document.createElement('div');
      overlay.style.cssText = `
        position: fixed; inset: 0; background: rgba(0,0,0,0.9);
        display: flex; align-items: center; justify-content: center;
        z-index: 999; cursor: pointer; padding: 20px;
      `;
      const img = document.createElement('img');
      img.src = imgSrc;
      img.style.cssText = 'max-width: 90vw; max-height: 90vh; border-radius: 8px;';
      overlay.appendChild(img);
      overlay.addEventListener('click', () => overlay.remove());
      document.body.appendChild(overlay);
    });
  });

  // ===== 5. ANIMATION À L'APPARITION (Intersection Observer) =====
  const animated = document.querySelectorAll('section');
  if ('IntersectionObserver' in window) {
    const io = new IntersectionObserver((entries) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.style.opacity = '1';
          entry.target.style.transform = 'translateY(0)';
        }
      });
    }, { threshold: 0.1 });

    animated.forEach(s => {
      s.style.opacity = '0';
      s.style.transform = 'translateY(30px)';
      s.style.transition = 'opacity 0.8s ease, transform 0.8s ease';
      io.observe(s);
    });
  }

  // ===== 6. ANIMATION DES DREAM CARDS (1 seul ouvert à la fois) =====
  const dreamCards = document.querySelectorAll('.dream__card');
  dreamCards.forEach(card => {
    card.addEventListener('toggle', () => {
      if (card.open) {
        dreamCards.forEach(c => {
          if (c !== card) c.open = false;
        });
      }
    });
  });

})();