// ============================================
// MARIA365.ONLINE - JavaScript Principal
// ============================================

document.addEventListener('DOMContentLoaded', function () {

  // --- Menú móvil ---
  const navToggle = document.querySelector('.nav-toggle');
  const navMenu = document.querySelector('.nav-menu');

  if (navToggle && navMenu) {
    navToggle.addEventListener('click', function () {
      navMenu.classList.toggle('abierto');
      navToggle.textContent = navMenu.classList.contains('abierto') ? '✕' : '☰';
    });

    // Cerrar menú al hacer click en un link (excluye el toggle del dropdown)
    navMenu.querySelectorAll('a').forEach(function (link) {
      link.addEventListener('click', function () {
        navMenu.classList.remove('abierto');
        navToggle.textContent = '☰';
      });
    });
  }

  // --- Marcar link activo en navbar ---
  const paginaActual = window.location.pathname.split('/').pop();
  document.querySelectorAll('.nav-menu a').forEach(function (link) {
    const href = link.getAttribute('href').split('/').pop();
    if (href === paginaActual || (paginaActual === '' && href === 'index.html')) {
      link.classList.add('activo');
    }
  });

  // --- Tabs ---
  document.querySelectorAll('.tab-btn').forEach(function (btn) {
    btn.addEventListener('click', function () {
      const grupo = this.closest('.tabs-container');
      const target = this.dataset.tab;

      grupo.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('activo'));
      grupo.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('activo'));

      this.classList.add('activo');
      const panel = grupo.querySelector('#' + target);
      if (panel) panel.classList.add('activo');
    });
  });

  // --- Acordeón ---
  document.querySelectorAll('.acordeon-header').forEach(function (header) {
    header.addEventListener('click', function () {
      const body = this.nextElementSibling;
      const estaAbierto = this.classList.contains('activo');

      // Cerrar todos en el mismo grupo
      const parent = this.closest('.acordeon-group') || document;
      parent.querySelectorAll('.acordeon-header').forEach(h => {
        h.classList.remove('activo');
        h.nextElementSibling.classList.remove('activo');
      });

      // Abrir el clickeado si estaba cerrado
      if (!estaAbierto) {
        this.classList.add('activo');
        body.classList.add('activo');
      }
    });
  });

  // --- Oraciones toggle ---
  document.querySelectorAll('.oracion-header').forEach(function (header) {
    header.addEventListener('click', function () {
      const body = this.nextElementSibling;
      const icon = this.querySelector('.oracion-toggle-icon');
      const estaAbierto = body.classList.contains('activo');

      document.querySelectorAll('.oracion-body').forEach(b => b.classList.remove('activo'));
      document.querySelectorAll('.oracion-toggle-icon').forEach(i => i.textContent = '+');

      if (!estaAbierto) {
        body.classList.add('activo');
        if (icon) icon.textContent = '−';
      }
    });
  });

  // --- Scroll suave para anclas ---
  document.querySelectorAll('a[href^="#"]').forEach(function (anchor) {
    anchor.addEventListener('click', function (e) {
      const target = document.querySelector(this.getAttribute('href'));
      if (target) {
        e.preventDefault();
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

  // --- Animación de entrada para cards ---
  if ('IntersectionObserver' in window) {
    const observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.style.opacity = '1';
          entry.target.style.transform = 'translateY(0)';
          observer.unobserve(entry.target);
        }
      });
    }, { threshold: 0.1 });

    document.querySelectorAll('.card, .card-info, .mes-card, .dogma-card, .tema-card, .pais-item, .qs-stat, .pub-card').forEach(function (el) {
      el.style.opacity = '0';
      el.style.transform = 'translateY(20px)';
      el.style.transition = 'opacity 0.5s ease, transform 0.5s ease, box-shadow 0.3s ease';
      observer.observe(el);
    });
  }

  // --- Carruseles de publicaciones ---
  document.querySelectorAll('.pub-carousel-wrapper').forEach(function (wrapper) {
    const carousel = wrapper.querySelector('.pub-carousel');
    const btnPrev = wrapper.querySelector('.carousel-btn-prev');
    const btnNext = wrapper.querySelector('.carousel-btn-next');
    if (!carousel) return;

    var scrollAmount = 300;

    if (btnPrev) {
      btnPrev.addEventListener('click', function () {
        carousel.scrollBy({ left: -scrollAmount, behavior: 'smooth' });
      });
    }

    if (btnNext) {
      btnNext.addEventListener('click', function () {
        carousel.scrollBy({ left: scrollAmount, behavior: 'smooth' });
      });
    }
  });

});
