// Resizable sidebar implementation using Pointer Events for reliable dragging (mouse/touch/stylus).
(function () {
  const handle = document.getElementById('dragHandle');
  const sidebar = document.getElementById('sidebar');
  const container = document.querySelector('.app-content');
  const STORAGE_KEY = 'elig_sidebar_width';

  if (!handle || !sidebar || !container) return;

  // Minimum and maximum sizes (in px / pct)
  const MIN_WIDTH = 220;
  const MAX_WIDTH_PCT = 0.75; // max 75% of container width
  let dragging = false;
  let startClientX = 0;
  let startWidth = 0;
  let pointerId = null;

  // Initialize from persisted width if available
  function applyStored() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const px = parseInt(raw, 10);
      if (isNaN(px)) return;
      const maxAllowed = Math.floor(container.clientWidth * MAX_WIDTH_PCT);
      const widthToSet = Math.min(Math.max(px, MIN_WIDTH), maxAllowed);
      sidebar.style.flex = `0 0 ${widthToSet}px`;
      sidebar.style.width = `${widthToSet}px`;
    } catch (e) { /* ignore storage errors */ }
  }
  applyStored();

  function startDrag(e) {
    // Accept pointerdown only (mouse/touch/stylus)
    if (dragging) return;
    dragging = true;
    // For pointer events, store pointerId to release capture later
    pointerId = e.pointerId || null;
    startClientX = e.clientX !== undefined ? e.clientX : (e.touches && e.touches[0] && e.touches[0].clientX) || 0;
    startWidth = sidebar.getBoundingClientRect().width;

    // Freeze text selection & set dragging cursor
    document.body.style.userSelect = 'none';
    document.documentElement.classList.add('dragging');

    // Use pointer capture if available (improves reliability)
    try { if (pointerId && handle.setPointerCapture) handle.setPointerCapture(pointerId); } catch (err) {}

    window.addEventListener('pointermove', onDrag);
    window.addEventListener('pointerup', endDrag);
    window.addEventListener('pointercancel', endDrag);
  }

  function onDrag(e) {
    if (!dragging) return;
    // Prevent default scrolling on touch
    if (e.cancelable) e.preventDefault();

    const clientX = e.clientX !== undefined ? e.clientX : (e.touches && e.touches[0] && e.touches[0].clientX) || 0;
    const delta = clientX - startClientX;
    const containerRect = container.getBoundingClientRect();
    const maxAllowed = Math.floor(containerRect.width * MAX_WIDTH_PCT);
    let newWidth = Math.round(startWidth + delta);
    newWidth = Math.max(MIN_WIDTH, Math.min(newWidth, maxAllowed));

    // Apply width as fixed flex-basis so main column shrinks/grows properly
    sidebar.style.flex = `0 0 ${newWidth}px`;
    sidebar.style.width = `${newWidth}px`;

    // Persist
    try { localStorage.setItem(STORAGE_KEY, String(newWidth)); } catch (err) {}
    // update fade overlay if needed (results area)
    if (window.__elig_updateResultsFade) window.__elig_updateResultsFade();
  }

  function endDrag(e) {
    if (!dragging) return;
    dragging = false;
    document.body.style.userSelect = '';
    document.documentElement.classList.remove('dragging');

    // release pointer capture
    try { if (pointerId && handle.releasePointerCapture) handle.releasePointerCapture(pointerId); } catch (err) {}

    window.removeEventListener('pointermove', onDrag);
    window.removeEventListener('pointerup', endDrag);
    window.removeEventListener('pointercancel', endDrag);
  }

  // Keyboard accessibility: left/right to shrink/expand
  handle.addEventListener('keydown', (ev) => {
    const step = 20;
    if (ev.key === 'ArrowLeft' || ev.key === 'Left') {
      const cur = sidebar.getBoundingClientRect().width;
      const next = Math.max(MIN_WIDTH, cur - step);
      sidebar.style.flex = `0 0 ${next}px`;
      sidebar.style.width = `${next}px`;
      try { localStorage.setItem(STORAGE_KEY, String(next)); } catch (err) {}
      ev.preventDefault();
      if (window.__elig_updateResultsFade) window.__elig_updateResultsFade();
    } else if (ev.key === 'ArrowRight' || ev.key === 'Right') {
      const containerRect = container.getBoundingClientRect();
      const maxAllowed = Math.floor(containerRect.width * MAX_WIDTH_PCT);
      const cur = sidebar.getBoundingClientRect().width;
      const next = Math.min(maxAllowed, cur + step);
      sidebar.style.flex = `0 0 ${next}px`;
      sidebar.style.width = `${next}px`;
      try { localStorage.setItem(STORAGE_KEY, String(next)); } catch (err) {}
      ev.preventDefault();
      if (window.__elig_updateResultsFade) window.__elig_updateResultsFade();
    } else if (ev.key === 'Home') {
      sidebar.style.flex = `0 0 ${MIN_WIDTH}px`;
      sidebar.style.width = `${MIN_WIDTH}px`;
      try { localStorage.setItem(STORAGE_KEY, String(MIN_WIDTH)); } catch (err) {}
      ev.preventDefault();
      if (window.__elig_updateResultsFade) window.__elig_updateResultsFade();
    } else if (ev.key === 'End') {
      const containerRect = container.getBoundingClientRect();
      const maxAllowed = Math.floor(containerRect.width * MAX_WIDTH_PCT);
      sidebar.style.flex = `0 0 ${maxAllowed}px`;
      sidebar.style.width = `${maxAllowed}px`;
      try { localStorage.setItem(STORAGE_KEY, String(maxAllowed)); } catch (err) {}
      ev.preventDefault();
      if (window.__elig_updateResultsFade) window.__elig_updateResultsFade();
    }
  });

  // Attach pointerdown (works for mouse/touch/stylus)
  handle.addEventListener('pointerdown', startDrag, { passive: false });

  // Also support mouse for older browsers as fallback
  handle.addEventListener('mousedown', (e) => {
    // Let pointerdown handle if available
    if (window.PointerEvent) return;
    startDrag(e);
  });

  // Ensure stored width stays within new constraints on resize
  window.addEventListener('resize', () => {
    const curWidth = sidebar.getBoundingClientRect().width;
    const maxAllowed = Math.floor(container.clientWidth * MAX_WIDTH_PCT);
    if (curWidth > maxAllowed) {
      const newW = Math.floor(Math.max(MIN_WIDTH, Math.min(curWidth, maxAllowed)));
      sidebar.style.flex = `0 0 ${newW}px`;
      sidebar.style.width = `${newW}px`;
      try { localStorage.setItem(STORAGE_KEY, String(newW)); } catch (err) {}
    }
  });
})();
