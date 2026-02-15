document.addEventListener('click', e => {
  // Find the closest TD element from the click target
  const cell = e.target.closest('td');
  
  // Must be a TD, and not contain or be within a button/input
  if (
    cell &&
    cell.tagName === 'TD' &&
    !e.target.closest('button') &&
    !e.target.closest('input') &&
    !e.target.closest('a')
  ) {
    const text = cell.textContent.trim();
    if (text) {
      console.log('[Clipboard] Copying:', text);
      navigator.clipboard.writeText(text).then(() => {
        // Remove previous permanent highlight
        document.querySelectorAll('td.last-copied').forEach(td => td.classList.remove('last-copied'));

        // Add permanent highlight to this cell
        cell.classList.add('last-copied');

        // Add temporary copied flash effect
        cell.classList.add('copied');
        setTimeout(() => cell.classList.remove('copied'), 800);
      }).catch(err => {
        console.error('Clipboard copy failed:', err);
      });
    }
  }
});
