document.addEventListener('click', e => {
  const cell = e.target;

  // Must be a TD, and not inside or containing a button/input
  if (
    cell.tagName === 'TD' &&
    !cell.closest('button') &&
    !cell.querySelector('button') &&
    !cell.closest('input') &&
    !cell.querySelector('input')
  ) {
    const text = cell.textContent.trim();
    if (text) {
      console.log(text);
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
