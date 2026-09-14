// Disable DevTools and Right-click inspect across the entire application
(function () {
  // Disable Right Click context menu
  document.addEventListener('contextmenu', function (e) {
    e.preventDefault();
    return false;
  }, { capture: true });

  // Disable DevTools Keyboard Shortcuts
  document.addEventListener('keydown', function (e) {
    // F12
    if (e.keyCode === 123 || e.key === 'F12') {
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
    // Ctrl+Shift+I / Cmd+Option+I (Inspect)
    // Ctrl+Shift+J / Cmd+Option+J (Console)
    // Ctrl+Shift+C / Cmd+Option+C (Inspect Element)
    if ((e.ctrlKey || e.metaKey) && e.shiftKey && 
        (e.key === 'I' || e.key === 'i' || e.key === 'J' || e.key === 'j' || e.key === 'C' || e.key === 'c')) {
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
    // Ctrl+U / Cmd+Option+U (View Source)
    if ((e.ctrlKey || e.metaKey) && (e.key === 'U' || e.key === 'u')) {
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
    // Ctrl+S / Cmd+S (Save page)
    if ((e.ctrlKey || e.metaKey) && (e.key === 'S' || e.key === 's')) {
      e.preventDefault();
      e.stopPropagation();
      return false;
    }
  }, { capture: true });
})();
