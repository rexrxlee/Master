// Charts created inside collapsed sections need their visible dimensions on open.
document.addEventListener("toggle", event => {
  if (event.target.tagName !== "DETAILS" || !event.target.open) return;
  requestAnimationFrame(() => window.dispatchEvent(new Event("resize")));
}, true);
