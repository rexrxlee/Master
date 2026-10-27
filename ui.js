// Charts created inside collapsed sections need their visible dimensions on open.
document.addEventListener("toggle", event => {
  if (event.target.tagName !== "DETAILS" || !event.target.open) return;
  requestAnimationFrame(() => window.dispatchEvent(new Event("resize")));
}, true);

const insuranceNav = document.querySelector('.sidebar-nav');
if (insuranceNav && !insuranceNav.querySelector('a[href="insurance.html"]')) {
  const link = document.createElement('a'); link.className='nav-item'; link.href='insurance.html'; link.textContent='Insurance'; insuranceNav.appendChild(link);
}
