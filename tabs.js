const groups = ['operation', 'support'];

function selectGroup(selected) {
  groups.forEach(group => {
    const active = group === selected;
    const tab = document.getElementById(`tab-${group}`);
    const panel = document.getElementById(`panel-${group}`);
    tab.classList.toggle('active', active);
    tab.setAttribute('aria-selected', String(active));
    panel.classList.toggle('hidden', !active);

    const frame = panel.querySelector('iframe');
    if (active && !frame.src && frame.dataset.src) frame.src = frame.dataset.src;
  });
}

groups.forEach(group => {
  const tab = document.getElementById(`tab-${group}`);
  tab.setAttribute('role', 'tab');
  tab.addEventListener('click', () => selectGroup(group));
});

selectGroup('operation');
