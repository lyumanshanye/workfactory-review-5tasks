(() => {
  const tabs = [...document.querySelectorAll('[role="tab"]')];
  function activate(tab) {
    tabs.forEach(item => {
      const selected = item === tab;
      item.setAttribute('aria-selected', String(selected));
      item.tabIndex = selected ? 0 : -1;
      document.getElementById(item.getAttribute('aria-controls')).hidden = !selected;
    });
    history.replaceState(null, '', '#' + tab.getAttribute('aria-controls'));
  }
  tabs.forEach((tab, index) => {
    tab.addEventListener('click', () => activate(tab));
    tab.addEventListener('keydown', event => {
      let next;
      if (event.key === 'ArrowRight') next = tabs[(index + 1) % tabs.length];
      if (event.key === 'ArrowLeft') next = tabs[(index + tabs.length - 1) % tabs.length];
      if (event.key === 'Home') next = tabs[0];
      if (event.key === 'End') next = tabs[tabs.length - 1];
      if (next) { event.preventDefault(); activate(next); next.focus(); }
    });
  });
  function restoreTab() {
    const current = tabs.find(tab => '#' + tab.getAttribute('aria-controls') === location.hash);
    activate(current || tabs[0]);
  }
  restoreTab();
  window.addEventListener('hashchange', restoreTab);
  const query = document.getElementById('query');
  const group = document.getElementById('group');
  const rows = [...document.querySelectorAll('.rubric-row')];
  function filter() {
    const term = query.value.trim().toLocaleLowerCase();
    let visible = 0;
    rows.forEach(row => {
      row.hidden = !(row.textContent.toLocaleLowerCase().includes(term) && (group.value === 'all' || row.dataset.group === group.value));
      if (!row.hidden) visible++;
    });
    document.getElementById('result-count').textContent = `${visible} / ${rows.length} 条`;
    document.getElementById('empty').hidden = visible > 0;
  }
  query.addEventListener('input', filter);
  group.addEventListener('change', filter);
  document.getElementById('expanded').addEventListener('change', event => {
    rows.forEach(row => { row.open = event.target.checked; });
  });
  document.body.dataset.ready = 'true';
})();
