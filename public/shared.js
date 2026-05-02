// ═══════════════════════════════════════════════════════════════
// MGSDB2026 — Shared Sidebar + Layout
// ═══════════════════════════════════════════════════════════════

const NAV = [
  { section: 'Main', items: [
    { icon: '🏠', label: 'Dashboard', url: 'index.html' }
  ]},
  { section: 'People', items: [
    { icon: '👷', label: 'Workers', url: 'workers.html' },
    { icon: '💼', label: 'Employers', url: 'employers.html' },
    { icon: '🏭', label: 'Suppliers', url: 'suppliers.html' },
    { icon: '🏛️', label: 'Local Agencies', url: 'agencies.html' }
  ]},
  { section: 'Reference Data', items: [
    { icon: '🌍', label: 'Nationalities', url: 'nationalities.html' }
  ]},
  { section: 'Operations', items: [
    { icon: '🪪', label: 'Compliance Services', url: 'compliance.html' },
    { icon: '🧾', label: 'Accountant', url: 'accountant.html' },
    { icon: '📊', label: 'Reports', url: 'reports.html' }
  ]},
  { section: 'System', items: [
    { icon: '⚙️', label: 'Settings', url: 'settings.html' }
  ]}
];

function renderLayout(pageTitle, currentUrl) {
  const currentFile = currentUrl || (window.location.pathname.split('/').pop() || 'index.html');

  const sidebarHTML = `
    <div class="sidebar-header">
      <h2>📋 MGSDB2026</h2>
      <p>Massroua General Services</p>
    </div>
    <div class="sidebar-nav">
      ${NAV.map(sec => `
        <div class="nav-section">${sec.section}</div>
        ${sec.items.map(item => `
          <a href="${item.url}" class="nav-item ${item.url === currentFile ? 'active' : ''}">
            <span class="ni">${item.icon}</span>
            <span class="nl">${item.label}</span>
          </a>
        `).join('')}
      `).join('')}
    </div>
    <div class="sidebar-footer">
      <div class="cloud-status">
        <div class="cloud-dot ok" id="cloudDot"></div>
        <span id="cloudStatusTxt">Connected</span>
      </div>
    </div>
  `;

  const toolbarHTML = `
    <h1 id="toolbarTitle">${pageTitle}</h1>
    <button class="btn btn-light" onclick="location.reload()">🔄 Refresh</button>
  `;

  return { sidebarHTML, toolbarHTML };
}

function setCloud(ok, msg) {
  const dot = document.getElementById('cloudDot');
  const txt = document.getElementById('cloudStatusTxt');
  if (dot) dot.className = 'cloud-dot ' + (ok ? 'ok' : 'err');
  if (txt) txt.textContent = msg || (ok ? 'Connected' : 'Error');
}

// Mount the layout into a page
function mountLayout(pageTitle, contentHTML) {
  const { sidebarHTML, toolbarHTML } = renderLayout(pageTitle);
  document.body.innerHTML = `
    <div class="layout">
      <div class="sidebar">${sidebarHTML}</div>
      <div class="main">
        <div class="main-toolbar">${toolbarHTML}</div>
        <div class="main-content" id="mainContent">${contentHTML}</div>
      </div>
    </div>
  `;
}