<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Address Book — MGS</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>
<style>
  body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; background: #f5f5f5; }
  .topbar { background: #1976d2; color: white; padding: 14px 20px; font-size: 18px; font-weight: 600; }
  .container { padding: 20px; }
  .toolbar { display: flex; gap: 12px; align-items: center; margin-bottom: 16px; flex-wrap: wrap; }
  .btn-primary { background: #1976d2; color: white; border: none; padding: 9px 16px; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 500; }
  .btn-primary:hover { background: #1565c0; }
  .btn-secondary { background: #eee; color: #333; border: none; padding: 9px 14px; border-radius: 6px; cursor: pointer; font-size: 14px; }
  .btn-secondary:hover { background: #ddd; }
  #searchBox { flex: 1; min-width: 220px; padding: 9px 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 14px; }
  .count { color: #666; font-size: 13px; }
  .status { padding: 10px 14px; border-radius: 6px; background: #fff3cd; color: #856404; border: 1px solid #ffeeba; display: inline-block; margin-bottom: 12px; }
  .status.ok { background: #d4edda; color: #155724; border-color: #c3e6cb; }
  .status.err { background: #f8d7da; color: #721c24; border-color: #f5c6cb; }
  .contacts-table { width: 100%; background: white; border-collapse: collapse; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
  .contacts-table th { background: #f8f9fa; padding: 12px; text-align: left; font-size: 13px; color: #555; border-bottom: 2px solid #e0e0e0; }
  .contacts-table td { padding: 11px 12px; border-bottom: 1px solid #f0f0f0; font-size: 14px; }
  .contacts-table tr:hover { background: #f9f9f9; }
  .empty { padding: 40px; text-align: center; color: #888; background: white; border-radius: 8px; border: 2px dashed #ddd; }
  .row-btn { background: #eee; border: none; padding: 5px 10px; border-radius: 4px; cursor: pointer; margin-right: 4px; font-size: 12px; }
  .row-btn:hover { background: #ddd; }

  /* ===== Popup ===== */
  .modal-backdrop { position: fixed; inset: 0; background: rgba(0,0,0,0.35); z-index: 999; display: none; }
  .modal { position: fixed; top: 60px; left: 50%; transform: translateX(-50%); width: 780px; max-width: 95vw; max-height: 88vh; background: white; border-radius: 10px; box-shadow: 0 12px 40px rgba(0,0,0,0.25); z-index: 1000; display: none; flex-direction: column; overflow: hidden; }
  .modal.open { display: flex; }
  .modal-header { background: #1976d2; color: white; padding: 12px 18px; cursor: move; display: flex; align-items: center; justify-content: space-between; user-select: none; }
  .modal-header h3 { margin: 0; font-size: 16px; font-weight: 600; }
  .modal-header .close-btn { background: transparent; border: none; color: white; font-size: 22px; cursor: pointer; padding: 0 6px; }
  .modal-body { flex: 1; overflow-y: auto; padding: 14px 18px; }
  .modal-footer { padding: 12px 18px; border-top: 1px solid #eee; display: flex; justify-content: space-between; gap: 8px; background: #fafafa; }
  .footer-left { display: flex; align-items: center; gap: 8px; }
  .footer-right { display: flex; gap: 8px; }

  /* Reorder toggle */
  .reorder-toggle { font-size: 12px; padding: 6px 12px; border-radius: 5px; background: #f0f0f0; color: #555; border: 1px solid #ddd; cursor: pointer; }
  .reorder-toggle.on { background: #fff3cd; color: #856404; border-color: #ffeeba; }

  /* ===== Accordion ===== */
  .group { border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 10px; overflow: hidden; }
  .group-header { background: #f8f9fa; padding: 12px 14px; cursor: pointer; display: flex; align-items: center; justify-content: space-between; font-weight: 500; font-size: 14px; color: #333; }
  .group-header:hover { background: #eef3f7; }
  .group-header .ar { color: #666; font-size: 13px; font-weight: 400; margin-left: 8px; }
  .group-header .chevron { transition: transform 0.2s; display: inline-block; margin-right: 6px; }
  .group.open .group-header { background: #e3f2fd; color: #1976d2; }
  .group.open .chevron { transform: rotate(90deg); }
  .group-body { display: none; padding: 12px; background: white; }
  .group.open .group-body { display: block; }

  /* ===== Field rows ===== */
  .field-row {
    display: grid;
    grid-template-columns: 22px 1fr 1.4fr 60px;
    gap: 8px;
    align-items: center;
    padding: 6px 8px;
    border-radius: 5px;
    border: 1px solid transparent;
    margin-bottom: 4px;
    background: white;
  }
  .field-row:hover { background: #fafafa; }
  .field-row.dragging { opacity: 0.4; }
  .field-row.drag-over { background: #fff8e1; border-color: #ffc107; }
  .drag-handle { color: #ccc; cursor: grab; user-select: none; text-align: center; font-size: 16px; line-height: 1; }
  .drag-handle:active { cursor: grabbing; color: #1976d2; }
  .field-row .label-en { font-size: 13px; color: #444; }
  .field-row .label-ar { font-size: 12px; color: #888; direction: rtl; text-align: left; }
  .field-row input, .field-row textarea, .field-row select {
    width: 100%; padding: 6px 9px; border: 1px solid #ccc; border-radius: 5px; font-size: 13px; font-family: inherit; box-sizing: border-box;
  }
  .field-row input:focus, .field-row textarea:focus { outline: none; border-color: #1976d2; }
  .field-row textarea { min-height: 30px; resize: vertical; }
  .arrows { display: flex; gap: 2px; }
  .arrow-btn {
    width: 26px; height: 26px; padding: 0; border: 1px solid #ddd; background: #fafafa; border-radius: 4px; cursor: pointer; font-size: 12px; line-height: 1; color: #666;
  }
  .arrow-btn:hover { background: #e3f2fd; border-color: #1976d2; color: #1976d2; }

  /* ===== Custom field add ===== */
  .add-field-row { margin-top: 10px; padding-top: 8px; border-top: 1px dashed #ddd; text-align: center; }
  .add-field-btn { background: #fff; border: 1px dashed #1976d2; color: #1976d2; padding: 7px 14px; border-radius: 6px; cursor: pointer; font-size: 13px; }
  .add-field-btn:hover { background: #e3f2fd; }

  /* ===== Client Type section (above Identity) ===== */
  .client-type-section {
    background: #fff8e1;
    border: 1px solid #ffe082;
    border-radius: 8px;
    padding: 12px 14px;
    margin-bottom: 12px;
  }
  .client-type-section .ct-header {
    font-size: 13px;
    font-weight: 500;
    color: #5d4037;
    margin-bottom: 8px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .client-type-section .ct-header .ar { color: #8d6e63; font-weight: 400; font-size: 12px; }
  .ct-radios { display: flex; gap: 18px; margin-bottom: 6px; }
  .ct-radios label { display: flex; align-items: center; gap: 6px; font-size: 13px; cursor: pointer; }
  .ct-company-fields {
    display: none;
    margin-top: 10px;
    padding-top: 10px;
    border-top: 1px dashed #ffd54f;
  }
  .ct-company-fields.show { display: block; }
  .ct-field { display: grid; grid-template-columns: 1fr 1.4fr; gap: 8px; align-items: center; margin-bottom: 6px; }
  .ct-field label { font-size: 13px; color: #444; }
  .ct-field label .ar { display: block; color: #888; font-size: 12px; direction: rtl; text-align: left; }
  .ct-field input {
    width: 100%; padding: 6px 9px; border: 1px solid #ccc; border-radius: 5px; font-size: 13px; box-sizing: border-box;
  }

  /* ===== Arabic-value inputs (colored background, RTL) ===== */
  .field-row.is-arabic { background: #fffbea; }
  .field-row.is-arabic:hover { background: #fff7d6; }
  .field-row.is-arabic input,
  .field-row.is-arabic textarea {
    background: #fffde7;
    border-color: #f6d97a;
    direction: rtl;
    text-align: right;
  }
  .ct-field input.arabic-input {
    background: #fffde7;
    border-color: #f6d97a;
    direction: rtl;
    text-align: right;
  }

  /* ===== Tabs ===== */
  .tabs { display: flex; gap: 4px; border-bottom: 2px solid #e0e0e0; margin-bottom: 16px; overflow-x: auto; overflow-y: hidden; }
  .tabs::-webkit-scrollbar { height: 4px; }
  .tabs::-webkit-scrollbar-thumb { background: #ccc; border-radius: 2px; }
  .tab { padding: 10px 18px; cursor: pointer; font-size: 14px; color: #555; border: none; background: transparent; border-bottom: 3px solid transparent; margin-bottom: -2px; white-space: nowrap; }
  .tab:hover { color: #1976d2; }
  .tab.active { color: #1976d2; border-bottom-color: #1976d2; font-weight: 500; }
  .tab-pane { display: none; }
  .tab-pane.active { display: block; }
  .tab.tab-agency { color: #ad1457; }
  .tab.tab-agency.active { color: #ad1457; border-bottom-color: #ad1457; }
  .tab.tab-agency .ag-count { display: inline-block; background: #fce4ec; color: #ad1457; padding: 1px 6px; border-radius: 8px; font-size: 11px; margin-left: 4px; font-weight: 500; }
  .tab.tab-agency.active .ag-count { background: #ad1457; color: white; }
  .tab.tab-add { color: #2e7d32; font-weight: 500; }
  .tab.tab-add:hover { color: #1b5e20; }

  /* ===== Card grid (Imported Contacts) ===== */
  .card-grid {
    display: grid;
    grid-template-columns: repeat(8, 1fr);
    gap: 8px;
    margin-bottom: 16px;
  }
  .contact-card {
    background: white;
    border: 1px solid #e0e0e0;
    border-radius: 6px;
    padding: 8px 9px;
    font-size: 11px;
    cursor: pointer;
    transition: all 0.15s;
    min-height: 80px;
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }
  .contact-card:hover {
    border-color: #1976d2;
    box-shadow: 0 2px 6px rgba(25,118,210,0.18);
    transform: translateY(-1px);
  }
  .contact-card .cc-id { font-size: 10px; color: #888; font-weight: 500; margin-bottom: 2px; }
  .contact-card .cc-name { font-size: 12px; font-weight: 500; color: #222; line-height: 1.25; margin-bottom: 4px; word-break: break-word; max-height: 30px; overflow: hidden; }
  .contact-card .cc-company { font-size: 10px; color: #1565c0; font-weight: 500; line-height: 1.2; margin-bottom: 3px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .contact-card .cc-phone { font-size: 10px; color: #555; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .contact-card .cc-email { font-size: 10px; color: #888; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .contact-card .cc-tags { margin-top: auto; padding-top: 4px; display: flex; flex-wrap: wrap; gap: 2px; }
  .contact-card .cc-tag { background: #e3f2fd; color: #1565c0; padding: 1px 5px; border-radius: 3px; font-size: 9px; font-weight: 500; }
  .contact-card .cc-tag.unassigned { background: #fff3cd; color: #856404; }
  .contact-card .cc-tag.detected { background: #fce4ec; color: #ad1457; }

  /* New/Edited tab — per-card action buttons */
  .cc-actions { margin-top: 6px; padding-top: 6px; border-top: 1px dashed #e0e0e0; display: flex; flex-wrap: wrap; gap: 4px; align-items: center; }
  .cc-btn { font-size: 10px; padding: 3px 7px; border-radius: 4px; border: 1px solid; cursor: pointer; font-weight: 500; line-height: 1.2; }
  .cc-btn:disabled { opacity: 0.6; cursor: wait; }
  .cc-push { background: #e3f2fd; color: #1565c0; border-color: #90caf9; }
  .cc-push:hover:not(:disabled) { background: #bbdefb; }
  .cc-done { background: #e8f5e9; color: #2e7d32; border-color: #a5d6a7; }
  .cc-done:hover:not(:disabled) { background: #c8e6c9; }
  .cc-pushed-badge { font-size: 9px; color: #2e7d32; background: #e8f5e9; padding: 2px 6px; border-radius: 3px; border: 1px solid #a5d6a7; }
  .cc-mobile-badge { font-size: 9px; color: #0a3678; background: #e7f1ff; padding: 2px 6px; border-radius: 3px; border: 1px solid #b6d4fe; }

  /* GPS corner pins on contact cards */
  .contact-card { position: relative; }
  .cc-gps-corner { position: absolute; top: 4px; right: 4px; display: flex; gap: 2px; }
  .cc-gps-pin { font-size: 14px; cursor: pointer; line-height: 1; padding: 3px 5px; border-radius: 50%; background: rgba(255,255,255,0.85); border: 1px solid #cfe2ff; transition: transform .15s, background .15s; }
  .cc-gps-pin:hover { background: #cfe2ff; transform: scale(1.15); }

  /* Pagination */
  .pagination { display: flex; align-items: center; justify-content: center; gap: 6px; margin: 16px 0; flex-wrap: wrap; }
  .page-btn { background: white; border: 1px solid #ddd; padding: 6px 11px; border-radius: 5px; cursor: pointer; font-size: 13px; color: #555; min-width: 34px; }
  .page-btn:hover { background: #f5f5f5; border-color: #999; }
  .page-btn.active { background: #1976d2; color: white; border-color: #1976d2; }
  .page-btn:disabled { opacity: 0.4; cursor: not-allowed; }
  .page-info { color: #777; font-size: 13px; margin: 0 10px; }

  /* Import toolbar */
  .import-toolbar { display: flex; gap: 10px; align-items: center; margin-bottom: 12px; flex-wrap: wrap; }
  .import-toolbar input[type="text"] { flex: 1; min-width: 200px; padding: 8px 12px; border: 1px solid #ccc; border-radius: 6px; font-size: 13px; }

  /* Responsive: shrink grid on smaller screens */
  @media (max-width: 1100px) { .card-grid { grid-template-columns: repeat(6, 1fr); } }
  @media (max-width: 800px)  { .card-grid { grid-template-columns: repeat(4, 1fr); } }
  @media (max-width: 540px)  { .card-grid { grid-template-columns: repeat(2, 1fr); } }

  /* ===== Import preview modal ===== */
  .import-modal {
    position: fixed; top: 50%; left: 50%; transform: translate(-50%,-50%);
    width: 640px; max-width: 95vw; max-height: 85vh;
    background: white; border-radius: 10px; box-shadow: 0 12px 40px rgba(0,0,0,0.25);
    z-index: 1100; display: none; flex-direction: column; overflow: hidden;
  }
  .import-modal.open { display: flex; }
  .import-modal .imh { background: #1976d2; color: white; padding: 12px 18px; display: flex; justify-content: space-between; align-items: center; }
  .import-modal .imh h3 { margin: 0; font-size: 16px; font-weight: 600; }
  .import-modal .imh button { background: transparent; border: none; color: white; font-size: 22px; cursor: pointer; }
  .import-modal .imb { padding: 18px; overflow-y: auto; flex: 1; }
  .import-modal .imf { padding: 12px 18px; border-top: 1px solid #eee; display: flex; justify-content: flex-end; gap: 8px; background: #fafafa; }
  .stats-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; margin-bottom: 16px; }
  .stat-box { background: #f5f5f5; padding: 12px; border-radius: 6px; }
  .stat-box .stat-num { font-size: 22px; font-weight: 600; color: #1976d2; }
  .stat-box .stat-label { font-size: 12px; color: #666; margin-top: 2px; }
  .sample-list { background: #fafafa; border-radius: 6px; padding: 10px 12px; max-height: 220px; overflow-y: auto; font-size: 12px; }
  .sample-list .sample-row { padding: 4px 0; border-bottom: 1px solid #eee; }
  .sample-list .sample-row:last-child { border-bottom: none; }
  .sample-list .sample-prefix { display: inline-block; background: #e3f2fd; color: #1565c0; padding: 1px 6px; border-radius: 3px; font-size: 10px; font-weight: 500; margin-right: 6px; }
  .progress-bar { height: 6px; background: #eee; border-radius: 3px; overflow: hidden; margin: 10px 0; }
  .progress-fill { height: 100%; background: #1976d2; transition: width 0.2s; }
</style>
</head>
<body>

<div class="topbar">📒 Address Book</div>

<div class="container">

  <div id="status" class="status">Connecting to Firebase…</div>

  <!-- Tabs -->
  <div class="tabs" id="tabBar">
    <button class="tab active" data-tab="contacts">Contacts</button>
    <button class="tab" data-tab="imported">Imported Contacts</button>
    <button class="tab" data-tab="unassigned" style="color:#856404;">⚠ Unassigned <span id="unassignedCountBadge" style="background:#fff3cd; color:#856404; padding:1px 6px; border-radius:8px; font-size:11px; margin-left:4px; font-weight:500;">0</span></button>
    <button class="tab" data-tab="newedited" style="color:#0d6efd;">✏️ New/Edited <span id="newEditedCountBadge" style="background:#cfe2ff; color:#0d6efd; padding:1px 6px; border-radius:8px; font-size:11px; margin-left:4px; font-weight:500;">0</span></button>
    <span id="agencyTabsSlot"></span><!-- dynamic agency tabs from Local_Agencies collection -->
    <button class="tab" data-tab="archive">🗄 Archive</button>
  </div>

  <!-- ============================================
       TAB 1: Contacts (card grid)
       ============================================ -->
  <div class="tab-pane active" id="pane-contacts">
    <div class="toolbar">
      <input id="searchBox" type="text" placeholder="Search by name, phone, email…">
      <span id="countLabel" class="count">0 contacts</span>
      <span style="font-size:12px; color:#888; margin-left:auto;">To add a new contact, open the agency tab where it belongs.</span>
      <button onclick="exportForGoogleContactsVCF()" style="background:#0d6efd; color:white; border:none; padding:8px 14px; border-radius:6px; cursor:pointer; font-weight:600; font-size:13px;">📥 Export for Google Contacts (vCard)</button>
    </div>

    <div id="contactsGrid" class="card-grid"></div>

    <div id="emptyState" class="empty" style="display:none;">
      No contacts yet. Click <b>＋ Add Contact</b> to create the first one.
    </div>

    <div id="contactsPagination" class="pagination"></div>
  </div>

  <!-- ============================================
       TAB 2: Imported Contacts (card grid)
       ============================================ -->
  <div class="tab-pane" id="pane-imported">
    <div class="import-toolbar">
      <button id="btnImport" class="btn-primary">⬇ Import from Google Contacts</button>
      <input id="impSearchBox" type="text" placeholder="Search imported…">
      <span id="impCountLabel" class="count">0 contacts</span>
    </div>

    <div id="importedGrid" class="card-grid"></div>

    <div id="importedEmpty" class="empty" style="display:none;">
      No imported contacts yet. Click <b>⬇ Import from Google Contacts</b> to bring them in.
    </div>

    <div id="importedPagination" class="pagination"></div>
  </div>

  <!-- ============================================
       TAB 3: Archive (card grid for archived contacts)
       ============================================ -->
  <div class="tab-pane" id="pane-archive">
    <div class="import-toolbar">
      <input id="archiveSearchBox" type="text" placeholder="Search archive…">
      <span id="archiveCountLabel" class="count">0 archived</span>
    </div>

    <div id="archiveGrid" class="card-grid"></div>

    <div id="archiveEmpty" class="empty" style="display:none;">
      No archived contacts.
    </div>

    <div id="archivePagination" class="pagination"></div>
  </div>

  <!-- ============================================
       TAB: Unassigned (no matching agency)
       ============================================ -->
  <div class="tab-pane" id="pane-unassigned">
    <div class="import-toolbar">
      <span style="font-size:15px; font-weight:600; color:#856404;">⚠ Unassigned contacts</span>
      <input id="unassignedSearchBox" type="text" placeholder="Search…">
      <span id="unassignedCountLabel" class="count">0 contacts</span>
    </div>

    <div style="background:#fff8e1; border:1px solid #ffe082; padding:10px 14px; border-radius:6px; margin-bottom:12px; font-size:13px; color:#5d4037;">
      These contacts have no prefix code, or their prefix doesn't match any registered agency. Add the matching agency in <b>local_agencies_manager.html</b> or edit the contact to fix its Prefix Codes.
    </div>

    <div id="unassignedGrid" class="card-grid"></div>

    <div id="unassignedEmpty" class="empty" style="display:none;">
      🎉 No unassigned contacts — every contact is linked to at least one agency.
    </div>

    <div id="unassignedPagination" class="pagination"></div>
  </div>

  <!-- ============================================
       TAB: New/Edited (contacts touched by the user)
       ============================================ -->
  <div class="tab-pane" id="pane-newedited">
    <div class="import-toolbar">
      <span style="font-size:15px; font-weight:600; color:#0d6efd;">✏️ New / Edited contacts</span>
      <input id="newEditedSearchBox" type="text" placeholder="Search…">
      <span id="newEditedCountLabel" class="count">0 contacts</span>
    </div>

    <div style="background:#e7f1ff; border:1px solid #b6d4fe; padding:10px 14px; border-radius:6px; margin-bottom:12px; font-size:13px; color:#0a3678;">
      Contacts you created from scratch or modified after import. Sorted by most recent edit. Imported-but-untouched contacts are not listed here.
    </div>

    <div id="newEditedGrid" class="card-grid"></div>

    <div id="newEditedEmpty" class="empty" style="display:none;">
      No new or edited contacts yet.
    </div>

    <div id="newEditedPagination" class="pagination"></div>
  </div>

  <!-- ============================================
       TAB: Agency (shared pane, filtered by selected agency)
       ============================================ -->
  <div class="tab-pane" id="pane-agency">
    <div class="import-toolbar">
      <button id="btnAddContactToAgency" class="btn-primary">＋ Add Contact to <span id="addCtaCode">this agency</span></button>
      <span id="agencyTitle" style="font-size:15px; font-weight:600; color:#ad1457;">Agency</span>
      <input id="agencySearchBox" type="text" placeholder="Search…">
      <span id="agencyCountLabel" class="count">0 contacts</span>
      <a href="local_agencies_manager.html" style="font-size:12px; color:#666; text-decoration:none; padding:6px 10px; border:1px solid #ddd; border-radius:5px;">⚙ Manage agencies</a>
    </div>

    <div id="agencyGrid" class="card-grid"></div>

    <div id="agencyEmpty" class="empty" style="display:none;">
      No contacts found with this agency code in their prefix.
    </div>

    <div id="agencyPagination" class="pagination"></div>
  </div>

</div>

<!-- Hidden file picker for CSV upload -->
<input type="file" id="csvFileInput" accept=".csv" style="display:none;">

<!-- Import preview modal -->
<div id="importModal" class="import-modal">
  <div class="imh">
    <h3 id="importModalTitle">Import from Google Contacts</h3>
    <button id="importModalClose">×</button>
  </div>
  <div class="imb" id="importModalBody">
    <!-- populated dynamically -->
  </div>
  <div class="imf" id="importModalFooter">
    <!-- populated dynamically -->
  </div>
</div>

<!-- ============================================================
     POPUP — Add / Edit Contact
     ============================================================ -->
<div id="modalBackdrop" class="modal-backdrop"></div>

<div id="contactModal" class="modal">
  <div class="modal-header" id="modalHeader">
    <h3 id="modalTitle">New Contact PopUp</h3>
    <button class="close-btn" id="modalClose">×</button>
  </div>
  <div class="modal-body">
    <!-- Client Type — always above Identity -->
    <div class="client-type-section">
      <div class="ct-header">
        <span>Client Type <span class="ar">/ نوع العميل</span></span>
        <label style="display:flex; align-items:center; gap:6px; font-size:13px; color:#5d4037; cursor:pointer; background:#fff; padding:4px 10px; border:1px solid #ffe082; border-radius:5px;">
          <input type="checkbox" id="archivedCheckbox" name="Archived">
          🗄 Archive this contact / أرشفة
        </label>
      </div>
      <div class="ct-radios">
        <label><input type="radio" name="ClientType" value="Personal" checked> Personal / شخصي</label>
        <label><input type="radio" name="ClientType" value="Company"> Company / شركة</label>
      </div>
      <div class="ct-company-fields" id="ctCompanyFields">
        <div class="ct-field">
          <label>Company Name (English)<span class="ar">اسم الشركة (إنجليزي)</span></label>
          <input type="text" name="CompanyName_EN">
        </div>
        <div class="ct-field">
          <label>Company Name (Arabic)<span class="ar">اسم الشركة (عربي)</span></label>
          <input type="text" name="CompanyName_AR" class="arabic-input">
        </div>
      </div>
    </div>

    <div id="groupsContainer"></div>
  </div>
  <div class="modal-footer">
    <div class="footer-left">
      <button class="reorder-toggle" id="reorderToggle">↕ Reorder mode: OFF</button>
    </div>
    <div class="footer-right">
      <button class="btn-secondary" id="cancelBtn">Cancel</button>
      <button class="btn-primary" id="saveBtn">Save / حفظ</button>
    </div>
  </div>
</div>

<!-- ============================================================
     FIREBASE INIT — auto-config from Firebase Hosting
     ============================================================ -->
<script type="module">
  import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
  import { getFirestore, collection, onSnapshot, addDoc, doc, getDoc, setDoc, updateDoc, deleteDoc, serverTimestamp }
    from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
  import { getStorage, ref as storageRef, uploadBytes, getDownloadURL, deleteObject }
    from "https://www.gstatic.com/firebasejs/10.12.2/firebase-storage.js";

  (async () => {
    try {
      const res = await fetch('/__/firebase/init.json');
      if (!res.ok) throw new Error('init.json HTTP ' + res.status);
      const firebaseConfig = await res.json();

      const app = initializeApp(firebaseConfig);
      const db = getFirestore(app);
      window._fb = {
        db, storage: getStorage(app),
        collection, onSnapshot, addDoc, doc, getDoc, setDoc, updateDoc, deleteDoc, serverTimestamp,
        storageRef, uploadBytes, getDownloadURL, deleteObject
      };
      document.dispatchEvent(new Event('firebase-ready'));
    } catch (e) {
      const s = document.getElementById('status');
      s.textContent = '✗ Firebase init failed: ' + e.message;
      s.classList.remove('ok'); s.classList.add('err');
    }
  })();
</script>

<script>
/* =====================================================
   APPS SCRIPT BRIDGE — pushes Firestore contacts into Google Contacts
   ===================================================== */
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxEbNo287tvEezLwy8vYHrNWxPXELlzzsGVg9yeA6XXDIkXdR0ILEKOTWbG88PxHiIf/exec';

/* =====================================================
   FIELD GROUPS — bilingual labels (English / Arabic)
   ===================================================== */
const GROUPS = [
  {
    key: 'idpassport',
    label_en: 'ID & Passport Copy',
    label_ar: 'صورة الهوية وجواز السفر',
    isCustom: true,
    fields: []
  },
  {
    key: 'navigation',
    label_en: 'Navigation',
    label_ar: 'الموقع والملاحة',
    isCustom: true,
    fields: []
  },
  {
    key: 'identity',
    label_en: 'Identity',
    label_ar: 'الهوية',
    fields: [
      { name: 'EmployerID',           en: 'Employer ID',           ar: 'رقم صاحب العمل' },
      { name: 'EmployerCode',         en: 'Employer Code',         ar: 'رمز صاحب العمل' },
      { name: 'DetectedCodes',        en: 'Prefix Codes',          ar: 'رموز البادئة (الوكالات)', isPrefixCodes: true },
      { name: 'EmployerName',         en: 'Employer Name',         ar: 'اسم صاحب العمل' },
      { name: 'EFN',                  en: 'Employer Full Name',    ar: 'الاسم الكامل لصاحب العمل' },
      { name: 'FirstName_EN',         en: 'First Name (English)',  ar: 'الاسم الأول (إنجليزي)' },
      { name: 'FirstName_AR',         en: 'First Name (Arabic)',   ar: 'الاسم الأول (عربي)', isArabic: true },
      { name: 'MiddleName_EN',        en: 'Middle Name (English)', ar: 'اسم الأب (إنجليزي)' },
      { name: 'MiddleName_AR',        en: 'Middle Name (Arabic)',  ar: 'اسم الأب (عربي)', isArabic: true },
      { name: 'LastName_EN',          en: 'Last Name (English)',   ar: 'اسم العائلة (إنجليزي)' },
      { name: 'LastName_AR',          en: 'Last Name (Arabic)',    ar: 'اسم العائلة (عربي)', isArabic: true },
      { name: 'ContactPerson',        en: 'Contact Person',        ar: 'الشخص المسؤول' },
      { name: 'BirthDate',            en: 'Birth Date',            ar: 'تاريخ الميلاد', type: 'date' },
      { name: 'Anniversary',          en: 'Anniversary',           ar: 'تاريخ الذكرى', type: 'date' },
      { name: 'Disabled',             en: 'Disabled',              ar: 'معطّل', type: 'checkbox' },
      { name: 'DisabledAt',           en: 'Disabled At',           ar: 'تاريخ التعطيل', type: 'date' },
      { name: 'EmployerDisabled',     en: 'Employer Disabled',     ar: 'صاحب العمل معطّل', type: 'checkbox' },
      { name: 'EmployerDisabledDate', en: 'Employer Disabled Date',ar: 'تاريخ تعطيل صاحب العمل', type: 'date' }
    ]
  },
  {
    key: 'phones_emails',
    label_en: 'Phones & Emails',
    label_ar: 'الهواتف والبريد الإلكتروني',
    fields: [
      { name: 'MobileNumber',   en: 'Mobile Number',   ar: 'رقم الجوال' },
      { name: 'MobileNumber2',  en: 'Mobile Number 2', ar: 'رقم الجوال 2' },
      { name: 'MobileNumber3',  en: 'Mobile Number 3', ar: 'رقم الجوال 3' },
      { name: 'MobileNumber4',  en: 'Mobile Number 4', ar: 'رقم الجوال 4' },
      { name: 'MobileNumber5',  en: 'Mobile Number 5', ar: 'رقم الجوال 5' },
      { name: 'MobileNumber6',  en: 'Mobile Number 6', ar: 'رقم الجوال 6' },
      { name: 'MobileNumber7',  en: 'Mobile Number 7', ar: 'رقم الجوال 7' },
      { name: 'MobileNumber8',  en: 'Mobile Number 8', ar: 'رقم الجوال 8' },
      { name: 'MobileNumber9',  en: 'Mobile Number 9', ar: 'رقم الجوال 9' },
      { name: 'MobileNumber10', en: 'Mobile Number 10',ar: 'رقم الجوال 10' },
      { name: 'ContactNumber',  en: 'Contact Number',  ar: 'رقم الاتصال' },
      { name: 'EmployerContact',en: 'Employer Contact',ar: 'اتصال صاحب العمل' },
      { name: 'Email',          en: 'Email',           ar: 'البريد الإلكتروني', type: 'email' },
      { name: 'EmployerEmail',  en: 'Employer Email',  ar: 'بريد صاحب العمل', type: 'email' },
      { name: 'EmailAddress',   en: 'Email Address',   ar: 'عنوان البريد الإلكتروني', type: 'email' }
    ]
  },
  {
    key: 'address',
    label_en: 'Address',
    label_ar: 'العنوان',
    fields: [
      { name: 'Address',          en: 'Address',          ar: 'العنوان' },
      { name: 'EmployerAddress',  en: 'Employer Address', ar: 'عنوان صاحب العمل' },
      { name: 'City',             en: 'City',             ar: 'المدينة' },
      { name: 'EmployerCity',     en: 'Employer City',    ar: 'مدينة صاحب العمل' },
      { name: 'Country',          en: 'Country',          ar: 'البلد' }
    ]
  },
  {
    key: 'agency',
    label_en: 'Agency',
    label_ar: 'الوكالة',
    fields: [
      { name: 'AgencyID',   en: 'Agency ID',   ar: 'رقم الوكالة' },
      { name: 'AgencyCode', en: 'Agency Code', ar: 'رمز الوكالة' }
    ]
  },
  {
    key: 'folder',
    label_en: 'Folder (Drive)',
    label_ar: 'مجلد جوجل درايف',
    fields: [
      { name: 'EmployerFolder',     en: 'Employer Folder',      ar: 'مجلد صاحب العمل' },
      { name: 'EmployerFolderName', en: 'Employer Folder Name', ar: 'اسم مجلد صاحب العمل' },
      { name: 'EmployerFolderID',   en: 'Employer Folder ID',   ar: 'رقم مجلد صاحب العمل' },
      { name: 'FolderUrl',          en: 'Folder URL',           ar: 'رابط المجلد' },
      { name: 'EmployerFolderUrl',  en: 'Employer Folder URL',  ar: 'رابط مجلد صاحب العمل' }
    ]
  },
  {
    key: 'documents',
    label_en: 'Documents',
    label_ar: 'المستندات',
    fields: [
      { name: 'RequiredDocuments', en: 'Required Documents', ar: 'المستندات المطلوبة', type: 'textarea' },
      { name: 'UploadedDocuments', en: 'Uploaded Documents', ar: 'المستندات المرفوعة', type: 'textarea' }
    ]
  },
  {
    key: 'google_sync',
    label_en: 'Google Sync',
    label_ar: 'مزامنة جوجل',
    fields: [
      { name: 'GoogleContactID', en: 'Google Contact ID', ar: 'رقم جهة الاتصال جوجل' },
      { name: 'SyncStatus',      en: 'Sync Status',       ar: 'حالة المزامنة' },
      { name: 'LastSyncDate',    en: 'Last Sync Date',    ar: 'تاريخ آخر مزامنة', type: 'date' },
      { name: 'Photo',           en: 'Photo URL',         ar: 'رابط الصورة' }
    ]
  },
  {
    key: 'notes',
    label_en: 'Notes',
    label_ar: 'ملاحظات',
    fields: [
      { name: 'Notes',         en: 'Notes',          ar: 'ملاحظات', type: 'textarea' },
      { name: 'EmployerNotes', en: 'Employer Notes', ar: 'ملاحظات صاحب العمل', type: 'textarea' }
    ]
  }
];

/* =====================================================
   STATE
   ===================================================== */
let allContacts = [];
let currentEditingId = null;
let customFields = [];        // [{ name, en, ar, group }]
let fieldOrder = {};          // { groupKey: [fieldName, ...] }

/* =====================================================
   STATE — agencies
   ===================================================== */
let localAgencies = [];          // [{ _id, code, nameEn, nameAr }]
let currentAgencyCode = null;    // the code of the currently-viewed agency tab

/* =====================================================
   FIREBASE READY
   ===================================================== */
document.addEventListener('firebase-ready', async () => {
  const s = document.getElementById('status');
  s.textContent = '✓ Connected — loading contacts…';
  s.classList.add('ok');

  const { db, collection, onSnapshot, doc, getDoc } = window._fb;

  // Load custom fields
  try {
    const cfDoc = await getDoc(doc(db, '_address_book_meta', 'custom_fields'));
    if (cfDoc.exists()) customFields = cfDoc.data().fields || [];
  } catch(e) {}

  // Load saved field order
  try {
    const orderDoc = await getDoc(doc(db, '_address_book_meta', 'field_order'));
    if (orderDoc.exists()) fieldOrder = orderDoc.data().orders || {};
  } catch(e) {}

  // Live local agencies — drives the dynamic agency tabs.
  // Source of truth is local_agencies_manager.html (collection: Local_Agencies).
  // Field-name variants accepted: AgencyCode, agencyCode, Code, code.
  onSnapshot(collection(db, 'Local_Agencies'), (snap) => {
    const seen = new Set();
    localAgencies = snap.docs
      .map(d => {
        const data = d.data() || {};
        const codeRaw = data.AgencyCode || data.agencyCode || data.Code || data.code || '';
        const code = String(codeRaw).trim().toUpperCase();
        return { _id: d.id, ...data, AgencyCode: code };
      })
      .filter(a => {
        if (!a.AgencyCode) return false;
        if (seen.has(a.AgencyCode)) return false;
        seen.add(a.AgencyCode);
        return true;
      });
    console.log('[Address Book] Local_Agencies loaded:', localAgencies.map(a => a.AgencyCode));
    renderAgencyTabs();
    updateUnassignedBadge();
    const unassignedPane = document.getElementById('pane-unassigned');
    if (unassignedPane && unassignedPane.classList.contains('active')) renderUnassigned();
  });

  // Live contacts
  onSnapshot(collection(db, 'address_book'), (snap) => {
    allContacts = snap.docs.map(d => ({ _id: d.id, ...d.data() }));
    render();
    renderAgencyTabs(); // refresh counts on tabs
    updateUnassignedBadge();
    updateNewEditedBadge();
    const newEditedPane = document.getElementById('pane-newedited');
    if (newEditedPane && newEditedPane.classList.contains('active')) renderNewEdited();
    s.style.display = 'none';
  }, (err) => {
    s.textContent = '✗ Error: ' + err.message;
    s.classList.remove('ok'); s.classList.add('err');
  });
});

/* =====================================================
   TABLE RENDER
   ===================================================== */
/* =====================================================
   SHARED CARD HTML — used by both Contacts and Imported tabs
   ===================================================== */
function contactCardHtml(c) {
  const id    = c.ContactID || c.EmployerID || (c._id ? c._id.slice(0, 6) : '');
  const efn   = c.EFN || c.EmployerName ||
                [c.FirstName_EN, c.MiddleName_EN, c.LastName_EN].filter(Boolean).join(' ') ||
                '(no name)';
  const isCompany   = c.ClientType === 'Company';
  const companyName = isCompany ? (c.CompanyName_EN || c.CompanyName_AR || '') : '';
  const phone = c.MobileNumber || c.MobileNumber2 || c.ContactNumber || '';
  const email = c.Email || c.EmployerEmail || c.EmailAddress || '';

  // Tags: prefer assigned agencies, fall back to detected codes (still unassigned)
  const agencies = Array.isArray(c.agencies) ? c.agencies : [];
  const detected = Array.isArray(c.DetectedCodes) ? c.DetectedCodes : [];
  let tagsHtml = '';
  if (agencies.length) {
    tagsHtml = agencies.map(a => {
      const code = typeof a === 'string' ? a : (a.code || '');
      return `<span class="cc-tag">${escapeHtml(code)}</span>`;
    }).join('');
  } else if (detected.length) {
    tagsHtml = detected.map(code => `<span class="cc-tag detected">${escapeHtml(code)}</span>`).join('');
  } else {
    tagsHtml = `<span class="cc-tag unassigned">Unassigned</span>`;
  }

  // GPS corner icons (top-right) — visible only when location stored.
  // Use onclick (not <a>) so the card's openModal doesn't fire.
  const home = c.HomeLocation || c.Location;
  const work = c.WorkLocation;
  const gpsIcons = [];
  if (home && home.lat != null && home.lng != null) {
    gpsIcons.push(`<span class="cc-gps-pin" title="Navigate to Home GPS" onclick="event.stopPropagation();window.open('https://www.google.com/maps/dir/?api=1&destination=${home.lat},${home.lng}&dir_action=navigate&travelmode=driving','_blank')">🏠</span>`);
  }
  if (work && work.lat != null && work.lng != null) {
    gpsIcons.push(`<span class="cc-gps-pin" title="Navigate to Work GPS" onclick="event.stopPropagation();window.open('https://www.google.com/maps/dir/?api=1&destination=${work.lat},${work.lng}&dir_action=navigate&travelmode=driving','_blank')">🏢</span>`);
  }
  const gpsCornerHtml = gpsIcons.length ? `<div class="cc-gps-corner">${gpsIcons.join('')}</div>` : '';

  return `
    <div class="contact-card" onclick="openModal('${c._id}')">
      ${gpsCornerHtml}
      <div class="cc-id">#${escapeHtml(String(id))}</div>
      <div class="cc-name">${escapeHtml(efn)}</div>
      ${companyName ? `<div class="cc-company">🏢 ${escapeHtml(companyName)}</div>` : ''}
      ${phone ? `<div class="cc-phone">📞 ${escapeHtml(phone)}</div>` : ''}
      ${email ? `<div class="cc-email">✉ ${escapeHtml(email)}</div>` : ''}
      <div class="cc-tags">${tagsHtml}</div>
    </div>
  `;
}

/* =====================================================
   TABLE RENDER  →  card grid for Contacts tab
   ===================================================== */
const CONTACTS_PAGE_SIZE = 120;   // 8 cols × 15 rows
let contactsPage = 1;

function render() {
  const q = (document.getElementById('searchBox').value || '').toLowerCase();
  const list = allContacts
    .filter(c => !c.Archived)                                     // exclude archived
    .filter(c => !q || JSON.stringify(c).toLowerCase().includes(q));

  document.getElementById('countLabel').textContent = list.length + ' contacts';

  const grid  = document.getElementById('contactsGrid');
  const empty = document.getElementById('emptyState');
  const pag   = document.getElementById('contactsPagination');

  if (list.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    pag.innerHTML = '';
    const importedPane = document.getElementById('pane-imported');
    if (importedPane && importedPane.classList.contains('active')) renderImported();
    const archivePane = document.getElementById('pane-archive');
    if (archivePane && archivePane.classList.contains('active')) renderArchive();
    return;
  }
  empty.style.display = 'none';

  const totalPages = Math.max(1, Math.ceil(list.length / CONTACTS_PAGE_SIZE));
  if (contactsPage > totalPages) contactsPage = totalPages;
  const start = (contactsPage - 1) * CONTACTS_PAGE_SIZE;
  const pageItems = list.slice(start, start + CONTACTS_PAGE_SIZE);

  grid.innerHTML = pageItems.map(contactCardHtml).join('');

  pag.innerHTML = paginationHtml(contactsPage, totalPages, list.length);
  pag.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      contactsPage = parseInt(btn.dataset.page, 10);
      render();
      document.getElementById('pane-contacts').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });

  // Also refresh the imported grid if its tab is active
  const importedPane = document.getElementById('pane-imported');
  if (importedPane && importedPane.classList.contains('active')) renderImported();
  const archivePane = document.getElementById('pane-archive');
  if (archivePane && archivePane.classList.contains('active')) renderArchive();
  const unassignedPane = document.getElementById('pane-unassigned');
  if (unassignedPane && unassignedPane.classList.contains('active')) renderUnassigned();
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
    .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

document.getElementById('searchBox').addEventListener('input', () => {
  contactsPage = 1;
  render();
});

/* =====================================================
   TABS — switch between Contacts and Imported Contacts
   ===================================================== */
document.querySelectorAll('.tab').forEach(t => {
  t.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.tab-pane').forEach(x => x.classList.remove('active'));
    t.classList.add('active');
    document.getElementById('pane-' + t.dataset.tab).classList.add('active');
    if (t.dataset.tab === 'imported')   renderImported();
    if (t.dataset.tab === 'archive')    renderArchive();
    if (t.dataset.tab === 'unassigned') renderUnassigned();
    if (t.dataset.tab === 'newedited')  renderNewEdited();
  });
});

/* =====================================================
   UNASSIGNED TAB — contacts with no matching agency
   ===================================================== */
const UNASSIGNED_PAGE_SIZE = 120;
let unassignedPage = 1;

document.getElementById('unassignedSearchBox').addEventListener('input', () => {
  unassignedPage = 1;
  renderUnassigned();
});

function isUnassigned(c) {
  if (c.Archived) return false;
  const codes = Array.isArray(c.DetectedCodes) ? c.DetectedCodes : [];
  if (codes.length === 0) return true;
  const registered = new Set(localAgencies.map(a => a.AgencyCode));
  // Unassigned if NONE of the contact's codes match a registered agency
  return !codes.some(code => registered.has(code));
}

function renderUnassigned() {
  const q = (document.getElementById('unassignedSearchBox').value || '').toLowerCase();
  const list = allContacts
    .filter(isUnassigned)
    .filter(c => !q || JSON.stringify(c).toLowerCase().includes(q));

  document.getElementById('unassignedCountLabel').textContent = list.length + ' contacts';
  document.getElementById('unassignedCountBadge').textContent = list.length;

  const grid  = document.getElementById('unassignedGrid');
  const empty = document.getElementById('unassignedEmpty');
  const pag   = document.getElementById('unassignedPagination');

  if (list.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    pag.innerHTML = '';
    return;
  }
  empty.style.display = 'none';

  const totalPages = Math.max(1, Math.ceil(list.length / UNASSIGNED_PAGE_SIZE));
  if (unassignedPage > totalPages) unassignedPage = totalPages;
  const start = (unassignedPage - 1) * UNASSIGNED_PAGE_SIZE;
  const pageItems = list.slice(start, start + UNASSIGNED_PAGE_SIZE);

  grid.innerHTML = pageItems.map(contactCardHtml).join('');

  pag.innerHTML = paginationHtml(unassignedPage, totalPages, list.length);
  pag.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      unassignedPage = parseInt(btn.dataset.page, 10);
      renderUnassigned();
      document.getElementById('pane-unassigned').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

/* Always keep the Unassigned badge fresh */
function updateUnassignedBadge() {
  const n = allContacts.filter(isUnassigned).length;
  const el = document.getElementById('unassignedCountBadge');
  if (el) el.textContent = n;
}

/* =====================================================
   NEW/EDITED TAB — contacts created or modified by user
   ===================================================== */
const NEWEDITED_PAGE_SIZE = 120;
let newEditedPage = 1;

document.getElementById('newEditedSearchBox').addEventListener('input', () => {
  newEditedPage = 1;
  renderNewEdited();
});

function isUserTouched(c) {
  return c && c._userEdited === true && !c.Archived;
}

function tsValue(c) {
  // Most-recent activity timestamp; works with Firestore Timestamp objects or plain numbers
  const t = c._updatedAt || c._createdAt;
  if (!t) return 0;
  if (typeof t.toMillis === 'function') return t.toMillis();
  if (typeof t === 'number') return t;
  if (t.seconds) return t.seconds * 1000;
  return 0;
}

/**
 * Card markup used only inside the New/Edited tab — same body as a regular card,
 * plus a row of action buttons (Push to Google, Done) at the bottom.
 */
function newEditedCardHtml(c) {
  const baseCard = contactCardHtml(c);
  const pushed = c._googleResourceName ? true : false;
  const pushedBadge = pushed
    ? `<div class="cc-pushed-badge" title="Pushed to Google on ${escapeHtml(formatPushedAt(c._googlePushedAt))}">✅ Pushed</div>`
    : '';
  const mobileBadge = c._addedViaMobile
    ? `<div class="cc-mobile-badge" title="Added or edited from mobile app">📱 Mobile</div>`
    : '';
  const actions = `
    <div class="cc-actions" onclick="event.stopPropagation()">
      <button type="button" class="cc-btn cc-push" onclick="pushContactCardToGoogle('${c._id}', this)">
        ${pushed ? '🔄 Re-push' : '📤 Push to Google'}
      </button>
      <button type="button" class="cc-btn cc-done" onclick="markContactDone('${c._id}', this)">
        ✓ Done
      </button>
      ${mobileBadge}
      ${pushedBadge}
    </div>
  `;
  // Inject the actions inside the card by inserting right before the closing div of the card
  return baseCard.replace(/(<\/div>\s*)$/, actions + '$1');
}

function formatPushedAt(t) {
  if (!t) return 'unknown';
  if (typeof t.toDate === 'function') return t.toDate().toLocaleString();
  if (t.seconds) return new Date(t.seconds * 1000).toLocaleString();
  if (typeof t === 'number') return new Date(t).toLocaleString();
  return String(t);
}

window.pushContactCardToGoogle = async function(firestoreId, btn) {
  const contact = allContacts.find(c => c._id === firestoreId);
  if (!contact) { alert('Contact not found in cache.'); return; }
  if (!btn) return;

  console.log('=== PUSH DEBUG START ===');
  console.log('Firestore _id:', firestoreId);
  console.log('ContactID:', contact.ContactID);
  console.log('EFN:', contact.EFN);
  console.log('_googleResourceName in cache:', contact._googleResourceName);
  console.log('_googlePushedAt in cache:', contact._googlePushedAt);
  console.log('Full contact object:', contact);

  const originalLabel = btn.textContent;
  btn.disabled = true;
  btn.textContent = '⏳ Pushing…';

  try {
    const payload = {
      action: 'pushContactToGoogle',
      contact: {
        ContactID:          contact.ContactID || '',
        GoogleResourceName: contact._googleResourceName || '',
        FirstName_EN:       contact.FirstName_EN  || '',
        MiddleName_EN:      contact.MiddleName_EN || '',
        LastName_EN:        contact.LastName_EN   || '',
        FirstName_AR:       contact.FirstName_AR  || '',
        MiddleName_AR:      contact.MiddleName_AR || '',
        LastName_AR:        contact.LastName_AR   || '',
        EFN:                contact.EFN || '',
        CompanyName_EN:     contact.CompanyName_EN || '',
        ClientType:         contact.ClientType || 'Personal',
        MobileNumber:       contact.MobileNumber  || '',
        MobileNumber2:      contact.MobileNumber2 || '',
        MobileNumber3:      contact.MobileNumber3 || '',
        MobileNumber4:      contact.MobileNumber4 || '',
        MobileNumber5:      contact.MobileNumber5 || '',
        MobileNumber6:      contact.MobileNumber6 || '',
        MobileNumber7:      contact.MobileNumber7 || '',
        MobileNumber8:      contact.MobileNumber8 || '',
        MobileNumber9:      contact.MobileNumber9 || '',
        MobileNumber10:     contact.MobileNumber10 || '',
        Email:              contact.Email || '',
        EmployerEmail:      contact.EmployerEmail || '',
        EmailAddress:       contact.EmailAddress  || '',
        Address:            contact.Address || '',
        City:               contact.City    || '',
        Country:            contact.Country || '',
        BirthDate:          contact.BirthDate || '',
        Anniversary:        contact.Anniversary || '',
        EmployerNotes:      contact.EmployerNotes || '',
        DetectedCodes:      Array.isArray(contact.DetectedCodes) ? contact.DetectedCodes : []
      }
    };
    console.log('Payload sent to Apps Script:', payload);

    const res = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      redirect: 'follow',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    if (!res.ok) throw new Error('HTTP ' + res.status);
    const data = await res.json();
    console.log('Apps Script response:', data);

    if (!data || data.success !== true) {
      throw new Error(data && data.error ? data.error : 'Apps Script did not return success');
    }

    const { db, doc, updateDoc, serverTimestamp } = window._fb;
    await updateDoc(doc(db, 'address_book', firestoreId), {
      _googleResourceName: data.resourceName,
      _googlePushedAt:     serverTimestamp()
    });
    console.log('Stamped _googleResourceName:', data.resourceName);
    console.log('=== PUSH DEBUG END ===');

    btn.textContent = '✅ ' + (data.action === 'updated' ? 'Updated' : (data.action || 'Pushed'));
    setTimeout(() => { btn.textContent = originalLabel; btn.disabled = false; }, 2500);
  } catch (err) {
    console.error('Push failed:', err);
    alert('Push failed: ' + (err.message || err));
    btn.textContent = originalLabel;
    btn.disabled = false;
  }
};

/**
 * Remove a contact from the New/Edited tab by clearing the _userEdited flag.
 * The doc stays in Firestore — only the "review queue" is emptied.
 */
window.markContactDone = async function(firestoreId, btn) {
  if (!confirm('Remove this contact from the New/Edited list? It will stay in Firestore as usual.')) return;
  if (btn) { btn.disabled = true; btn.textContent = '⏳'; }
  try {
    const { db, doc, updateDoc } = window._fb;
    await updateDoc(doc(db, 'address_book', firestoreId), { _userEdited: false });
    // onSnapshot will re-render and the card will disappear from this tab
  } catch (err) {
    alert('Mark done failed: ' + (err.message || err));
    if (btn) { btn.disabled = false; btn.textContent = '✓ Done'; }
  }
};

function renderNewEdited() {
  const q = (document.getElementById('newEditedSearchBox').value || '').toLowerCase();
  const list = allContacts
    .filter(isUserTouched)
    .filter(c => !q || JSON.stringify(c).toLowerCase().includes(q))
    .sort((a, b) => tsValue(b) - tsValue(a));   // newest first

  document.getElementById('newEditedCountLabel').textContent = list.length + ' contacts';
  document.getElementById('newEditedCountBadge').textContent = list.length;

  const grid  = document.getElementById('newEditedGrid');
  const empty = document.getElementById('newEditedEmpty');
  const pag   = document.getElementById('newEditedPagination');

  if (list.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    pag.innerHTML = '';
    return;
  }
  empty.style.display = 'none';

  const totalPages = Math.max(1, Math.ceil(list.length / NEWEDITED_PAGE_SIZE));
  if (newEditedPage > totalPages) newEditedPage = totalPages;
  const start = (newEditedPage - 1) * NEWEDITED_PAGE_SIZE;
  const pageItems = list.slice(start, start + NEWEDITED_PAGE_SIZE);

  grid.innerHTML = pageItems.map(newEditedCardHtml).join('');

  pag.innerHTML = paginationHtml(newEditedPage, totalPages, list.length);
  pag.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      newEditedPage = parseInt(btn.dataset.page, 10);
      renderNewEdited();
      document.getElementById('pane-newedited').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

/* Always keep the New/Edited badge fresh */
function updateNewEditedBadge() {
  const n = allContacts.filter(isUserTouched).length;
  const el = document.getElementById('newEditedCountBadge');
  if (el) el.textContent = n;
}

/* =====================================================
   ARCHIVE TAB — card grid + pagination
   ===================================================== */
const ARCHIVE_PAGE_SIZE = 120;
let archivePage = 1;

document.getElementById('archiveSearchBox').addEventListener('input', () => {
  archivePage = 1;
  renderArchive();
});

function renderArchive() {
  const q = (document.getElementById('archiveSearchBox').value || '').toLowerCase();
  const list = allContacts
    .filter(c => c.Archived === true)
    .filter(c => !q || JSON.stringify(c).toLowerCase().includes(q));

  document.getElementById('archiveCountLabel').textContent = list.length + ' archived';

  const grid  = document.getElementById('archiveGrid');
  const empty = document.getElementById('archiveEmpty');
  const pag   = document.getElementById('archivePagination');

  if (list.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    pag.innerHTML = '';
    return;
  }
  empty.style.display = 'none';

  const totalPages = Math.max(1, Math.ceil(list.length / ARCHIVE_PAGE_SIZE));
  if (archivePage > totalPages) archivePage = totalPages;
  const start = (archivePage - 1) * ARCHIVE_PAGE_SIZE;
  const pageItems = list.slice(start, start + ARCHIVE_PAGE_SIZE);

  grid.innerHTML = pageItems.map(contactCardHtml).join('');

  pag.innerHTML = paginationHtml(archivePage, totalPages, list.length);
  pag.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      archivePage = parseInt(btn.dataset.page, 10);
      renderArchive();
      document.getElementById('pane-archive').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

/* =====================================================
   AGENCY TABS — dynamic tabs based on local_agencies collection
   ===================================================== */
function renderAgencyTabs() {
  const slot = document.getElementById('agencyTabsSlot');
  if (!slot) return;

  slot.innerHTML = localAgencies.map(a => {
    const code = a.AgencyCode;
    const count = allContacts.filter(c =>
      !c.Archived && Array.isArray(c.DetectedCodes) && c.DetectedCodes.includes(code)
    ).length;
    const isActive = (currentAgencyCode === code) ? ' active' : '';
    return `<button class="tab tab-agency${isActive}" data-tab="agency" data-agency-code="${escapeHtml(code)}">
      ${escapeHtml(code)}<span class="ag-count">${count}</span>
    </button>`;
  }).join('');

  // Wire click handlers on the freshly-rendered tabs
  slot.querySelectorAll('.tab-agency').forEach(t => {
    t.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(x => x.classList.remove('active'));
      t.classList.add('active');
      document.getElementById('pane-agency').classList.add('active');
      currentAgencyCode = t.dataset.agencyCode;
      agencyPage = 1;
      renderAgency();
    });
  });
}

/* ===== Agency tab pane ===== */
const AGENCY_PAGE_SIZE = 120;
let agencyPage = 1;

document.getElementById('agencySearchBox').addEventListener('input', () => {
  agencyPage = 1;
  renderAgency();
});

function renderAgency() {
  if (!currentAgencyCode) return;
  const agency = localAgencies.find(a => a.AgencyCode === currentAgencyCode);
  const q = (document.getElementById('agencySearchBox').value || '').toLowerCase();
  const list = allContacts
    .filter(c => !c.Archived)
    .filter(c => Array.isArray(c.DetectedCodes) && c.DetectedCodes.includes(currentAgencyCode))
    .filter(c => !q || JSON.stringify(c).toLowerCase().includes(q));

  // Try common name field patterns
  const nameEn = agency ? (agency.AgencyName_EN || agency.AgencyNameEnglish || agency.AgencyName || agency.NameEn || agency.Name || '') : '';
  const nameAr = agency ? (agency.AgencyName_AR || agency.AgencyNameArabic || agency.NameAr || '') : '';
  const title  = currentAgencyCode + (nameEn ? ' — ' + nameEn : '') + (nameAr ? ' / ' + nameAr : '');

  document.getElementById('agencyTitle').textContent = title;
  document.getElementById('agencyCountLabel').textContent = list.length + ' contacts';
  document.getElementById('addCtaCode').textContent = currentAgencyCode;

  const grid  = document.getElementById('agencyGrid');
  const empty = document.getElementById('agencyEmpty');
  const pag   = document.getElementById('agencyPagination');

  if (list.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    pag.innerHTML = '';
    return;
  }
  empty.style.display = 'none';

  const totalPages = Math.max(1, Math.ceil(list.length / AGENCY_PAGE_SIZE));
  if (agencyPage > totalPages) agencyPage = totalPages;
  const start = (agencyPage - 1) * AGENCY_PAGE_SIZE;
  const pageItems = list.slice(start, start + AGENCY_PAGE_SIZE);

  grid.innerHTML = pageItems.map(contactCardHtml).join('');

  pag.innerHTML = paginationHtml(agencyPage, totalPages, list.length);
  pag.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      agencyPage = parseInt(btn.dataset.page, 10);
      renderAgency();
      document.getElementById('pane-agency').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

/* Agencies are managed in local_agencies_manager.html — this address book only reads. */

/* =====================================================
   IMPORTED CONTACTS — card grid + pagination
   ===================================================== */
const IMPORT_PAGE_SIZE = 120;   // 8 columns × 15 rows
let importedPage = 1;

document.getElementById('impSearchBox').addEventListener('input', () => {
  importedPage = 1;
  renderImported();
});

document.getElementById('btnImport').addEventListener('click', () => {
  document.getElementById('csvFileInput').click();
});

function renderImported() {
  const q = (document.getElementById('impSearchBox').value || '').toLowerCase();
  const list = allContacts
    .filter(c => !c.Archived)                                     // exclude archived
    .filter(c => !q || JSON.stringify(c).toLowerCase().includes(q));

  document.getElementById('impCountLabel').textContent = list.length + ' contacts';

  const grid  = document.getElementById('importedGrid');
  const empty = document.getElementById('importedEmpty');
  const pag   = document.getElementById('importedPagination');

  if (list.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'block';
    pag.innerHTML = '';
    return;
  }
  empty.style.display = 'none';

  const totalPages = Math.max(1, Math.ceil(list.length / IMPORT_PAGE_SIZE));
  if (importedPage > totalPages) importedPage = totalPages;
  const start = (importedPage - 1) * IMPORT_PAGE_SIZE;
  const pageItems = list.slice(start, start + IMPORT_PAGE_SIZE);

  grid.innerHTML = pageItems.map(contactCardHtml).join('');

  pag.innerHTML = paginationHtml(importedPage, totalPages, list.length);
  pag.querySelectorAll('[data-page]').forEach(btn => {
    btn.addEventListener('click', () => {
      importedPage = parseInt(btn.dataset.page, 10);
      renderImported();
      document.getElementById('pane-imported').scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

function paginationHtml(current, total, totalItems) {
  if (total <= 1) return `<span class="page-info">${totalItems} contacts on this page</span>`;

  let html = '';
  html += `<button class="page-btn" data-page="${Math.max(1, current - 1)}" ${current === 1 ? 'disabled' : ''}>‹ Prev</button>`;

  // Smart page numbers: always show first, last, current ±2
  const pages = new Set([1, total, current - 1, current, current + 1]);
  pages.add(current - 2); pages.add(current + 2);
  const sorted = [...pages].filter(p => p >= 1 && p <= total).sort((a,b) => a - b);

  let prev = 0;
  sorted.forEach(p => {
    if (prev && p - prev > 1) html += `<span class="page-info">…</span>`;
    html += `<button class="page-btn ${p === current ? 'active' : ''}" data-page="${p}">${p}</button>`;
    prev = p;
  });

  html += `<button class="page-btn" data-page="${Math.min(total, current + 1)}" ${current === total ? 'disabled' : ''}>Next ›</button>`;
  html += `<span class="page-info">Page ${current} of ${total} • ${totalItems} contacts</span>`;
  return html;
}

// Re-render imported cards whenever data updates is handled inside render()

/* =====================================================
   POPUP
   ===================================================== */
const modal      = document.getElementById('contactModal');
const backdrop   = document.getElementById('modalBackdrop');
const btnClose   = document.getElementById('modalClose');
const btnCancel  = document.getElementById('cancelBtn');
const btnSave    = document.getElementById('saveBtn');
const btnReorder = document.getElementById('reorderToggle');

// New "+ Add Contact" button lives on the Agency tab and pre-fills the agency code
document.getElementById('btnAddContactToAgency').addEventListener('click', () => {
  if (!currentAgencyCode) {
    alert('Please open an agency tab first.');
    return;
  }
  openModal(null, { presetAgencyCode: currentAgencyCode });
});

btnClose.addEventListener('click', closeModal);
btnCancel.addEventListener('click', closeModal);
backdrop.addEventListener('click', closeModal);
btnSave.addEventListener('click', saveContact);
btnReorder.addEventListener('click', toggleReorderMode);

let reorderMode = true; // default ON — arrows always visible; drag handle shown only in this mode

window.openModal = function(id, options) {
  options = options || {};
  currentEditingId = id;
  const existing = id ? (allContacts.find(c => c._id === id) || {}) : {};
  // For new contacts: pre-fill DetectedCodes with the agency we're adding under
  const data = id ? existing : (options.presetAgencyCode ? { DetectedCodes: [options.presetAgencyCode] } : {});

  // Modal title
  let title;
  if (id) {
    title = 'Edit Contact / تعديل جهة اتصال';
  } else if (options.presetAgencyCode) {
    title = `Add Contact to ${options.presetAgencyCode}`;
  } else {
    title = 'New Contact PopUp';
  }
  document.getElementById('modalTitle').textContent = title;

  // Populate Client Type section
  const ct = data.ClientType || 'Personal';
  document.querySelectorAll('input[name="ClientType"]').forEach(r => r.checked = (r.value === ct));
  document.querySelector('input[name="CompanyName_EN"]').value = data.CompanyName_EN || '';
  document.querySelector('input[name="CompanyName_AR"]').value = data.CompanyName_AR || '';
  document.getElementById('ctCompanyFields').classList.toggle('show', ct === 'Company');

  // Populate Archived checkbox + remember original state
  const archCb = document.getElementById('archivedCheckbox');
  archCb.checked = !!data.Archived;
  archCb.dataset.original = data.Archived ? '1' : '0';

  // Remember preset agency for save
  modal.dataset.presetAgencyCode = options.presetAgencyCode || '';

  buildForm(data);
  modal.classList.add('open');
  backdrop.style.display = 'block';
  modal.style.left = '50%';
  modal.style.top = '60px';
  modal.style.transform = 'translateX(-50%)';
};

// Wire Client Type radio toggle (show/hide company fields)
document.querySelectorAll('input[name="ClientType"]').forEach(r => {
  r.addEventListener('change', () => {
    document.getElementById('ctCompanyFields').classList.toggle('show', r.value === 'Company' && r.checked);
  });
});

function closeModal() {
  modal.classList.remove('open');
  backdrop.style.display = 'none';
  currentEditingId = null;
}

function toggleReorderMode() {
  reorderMode = !reorderMode;
  btnReorder.textContent = '↕ Reorder mode: ' + (reorderMode ? 'ON' : 'OFF');
  btnReorder.classList.toggle('on', reorderMode);
  const data = collectFormData();
  buildForm(data);
}
// initialize button label
btnReorder.textContent = '↕ Reorder mode: ON';
btnReorder.classList.add('on');

/* Modal drag */
(function makeDraggable() {
  const header = document.getElementById('modalHeader');
  let dragging = false, sx=0, sy=0, startLeft=0, startTop=0;
  header.addEventListener('mousedown', (e) => {
    if (e.target.tagName === 'BUTTON') return;
    dragging = true;
    const rect = modal.getBoundingClientRect();
    modal.style.left = rect.left + 'px';
    modal.style.top  = rect.top + 'px';
    modal.style.transform = 'none';
    sx = e.clientX; sy = e.clientY;
    startLeft = rect.left; startTop = rect.top;
    e.preventDefault();
  });
  document.addEventListener('mousemove', (e) => {
    if (!dragging) return;
    modal.style.left = (startLeft + e.clientX - sx) + 'px';
    modal.style.top  = (startTop  + e.clientY - sy) + 'px';
  });
  document.addEventListener('mouseup', () => dragging = false);
})();

/* =====================================================
   APPLY SAVED ORDER TO A GROUP
   Returns fields array in saved order, appending any new fields
   ===================================================== */
function orderedFieldsFor(group) {
  // Start with group definition fields + any custom fields for this group
  const all = [
    ...group.fields,
    ...customFields.filter(cf => cf.group === group.key)
  ];

  const savedOrder = fieldOrder[group.key];
  if (!savedOrder || !savedOrder.length) return all;

  // Re-order by savedOrder, append any fields not in savedOrder at the end
  const byName = Object.fromEntries(all.map(f => [f.name, f]));
  const ordered = [];
  savedOrder.forEach(name => {
    if (byName[name]) {
      ordered.push(byName[name]);
      delete byName[name];
    }
  });
  // Append any remaining (newly added) fields
  Object.values(byName).forEach(f => ordered.push(f));
  return ordered;
}

/* =====================================================
   FORM BUILD
   ===================================================== */
function buildForm(data) {
  const container = document.getElementById('groupsContainer');

  // Preserve currently open group across rebuilds; default to first group on initial build
  let openKey = null;
  const currentOpen = container.querySelector('.group.open');
  if (currentOpen) openKey = currentOpen.dataset.groupKey;
  if (!openKey) openKey = GROUPS[0].key;

  container.innerHTML = '';

  GROUPS.forEach((g) => {
    const fields = orderedFieldsFor(g);
    const groupDiv = document.createElement('div');
    groupDiv.className = 'group' + (g.key === openKey ? ' open' : '');
    groupDiv.dataset.groupKey = g.key;
    if (g.isCustom && g.key === 'idpassport') {
      groupDiv.innerHTML = `
        <div class="group-header" data-group="${g.key}">
          <div>
            <span class="chevron">▶</span>
            ${escapeHtml(g.label_en)}
            <span class="ar">${escapeHtml(g.label_ar)}</span>
          </div>
          <span style="font-size:12px;color:#888;">Drop files here</span>
        </div>
        <div class="group-body">
          ${renderIDPassportPanel(data)}
        </div>
      `;
    } else if (g.isCustom && g.key === 'navigation') {
      groupDiv.innerHTML = `
        <div class="group-header" data-group="${g.key}">
          <div>
            <span class="chevron">▶</span>
            ${escapeHtml(g.label_en)}
            <span class="ar">${escapeHtml(g.label_ar)}</span>
          </div>
          <span style="font-size:12px;color:#888;">GPS locations</span>
        </div>
        <div class="group-body">
          ${renderNavigationPanel(data)}
        </div>
      `;
    } else {
      groupDiv.innerHTML = `
        <div class="group-header" data-group="${g.key}">
          <div>
            <span class="chevron">▶</span>
            ${escapeHtml(g.label_en)}
            <span class="ar">${escapeHtml(g.label_ar)}</span>
          </div>
          <span style="font-size:12px;color:#888;">${fields.length} fields</span>
        </div>
        <div class="group-body">
          <div class="fields-list" data-group="${g.key}">
            ${fields.map(f => fieldRowHtml(f, data, g.key)).join('')}
          </div>
          <div class="add-field-row">
            <button type="button" class="add-field-btn" onclick="addCustomField('${g.key}')">
              ＋ Add Custom Field / إضافة حقل جديد
            </button>
          </div>
        </div>
      `;
    }
    container.appendChild(groupDiv);
  });

  // Accordion behavior
  container.querySelectorAll('.group-header').forEach(h => {
    h.addEventListener('click', () => {
      const parent = h.parentElement;
      const isOpen = parent.classList.contains('open');
      container.querySelectorAll('.group').forEach(g => g.classList.remove('open'));
      if (!isOpen) parent.classList.add('open');
    });
  });

  // Wire drag-and-drop
  container.querySelectorAll('.fields-list').forEach(list => wireDragDrop(list));

  // Auto-compute full names from the 6 name parts. Runs once now (initial populate)
  // and again on every keystroke in any of the part fields. Manual edits to EFN
  // or EmployerName are kept only until the user types in a part field again.
  wireFullNameAutoCompute(container);
}

function wireFullNameAutoCompute(container) {
  const efnEl  = container.querySelector('input[name="EFN"], textarea[name="EFN"]');
  const empEl  = container.querySelector('input[name="EmployerName"], textarea[name="EmployerName"]');
  const f_en   = container.querySelector('input[name="FirstName_EN"]');
  const m_en   = container.querySelector('input[name="MiddleName_EN"]');
  const l_en   = container.querySelector('input[name="LastName_EN"]');
  const f_ar   = container.querySelector('input[name="FirstName_AR"]');
  const m_ar   = container.querySelector('input[name="MiddleName_AR"]');
  const l_ar   = container.querySelector('input[name="LastName_AR"]');

  function joinParts(a, b, c) {
    return [a && a.value, b && b.value, c && c.value].filter(Boolean).join(' ').trim();
  }
  function refreshEN() { if (efnEl) efnEl.value = joinParts(f_en, m_en, l_en); }
  function refreshAR() { if (empEl) empEl.value = joinParts(f_ar, m_ar, l_ar); }

  [f_en, m_en, l_en].forEach(el => el && el.addEventListener('input', refreshEN));
  [f_ar, m_ar, l_ar].forEach(el => el && el.addEventListener('input', refreshAR));

  // Initial sync — only auto-fill if the target field is currently empty,
  // so we don't overwrite existing stored values.
  if (efnEl && !efnEl.value.trim()) refreshEN();
  if (empEl && !empEl.value.trim()) refreshAR();
}

/* =====================================================
   ID & PASSPORT COPY PANEL — drag&drop, uploads to Firebase Storage
   ===================================================== */
function renderIDPassportPanel(data) {
  const cid = String(data.ContactID || '').trim() || '(will be assigned on save)';
  const idURL  = data.IDCopyURL || '';
  const idName = data.IDCopyName || '';
  const psURL  = data.PassportCopyURL || '';
  const psName = data.PassportCopyName || '';
  const styleCommon =
    'border:2px dashed #4caf50;border-radius:10px;padding:18px;text-align:center;'+
    'background:#f1f8e9;cursor:pointer;transition:all .2s;';
  function slot(kind, label, currentURL, currentName) {
    const slotId = 'idp-' + kind;
    const has = currentURL ? true : false;
    return `
      <div class="idp-slot" data-kind="${kind}" id="${slotId}"
           style="${styleCommon}"
           ondragover="event.preventDefault();this.style.background='#dcedc8';"
           ondragleave="this.style.background='#f1f8e9';"
           ondrop="handleIDPDrop(event,'${kind}')"
           onclick="document.getElementById('${slotId}-input').click()">
        <div style="font-size:28px;line-height:1;">📎</div>
        <div style="font-weight:700;color:#33691e;margin-top:4px;">${label}</div>
        <div style="font-size:11px;color:#558b2f;margin-top:2px;">Click or drag a file (JPG, PNG, PDF) — max 10 MB</div>
        ${has ? `
          <div class="idp-preview" style="margin-top:10px;padding:8px 10px;background:white;border:1px solid #c8e6c9;border-radius:6px;display:flex;align-items:center;gap:8px;justify-content:space-between;">
            <a href="${currentURL}" target="_blank" rel="noopener" style="flex:1;color:#1976d2;font-size:12px;word-break:break-all;text-decoration:none;">📂 ${escapeHtml(currentName || 'View file')}</a>
            <button type="button" onclick="event.stopPropagation();removeIDPFile('${kind}')" style="background:#e57373;color:white;border:none;padding:4px 8px;border-radius:4px;cursor:pointer;font-size:11px;">✕ Remove</button>
          </div>
        ` : `<div class="idp-preview" style="display:none;"></div>`}
        <div class="idp-status" style="margin-top:6px;font-size:11px;color:#888;"></div>
        <input type="file" id="${slotId}-input" accept="image/*,application/pdf" style="display:none;"
               onchange="if(this.files[0]){handleIDPFile(this.files[0],'${kind}');this.value='';}">
      </div>`;
  }
  return `
    <div style="padding:6px 8px 10px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;font-size:12px;color:#666;">
        <span>Contact ID: <b>${escapeHtml(cid)}</b></span>
        <span style="color:#33691e;font-weight:600;">Active marker = ID or Passport uploaded</span>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;">
        ${slot('id',       '🪪 Drop ID Copy here',       idURL, idName)}
        ${slot('passport', '📕 Drop Passport Copy here', psURL, psName)}
      </div>
    </div>
  `;
}

window.handleIDPDrop = function(e, kind) {
  e.preventDefault();
  const slot = document.getElementById('idp-' + kind);
  if (slot) slot.style.background = '#f1f8e9';
  const files = e.dataTransfer && e.dataTransfer.files ? e.dataTransfer.files : [];
  if (files && files[0]) handleIDPFile(files[0], kind);
};

window.handleIDPFile = async function(file, kind) {
  if (!currentEditingId) {
    alert('Save the contact once before uploading ID/Passport so it has a ContactID.');
    return;
  }
  if (file.size > 10 * 1024 * 1024) { alert('File too large. Max 10 MB.'); return; }
  const slot = document.getElementById('idp-' + kind);
  const statusEl = slot ? slot.querySelector('.idp-status') : null;
  if (statusEl) { statusEl.textContent = '⏳ Uploading…'; statusEl.style.color = '#1976d2'; }

  try {
    const { storage, storageRef, uploadBytes, getDownloadURL, db, doc, updateDoc, serverTimestamp } = window._fb;
    const contact = allContacts.find(c => c._id === currentEditingId);
    const contactID = (contact && contact.ContactID) ? String(contact.ContactID) : currentEditingId;
    const codes = Array.isArray(contact && contact.DetectedCodes) ? contact.DetectedCodes : [];
    const primaryCode = (codes[0] || 'NOCODE').toUpperCase();
    const efn =
      (contact && contact.EFN)
      || [contact && contact.FirstName_EN, contact && contact.MiddleName_EN, contact && contact.LastName_EN].filter(Boolean).join(' ').trim()
      || (contact && contact.EmployerName)
      || 'Unknown';
    const today = new Date().toISOString().slice(0, 10);
    // Strip characters Firebase Storage doesn't accept in path segments
    const safeEFN = String(efn).replace(/[\/\\#?*\n\r]/g, ' ').replace(/\s+/g, ' ').trim();
    const folderName = `${primaryCode}(${contactID}) - ${safeEFN} - ${today}`;
    const ext = (file.name.match(/\.([a-zA-Z0-9]+)$/) || [,'bin'])[1];
    const typeLabel = kind === 'id' ? 'ID' : 'Passport';
    const fileName = `${primaryCode}-${contactID}-${typeLabel}.${ext}`;
    const path = `employers/${folderName}/${fileName}`;
    const r = storageRef(storage, path);
    await uploadBytes(r, file, { contentType: file.type || 'application/octet-stream' });
    const url = await getDownloadURL(r);

    // Stamp the contact + flag as active. For each agency code in DetectedCodes,
    // create the implicit employer "folder" by writing an employer_settings doc.
    const fieldMap = kind === 'id'
      ? { IDCopyURL: url, IDCopyName: file.name, IDCopyPath: path, IDCopyUploadedAt: new Date().toISOString() }
      : { PassportCopyURL: url, PassportCopyName: file.name, PassportCopyPath: path, PassportCopyUploadedAt: new Date().toISOString() };
    fieldMap._userEdited = true;
    fieldMap._activeEmployer = true;
    fieldMap._activatedAt = serverTimestamp();
    fieldMap.EmployerFolderName = folderName;
    await updateDoc(doc(db, 'address_book', currentEditingId), fieldMap);

    // Create per-(contact, agency) employer_settings markers so the Employers Manager shows them as Active
    const { setDoc } = window._fb;
    for (const code of codes) {
      const key = contactID + '-' + code;
      await setDoc(doc(db, 'employer_settings', key), {
        EmployerKey:        key,
        EmployerID:         contactID,
        ContactID:          contactID,
        AgencyCode:         code,
        EmployerFolderName: folderName,
        EmployerFolderID:   '',
        RequiredDocuments:  [],
        UploadedDocuments:  {},
        Disabled:           false,
        ActivatedAt:        new Date().toISOString()
      }, { merge: true });
    }

    if (statusEl) { statusEl.textContent = '✅ Uploaded'; statusEl.style.color = '#2e7d32'; }
    // Re-render the panel to show the preview/remove
    setTimeout(() => {
      const fresh = allContacts.find(c => c._id === currentEditingId) || {};
      Object.assign(fresh, fieldMap);
      const groupDiv = document.querySelector('.group[data-group-key="idpassport"] .group-body');
      if (groupDiv) groupDiv.innerHTML = renderIDPassportPanel(fresh);
    }, 400);
  } catch (err) {
    if (statusEl) { statusEl.textContent = '❌ ' + (err.message || err); statusEl.style.color = '#c62828'; }
    alert('Upload failed: ' + (err.message || err));
  }
};

window.removeIDPFile = async function(kind) {
  if (!currentEditingId) return;
  if (!confirm('Remove this file? It will be deleted from storage.')) return;
  try {
    const { storage, storageRef, deleteObject, db, doc, updateDoc } = window._fb;
    const contact = allContacts.find(c => c._id === currentEditingId) || {};
    const path = kind === 'id' ? contact.IDCopyPath : contact.PassportCopyPath;
    if (path) {
      try { await deleteObject(storageRef(storage, path)); } catch(_) {}
    }
    const fields = kind === 'id'
      ? { IDCopyURL: '', IDCopyName: '', IDCopyPath: '' }
      : { PassportCopyURL: '', PassportCopyName: '', PassportCopyPath: '' };
    await updateDoc(doc(db, 'address_book', currentEditingId), fields);
    Object.assign(contact, fields);
    const groupDiv = document.querySelector('.group[data-group-key="idpassport"] .group-body');
    if (groupDiv) groupDiv.innerHTML = renderIDPassportPanel(contact);
  } catch (err) {
    alert('Remove failed: ' + (err.message || err));
  }
};

/* =====================================================
   NAVIGATION PANEL — Home GPS + Work GPS (Capture, Open, Navigate, Share)
   ===================================================== */
function renderNavigationPanel(data) {
  // Migrate any legacy `Location` field to Home
  const home = data.HomeLocation || data.Location || null;
  const work = data.WorkLocation || null;
  const contactName = escapeHtml(data.EFN || [data.FirstName_EN, data.MiddleName_EN, data.LastName_EN].filter(Boolean).join(' ').trim() || data.EmployerName || 'Contact');

  function slot(kind, label, icon, loc) {
    const has = loc && loc.lat != null && loc.lng != null;
    const acc = (has && loc.accuracy != null) ? ` ±${Math.round(loc.accuracy)}m` : '';
    const captured = (has && loc.capturedAt) ? new Date(loc.capturedAt).toLocaleString() : '';
    const mapsUrl   = has ? `https://www.google.com/maps?q=${loc.lat},${loc.lng}` : '';
    const navUrl    = has ? `https://www.google.com/maps/dir/?api=1&destination=${loc.lat},${loc.lng}` : '';
    return `
      <div class="nav-slot" data-kind="${kind}" style="border:1px solid #d0e3ff;border-radius:10px;padding:14px;margin-bottom:10px;background:#f5f9ff;">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;">
          <div style="font-weight:700;color:#1e40af;font-size:14px;">${icon} ${escapeHtml(label)}</div>
          <div style="font-size:11px;color:#666;">${has ? '🗺️ Saved' + acc : 'No GPS yet'}</div>
        </div>
        <div id="nav-status-${kind}" style="font-size:12px;color:#555;margin-bottom:8px;">
          ${has
            ? `Coordinates: <code style="background:white;padding:2px 5px;border-radius:3px;">${loc.lat.toFixed(6)}, ${loc.lng.toFixed(6)}</code>${captured ? '<br><span style="font-size:10px;color:#888;">Captured: ' + escapeHtml(captured) + '</span>' : ''}`
            : 'Click <b>Capture</b> to record this contact GPS coordinates.'}
        </div>
        <div style="display:flex;gap:6px;flex-wrap:wrap;">
          <button type="button" class="nav-btn capture" id="nav-capture-${kind}" onclick="captureNavLocation('${kind}')"
                  style="background:#10b981;color:white;border:none;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">🗺️ Capture</button>
          ${has ? `
            <a class="nav-btn" href="${mapsUrl}" target="_blank" rel="noopener"
               style="background:#0d6efd;color:white;text-decoration:none;padding:6px 12px;border-radius:5px;font-size:12px;font-weight:600;">📂 Open</a>
            <a class="nav-btn" href="${navUrl}" target="_blank" rel="noopener"
               style="background:#7c3aed;color:white;text-decoration:none;padding:6px 12px;border-radius:5px;font-size:12px;font-weight:600;">🧭 Navigate</a>
            <button type="button" class="nav-btn"
                    onclick="shareNavLocation(${loc.lat}, ${loc.lng}, '${contactName} — ${escapeHtml(label)}')"
                    style="background:#f59e0b;color:white;border:none;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">📤 Share</button>
            <button type="button" class="nav-btn"
                    onclick="editNavLocation('${kind}')"
                    style="background:#e0e7ff;color:#3730a3;border:1px solid #c7d2fe;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">✏️ Edit</button>
          ` : `
            <button type="button" class="nav-btn"
                    onclick="editNavLocation('${kind}')"
                    style="background:#e0e7ff;color:#3730a3;border:1px solid #c7d2fe;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">✏️ Paste Link / Coordinates</button>
          `}
        </div>
      </div>
    `;
  }

  return `
    <div style="padding:6px 8px 4px;">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;font-size:12px;color:#666;">
        <span>Home / Work GPS coordinates for this contact.</span>
        <span style="color:#1e40af;font-weight:600;">Click Capture to record current position.</span>
      </div>
      ${slot('home', 'Home GPS',  '🏠', home)}
      ${slot('work', 'Work GPS',  '🏢', work)}
    </div>
  `;
}

window.captureNavLocation = function(kind) {
  if (!currentEditingId) {
    alert('Save the contact once before capturing GPS so it has a ContactID.');
    return;
  }
  if (!navigator.geolocation) { alert('Geolocation not supported by this browser.'); return; }

  const statusEl = document.getElementById('nav-status-' + kind);
  const btn      = document.getElementById('nav-capture-' + kind);
  if (!statusEl || !btn) return;

  const TARGET_ACC  = 5;
  const MAX_WAIT_MS = 60000;
  let best = null, watchId = null, stopped = false;
  const startedAt = Date.now();

  btn.disabled = true;
  btn.textContent = '⏳ Acquiring…';
  statusEl.innerHTML = `
    <div id="nav-live-${kind}" style="padding:8px;background:#fff3cd;border:1px solid #ffe69c;border-radius:6px;font-size:12px;">
      ⏳ Waiting for accuracy ≤ ${TARGET_ACC}m. (Higher accuracy outdoors.)
    </div>
    <button type="button" onclick="finishNavCapture('${kind}')" style="margin-top:6px;background:#6c757d;color:white;border:none;padding:5px 10px;border-radius:5px;font-size:11px;cursor:pointer;">Use Current Reading</button>
  `;

  async function commit(reason) {
    if (stopped) return;
    stopped = true;
    if (watchId != null) navigator.geolocation.clearWatch(watchId);
    if (!best) {
      document.getElementById('nav-live-' + kind).innerHTML = '❌ Could not get any GPS reading.';
      btn.disabled = false; btn.textContent = '🗺️ Try Again';
      return;
    }
    const reading = { lat: best.lat, lng: best.lng, accuracy: best.acc, capturedAt: new Date().toISOString() };
    // Persist to Firestore immediately
    try {
      const { db, doc, updateDoc, serverTimestamp } = window._fb;
      const field = kind === 'home' ? 'HomeLocation' : 'WorkLocation';
      const patch = { _userEdited: true, _updatedAt: serverTimestamp() };
      patch[field] = reading;
      await updateDoc(doc(db, 'address_book', currentEditingId), patch);
      // Reflect in local cache + re-render this panel only
      const contact = allContacts.find(c => c._id === currentEditingId) || {};
      contact[field] = reading;
      const body = document.querySelector('.group[data-group-key="navigation"] .group-body');
      if (body) body.innerHTML = renderNavigationPanel(contact);
    } catch (err) {
      alert('Save failed: ' + (err.message || err));
      btn.disabled = false; btn.textContent = '🗺️ Capture';
    }
  }
  window['finishNavCapture'] = (k) => commit('using current reading');
  window['_navCaptureCommit_' + kind] = commit;

  watchId = navigator.geolocation.watchPosition(
    pos => {
      if (stopped) return;
      const lat = pos.coords.latitude, lng = pos.coords.longitude, acc = pos.coords.accuracy;
      if (!best || acc < best.acc) best = { lat, lng, acc };
      const live = document.getElementById('nav-live-' + kind);
      if (live) live.innerHTML = `⏳ Current: ±${Math.round(acc)}m · Best: ±${Math.round(best.acc)}m · Target ≤ ${TARGET_ACC}m`;
      if (acc <= TARGET_ACC) commit('target reached');
      else if (Date.now() - startedAt > MAX_WAIT_MS) commit('60s timeout — best used');
    },
    err => {
      const msg = err.code === 1 ? 'Permission denied' : err.code === 2 ? 'Position unavailable' : err.code === 3 ? 'Timed out' : (err.message || 'Unknown');
      const live = document.getElementById('nav-live-' + kind);
      if (live) { live.innerHTML = '❌ ' + msg; live.style.background = '#f8d7da'; }
      stopped = true;
      if (watchId != null) navigator.geolocation.clearWatch(watchId);
      btn.disabled = false; btn.textContent = '🗺️ Try Again';
    },
    { enableHighAccuracy: true, timeout: 30000, maximumAge: 0 }
  );
};

window.shareNavLocation = async function(lat, lng, title) {
  const url = `https://www.google.com/maps?q=${lat},${lng}`;
  const text = `📍 ${title}\n${url}`;
  if (navigator.share) {
    try { await navigator.share({ title, text, url }); return; }
    catch (e) { if (e.name === 'AbortError') return; }
  }
  try {
    await navigator.clipboard.writeText(text);
    alert('📋 Link copied to clipboard');
  } catch (e) {
    prompt('Copy this location link:', text);
  }
};

window.removeNavLocation = async function(kind) {
  if (!currentEditingId) return;
  const label = kind === 'home' ? 'Home' : 'Work';
  if (!confirm('Remove the saved ' + label + ' GPS?')) return;
  try {
    const { db, doc, updateDoc } = window._fb;
    const field = kind === 'home' ? 'HomeLocation' : 'WorkLocation';
    const patch = {};
    patch[field] = null;
    await updateDoc(doc(db, 'address_book', currentEditingId), patch);
    const contact = allContacts.find(c => c._id === currentEditingId) || {};
    contact[field] = null;
    const body = document.querySelector('.group[data-group-key="navigation"] .group-body');
    if (body) body.innerHTML = renderNavigationPanel(contact);
  } catch (err) {
    alert('Remove failed: ' + (err.message || err));
  }
};

/**
 * Parse a pasted Google Maps URL or plain "lat,lng" string into {lat, lng}.
 * Supports common formats:
 *   https://www.google.com/maps?q=LAT,LNG
 *   https://www.google.com/maps/place/Name/@LAT,LNG,zoom
 *   https://www.google.com/maps/dir/?destination=LAT,LNG
 *   https://maps.google.com/?ll=LAT,LNG
 *   https://www.google.com/maps?q=loc:LAT,LNG
 *   LAT,LNG (raw)
 * Returns null if it can't extract numeric coordinates.
 */
function parseLocationInput(text) {
  if (!text) return null;
  const s = String(text).trim();
  // First, simple "lat, lng" anywhere — most reliable
  const m = s.match(/(-?\d{1,3}\.\d+)\s*,\s*(-?\d{1,3}\.\d+)/);
  if (m) {
    const lat = parseFloat(m[1]);
    const lng = parseFloat(m[2]);
    if (Math.abs(lat) <= 90 && Math.abs(lng) <= 180) return { lat, lng };
  }
  return null;
}

window.editNavLocation = async function(kind) {
  if (!currentEditingId) {
    alert('Save the contact once before editing GPS so it has a ContactID.');
    return;
  }
  const label = kind === 'home' ? 'Home' : 'Work';
  const contact = allContacts.find(c => c._id === currentEditingId) || {};
  const field   = kind === 'home' ? 'HomeLocation' : 'WorkLocation';
  const existing = contact[field] || (kind === 'home' ? contact.Location : null);
  const current = existing && existing.lat != null ? (existing.lat + ', ' + existing.lng) : '';

  const slot = document.querySelector('.nav-slot[data-kind="' + kind + '"]');
  if (!slot) return;

  // Replace slot inner with an inline form
  slot.innerHTML = `
    <div style="font-weight:700;color:#1e40af;font-size:14px;margin-bottom:8px;">${kind === 'home' ? '🏠' : '🏢'} Edit ${label} GPS</div>
    <div style="font-size:12px;color:#555;margin-bottom:6px;">
      Paste a Google Maps link, or type coordinates like <code style="background:white;padding:1px 4px;border-radius:3px;">33.8463, 35.9019</code>:
    </div>
    <textarea id="nav-input-${kind}" rows="2"
              style="width:100%;padding:8px;border:1px solid #c7d2fe;border-radius:5px;font-family:inherit;font-size:13px;resize:vertical;"
              placeholder="Paste WhatsApp/Email link, or type lat, lng…">${escapeHtml(current)}</textarea>
    <div id="nav-input-status-${kind}" style="font-size:11px;color:#666;margin-top:6px;min-height:14px;"></div>
    <div style="display:flex;gap:6px;margin-top:8px;flex-wrap:wrap;">
      <button type="button" onclick="saveNavLocationFromInput('${kind}')"
              style="background:#10b981;color:white;border:none;padding:6px 14px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">💾 Save</button>
      ${existing ? `<button type="button" onclick="clearNavLocation('${kind}')"
              style="background:#fee2e2;color:#991b1b;border:1px solid #fecaca;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">🗑️ Clear</button>` : ''}
      <button type="button" onclick="cancelNavLocationEdit()"
              style="background:#f3f4f6;color:#444;border:1px solid #d1d5db;padding:6px 12px;border-radius:5px;cursor:pointer;font-size:12px;font-weight:600;">Cancel</button>
    </div>
    <div style="font-size:11px;color:#888;margin-top:8px;">
      Tip: short links (maps.app.goo.gl) need to be opened first — tap the link, then copy the URL from the address bar.
    </div>
  `;
};

window.saveNavLocationFromInput = async function(kind) {
  const input = document.getElementById('nav-input-' + kind);
  const status = document.getElementById('nav-input-status-' + kind);
  const parsed = parseLocationInput(input.value);
  if (!parsed) {
    status.textContent = '❌ Could not find lat,lng coordinates in that text. Try opening the link first and copying the full URL.';
    status.style.color = '#c62828';
    return;
  }
  status.textContent = '⏳ Saving…';
  status.style.color = '#0066cc';
  try {
    const { db, doc, updateDoc, serverTimestamp } = window._fb;
    const field = kind === 'home' ? 'HomeLocation' : 'WorkLocation';
    const reading = { lat: parsed.lat, lng: parsed.lng, source: 'pasted', capturedAt: new Date().toISOString() };
    const patch = { _userEdited: true, _updatedAt: serverTimestamp() };
    patch[field] = reading;
    await updateDoc(doc(db, 'address_book', currentEditingId), patch);
    const contact = allContacts.find(c => c._id === currentEditingId) || {};
    contact[field] = reading;
    const body = document.querySelector('.group[data-group-key="navigation"] .group-body');
    if (body) body.innerHTML = renderNavigationPanel(contact);
  } catch (err) {
    status.textContent = '❌ Save failed: ' + (err.message || err);
    status.style.color = '#c62828';
  }
};

window.clearNavLocation = async function(kind) {
  if (!currentEditingId) return;
  if (!confirm('Clear the saved ' + (kind === 'home' ? 'Home' : 'Work') + ' GPS?')) return;
  try {
    const { db, doc, updateDoc } = window._fb;
    const field = kind === 'home' ? 'HomeLocation' : 'WorkLocation';
    const patch = {};
    patch[field] = null;
    await updateDoc(doc(db, 'address_book', currentEditingId), patch);
    const contact = allContacts.find(c => c._id === currentEditingId) || {};
    contact[field] = null;
    const body = document.querySelector('.group[data-group-key="navigation"] .group-body');
    if (body) body.innerHTML = renderNavigationPanel(contact);
  } catch (err) {
    alert('Clear failed: ' + (err.message || err));
  }
};

window.cancelNavLocationEdit = function() {
  const contact = allContacts.find(c => c._id === currentEditingId) || {};
  const body = document.querySelector('.group[data-group-key="navigation"] .group-body');
  if (body) body.innerHTML = renderNavigationPanel(contact);
};

function fieldRowHtml(f, data, groupKey) {
  let v = data[f.name] != null ? data[f.name] : '';
  if (Array.isArray(v)) v = v.join(' ');                       // arrays → space-separated text
  const type = f.type || 'text';
  let input;
  if (type === 'textarea') {
    input = `<textarea name="${f.name}" rows="1">${escapeHtml(v)}</textarea>`;
  } else if (type === 'checkbox') {
    input = `<input name="${f.name}" type="checkbox" ${v === true || v === 'TRUE' || v === 'true' ? 'checked':''}>`;
  } else {
    const extra = f.isPrefixCodes ? ' placeholder="MGSC LRA NCS" style="text-transform:uppercase;"' : '';
    input = `<input name="${f.name}" type="${type}" value="${escapeHtml(v)}"${extra}>`;
  }
  return `
    <div class="field-row${f.isArabic ? ' is-arabic' : ''}" data-field="${f.name}" data-group="${groupKey}" draggable="${reorderMode}">
      <span class="drag-handle" title="Drag to reorder" style="display:${reorderMode ? 'inline-block':'none'};" tabindex="-1">⋮⋮</span>
      <div>
        <div class="label-en">${escapeHtml(f.en)}</div>
        <div class="label-ar">${escapeHtml(f.ar)}</div>
      </div>
      ${input}
      <div class="arrows">
        <button type="button" class="arrow-btn" title="Move up"   tabindex="-1" onclick="moveField('${groupKey}','${f.name}',-1)">▲</button>
        <button type="button" class="arrow-btn" title="Move down" tabindex="-1" onclick="moveField('${groupKey}','${f.name}', 1)">▼</button>
      </div>
    </div>
  `;
}

/* =====================================================
   REORDER — up/down arrows
   ===================================================== */
window.moveField = async function(groupKey, fieldName, delta) {
  const g = GROUPS.find(g => g.key === groupKey);
  const current = orderedFieldsFor(g).map(f => f.name);
  const i = current.indexOf(fieldName);
  if (i === -1) return;
  const j = i + delta;
  if (j < 0 || j >= current.length) return;
  [current[i], current[j]] = [current[j], current[i]];
  fieldOrder[groupKey] = current;
  await persistOrder();
  const data = collectFormData();
  buildForm(data);
};

/* =====================================================
   REORDER — drag and drop
   ===================================================== */
function wireDragDrop(list) {
  let draggedEl = null;
  list.querySelectorAll('.field-row').forEach(row => {
    row.addEventListener('dragstart', (e) => {
      draggedEl = row;
      row.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move';
      try { e.dataTransfer.setData('text/plain', row.dataset.field); } catch(_) {}
    });
    row.addEventListener('dragend', () => {
      row.classList.remove('dragging');
      list.querySelectorAll('.field-row').forEach(r => r.classList.remove('drag-over'));
      draggedEl = null;
    });
    row.addEventListener('dragover', (e) => {
      e.preventDefault();
      if (!draggedEl || draggedEl === row) return;
      list.querySelectorAll('.field-row').forEach(r => r.classList.remove('drag-over'));
      row.classList.add('drag-over');
    });
    row.addEventListener('dragleave', () => row.classList.remove('drag-over'));
    row.addEventListener('drop', async (e) => {
      e.preventDefault();
      row.classList.remove('drag-over');
      if (!draggedEl || draggedEl === row) return;
      const groupKey = row.dataset.group;
      const g = GROUPS.find(g => g.key === groupKey);
      const current = orderedFieldsFor(g).map(f => f.name);
      const fromIdx = current.indexOf(draggedEl.dataset.field);
      const toIdx   = current.indexOf(row.dataset.field);
      if (fromIdx === -1 || toIdx === -1) return;
      const [moved] = current.splice(fromIdx, 1);
      current.splice(toIdx, 0, moved);
      fieldOrder[groupKey] = current;
      await persistOrder();
      const data = collectFormData();
      buildForm(data);
    });
  });
}

async function persistOrder() {
  const { db, doc, setDoc } = window._fb;
  try {
    await setDoc(doc(db, '_address_book_meta', 'field_order'), { orders: fieldOrder });
  } catch (e) {
    console.warn('Could not save field order:', e);
  }
}

/* =====================================================
   ADD CUSTOM FIELD
   ===================================================== */
window.addCustomField = async function(groupKey) {
  const en = prompt('Field name (English):');
  if (!en) return;
  const ar = prompt('اسم الحقل (عربي):');
  if (!ar) return;
  const safe = en.replace(/[^a-zA-Z0-9_]/g, '_');
  customFields.push({ name: safe, en, ar, group: groupKey });

  const { db, doc, setDoc } = window._fb;
  try {
    await setDoc(doc(db, '_address_book_meta', 'custom_fields'), { fields: customFields });
  } catch(e) {
    alert('Could not save custom field definition: ' + e.message);
    return;
  }

  const data = collectFormData();
  buildForm(data);
};

/* =====================================================
   SAVE CONTACT
   ===================================================== */
function collectFormData() {
  const data = {};
  document.querySelectorAll('#contactModal [name]').forEach(el => {
    if (el.type === 'checkbox') data[el.name] = el.checked;
    else if (el.type === 'radio') { if (el.checked) data[el.name] = el.value; }
    else data[el.name] = el.value;
  });
  return data;
}

async function saveContact() {
  const data = collectFormData();

  // Convert Archived from checkbox value to true/false
  data.Archived = document.getElementById('archivedCheckbox').checked;

  // Convert DetectedCodes from space-separated text input back to a deduped uppercase array
  if (typeof data.DetectedCodes === 'string') {
    data.DetectedCodes = [...new Set(
      data.DetectedCodes.toUpperCase().split(/[\s,]+/).filter(s => /^[A-Z]{2,6}$/.test(s))
    )];
  }

  // If we opened this modal as "Add to <Agency>", make sure that code is in DetectedCodes
  const presetAgency = modal.dataset.presetAgencyCode || '';
  if (!currentEditingId && presetAgency) {
    const codes = Array.isArray(data.DetectedCodes) ? data.DetectedCodes.slice() : [];
    if (!codes.includes(presetAgency)) codes.push(presetAgency);
    data.DetectedCodes = codes;
  }

  // Password gate: changing from Archived=true to Archived=false
  const wasArchived = document.getElementById('archivedCheckbox').dataset.original === '1';
  if (wasArchived && !data.Archived) {
    const pwd = prompt('Enter admin password to UN-archive this contact:\nأدخل كلمة المرور لإلغاء أرشفة جهة الاتصال:');
    if (pwd !== 'masr3434') {
      alert('Wrong password. Contact remains archived.\nكلمة مرور خاطئة.');
      document.getElementById('archivedCheckbox').checked = true;
      data.Archived = true;
      return;
    }
  }

  const { db, collection, addDoc, doc, updateDoc, serverTimestamp } = window._fb;
  try {
    btnSave.disabled = true;
    btnSave.textContent = 'Saving…';
    if (currentEditingId) {
      await updateDoc(doc(db, 'address_book', currentEditingId), { ...data, _userEdited: true, _updatedAt: serverTimestamp() });
    } else {
      // Auto-assign next ContactID
      const existingMax = allContacts.reduce((max, c) => {
        const n = parseInt(c.ContactID, 10);
        return (Number.isFinite(n) && n > max) ? n : max;
      }, 0);
      const newId = String(existingMax + 1).padStart(4, '0');
      await addDoc(collection(db, 'address_book'), { ...data, ContactID: newId, _userEdited: true, _createdAt: serverTimestamp() });
    }
    closeModal();
  } catch (e) {
    alert('Save failed: ' + e.message);
  } finally {
    btnSave.disabled = false;
    btnSave.textContent = 'Save / حفظ';
  }
}

/* =====================================================
   GOOGLE CONTACTS CSV IMPORT
   ===================================================== */
const importModal       = document.getElementById('importModal');
const importModalBody   = document.getElementById('importModalBody');
const importModalFooter = document.getElementById('importModalFooter');
document.getElementById('importModalClose').addEventListener('click', () => importModal.classList.remove('open'));

document.getElementById('csvFileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;

  showImportModal('Parsing CSV…', `<p>Reading <b>${escapeHtml(file.name)}</b>…</p>`, '');

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete: (results) => {
      const rows = results.data || [];
      const parsed = rows.map(parseGoogleContactRow).filter(c => c);
      showImportPreview(parsed, file.name);
    },
    error: (err) => {
      showImportModal('Error', `<p style="color:#c62828;">CSV parse failed: ${escapeHtml(err.message)}</p>`,
        `<button class="btn-secondary" onclick="document.getElementById('importModal').classList.remove('open')">Close</button>`);
    }
  });

  // reset so same file can be re-picked
  e.target.value = '';
});

/* =====================================================
   EXPORT ALL CONTACTS AS GOOGLE-COMPATIBLE CSV
   ===================================================== */
   /* =====================================================
   EXPORT ALL CONTACTS AS vCard (.vcf) — best for Google Contacts
   ===================================================== */
function exportForGoogleContactsVCF() {
  if (!allContacts || allContacts.length === 0) {
    alert('No contacts to export.');
    return;
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/\\/g, '\\\\')
      .replace(/\n/g, '\\n')
      .replace(/,/g, '\\,')
      .replace(/;/g, '\\;');
  }

  const lines = [];
  const list = allContacts.filter(c => !c.Archived);

  list.forEach(c => {
    const first  = c.FirstName_EN  || '';
    const middle = c.MiddleName_EN || '';
    const last   = c.LastName_EN   || '';
    const prefix = Array.isArray(c.DetectedCodes) ? c.DetectedCodes.join(' ') : '';
    const fn     = [prefix, first, middle, last].filter(Boolean).join(' ').trim() || c.EFN || 'Unnamed';

    lines.push('BEGIN:VCARD');
    lines.push('VERSION:3.0');
    // N: LastName;FirstName;MiddleName;Prefix;Suffix
    lines.push('N:' + esc(last) + ';' + esc(first) + ';' + esc(middle) + ';' + esc(prefix) + ';');
    lines.push('FN:' + esc(fn));

    if (c.ClientType === 'Company' && c.CompanyName_EN) {
      
function exportForGoogleContacts() {
  if (!allContacts || allContacts.length === 0) {
    alert('No contacts to export.');
    return;
  }

  const headers = [
    'Name Prefix',
    'First Name', 'Middle Name', 'Last Name',
    'Nickname',
    'Organization Name',
    'Notes',
    'Phone 1 - Type', 'Phone 1 - Value',
    'Phone 2 - Type', 'Phone 2 - Value',
    'Phone 3 - Type', 'Phone 3 - Value',
    'Phone 4 - Type', 'Phone 4 - Value',
    'Phone 5 - Type', 'Phone 5 - Value',
    'Email 1 - Type', 'Email 1 - Value',
    'Email 2 - Type', 'Email 2 - Value',
    'Email 3 - Type', 'Email 3 - Value',
    'Address 1 - Street',
    'Address 1 - City',
    'Address 1 - Country',
    'Birthday',
    'Event 1 - Label', 'Event 1 - Value'
  ];

  function csv(v) {
    if (v == null) return '';
    const s = String(v).replace(/"/g, '""');
    return /[",\r\n]/.test(s) ? '"' + s + '"' : s;
  }

  const rows = allContacts
    .filter(c => !c.Archived)
    .map(c => {
      const prefix = Array.isArray(c.DetectedCodes) ? c.DetectedCodes.join(' ') : '';
      const notes = 'MGSDB ContactID: ' + (c.ContactID || '') + (c.EmployerNotes ? '\n' + c.EmployerNotes : '');
      return [
        csv(prefix),
        csv(c.FirstName_EN || ''),
        csv(c.MiddleName_EN || ''),
        csv(c.LastName_EN || ''),
        '',
        csv(c.ClientType === 'Company' ? (c.CompanyName_EN || '') : ''),
        csv(notes),
        c.MobileNumber  ? 'Mobile' : '', csv(c.MobileNumber  || ''),
        c.MobileNumber2 ? 'Mobile' : '', csv(c.MobileNumber2 || ''),
        c.MobileNumber3 ? 'Mobile' : '', csv(c.MobileNumber3 || ''),
        c.MobileNumber4 ? 'Mobile' : '', csv(c.MobileNumber4 || ''),
        c.MobileNumber5 ? 'Mobile' : '', csv(c.MobileNumber5 || ''),
        c.Email         ? 'Home'   : '', csv(c.Email         || ''),
        c.EmployerEmail ? 'Work'   : '', csv(c.EmployerEmail || ''),
        c.EmailAddress  ? 'Other'  : '', csv(c.EmailAddress  || ''),
        csv(c.Address   || ''),
        csv(c.City      || ''),
        csv(c.Country   || ''),
        csv(c.BirthDate || ''),
        c.Anniversary   ? 'IKAMA'  : '', csv(c.Anniversary   || '')
      ].join(',');
    });

  const csvBody = [headers.join(',')].concat(rows).join('\r\n');
  const blob = new Blob(['\uFEFF' + csvBody], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = 'address_book_for_google_' + stamp + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);

  alert('✅ Exported ' + rows.length + ' contacts.\n\nImport at contacts.google.com → left sidebar → Import.');
}

/* Parse one Google CSV row into our contact shape */
function parseGoogleContactRow(row) {
  // Google CSV columns vary slightly; we read by exact name with fallbacks
  const get = (k) => (row[k] != null ? String(row[k]).trim() : '');

  const firstName  = get('First Name');
  const middleName = get('Middle Name');
  const lastName   = get('Last Name');
  const namePrefix = get('Name Prefix') || get('Title');

  // Collect up to 10 phone numbers from "Phone N - Value"
  const phones = [];
  for (let i = 1; i <= 10; i++) {
    const v = get('Phone ' + i + ' - Value');
    if (v) {
      // Google sometimes puts multiple numbers separated by ":::" in one field
      v.split(/\s*:::\s*/).forEach(p => { if (p.trim()) phones.push(p.trim()); });
    }
  }

  // Collect up to 5 emails
  const emails = [];
  for (let i = 1; i <= 5; i++) {
    const v = get('E-mail ' + i + ' - Value');
    if (v) v.split(/\s*:::\s*/).forEach(e => { if (e.trim()) emails.push(e.trim()); });
  }

  // Address (use Address 1)
  const address = get('Address 1 - Formatted') || get('Address 1 - Street');
  const city    = get('Address 1 - City');
  const country = get('Address 1 - Country');

  const org     = get('Organization 1 - Name') || get('Organization Name');
  const notes   = get('Notes');
  const birth   = get('Birthday');

  // Build the EFN from name parts
  const efn = [firstName, middleName, lastName].filter(Boolean).join(' ').trim();
  if (!efn && !phones.length && !emails.length) return null; // skip completely empty rows

  // Detect prefix codes in name (e.g. "MGSC LRA - John")
  const rawPrefixSource = `${namePrefix} ${firstName}`.trim();
  const detectedCodes = detectPrefixCodes(rawPrefixSource) || detectPrefixCodes(efn) || [];

  const contact = {
    EFN: efn,
    EmployerName: efn,
    FirstName_EN: firstName,
    MiddleName_EN: middleName,
    LastName_EN: lastName,
    BirthDate: birth || '',
    Address: address,
    City: city,
    Country: country,
    EmployerNotes: notes,
    ClientType: org ? 'Company' : 'Personal',
    CompanyName_EN: org,
    RawPrefix: namePrefix,
    DetectedCodes: detectedCodes,           // array of strings like ["LRA","NCS"]
    agencies: [],                            // will be filled later when agencies are created
    SyncSource: 'GoogleContactsCSV',
    _imported: true
  };

  phones.forEach((p, i) => {
    if (i === 0) contact.MobileNumber = p;
    else contact['MobileNumber' + (i + 1)] = p;
  });
  emails.forEach((em, i) => {
    if (i === 0) contact.Email = em;
    else if (i === 1) contact.EmployerEmail = em;
    else if (i === 2) contact.EmailAddress = em;
  });

  return contact;
}

/* Detect uppercase prefix codes (2-6 letters) at start of name */
function detectPrefixCodes(str) {
  if (!str) return [];
  // Match codes like MGSC, LRA, NCS, MGSS at the start, optionally several separated by spaces, before a hyphen or end
  const m = str.match(/^([A-Z]{2,6}(?:\s+[A-Z]{2,6})*)/);
  if (!m) return [];
  return m[1].split(/\s+/);
}

/* Show import preview */
function showImportPreview(parsed, fileName) {
  const total = parsed.length;
  const withPrefix = parsed.filter(c => (c.DetectedCodes || []).length > 0).length;
  const withPhone = parsed.filter(c => c.MobileNumber).length;
  const withEmail = parsed.filter(c => c.Email).length;

  // existing contacts to compute next ContactID
  const existingMax = allContacts.reduce((max, c) => {
    const n = parseInt(c.ContactID, 10);
    return (Number.isFinite(n) && n > max) ? n : max;
  }, 0);
  const startID = existingMax + 1;
  const endID = existingMax + total;

  // sample preview (first 10)
  const sample = parsed.slice(0, 10).map(c => `
    <div class="sample-row">
      ${(c.DetectedCodes || []).map(code => `<span class="sample-prefix">${escapeHtml(code)}</span>`).join('')}
      <b>${escapeHtml(c.EFN || '(no name)')}</b>
      ${c.MobileNumber ? `<span style="color:#666;"> · ${escapeHtml(c.MobileNumber)}</span>` : ''}
    </div>
  `).join('');

  const body = `
    <div style="margin-bottom: 14px; color:#555;">
      File: <b>${escapeHtml(fileName)}</b>
    </div>

    <div class="stats-grid">
      <div class="stat-box"><div class="stat-num">${total}</div><div class="stat-label">contacts to import</div></div>
      <div class="stat-box"><div class="stat-num">${withPrefix}</div><div class="stat-label">with detected agency prefix</div></div>
      <div class="stat-box"><div class="stat-num">${withPhone}</div><div class="stat-label">with phone</div></div>
      <div class="stat-box"><div class="stat-num">${withEmail}</div><div class="stat-label">with email</div></div>
    </div>

    <div style="margin-bottom:8px; font-size:13px; color:#555;">
      ContactIDs will be assigned: <b>${String(startID).padStart(4,'0')}</b> → <b>${String(endID).padStart(4,'0')}</b>
    </div>

    <div style="font-size:13px; font-weight:500; margin: 10px 0 6px;">First 10 contacts (preview):</div>
    <div class="sample-list">${sample}</div>
  `;

  const footer = `
    <button class="btn-secondary" onclick="document.getElementById('importModal').classList.remove('open')">Cancel</button>
    <button class="btn-primary" id="btnConfirmImport">Import all ${total}</button>
  `;

  showImportModal('Import preview', body, footer);

  document.getElementById('btnConfirmImport').addEventListener('click', () => doImport(parsed, startID));
}

/* Batch-write to Firestore */
async function doImport(parsed, startID) {
  const total = parsed.length;
  const body = `
    <p>Saving <b>${total}</b> contacts to Firebase…</p>
    <div class="progress-bar"><div class="progress-fill" id="impFill" style="width:0%;"></div></div>
    <p id="impStatus" style="font-size:13px; color:#666;">Preparing…</p>
  `;
  showImportModal('Importing…', body, `<button class="btn-secondary" disabled>Please wait…</button>`);

  const { db, collection, addDoc, serverTimestamp } = window._fb;
  const colRef = collection(db, 'address_book');

  let done = 0, errors = 0;
  for (let i = 0; i < parsed.length; i++) {
    const c = parsed[i];
    const id = String(startID + i).padStart(4, '0');
    try {
      await addDoc(colRef, {
        ...c,
        ContactID: id,
        _createdAt: serverTimestamp()
      });
      done++;
    } catch (e) {
      errors++;
      console.warn('Import row failed:', e);
    }
    if (i % 20 === 0 || i === parsed.length - 1) {
      const pct = Math.round(((i + 1) / parsed.length) * 100);
      document.getElementById('impFill').style.width = pct + '%';
      document.getElementById('impStatus').textContent =
        `Saved ${done} of ${total}${errors ? ` · ${errors} errors` : ''}`;
    }
  }

  // Done
  showImportModal(
    'Import complete',
    `<p style="color:#155724;"><b>✓ Imported ${done} contacts.</b></p>
     ${errors ? `<p style="color:#c62828;">${errors} rows failed (see console).</p>` : ''}`,
    `<button class="btn-primary" onclick="document.getElementById('importModal').classList.remove('open')">Done</button>`
  );
}

function showImportModal(title, bodyHtml, footerHtml) {
  document.getElementById('importModalTitle').textContent = title;
  importModalBody.innerHTML = bodyHtml;
  importModalFooter.innerHTML = footerHtml;
  importModal.classList.add('open');
}
</script>

</body>
</html>
