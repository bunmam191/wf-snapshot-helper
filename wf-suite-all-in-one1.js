// ==UserScript==
// @name         WF – Lead + Snapshot (Ire per URL) + Dimensions (Backend) + Save (All-in-one) - CHUẨN 100%
// @namespace    wf-suite-all-in-one
// @version      1.2.3-bookmark
// @description  Tag LE/NLE dưới thumbnail, bulk LE/NLE, bulk Dimensions TRUE/FALSE (có Save Tags), snapshot SKU+MFG+IreID+Dimensions (theo từng URL, chỉ Ire ID), Agent Notes "Has Dimensions" (snapshot + backend), và nút Save Tags tiện lợi.
// @match        https://admin.wayfair.com/d/product-merchandising/*
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// ==/UserScript==

(function () {
  'use strict';

  /* =========================================================
   *  POLYFILL CHO MÔI TRƯỜNG KHÔNG CÓ TAMPERMONKEY (BOOKMARKLET)
   * ======================================================= */

  // Polyfill GM_getValue / GM_setValue bằng localStorage
  if (typeof GM_getValue === 'undefined') {
    window.GM_getValue = function (key, defaultValue) {
      try {
        const raw = localStorage.getItem(key);
        if (raw === null || raw === undefined) return defaultValue;
        return JSON.parse(raw);
      } catch (e) {
        return defaultValue;
      }
    };

    window.GM_setValue = function (key, value) {
      try {
        localStorage.setItem(key, JSON.stringify(value));
      } catch (e) {
        // ignore nếu đầy quota
      }
    };
  }

  // Polyfill unsafeWindow cho môi trường bookmarklet
  if (typeof unsafeWindow === 'undefined') {
    window.unsafeWindow = window;
  }

  /* =========================================================
   *  BACKEND DIMENSIONS MAP (từ GetProductMediaAssetMetadata)
   * ======================================================= */

  // uuidHasDimensions: asset UUID → boolean
  // uuidToIre: asset UUID → Ire ID
  // ireHasDimensions: Ire ID (string) → boolean
  const uuidHasDimensions = {};
  const uuidToIre = {};
  const ireHasDimensions = {};

  function markIreFromUuid(uuid) {
    const hasDim = uuidHasDimensions[uuid];
    const ireId = uuidToIre[uuid];
    if (ireId && typeof hasDim === 'boolean') {
      ireHasDimensions[String(ireId)] = hasDim;
    }
  }

  // Patch fetch để đọc HAS_DIMENSIONS & PREVIEW_IRE_ID từ backend
  (function patchFetch() {
    try {
      const w = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
      if (w.__wfSnapshotFetchPatched) return;
      if (typeof w.fetch !== 'function') return;
      w.__wfSnapshotFetchPatched = true;

      const originalFetch = w.fetch;

      w.fetch = function (input, init) {
        const fetchPromise = originalFetch.apply(this, arguments);

        try {
          const url = typeof input === 'string' ? input : input && input.url;
          if (!url) return fetchPromise;

          const isGraphQL =
            url.includes('/federation/internal/graphql') ||
            url.includes('GetProductMediaAssetMetadata');
          const isAssetDetails =
            url.includes('/d/product-merchandising/api/getAssetDetails') ||
            url.includes('/api/getAssetDetails');

          if (!isGraphQL && !isAssetDetails) {
            return fetchPromise;
          }

          return fetchPromise.then(async (response) => {
            try {
              const cloned = response.clone();
              const contentType =
                cloned.headers && cloned.headers.get('content-type');
              if (!contentType || !contentType.includes('application/json')) {
                return response;
              }

              const data = await cloned.json();

              // 1) getAssetDetails: map UUID -> Ire ID (chỉ type_id = 1 = image)
              if (isAssetDetails && data && Array.isArray(data.asset_details)) {
                for (const asset of data.asset_details) {
                  if (!asset) continue;
                  if (asset.type_id !== 1) continue;
                  const uuid = String(asset.id || '').trim();
                  const ireId = String(asset.image_resource_id || '').trim();
                  if (!uuid || !ireId) continue;

                  uuidToIre[uuid] = ireId;
                  markIreFromUuid(uuid);
                }
              }

              // 2) GraphQL: đọc HAS_DIMENSIONS & PREVIEW_IRE_ID
              if (
                isGraphQL &&
                data &&
                data.data &&
                data.data.getProductMediaAssetMetadata
              ) {
                const payload = data.data.getProductMediaAssetMetadata;
                if (payload && Array.isArray(payload.assets)) {
                  for (const asset of payload.assets) {
                    if (!asset) continue;
                    const uuid = String(asset.assetId || '').trim();
                    if (!uuid) continue;

                    let hasDimVal = null;
                    let previewIreId = null;

                    if (Array.isArray(asset.metadata)) {
                      for (const meta of asset.metadata) {
                        if (!meta) continue;
                        const field = meta.metadataFieldName;
                        const rawVal = (meta.value && meta.value[0]) || '';

                        if (field === 'HAS_DIMENSIONS') {
                          const low = String(rawVal).toLowerCase();
                          if (low === 'true') hasDimVal = true;
                          else if (low === 'false') hasDimVal = false;
                        } else if (field === 'PREVIEW_IRE_ID') {
                          previewIreId = String(rawVal || '').trim();
                        }
                      }
                    }

                    if (hasDimVal !== null) {
                      uuidHasDimensions[uuid] = hasDimVal;

                      if (previewIreId) {
                        ireHasDimensions[previewIreId] = hasDimVal;
                      }

                      markIreFromUuid(uuid);
                    }
                  }
                }
              }
            } catch (e) {
              // im lặng, không phá trang
            }

            return response;
          });
        } catch (e) {
          return fetchPromise;
        }
      };
    } catch (e) {
      // nếu patch fail thì thôi
    }
  })();

  /* =========================================================
   *  COMMON HELPERS
   * ======================================================= */

  const GROUP_CFG = {
    sectionSelector: '[data-test-id^="ocid-tile-"]'
  };

  // 🔑 Storage key theo từng trang (mỗi URL = 1 snapshot riêng)
  function getStorageKey() {
    let raw = location.pathname + location.search;
    raw = raw.replace(/[^\w]+/g, '_');
    return 'wf_snapshot_v2_' + raw;
  }

  function isPreviewOpen() {
    return !!document.querySelector(
      'button[data-test-id="asset-mark-lead-eligible-button"], button[data-test-id="asset-use-as-lead-button"]'
    );
  }

  function isReviewPage() {
    return Array.from(document.querySelectorAll('h2'))
      .some(h => h.textContent.includes('Review & Submit'));
  }

  function updatePageState() {
    const preview = isPreviewOpen();
    const review  = isReviewPage();
    document.body.classList.toggle('wf-preview-open', preview);
    document.body.classList.toggle('wf-review-submit', review);
  }

  function isInsidePreview(el) {
    return !!el.closest('[role="dialog"], [data-hb-id="Modal"]');
  }

  function isInsideReview(el) {
    return !!el.closest('[data-hb-id="Table"]') ||
           document.body.classList.contains('wf-review-submit');
  }

  function sleep(ms) {
    return new Promise(res => setTimeout(res, ms));
  }

  function waitForElement(selector, timeout = 4000) {
    return new Promise((resolve, reject) => {
      const existing = document.querySelector(selector);
      if (existing) return resolve(existing);

      const obs = new MutationObserver(() => {
        const found = document.querySelector(selector);
        if (found) {
          obs.disconnect();
          resolve(found);
        }
      });

      obs.observe(document.body, { childList: true, subtree: true });

      setTimeout(() => {
        obs.disconnect();
        reject(new Error('Timeout waiting for ' + selector));
      }, timeout);
    });
  }

  function getPreviewRoot() {
    return document.querySelector('[role="dialog"], [data-hb-id="Modal"]') || document.body;
  }

  /* =========================================================
   *  CSS
   * ======================================================= */

  function injectCSS() {
    if (document.getElementById('wf-lead-style')) return;

    const s = document.createElement('style');
    s.id = 'wf-lead-style';
    s.textContent = `
      .wf-thumb-wrap { position: relative; }

      .wf-lead-panel {
        margin-top: 4px;
        display: flex;
        justify-content: center;
        gap: 6px;
      }

      .wf-lead-panel button {
        all: unset;
        cursor: pointer;
        font-size: 11px;
        padding: 3px 9px;
        border-radius: 6px;
        border: 1px solid #d0d0d0;
        background: #ffffff;
        color: #333;
        font-weight: 500;
        white-space: nowrap;
      }

      .wf-lead-panel button[data-act="mark-lead"] {
        border-color: #7b189f;
        color: #7b189f;
      }

      .wf-lead-panel button[data-act="unmark-lead"] {
        border-color: #d32f2f;
        color: #d32f2f;
      }

      .wf-lead-panel button:hover {
        background: #f5f5f5;
      }

      .wf-thumb-check {
        position: absolute;
        bottom: 8px;
        right: 8px;
        width: 24px;
        height: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 0;
        border-radius: 4px;
        z-index: 5;
        opacity: 1;
        pointer-events: auto;
      }

      .wf-thumb-check input[type="checkbox"] {
        appearance: none;
        -webkit-appearance: none;
        width: 20px;
        height: 20px;
        margin: 0;
        border-radius: 4px;
        background: #fff;
        border: 2px solid #66256a;
        box-sizing: border-box;
        position: relative;
        cursor: pointer;
      }

      .wf-thumb-check input[type="checkbox"]::after {
        content: "";
        position: absolute;
        left: 7px;
        top: 4px;
        width: 5px;
        height: 10px;
        border-right: 2px solid transparent;
        border-bottom: 2px solid transparent;
        transform: rotate(45deg);
      }

      .wf-thumb-check input[type="checkbox"]:checked {
        background: #66256a;
        border-color: #66256a;
        box-shadow: 0 0 0 2px #fff;
      }
      .wf-thumb-check input[type="checkbox"]:checked::after {
        border-right-color: #fff;
        border-bottom-color: #fff;
      }

      .wf-thumb-selected { outline: none !important; }

      .wf-preview-open .wf-lead-panel,
      .wf-preview-open .wf-thumb-check,
      .wf-preview-open .wf-group-toolbar,
      .wf-review-submit .wf-lead-panel,
      .wf-review-submit .wf-thumb-check,
      .wf-review-submit #wf-lead-toolbar,
      .wf-review-submit .wf-group-toolbar {
        display: none !important;
      }

      .wf-hide-preview [role="dialog"],
      .wf-hide-preview [data-hb-id="Modal"] {
        opacity: 0 !important;
        pointer-events: none !important;
      }

      #wf-lead-toolbar {
        display: flex;
        align-items: center;
        gap: 6px;
        margin-left: 16px;
        font-size: 11px;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background: rgba(0,0,0,0.03);
        border-radius: 999px;
        padding: 3px 10px;
      }
      #wf-lead-toolbar button {
        all: unset;
        cursor: pointer;
        background: #7b189f;
        color: #fff;
        padding: 3px 10px;
        border-radius: 999px;
        font-size: 14px;
        font-weight: 600;
      }
      #wf-lead-toolbar button[data-bulk="stop"] { background: #999; }
      #wf-lead-toolbar button[data-bulk="clear"] { background: #d32f2f; }
      #wf-lead-toolbar button:hover { filter: brightness(1.1); }
      #wf-lead-toolbar .wf-counter {
        font-size: 11px;
        color: #444;
        margin-left: 4px;
      }
      #wf-lead-toolbar .wf-progress {
        font-size: 11px;
        color: #777;
        margin-left: 8px;
        min-width: 90px;
      }

      .wf-group-toolbar {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        font-size: 13px;
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background: rgba(102,37,106,0.06);
        border-radius: 999px;
        padding: 4px 12px;
        margin: 6px 0 8px;
      }
      .wf-group-toolbar label {
        display: inline-flex;
        align-items: center;
        gap: 6px;
        cursor: pointer;
        font-weight: 600;
      }
      .wf-group-toolbar input[type="checkbox"] {
        width: 18px;
        height: 18px;
        margin: 0;
      }
      .wf-group-toolbar button {
        all: unset;
        cursor: pointer;
        font-size: 12px;
        padding: 3px 10px;
        border-radius: 999px;
        background: #eee;
      }
      .wf-group-toolbar button:hover {
        filter: brightness(1.05);
      }

      #wf-pmp-save-button {
        position: fixed;
        top: 50%;
        right: 40px;
        transform: translateY(-50%);
        padding: 20px 40px;
        background-color: #28a745;
        color: #fff;
        border: none;
        border-radius: 5px;
        cursor: pointer;
        z-index: 9999;
      }
    `;
    document.head.appendChild(s);
  }

  /* =========================================================
   *  THUMBNAIL PANELS (LE/NLE + CHECKBOX)
   * ======================================================= */

  function addPanels() {
    injectCSS();
    updatePageState();
    if (document.body.classList.contains('wf-review-submit')) return;

    const imgs = document.querySelectorAll('img[data-test-id^="img-fluid-"]');
    imgs.forEach(img => {
      if (isInsidePreview(img) || isInsideReview(img)) return;

      const wrap = img.closest('div');
      if (!wrap) return;
      if (wrap.dataset.wfThumbWrap === '1') return;

      wrap.dataset.wfThumbWrap = '1';
      wrap.classList.add('wf-thumb-wrap');

      const panel = document.createElement('div');
      panel.className = 'wf-lead-panel';
      panel.innerHTML = `
        <button data-act="mark-lead" title="Mark as Lead Eligible">Tag LE</button>
        <button data-act="unmark-lead" title="Mark as Not Lead Eligible">Tag NLE</button>
      `;
      panel.addEventListener('click', e => {
        const btn = e.target.closest('button');
        if (!btn) return;
        e.stopPropagation();
        e.preventDefault();
        handleAction(btn.dataset.act, img);
      });

      const label = document.createElement('label');
      label.className = 'wf-thumb-check';
      label.innerHTML = `<input type="checkbox" data-wf-select>`;
      const chk = label.querySelector('input');

      label.addEventListener('click', e => e.stopPropagation());

      chk.addEventListener('change', () => {
        if (chk.checked) wrap.classList.add('wf-thumb-selected');
        else wrap.classList.remove('wf-thumb-selected');
        updateSelectedCounter();
      });

      wrap.appendChild(label);

      const card = img.closest('[data-test-id="asset-card-container"]');
      if (card && card.parentElement) {
        const outer = card.parentElement;
        const afterCard = card.nextSibling;
        if (afterCard) outer.insertBefore(panel, afterCard);
        else outer.appendChild(panel);
      } else {
        const parent = wrap.parentElement || wrap;
        const afterWrap = wrap.nextSibling;
        if (afterWrap) parent.insertBefore(panel, afterWrap);
        else parent.appendChild(panel);
      }
    });

    updateSelectedCounter();
  }

  function getSelectedImgs() {
    const inputs = document.querySelectorAll('input[data-wf-select]:checked');
    const list = [];
    inputs.forEach(inp => {
      const wrap = inp.closest('.wf-thumb-wrap');
      if (!wrap) return;
      const img = wrap.querySelector('img[data-test-id^="img-fluid-"]');
      if (!img) return;
      if (isInsidePreview(img) || isInsideReview(img)) return;
      list.push(img);
    });
    return list;
  }

  function clearAllSelections() {
    document.querySelectorAll('.wf-thumb-wrap input[type="checkbox"][data-wf-select]')
      .forEach(cb => {
        cb.checked = false;
        cb.dispatchEvent(new Event('change', { bubbles: true }));
      });

    document.querySelectorAll('.wf-group-toolbar input[type="checkbox"][data-wf-group-master]')
      .forEach(master => {
        master.checked = false;
        master.dispatchEvent(new Event('change', { bubbles: true }));
      });

    updateSelectedCounter();
  }

  function updateSelectedCounter() {
    const el = document.getElementById('wf-count-selected');
    if (!el) return;
    el.textContent = 'Select: ' + getSelectedImgs().length;
  }

  /* =========================================================
   *  PER-SECTION TOOLBARS
   * ======================================================= */

  function ensureSectionControls() {
    const sel = GROUP_CFG.sectionSelector;
    if (!sel) return;

    const sections = document.querySelectorAll(sel);
    sections.forEach(sec => {
      if (!sec || sec.dataset.wfGroupInit === '1') return;
      if (!sec.querySelector('input[data-wf-select]')) return;

      sec.dataset.wfGroupInit = '1';

      const bar = document.createElement('div');
      bar.className = 'wf-group-toolbar';
      bar.innerHTML = `
        <label title="Select / unselect all thumbnails in this section">
          <input type="checkbox" data-wf-group-master>
          <span>Select all in section</span>
        </label>
        <button type="button" data-wf-group-clear>Clear</button>
      `;

      bar.addEventListener('change', e => {
        const master = e.target.closest('input[data-wf-group-master]');
        if (!master) return;
        e.stopPropagation();

        const checked = master.checked;
        const checkboxes = sec.querySelectorAll('input[data-wf-select]');
        checkboxes.forEach(chk => {
          chk.checked = checked;
          const wrap = chk.closest('.wf-thumb-wrap');
          if (!wrap) return;
          if (checked) wrap.classList.add('wf-thumb-selected');
          else wrap.classList.remove('wf-thumb-selected');
        });
        updateSelectedCounter();
      });

      bar.addEventListener('click', e => {
        const btn = e.target.closest('button[data-wf-group-clear]');
        if (!btn) return;
        e.preventDefault();
        e.stopPropagation();

        const master = bar.querySelector('input[data-wf-group-master]');
        if (master) master.checked = false;

        const checkboxes = sec.querySelectorAll('input[data-wf-select]');
        checkboxes.forEach(chk => {
          chk.checked = false;
          const wrap = chk.closest('.wf-thumb-wrap');
          if (wrap) wrap.classList.remove('wf-thumb-selected');
        });
        updateSelectedCounter();
      });

      sec.insertBefore(bar, sec.firstChild);
    });
  }

  function ensureListingSectionControls() {
    const headers = document.querySelectorAll('#association-header');
    headers.forEach(header => {
      const tile = header.closest('[data-test-id^="entity-tile-"]');
      if (!tile) return;
      if (tile.dataset.wfListingGroupInit === '1') return;
      if (!tile.querySelector('input[data-wf-select]')) return;

      tile.dataset.wfListingGroupInit = '1';

      const bar = document.createElement('div');
      bar.className = 'wf-group-toolbar';
      bar.innerHTML = `
        <label title="Select / unselect all thumbnails in this section">
          <input type="checkbox" data-wf-group-master>
          <span>Select all in section</span>
        </label>
          <button type="button" data-wf-group-clear>Clear</button>
      `;

      bar.addEventListener('change', e => {
        const master = e.target.closest('input[data-wf-group-master]');
        if (!master) return;
        e.stopPropagation();

        const checked = master.checked;
        const checkboxes = tile.querySelectorAll('input[data-wf-select]');
        checkboxes.forEach(chk => {
          chk.checked = checked;
          const wrap = chk.closest('.wf-thumb-wrap');
          if (!wrap) return;
          if (checked) wrap.classList.add('wf-thumb-selected');
          else wrap.classList.remove('wf-thumb-selected');
        });
        updateSelectedCounter();
      });

      bar.addEventListener('click', e => {
        const btn = e.target.closest('button[data-wf-group-clear]');
        if (!btn) return;
        e.preventDefault();
        e.stopPropagation();

        const master = bar.querySelector('input[data-wf-group-master]');
        if (master) master.checked = false;

        const checkboxes = tile.querySelectorAll('input[data-wf-select]');
        checkboxes.forEach(chk => {
          chk.checked = false;
          const wrap = chk.closest('.wf-thumb-wrap');
          if (wrap) wrap.classList.remove('wf-thumb-selected');
        });
        updateSelectedCounter();
      });

      const headerBox = header.closest('[data-hb-id="BoxV3"]');
      const parent = headerBox ? headerBox.parentElement : tile;
      parent.insertBefore(bar, headerBox || parent.firstChild);
    });
  }

  /* =========================================================
   *  LE/NLE HELPER
   * ======================================================= */

  async function clickMarkNotLeadEligibleInPreview(previewRoot, maxWaitMs = 2500) {
    const start = Date.now();
    while (Date.now() - start < maxWaitMs) {
      const btns = Array.from(previewRoot.querySelectorAll('button'));
      const btn = btns.find(b =>
        /Mark as Not Lead Eligible/i.test(b.innerText || b.textContent || '')
      );
      if (btn) {
        btn.click();
        return true;
      }
      await sleep(200);
    }
    return false;
  }

  async function handleAction(action, img, fromBulk = false) {
    const container = img.closest('[data-test-id="asset-card-container"]') ||
                      img.closest('.wf-thumb-wrap') ||
                      img.parentElement;

    const txt = (container && container.textContent || '').toLowerCase();
    const hasLeadEligible = txt.includes('lead eligible');
    const hasLeadOverride = txt.includes('lead override');

    if (action === 'unmark-lead') {
      if (!hasLeadEligible && !hasLeadOverride) return false;
    }

    if (action === 'mark-lead') {
      if (hasLeadOverride || hasLeadEligible) return false;
    }

    const manual = isPreviewOpen();
    let performed = false;

    try {
      if (!manual) document.body.classList.add('wf-hide-preview');

      if (action === 'mark-lead') {
        img.click();
        await waitForElement(
          'button[data-test-id="asset-mark-lead-eligible-button"], button[data-test-id="asset-use-as-lead-button"]'
        );
        updatePageState();

        const previewRoot = getPreviewRoot();
        const allBtns = Array.from(previewRoot.querySelectorAll('button'));
        const markBtnLead = allBtns.find(b =>
          /Mark as Lead Eligible/i.test(b.innerText || b.textContent || '')
        );
        const btnUseAsLead = allBtns.find(b =>
          /Use as Lead/i.test(b.innerText || b.textContent || '')
        );

        if (markBtnLead) {
          markBtnLead.click();
          performed = true;
        } else if (btnUseAsLead) {
          btnUseAsLead.click();
          performed = true;
        }

      } else if (action === 'unmark-lead') {
        if (hasLeadOverride) {
          img.click();
          await waitForElement('button[data-test-id="asset-use-as-lead-button"]');
          updatePageState();

          let previewRoot = getPreviewRoot();
          let btns1 = Array.from(previewRoot.querySelectorAll('button'));
          let btnDoNotUse = btns1.find(b =>
            /Do Not Use as Lead/i.test(b.innerText || b.textContent || '')
          );

          if (btnDoNotUse) {
            btnDoNotUse.click();
            await sleep(500);
          } else {
            console.warn('WF Lead Helper: không tìm thấy "Do Not Use as Lead" trong preview');
          }

          img.click();
          await waitForElement(
            'button[data-test-id="asset-mark-lead-eligible-button"], button[data-test-id="asset-use-as-lead-button"]'
          );
          updatePageState();

          previewRoot = getPreviewRoot();
          const ok = await clickMarkNotLeadEligibleInPreview(previewRoot);
          performed = ok;

        } else if (hasLeadEligible) {
          img.click();
          await waitForElement(
            'button[data-test-id="asset-mark-lead-eligible-button"], button[data-test-id="asset-use-as-lead-button"]'
          );
          updatePageState();

          const previewRoot = getPreviewRoot();
          const ok = await clickMarkNotLeadEligibleInPreview(previewRoot);
          performed = ok;
        }
      }

      setTimeout(() => {
        if (isPreviewOpen()) {
          document.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true
          }));
        }
      }, fromBulk ? 250 : 300);

    } catch (e) {
      console.error('WF Lead Helper error', e);
    } finally {
      setTimeout(() => {
        updatePageState();
        if (!manual) document.body.classList.remove('wf-hide-preview');
      }, fromBulk ? 450 : 650);
    }

    return performed;
  }

  /* =========================================================
   *  SAVE TAGS HELPERS
   * ======================================================= */

  function waitForEnabledSaveTags(timeout = 2000) {
    return new Promise((resolve, reject) => {
      const interval = 200;
      let elapsed = 0;
      const check = setInterval(() => {
        const buttons = Array.from(document.querySelectorAll('button'));
        const saveBtn = buttons.find(btn => {
          return (btn.innerText || '').trim() === 'Save Tags' && !btn.disabled;
        });
        if (saveBtn) {
          clearInterval(check);
          resolve(saveBtn);
        } else if ((elapsed += interval) >= timeout) {
          clearInterval(check);
          reject(new Error('Không tìm thấy nút Save Tags đang bật'));
        }
      }, interval);
    });
  }

  function waitForPopupSave(timeout = 2000) {
    return new Promise((resolve, reject) => {
      const interval = 200;
      let elapsed = 0;
      const check = setInterval(() => {
        const saveBtns = Array.from(document.querySelectorAll('button'));
        const validSave = saveBtns.find(btn => {
          const span = btn.querySelector('span');
          const visible = btn.offsetParent !== null;
          return span && (span.textContent || '').trim() === 'Save' && !btn.disabled && visible;
        });
        if (validSave) {
          clearInterval(check);
          resolve(validSave);
        } else if ((elapsed += interval) >= timeout) {
          clearInterval(check);
          reject(new Error('Không tìm thấy nút Save trong popup'));
        }
      }, interval);
    });
  }

  function waitForElementSimple(selector, timeout = 2000) {
    return new Promise((resolve, reject) => {
      const interval = 200;
      let elapsed = 0;
      const timer = setInterval(() => {
        const el = document.querySelector(selector);
        if (el) {
          clearInterval(timer);
          resolve(el);
        } else if ((elapsed += interval) >= timeout) {
          clearInterval(timer);
          reject(new Error('Không tìm thấy phần tử: ' + selector));
        }
      }, interval);
    });
  }

  function createSaveButtonPMP() {
    if (document.getElementById('wf-pmp-save-button')) return;

    const btn = document.createElement('button');
    btn.id = 'wf-pmp-save-button';
    btn.innerText = 'Save';
    document.body.appendChild(btn);

    btn.addEventListener('click', async () => {
      try {
        const saveTagsBtn = await waitForEnabledSaveTags();
        saveTagsBtn.click();

        const popupSaveBtn = await waitForPopupSave();
        await sleep(50);
        popupSaveBtn.click();

        const closeBtn = await waitForElementSimple('button[aria-label="Close"]');
        await sleep(50);
        closeBtn.click();
      } catch (err) {
        return;
      }
    });
  }

  /* =========================================================
   *  DIMENSIONS HELPER (TRUE / FALSE + SAVE TAGS)
   * ======================================================= */

  function readHasDimensionsFromPreview(previewRoot) {
    const wrapper = previewRoot.querySelector('div[data-test-id="dropdown-tag-HAS_DIMENSIONS"]');
    if (!wrapper) return null;

    const valEl = wrapper.querySelector('[data-test-id="dropdown-tag-HAS_DIMENSIONS-singleLineMultiValues"]');
    if (valEl) {
      const txt = (valEl.textContent || '').trim().toLowerCase();
      if (txt === 'true') return true;
      if (txt === 'false') return false;
    }

    const input = wrapper.querySelector('input[data-test-id="dropdown-tag-HAS_DIMENSIONS-input"]');
    if (input && input.value) {
      const v = input.value.trim().toLowerCase();
      if (v === 'true') return true;
      if (v === 'false') return false;
    }

    const menu = previewRoot.querySelector('ul[data-test-id="dropdown-tag-HAS_DIMENSIONS-menu"]');
    if (menu) {
      const selected = menu.querySelector('li[aria-selected="true"]') || menu.querySelector('li');
      if (selected) {
        const t = (selected.textContent || '').trim().toLowerCase();
        if (t === 'true') return true;
        if (t === 'false') return false;
      }
    }

    return null;
  }

  function readAssetIdFromPreview(previewRoot) {
    const p = previewRoot.querySelector('p[data-test-id="legacy-asset-id-text"]');
    if (!p) return null;
    const m = (p.textContent || '').match(/Ire ID:\s*(\d+)/i);
    return m ? m[1] : null;
  }

  async function setHasDimensions(previewRoot, valueTrue) {
    const wrapper = previewRoot.querySelector('div[data-test-id="dropdown-tag-HAS_DIMENSIONS"]');
    if (!wrapper) {
      console.warn('WF Dim Helper: không tìm thấy dropdown HAS_DIMENSIONS');
      return false;
    }

    try {
      const combo = wrapper.querySelector('[role="combobox"]');
      const input = wrapper.querySelector('input[data-test-id="dropdown-tag-HAS_DIMENSIONS-input"]');

      if (combo) {
        combo.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        combo.dispatchEvent(new MouseEvent('mouseup',   { bubbles: true }));
        combo.click();
      } else if (input) {
        input.focus();
        input.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        input.dispatchEvent(new MouseEvent('mouseup',   { bubbles: true }));
        input.click();
      } else {
        wrapper.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        wrapper.dispatchEvent(new MouseEvent('mouseup',   { bubbles: true }));
        wrapper.click();
      }

      if (input) {
        const ev = new KeyboardEvent('keydown', {
          key: 'ArrowDown',
          code: 'ArrowDown',
          keyCode: 40,
          which: 40,
          bubbles: true
        });
        input.dispatchEvent(ev);
      }

      let menu;
      try {
        menu = await waitForElement('ul[data-test-id="dropdown-tag-HAS_DIMENSIONS-menu"]', 2000);
      } catch (e) {
        console.warn('WF Dim Helper: không tìm thấy menu HAS_DIMENSIONS sau khi click combobox');
        return false;
      }

      const items = Array.from(menu.querySelectorAll('li'));
      if (!items.length) {
        console.warn('WF Dim Helper: menu HAS_DIMENSIONS không có item');
        return false;
      }

      const desiredText = valueTrue ? 'true' : 'false';
      let targetItem = items.find(li =>
        (li.textContent || '').trim().toLowerCase() === desiredText
      );

      if (!targetItem) {
        targetItem = valueTrue
          ? menu.querySelector('[data-test-id="dropdown-tag-HAS_DIMENSIONS-menu-item-0"]')
          : menu.querySelector('[data-test-id="dropdown-tag-HAS_DIMENSIONS-menu-item-1"]');
      }

      if (!targetItem) {
        console.warn('WF Dim Helper: không tìm được item True/False trong menu');
        return false;
      }

      targetItem.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
      targetItem.dispatchEvent(new MouseEvent('mouseup',   { bubbles: true }));
      targetItem.click();

      const start = Date.now();
      while (Date.now() - start < 1500) {
        const v = readHasDimensionsFromPreview(previewRoot);
        if (v === valueTrue) return true;
        await sleep(120);
      }

    } catch (e) {
      console.error('WF Dim Helper: lỗi khi set HAS_DIMENSIONS', e);
      return false;
    }

    return false;
  }

  // ✅ luôn đóng preview & chờ đóng xong trước khi qua asset tiếp theo
  async function handleDimensionAction(img, valueTrue, fromBulk = false) {
    let performed = false;

    try {
      img.click();
      await waitForElement(
        'div[data-test-id="dropdown-tag-HAS_DIMENSIONS"], button[data-test-id="asset-mark-lead-eligible-button"], button[data-test-id="asset-use-as-lead-button"]'
      );
      updatePageState();

      const previewRoot = getPreviewRoot();
      const currentDim = readHasDimensionsFromPreview(previewRoot);

      if (currentDim === valueTrue) {
        console.log('WF Dim Helper: HAS_DIMENSIONS đã đúng, skip asset này.');

        document.dispatchEvent(new KeyboardEvent('keydown', {
          key: 'Escape',
          code: 'Escape',
          keyCode: 27,
          which: 27,
          bubbles: true
        }));

        const startSkip = Date.now();
        while (isPreviewOpen() && Date.now() - startSkip < 1500) {
          await sleep(100);
        }

        return false;
      }

      const okDim = await setHasDimensions(previewRoot, valueTrue);
      if (!okDim) {
        console.warn('WF Dim Helper: setHasDimensions thất bại');
      } else {
        try {
          const saveTagsBtn = await waitForEnabledSaveTags(3000);
          saveTagsBtn.click();

          const popupSaveBtn = await waitForPopupSave(3000);
          await sleep(80);
          popupSaveBtn.click();

          try {
            const closeBtn = await waitForElementSimple('button[aria-label="Close"]', 3000);
            await sleep(80);
            closeBtn.click();
          } catch (e) {
            // không có Close cũng kệ
          }
        } catch (e) {
          console.warn('WF Dim Helper: Save Tags flow cho Dimensions thất bại', e);
        }
        performed = true;
      }

      if (isPreviewOpen()) {
        document.dispatchEvent(new KeyboardEvent('keydown', {
          key: 'Escape',
          code: 'Escape',
          keyCode: 27,
          which: 27,
          bubbles: true
        }));

        const startClose = Date.now();
        while (isPreviewOpen() && Date.now() - startClose < 1500) {
          await sleep(100);
        }
      }

    } catch (e) {
      console.error('WF Dim Helper error', e);
    } finally {
      setTimeout(() => {
        updatePageState();
      }, fromBulk ? 450 : 650);
    }

    return performed;
  }

  /* =========================================================
   *  BULK TOOLBAR (LE / NLE / DIM TRUE / DIM FALSE)
   * ======================================================= */

  let bulkRunning = false;
  let bulkCancelled = false;

  function ensureToolbar() {
    if (document.body.classList.contains('wf-review-submit')) return;
    if (document.getElementById('wf-lead-toolbar')) return;

    const logoBtn = document.querySelector('button[data-test-id="adminHomeLogo"]');
    if (!logoBtn) return;
    const headerBox = logoBtn.closest('div[data-hb-id="BoxV3"]') || logoBtn.parentElement;
    if (!headerBox) return;

    const bar = document.createElement('div');
    bar.id = 'wf-lead-toolbar';
    bar.innerHTML = `
      <button data-bulk="le">Tag LE</button>
      <button data-bulk="nle">Tag NLE</button>
      <button data-bulk="dim-true">Dim TRUE</button>
      <button data-bulk="dim-false">Dim FALSE</button>
      <button data-bulk="clear">Clear</button>
      <button data-bulk="stop">Stop</button>
      <span class="wf-counter" id="wf-count-selected">Select: 0</span>
      <span class="wf-progress" id="wf-progress-info"></span>
    `;

    bar.addEventListener('click', async e => {
      const btn = e.target.closest('button[data-bulk]');
      if (!btn) return;
      e.stopPropagation();
      e.preventDefault();

      const type = btn.dataset.bulk;

      if (type === 'clear') {
        clearAllSelections();
        const pi = document.getElementById('wf-progress-info');
        if (pi) pi.textContent = '';
        return;
      }

      if (type === 'stop') {
        bulkCancelled = true;
        return;
      }

      if (!['le', 'nle', 'dim-true', 'dim-false'].includes(type)) return;
      if (bulkRunning) return;

      const imgs = getSelectedImgs();
      if (!imgs.length) {
        updateSelectedCounter();
        return;
      }

      const progressEl = document.getElementById('wf-progress-info');
      const total = imgs.length;
      if (progressEl) progressEl.textContent = `Running: 0 / ${total}`;

      bulkRunning = true;
      bulkCancelled = false;

      try {
        for (let i = 0; i < imgs.length; i++) {
          if (bulkCancelled) break;
          const img = imgs[i];

          if (progressEl) {
            progressEl.textContent = `Running: ${i + 1} / ${total}`;
          }

          if (type === 'le' || type === 'nle') {
            await handleAction(
              type === 'le' ? 'mark-lead' : 'unmark-lead',
              img,
              true
            );
            await sleep(500);
          } else if (type === 'dim-true' || type === 'dim-false') {
            await handleDimensionAction(
              img,
              type === 'dim-true',
              true
            );
            await sleep(500);
          }
        }
      } finally {
        bulkRunning = false;
        bulkCancelled = false;
        const pi = document.getElementById('wf-progress-info');
        if (pi) pi.textContent = '';
        clearAllSelections();
      }
    });

    headerBox.appendChild(bar);
    updateSelectedCounter();
  }

  /* =========================================================
   *  SNAPSHOT HELPER (per-URL, Ire ID only, backend dim)
   * ======================================================= */

  function loadSnapshot() {
    return GM_getValue(getStorageKey(), []); // array of row objects
  }

  function saveSnapshot(rows) {
    GM_setValue(getStorageKey(), rows);
  }

  function findRow(rows, sku, mfg, assetId) {
    return rows.find(
      (r) =>
        r.sku === sku &&
        r.manufacturerPartId === mfg &&
        r.assetId === assetId
    );
  }

  function getTypeFromCard(card) {
    const knownTypes = [
      'Image',
      'Non Photo',
      'Environmental',
      'Silo',
      'Detail Shot',
      'Lifestyle',
      'Floor & Wall'
    ];

    const ps = Array.from(card.querySelectorAll('p'));
    for (const p of ps) {
      const text = (p.textContent || '').trim();
      if (knownTypes.includes(text)) {
        return text;
      }
    }

    const allText = ps
      .map(p => (p.textContent || '').trim())
      .filter(t => t &&
        !/^Ire ID:/i.test(t) &&
        !/^\d+\s+Images Shown on Site/i.test(t)
      )
      .join(' | ');

    return allText || 'Image';
  }

  function getLeadFlagsFromCard(card) {
    const text = card.textContent || '';
    const hasOverride = /Lead Override/i.test(text);
    const hasLeadEligible = /Lead Eligible/i.test(text);
    return { hasOverride, hasLeadEligible };
  }

  function getIreIdFromCard(card) {
    const legacyPs = Array.from(
      card.querySelectorAll('p[data-test-id="legacy-asset-id-for-asset-card"]')
    );
    for (const p of legacyPs) {
      const txt = (p.textContent || '').trim();
      const match = txt.match(/^\s*Ire ID:\s*(\d+)/i);
      if (match) return match[1];
    }
    return null;
  }

  function snapshotPage(options) {
    const opts = options || {};
    const silent = !!opts.silent;

    const rows = loadSnapshot();

    const cards = Array.from(
      document.querySelectorAll('div[data-test-id^="asset-"][data-test-id*="-status-"]')
    );

    if (!cards.length) {
      if (!silent) {
        alert(
          'Không tìm thấy asset card nào trên trang.\n' +
            'Hãy kiểm tra lại xem bạn có đang ở trang Variant Media / Listing Associations không.'
        );
      }
      return;
    }

    let added = 0;

    for (const card of cards) {
      const testId = card.getAttribute('data-test-id');
      if (!testId) continue;

      const parts = testId.split('-');
      if (parts.length < 4) continue;

      const sku = parts[1];
      const mfg = parts[2];

      const assetId = getIreIdFromCard(card);
      if (!assetId) continue;

      const type = getTypeFromCard(card);
      const { hasOverride, hasLeadEligible } = getLeadFlagsFromCard(card);

      let row = findRow(rows, sku, mfg, assetId);

      let backendDim;
      if (Object.prototype.hasOwnProperty.call(ireHasDimensions, String(assetId))) {
        backendDim = !!ireHasDimensions[String(assetId)];
      }

      if (!row) {
        row = {
          sku,
          manufacturerPartId: mfg,
          assetId,
          initial: {
            hasOverride,
            hasLeadEligible,
            type,
            dimension: typeof backendDim === 'boolean' ? backendDim : undefined
          },
          current: {
            hasOverride,
            hasLeadEligible,
            type,
            dimension: typeof backendDim === 'boolean' ? backendDim : undefined
          }
        };
        rows.push(row);
        added++;
      } else {
        const curDim = row.current ? row.current.dimension : undefined;
        row.current = {
          hasOverride,
          hasLeadEligible,
          type,
          dimension: curDim
        };

        if (row.initial && typeof row.initial.dimension === 'undefined' && typeof backendDim === 'boolean') {
          row.initial.dimension = backendDim;
        }
      }
    }

    saveSnapshot(rows);

    if (!silent) {
      const perMfg = {};
      for (const r of rows) {
        perMfg[r.manufacturerPartId] =
          (perMfg[r.manufacturerPartId] || 0) + 1;
      }
      const perMfgLines = Object.entries(perMfg)
        .slice(0, 10)
        .map(([m, c]) => `  - ${m}: ${c} asset(s)`)
        .join('\n');

      alert(
        'Snapshot hoàn tất.\n' +
          `Thêm mới: ${added} asset(s).\n` +
          `Tổng snapshot hiện có: ${rows.length} dòng (SKU + ManufacturerPartID + Ire ID).\n\n` +
          'Thống kê theo ManufacturerPartID (tối đa 10 dòng):\n' +
          (perMfgLines || '  (chưa có dữ liệu)')
      );
    }
  }

  function updateSnapshotDimensionByAssetId(assetId, dimValue) {
    if (!assetId) return;
    const rows = loadSnapshot();
    if (!rows.length) return;

    let changed = false;
    for (const r of rows) {
      if (r.assetId !== assetId) continue;

      if (!r.initial) r.initial = {};
      if (!r.current) r.current = {};

      if (dimValue === true || dimValue === false) {
        if (typeof r.initial.dimension === 'undefined') {
          r.initial.dimension = dimValue;
        }
        r.current.dimension = dimValue;
      } else {
        r.current.dimension = null;
      }
      changed = true;
    }

    if (changed) {
      saveSnapshot(rows);
    }
  }

  async function exportTSV() {
    try {
      snapshotPage({ silent: true });
    } catch (e) {
      console.warn('WF Snapshot Helper: refresh snapshot trước khi export bị lỗi', e);
    }

    const rows = loadSnapshot();
    if (!rows.length) {
      alert('Chưa có dữ liệu snapshot nào để xuất.');
      return;
    }

    const sortedRows = [...rows].sort((a, b) => {
      if (a.manufacturerPartId !== b.manufacturerPartId) {
        return a.manufacturerPartId.localeCompare(b.manufacturerPartId);
      }
      if (a.sku !== b.sku) {
        return a.sku.localeCompare(b.sku);
      }
      return String(a.assetId).localeCompare(String(b.assetId));
    });

    const header = [
      'SKU',
      'ManufacturerPartID',
      'AssetID',
      'Agent Notes',
      'Details Note',
      'Mark as Lead Override?',
      'Tagged as Lead?',
      'Mark as Not Lead Eligible?',
      'Type ban đầu',
      'Đổi thành'
    ];

    const lines = [header.join('\t')];

    for (const r of sortedRows) {
      const initType = (r.initial && r.initial.type) || 'Image';
      const curType = (r.current && r.current.type) || 'Image';

      const hasOverrideInit = !!(r.initial && r.initial.hasOverride);
      const hasLeadEligInit = !!(r.initial && r.initial.hasLeadEligible);
      const hasOverrideCur = !!(r.current && r.current.hasOverride);
      const hasLeadEligCur = !!(r.current && r.current.hasLeadEligible);

      const colT = hasOverrideInit ? 'Yes' : 'No';
      const colU = hasOverrideInit || hasLeadEligInit ? 'Yes' : 'No';

      let colV = 'No';
      if (colU === 'Yes' && !(hasOverrideCur || hasLeadEligCur)) {
        colV = 'Yes';
      }

      let detailsNote = '';
      let typeFrom = '';
      let typeTo = '';

      if (initType !== curType) {
        if (initType === 'Image') {
          detailsNote = 'Image type added';
        } else {
          detailsNote = 'Image type corrected';
        }
        typeFrom = initType;
        typeTo = curType;
      }

      let curDim;
      if (r.current && typeof r.current.dimension !== 'undefined') {
        curDim = r.current.dimension;
      } else if (Object.prototype.hasOwnProperty.call(ireHasDimensions, String(r.assetId))) {
        curDim = !!ireHasDimensions[String(r.assetId)];
      }

      const agentNotes = curDim === true ? 'Has Dimensions' : '';

      lines.push(
        [
          r.sku,
          r.manufacturerPartId,
          r.assetId,
          agentNotes,
          detailsNote,
          colT,
          colU,
          colV,
          typeFrom,
          typeTo
        ].join('\t')
      );
    }

    const tsv = lines.join('\n');

    try {
      await navigator.clipboard.writeText(tsv);
      alert(
        `Đã copy ${sortedRows.length} dòng vào clipboard.\n` +
          'Dán trực tiếp vào Google Sheet (Ctrl+V).\n' +
          'Mỗi URL snapshot riêng, cột Details Note thể hiện mọi thay đổi Image Type giữa các lần snapshot.\n' +
          'Current đã được cập nhật theo DOM tại thời điểm bạn bấm Alt+1.'
      );
    } catch (err) {
      console.error('Clipboard error:', err);
      alert(
        'Không copy được vào clipboard (trình duyệt chặn?).\n' +
          'Hãy kiểm tra lại quyền truy cập clipboard cho Tampermonkey.'
      );
    }
  }

  function clearSnapshot() {
    GM_setValue(getStorageKey(), []);
    alert('Đã xóa toàn bộ snapshot (SKU + ManufacturerPartID + Ire ID) cho trang hiện tại.');
  }

  // Hotkeys
  document.addEventListener('keydown', (e) => {
    const tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea') return;

    if (e.altKey && e.code === 'Backquote') {
      e.preventDefault();
      snapshotPage();
    } else if (e.altKey && e.key === '1') {
      e.preventDefault();
      exportTSV();
    } else if (e.altKey && e.key === '5') {
      e.preventDefault();
      clearSnapshot();
    }
  });

  console.log(
    '[WF Snapshot Helper] Loaded. Alt+` snapshot per URL, Alt+1 export TSV (có refresh current), Alt+5 clear snapshot for current URL.'
  );

  /* =========================================================
   *  HOOK SAVE TAGS → CẬP NHẬT DIMENSIONS VÀO SNAPSHOT
   * ======================================================= */

  function setupSaveTagsHook() {
    document.addEventListener('click', (e) => {
      const btn = e.target.closest('button');
      if (!btn) return;
      const txt = (btn.innerText || btn.textContent || '').trim();
      if (txt !== 'Save Tags') return;
      if (btn.disabled) return;

      if (!isPreviewOpen()) return;

      try {
        const previewRoot = getPreviewRoot();
        const dimVal = readHasDimensionsFromPreview(previewRoot);
        const assetId = readAssetIdFromPreview(previewRoot); // Ire ID
        if (assetId) {
          updateSnapshotDimensionByAssetId(assetId, dimVal);
        }
      } catch (err) {
        console.error('WF SaveTags hook error:', err);
      }
    }, true);
  }

  /* =========================================================
   *  INIT
   * ======================================================= */

  function init() {
    injectCSS();
    addPanels();
    ensureToolbar();
    ensureSectionControls();
    ensureListingSectionControls();
    updatePageState();
    createSaveButtonPMP();
    setupSaveTagsHook();

    setInterval(() => {
      try {
        addPanels();
        ensureToolbar();
        ensureSectionControls();
        ensureListingSectionControls();
        createSaveButtonPMP();
        updatePageState();
      } catch (e) {
        console.error('WF Helper tick error', e);
      }
    }, 1500);

    const stateObserver = new MutationObserver(() => {
      updatePageState();
    });
    stateObserver.observe(document.body, { childList: true, subtree: true });
  }

  init();

})();
