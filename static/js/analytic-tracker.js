/**
 * ============================================================
 * RELIEVV PDF-TO-EXCEL — COMPREHENSIVE ANALYTICS TRACKER
 * analytics_tracker.js
 *
 * DROP THIS SCRIPT TAG AT THE BOTTOM OF YOUR HTML, AFTER
 * script.js. IT AUTO-ATTACHES TO ALL EXISTING DOM ELEMENTS.
 *
 * Events tracked:
 *  - file_upload        → user selects / drags a PDF
 *  - conversion_start   → clicks Convert button
 *  - conversion_success → server returns result
 *  - conversion_error   → server returns error
 *  - file_download      → clicks Download button
 *  - reset_tool         → clicks Reset
 *  - tab_view           → switches info tabs (Features/Safety/Notes)
 *  - faq_expand         → clicks an FAQ item
 *  - nav_click          → clicks navbar links
 *  - page_visit         → on load (+ returning user, visit count)
 *  - time_on_page       → milestones at 30s, 60s, 120s, 300s
 *  - scroll_depth       → 25%, 50%, 75%, 90%, 100%
 *  - pdf_page_nav       → prev/next PDF page
 *  - sheet_option_change→ one sheet vs multiple sheets toggle
 *  - formatting_toggle  → include formatting checkbox
 * ============================================================
 */

(function () {
    'use strict';
  
    // ─── HELPERS ───────────────────────────────────────────────
  
    function sid() {
      try { return localStorage.getItem('rv_session_id') || 'unknown'; }
      catch (e) { return 'unknown'; }
    }
  
    function visitCount() {
      try { return parseInt(localStorage.getItem('rv_visit_count') || '1'); }
      catch (e) { return 1; }
    }
  
    function send(eventName, params) {
      if (typeof gtag !== 'function') return;
      gtag('event', eventName, Object.assign({
        session_id: sid(),
        visit_number: visitCount(),
        page_url: window.location.href,
        timestamp: new Date().toISOString()
      }, params));
    }
  
    function formatBytes(bytes) {
      if (!bytes) return 'unknown';
      if (bytes < 1024) return bytes + ' B';
      if (bytes < 1048576) return (bytes / 1024).toFixed(1) + ' KB';
      return (bytes / 1048576).toFixed(2) + ' MB';
    }
  
    function sizeCategory(bytes) {
      if (!bytes) return 'unknown';
      if (bytes < 102400) return 'small (<100KB)';
      if (bytes < 1048576) return 'medium (100KB–1MB)';
      if (bytes < 5242880) return 'large (1MB–5MB)';
      return 'very large (>5MB)';
    }
  
    // Reads conversion stats stored by script.js (if available)
    function getConversionStats() {
      try {
        return {
          total_conversions: parseInt(localStorage.getItem('rv_total_conversions') || '0'),
          today_conversions: parseInt(localStorage.getItem('rv_today_conversions') || '0')
        };
      } catch (e) { return {}; }
    }
  
    // ─── FILE UPLOAD EVENT ──────────────────────────────────────
    (function () {
      var fileInput = document.getElementById('fileInput');
      var uploadArea = document.getElementById('uploadArea');
  
      function onFileSelected(file, method) {
        if (!file) return;
        send('file_upload', {
          event_category: 'tool_usage',
          upload_method: method,         // 'browse' | 'drag_drop'
          file_name: file.name,
          file_size_bytes: file.size,
          file_size_label: formatBytes(file.size),
          file_size_category: sizeCategory(file.size),
          file_type: file.type || 'application/pdf'
        });
      }
  
      if (fileInput) {
        fileInput.addEventListener('change', function () {
          if (this.files && this.files[0]) {
            onFileSelected(this.files[0], 'browse');
          }
        });
      }
  
      if (uploadArea) {
        uploadArea.addEventListener('drop', function (e) {
          var files = e.dataTransfer && e.dataTransfer.files;
          if (files && files[0]) {
            onFileSelected(files[0], 'drag_drop');
          }
        });
        // Drag interaction curiosity tracking
        uploadArea.addEventListener('dragover', function () {
          send('drag_enter', { event_category: 'engagement', label: 'user dragged file over upload area' });
        });
      }
    })();
  
    // ─── CONVERT BUTTON ─────────────────────────────────────────
    (function () {
      var btn = document.getElementById('convertBtn');
      if (!btn) return;
  
      btn.addEventListener('click', function () {
        // Capture file info at click time
        var fileInput = document.getElementById('fileInput');
        var file = fileInput && fileInput.files && fileInput.files[0];
        var pageCount = parseInt(document.getElementById('summaryPageCount').textContent) || 0;
  
        send('conversion_start', {
          event_category: 'tool_usage',
          file_name: file ? file.name : 'no file',
          file_size_bytes: file ? file.size : 0,
          file_size_label: file ? formatBytes(file.size) : 'none',
          file_size_category: file ? sizeCategory(file.size) : 'none',
          page_count: pageCount,
          engine: 'convertapi'
        });
      });
    })();
  
    // ─── DOWNLOAD BUTTON & MODAL ─────────────────────────────────
    (function () {
      var downloadBtn  = document.getElementById('downloadBtn');
      var confirmBtn   = document.getElementById('confirmDownload');
  
      function fireDownload(trigger) {
        var pageCount = document.getElementById('modalPageCount')
                     ? parseInt(document.getElementById('modalPageCount').textContent) || 0
                     : 0;
        var cellCount = document.getElementById('modalCellCount')
                     ? parseInt(document.getElementById('modalCellCount').textContent.replace(/[^0-9]/g,'')) || 0
                     : 0;
        var procTime  = document.getElementById('modalTime')
                     ? document.getElementById('modalTime').textContent
                     : 'unknown';
        var fileSize  = document.getElementById('modalFileSize')
                     ? document.getElementById('modalFileSize').textContent
                     : 'unknown';
  
        // Increment local download counter
        try {
          var total = parseInt(localStorage.getItem('rv_total_conversions') || '0') + 1;
          localStorage.setItem('rv_total_conversions', total);
  
          var today = new Date().toDateString();
          var storedDate = localStorage.getItem('rv_today_date');
          var todayCount = storedDate === today
            ? parseInt(localStorage.getItem('rv_today_conversions') || '0') + 1
            : 1;
          localStorage.setItem('rv_today_date', today);
          localStorage.setItem('rv_today_conversions', todayCount);
        } catch (e) {}
  
        var stats = getConversionStats();
  
        send('file_download', {
          event_category: 'conversion',
          trigger: trigger,                // 'download_btn' | 'confirm_modal'
          pdf_pages: pageCount,
          cells_extracted: cellCount,
          processing_time: procTime,
          output_file_size: fileSize,
          user_total_downloads: stats.total_conversions,
          user_today_downloads: stats.today_conversions
        });
      }
  
      if (downloadBtn) {
        downloadBtn.addEventListener('click', function () {
          fireDownload('download_btn');
        });
      }
      if (confirmBtn) {
        confirmBtn.addEventListener('click', function () {
          fireDownload('confirm_modal');
        });
      }
    })();
  
    // ─── RESET TOOL ─────────────────────────────────────────────
    (function () {
      var btn = document.getElementById('resetBtn');
      if (!btn) return;
      btn.addEventListener('click', function () {
        send('reset_tool', {
          event_category: 'tool_usage',
          label: 'user reset the tool'
        });
      });
    })();
  
    // ─── PDF PAGE NAVIGATION ──────────────────────────────────────
    (function () {
      var prev = document.getElementById('prevPage');
      var next = document.getElementById('nextPage');
      if (prev) {
        prev.addEventListener('click', function () {
          send('pdf_page_nav', { event_category: 'engagement', direction: 'previous' });
        });
      }
      if (next) {
        next.addEventListener('click', function () {
          send('pdf_page_nav', { event_category: 'engagement', direction: 'next' });
        });
      }
    })();
  
    // ─── INFO TABS (Features / Safety / Notes) ────────────────────
    (function () {
      document.querySelectorAll('.info-tab').forEach(function (tab) {
        tab.addEventListener('click', function () {
          send('tab_view', {
            event_category: 'engagement',
            tab_name: tab.getAttribute('data-tab') || tab.textContent.trim()
          });
        });
      });
    })();
  
    // ─── FAQ EXPAND ───────────────────────────────────────────────
    (function () {
      document.querySelectorAll('.faq-item').forEach(function (item) {
        item.addEventListener('click', function () {
          var q = item.querySelector('h3');
          send('faq_expand', {
            event_category: 'engagement',
            question: q ? q.textContent.trim() : 'unknown'
          });
        });
      });
    })();
  
    // ─── NAVBAR LINK CLICKS ───────────────────────────────────────
    (function () {
      document.querySelectorAll('.navbar-nav .nav-link').forEach(function (link) {
        link.addEventListener('click', function () {
          send('nav_click', {
            event_category: 'navigation',
            link_text: link.textContent.trim(),
            link_href: link.getAttribute('href') || ''
          });
        });
      });
    })();
  
    // ─── SHEET LAYOUT TOGGLE ─────────────────────────────────────
    (function () {
      document.querySelectorAll('input[name="sheetLayout"]').forEach(function (radio) {
        radio.addEventListener('change', function () {
          send('sheet_option_change', {
            event_category: 'tool_config',
            value: radio.value === '1' ? 'one_sheet' : 'multiple_sheets'
          });
        });
      });
    })();
  
    // ─── FORMATTING CHECKBOX ─────────────────────────────────────
    (function () {
      var cb = document.getElementById('includeFormatting');
      if (!cb) return;
      cb.addEventListener('change', function () {
        send('formatting_toggle', {
          event_category: 'tool_config',
          enabled: cb.checked
        });
      });
    })();
  
    // ─── CONVERSION SUCCESS / ERROR HOOKS ────────────────────────
    // These hook into your existing script.js by observing the
    // progress label and download button state changes via MutationObserver.
    (function () {
      var downloadBtn = document.getElementById('downloadBtn');
      if (!downloadBtn) return;
  
      var conversionStartTime = null;
  
      // Watch when convert button is clicked to record start time
      var convertBtn = document.getElementById('convertBtn');
      if (convertBtn) {
        convertBtn.addEventListener('click', function () {
          conversionStartTime = Date.now();
        });
      }
  
      // Watch download button being enabled = conversion succeeded
      var observer = new MutationObserver(function (mutations) {
        mutations.forEach(function (m) {
          if (m.attributeName === 'disabled') {
            var isEnabled = !downloadBtn.disabled;
            if (isEnabled && conversionStartTime) {
              var elapsed = ((Date.now() - conversionStartTime) / 1000).toFixed(1);
              var pageCount = parseInt(document.getElementById('summaryPageCount').textContent) || 0;
  
              send('conversion_success', {
                event_category: 'conversion',
                processing_time_seconds: parseFloat(elapsed),
                pdf_pages: pageCount,
                label: 'converted in ' + elapsed + 's'
              });
  
              conversionStartTime = null;
            }
          }
        });
      });
  
      observer.observe(downloadBtn, { attributes: true, attributeFilter: ['disabled'] });
  
      // Watch progress label for error messages
      var progressLabel = document.getElementById('progressLabel');
      if (progressLabel) {
        var labelObserver = new MutationObserver(function () {
          var text = progressLabel.textContent.toLowerCase();
          if (text.includes('error') || text.includes('fail') || text.includes('could not')) {
            send('conversion_error', {
              event_category: 'error',
              error_message: progressLabel.textContent.trim()
            });
          }
        });
        labelObserver.observe(progressLabel, { childList: true, characterData: true, subtree: true });
      }
    })();
  
    // console.log('[Relievv Analytics] Tracker loaded — session:', sid(), '| visit #', visitCount());
  })();
