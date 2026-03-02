/* global Office, PowerPoint */
(function () {
  var REFERENCE_SHAPE_PREFIX = 'PPTAutomationReference_';
  var REFERENCE_LIST_SHAPE_PREFIX = 'PPTAutomationReferenceList_';
  var REFERENCE_HEADER_SHAPE_NAME = 'PPTAutomationReferenceHeader';
  var REFERENCE_BOX_SHAPE_NAME = 'PPTAutomationSourcesBox';
  var AUTO_RECOMMEND_IDLE_MS = 2400;
  var AUTO_RECOMMEND_POLL_MS = 1200;
  var AUTO_RECOMMEND_COOLDOWN_MS = 8000;
  var AUTO_RECOMMEND_SUPPRESSION_MS = 10000;

  if (!window.PPTAutomation) {
    window.PPTAutomation = {};
  }
  if (!window.PPTAutomation.uiState) {
    window.PPTAutomation.uiState = {
      latestPlan: null,
      latestSlideContext: null,
      pendingPlan: null,
      isRecommending: false,
      isAddingReference: false,
      isApplyingPlan: false,
      referenceDialogResolver: null,
      sourceLinksBySlide: {},
      previewSession: null,
      versionHistory: null
    };
  }
  if (typeof window.PPTAutomation.uiState.isApplyingPlan !== 'boolean') {
    window.PPTAutomation.uiState.isApplyingPlan = false;
  }
  if (typeof window.PPTAutomation.uiState.recommendationMonitorStarted !== 'boolean') {
    window.PPTAutomation.uiState.recommendationMonitorStarted = false;
  }
  if (typeof window.PPTAutomation.uiState.recommendationMonitorTicking !== 'boolean') {
    window.PPTAutomation.uiState.recommendationMonitorTicking = false;
  }
  if (typeof window.PPTAutomation.uiState.recommendationMonitorTimer !== 'number') {
    window.PPTAutomation.uiState.recommendationMonitorTimer = null;
  }
  if (typeof window.PPTAutomation.uiState.lastObservedActivitySignature !== 'string') {
    window.PPTAutomation.uiState.lastObservedActivitySignature = '';
  }
  if (typeof window.PPTAutomation.uiState.pendingActivitySignature !== 'string') {
    window.PPTAutomation.uiState.pendingActivitySignature = '';
  }
  if (typeof window.PPTAutomation.uiState.pendingActivityAt !== 'number') {
    window.PPTAutomation.uiState.pendingActivityAt = 0;
  }
  if (typeof window.PPTAutomation.uiState.lastRecommendationSignature !== 'string') {
    window.PPTAutomation.uiState.lastRecommendationSignature = '';
  }
  if (typeof window.PPTAutomation.uiState.lastRecommendationAt !== 'number') {
    window.PPTAutomation.uiState.lastRecommendationAt = 0;
  }
  if (typeof window.PPTAutomation.uiState.suppressAutoRecommendUntil !== 'number') {
    window.PPTAutomation.uiState.suppressAutoRecommendUntil = 0;
  }
  if (
    !window.PPTAutomation.uiState.previewSession ||
    typeof window.PPTAutomation.uiState.previewSession !== 'object'
  ) {
    window.PPTAutomation.uiState.previewSession = null;
  }
  if (
    !window.PPTAutomation.uiState.versionHistory ||
    typeof window.PPTAutomation.uiState.versionHistory !== 'object'
  ) {
    window.PPTAutomation.uiState.versionHistory = null;
  }

  function setStatus(message) {
    var node = document.getElementById('status');
    if (node) {
      node.textContent = message;
    }
  }

  function fetchBackendHealthSummary() {
    return fetch('/api/backend-health')
      .then(function (response) {
        if (!response.ok) {
          return response.text().then(function (t) {
            throw new Error('Backend health check failed: ' + t);
          });
        }
        return response.json();
      })
      .then(function (payload) {
        var provider = payload && typeof payload.provider === 'string' ? payload.provider : 'unknown';
        var model = payload && typeof payload.model === 'string' && payload.model.trim()
          ? payload.model.trim()
          : '(default)';
        return 'Backend LLM: ' + provider + ' ' + model;
      });
  }

  function nowMs() {
    return Date.now();
  }

  function isOverlayVisible(id) {
    var overlay = document.getElementById(id);
    if (!overlay) {
      return false;
    }
    return !overlay.classList.contains('hidden') && overlay.getAttribute('aria-hidden') !== 'true';
  }

  function clearPendingAutoRecommendation() {
    var uiState = window.PPTAutomation.uiState;
    uiState.pendingActivitySignature = '';
    uiState.pendingActivityAt = 0;
  }

  function resetRecommendationMonitorBaseline() {
    var uiState = window.PPTAutomation.uiState;
    uiState.lastObservedActivitySignature = '';
    clearPendingAutoRecommendation();
  }

  function suppressAutoRecommendations(durationMs) {
    var uiState = window.PPTAutomation.uiState;
    uiState.suppressAutoRecommendUntil = nowMs() + Math.max(0, Number(durationMs || 0));
    clearPendingAutoRecommendation();
    resetRecommendationMonitorBaseline();
  }

  function shouldAllowAutomaticRecommendation() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.isRecommending || uiState.isAddingReference || uiState.isApplyingPlan) {
      return false;
    }
    if (uiState.previewSession) {
      return false;
    }
    if (nowMs() < Number(uiState.suppressAutoRecommendUntil || 0)) {
      return false;
    }
    if (isOverlayVisible('previewOverlay') || isOverlayVisible('referenceOverlay')) {
      return false;
    }
    if (typeof window.PPTAutomation.collectSlideContext !== 'function') {
      return false;
    }
    return true;
  }

  function hidePreviewOverlay() {
    var overlay = document.getElementById('previewOverlay');
    var previewState = document.getElementById('overlayPreviewState');
    if (!overlay) return;
    overlay.classList.add('hidden');
    overlay.setAttribute('aria-hidden', 'true');
    if (previewState) {
      previewState.textContent = '';
    }
  }

  function hideReferenceOverlay() {
    var overlay = document.getElementById('referenceOverlay');
    if (!overlay) return;
    overlay.classList.add('hidden');
    overlay.setAttribute('aria-hidden', 'true');
  }

  function setReferenceOverlayError(message) {
    var errorNode = document.getElementById('referenceOverlayError');
    if (!errorNode) return;
    errorNode.textContent = message || '';
  }

  function resolveReferenceDialog(result) {
    var uiState = window.PPTAutomation.uiState;
    var resolver = uiState.referenceDialogResolver;
    uiState.referenceDialogResolver = null;
    hideReferenceOverlay();
    setReferenceOverlayError('');
    if (typeof resolver === 'function') {
      resolver(result);
    }
  }

  function showReferenceOverlay(payload) {
    var overlay = document.getElementById('referenceOverlay');
    var summary = document.getElementById('referenceSelectionSummary');
    var labelInput = document.getElementById('referenceLabelInput');
    var urlInput = document.getElementById('referenceUrlInput');
    var selectedLabel = payload && payload.selectedLabel ? payload.selectedLabel : 'Selected item';
    if (!overlay || !summary || !labelInput || !urlInput) {
      return Promise.reject(new Error('Reference form is unavailable.'));
    }

    summary.textContent = 'Selected item: ' + selectedLabel;
    labelInput.value = selectedLabel;
    urlInput.value = '';
    setReferenceOverlayError('');

    overlay.classList.remove('hidden');
    overlay.setAttribute('aria-hidden', 'false');
    try {
      labelInput.focus();
      labelInput.select();
    } catch (_error) {
      // no-op
    }

    return new Promise(function (resolve) {
      window.PPTAutomation.uiState.referenceDialogResolver = resolve;
    });
  }

  function showPreviewOverlay(plan, options) {
    var overlay = document.getElementById('previewOverlay');
    var summary = document.getElementById('overlaySummary');
    var previewState = document.getElementById('overlayPreviewState');
    var warnings = document.getElementById('overlayWarnings');
    var operations = document.getElementById('overlayOperations');
    if (!overlay || !summary || !warnings || !operations) return;
    var opts = options || {};

    summary.textContent = (plan && plan.summary) ? plan.summary : 'Review the generated plan before applying.';
    if (previewState) {
      previewState.textContent = opts.previewMessage || '';
    }
    warnings.innerHTML = '';
    operations.innerHTML = '';

    var planWarnings = plan && Array.isArray(plan.warnings) ? plan.warnings.slice() : [];
    if (Array.isArray(opts.runtimeWarnings) && opts.runtimeWarnings.length) {
      planWarnings = planWarnings.concat(opts.runtimeWarnings);
    }
    for (var i = 0; i < planWarnings.length; i += 1) {
      var wp = document.createElement('p');
      wp.textContent = '- ' + planWarnings[i];
      warnings.appendChild(wp);
    }

    var ops = plan && Array.isArray(plan.operations) ? plan.operations : [];
    if (!ops.length) {
      var empty = document.createElement('p');
      empty.className = 'overlay-op';
      empty.textContent = 'No operations in plan.';
      operations.appendChild(empty);
    } else {
      for (var j = 0; j < ops.length; j += 1) {
        var op = ops[j] || {};
        var item = document.createElement('div');
        item.className = 'overlay-op';
        var txt = summarizeOperationContent(op);
        item.textContent = (j + 1) + '. ' + (op.type || 'unknown') + ' -> ' + (op.target || 'auto-target') + ' | ' + txt.slice(0, 120);
        operations.appendChild(item);
      }
    }

    overlay.classList.remove('hidden');
    overlay.setAttribute('aria-hidden', 'false');
  }

  function summarizeOperationContent(op) {
    var content = op && op.content && typeof op.content === 'object' ? op.content : null;
    if (!content) return '(no content)';

    if (typeof content.text === 'string' && content.text.trim()) {
      return content.text.trim();
    }

    if (content.smartArt && typeof content.smartArt === 'object') {
      var smartItems = Array.isArray(content.smartArt.items) ? content.smartArt.items : [];
      var layout = typeof content.smartArt.layout === 'string' ? content.smartArt.layout : 'process';
      return 'SmartArt (' + layout + ', ' + smartItems.length + ' item(s))';
    }

    if (content.chart && typeof content.chart === 'object') {
      var chartType = typeof content.chart.type === 'string' ? content.chart.type : 'chart';
      return 'Chart (' + chartType + ')';
    }

    if (content.table && typeof content.table === 'object') {
      var tableRows = Array.isArray(content.table.rows) ? content.table.rows.length : 0;
      return 'Table (' + tableRows + ' row(s))';
    }

    if (Array.isArray(content.rows) && content.rows.length) {
      return 'Rows payload (' + content.rows.length + ' row(s))';
    }

    if (content.image && typeof content.image === 'object') {
      return 'Image payload';
    }

    return '(non-text payload)';
  }

  function renderRecommendations(recommendations, onSelect) {
    var recEl = document.getElementById('recommendations');
    if (!recEl) return;
    recEl.innerHTML = '';

    if (!recommendations || !recommendations.length) {
      recEl.textContent = 'No recommendations returned.';
      return;
    }

    for (var i = 0; i < recommendations.length; i += 1) {
      (function (rec) {
        var item = document.createElement('article');
        item.className = 'recommendation';

        var top = document.createElement('div');
        top.className = 'rec-top';

        var title = document.createElement('h3');
        title.textContent = rec.title || 'Recommendation';

        var typeChip = document.createElement('span');
        typeChip.className = 'rec-type';
        typeChip.textContent = rec.outputType || 'other';

        top.appendChild(title);
        top.appendChild(typeChip);

        var desc = document.createElement('p');
        desc.textContent = rec.description || '';

        var meta = document.createElement('p');
        meta.className = 'meta';
        var conf = typeof rec.confidence === 'number' ? rec.confidence.toFixed(2) : '0.00';
        meta.textContent = 'Confidence: ' + conf;

        var hintsRow = document.createElement('div');
        hintsRow.className = 'rec-hints';
        var hints = Array.isArray(rec.applyHints) ? rec.applyHints.slice(0, 4) : [];
        for (var h = 0; h < hints.length; h += 1) {
          var hint = document.createElement('span');
          hint.className = 'rec-hint';
          hint.textContent = hints[h];
          hintsRow.appendChild(hint);
        }

        var button = document.createElement('button');
        button.type = 'button';
        button.textContent = 'Generate Plan';
        button.addEventListener('click', function () {
          onSelect(rec);
        });

        item.appendChild(top);
        item.appendChild(desc);
        item.appendChild(meta);
        if (hints.length) {
          item.appendChild(hintsRow);
        }
        item.appendChild(button);
        recEl.appendChild(item);
      })(recommendations[i]);
    }
  }

  function buildSummaryContextWithImage(slideContext) {
    var objects = slideContext && Array.isArray(slideContext.objects) ? slideContext.objects : [];
    var summarized = [];
    for (var i = 0; i < objects.length && i < 20; i += 1) {
      var obj = objects[i] || {};
      summarized.push({
        id: obj.id,
        name: obj.name,
        type: obj.type,
        text: typeof obj.text === 'string' ? obj.text.slice(0, 200) : undefined,
        style: obj.style,
        table: obj.table,
        chart: obj.chart,
        bbox: obj.bbox
      });
    }

    var imageBase64 = null;
    if (slideContext && slideContext.rawSlide && typeof slideContext.rawSlide.imageBase64 === 'string') {
      imageBase64 = slideContext.rawSlide.imageBase64.slice(0, 180000);
    }

    return {
      slide: (slideContext && slideContext.slide) || {},
      selection: (slideContext && slideContext.selection) || { shapeIds: [] },
      themeHints: (slideContext && slideContext.themeHints) || {},
      objectCount: objects.length,
      objects: summarized,
      rawSlide: {
        ooxml: null,
        imageBase64: imageBase64,
        exportedAt: (slideContext && slideContext.rawSlide && slideContext.rawSlide.exportedAt) || null
      }
    };
  }

  function collectSelectedTextSnapshot() {
    if (
      !Office ||
      !Office.context ||
      !Office.context.document ||
      typeof Office.context.document.getSelectedDataAsync !== 'function' ||
      !Office.CoercionType ||
      !Office.CoercionType.Text
    ) {
      return Promise.resolve('');
    }

    return new Promise(function (resolve) {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
        if (!asyncResult || asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          resolve('');
          return;
        }

        var value = asyncResult.value;
        if (typeof value === 'string') {
          resolve(value.trim());
          return;
        }
        if (value && typeof value === 'object') {
          var candidate = value.value || value.text || value.data;
          resolve(typeof candidate === 'string' ? candidate.trim() : '');
          return;
        }
        resolve('');
      });
    });
  }

  function buildRecommendationActivitySignature(payload) {
    var selectedShapeIds = payload && Array.isArray(payload.selectedShapeIds)
      ? payload.selectedShapeIds
      : (payload && payload.selection && Array.isArray(payload.selection.shapeIds) ? payload.selection.shapeIds : []);
    var selectedText = payload && typeof payload.selectedText === 'string'
      ? payload.selectedText
      : (payload && payload.selection && typeof payload.selection.text === 'string' ? payload.selection.text : '');
    var slideId = payload && typeof payload.slideId === 'string'
      ? payload.slideId
      : (payload && payload.slide && typeof payload.slide.id === 'string' ? payload.slide.id : 'active');
    var objects = payload && Array.isArray(payload.objects) ? payload.objects : [];
    var ranked = [];
    var i;

    for (i = 0; i < objects.length; i += 1) {
      var item = objects[i] || {};
      var text = typeof item.text === 'string' ? squeezeInlineText(item.text).slice(0, 180) : '';
      var name = typeof item.name === 'string' ? squeezeInlineText(item.name).slice(0, 80) : '';
      if (!text && !name) {
        continue;
      }
      var top = Number.POSITIVE_INFINITY;
      var left = Number.POSITIVE_INFINITY;
      if (Array.isArray(item.bbox) && item.bbox.length >= 2) {
        top = Number(item.bbox[1]);
        left = Number(item.bbox[0]);
      } else {
        top = Number(item.top);
        left = Number(item.left);
      }
      ranked.push({
        key: (text || name) + '|' + name,
        top: Number.isFinite(top) ? top : Number.POSITIVE_INFINITY,
        left: Number.isFinite(left) ? left : Number.POSITIVE_INFINITY
      });
    }

    ranked.sort(function (a, b) {
      if (a.top !== b.top) {
        return a.top - b.top;
      }
      if (a.left !== b.left) {
        return a.left - b.left;
      }
      return a.key.localeCompare(b.key);
    });

    var parts = [
      String(slideId || 'active'),
      selectedShapeIds.join(','),
      squeezeInlineText(String(selectedText || '')).slice(0, 600)
    ];
    for (i = 0; i < ranked.length && i < 12; i += 1) {
      parts.push(ranked[i].key);
    }
    return parts.join('||');
  }

  function collectRecommendationActivitySnapshot() {
    if (typeof PowerPoint === 'undefined' || !PowerPoint || typeof PowerPoint.run !== 'function') {
      return Promise.resolve(null);
    }

    return PowerPoint.run(function (context) {
      var presentation = context.presentation;
      var slides = presentation.slides;
      var selectedShapes = null;
      var selectedTextRangeOrNull = null;
      var selectedTextRange = null;

      slides.load('items');
      if (presentation && typeof presentation.getSelectedShapes === 'function') {
        selectedShapes = presentation.getSelectedShapes();
        selectedShapes.load('items/id');
      }

      return context.sync()
        .then(function () {
          return resolveActiveSlide(context, slides);
        })
        .then(function (activeSlide) {
          if (!activeSlide) {
            return null;
          }

          activeSlide.load('id');
          var shapes = activeSlide.shapes;
          shapes.load('items/id,items/name,items/type,items/left,items/top');

          if (presentation && typeof presentation.getSelectedTextRangeOrNullObject === 'function') {
            selectedTextRangeOrNull = presentation.getSelectedTextRangeOrNullObject();
            selectedTextRangeOrNull.load('isNullObject,text');
          } else if (presentation && typeof presentation.getSelectedTextRange === 'function') {
            selectedTextRange = presentation.getSelectedTextRange();
            selectedTextRange.load('text');
          }

          return context.sync().then(function () {
            var shapeItems = shapes.items || [];
            var ranked = shapeItems.slice().sort(function (a, b) {
              var aTop = Number(a && a.top);
              var bTop = Number(b && b.top);
              if (aTop !== bTop) {
                return aTop - bTop;
              }
              var aLeft = Number(a && a.left);
              var bLeft = Number(b && b.left);
              return aLeft - bLeft;
            }).slice(0, 12);

            for (var i = 0; i < ranked.length; i += 1) {
              var shape = ranked[i];
              try {
                if (shape && shape.textFrame && shape.textFrame.textRange) {
                  shape.textFrame.textRange.load('text');
                }
              } catch (_error) {
                // best-effort only
              }
            }

            return context.sync().then(function () {
              return collectSelectedTextSnapshot().then(function (asyncSelectedText) {
                var selectedShapeIds = [];
                var selectedItems = selectedShapes && Array.isArray(selectedShapes.items) ? selectedShapes.items : [];
                for (var j = 0; j < selectedItems.length; j += 1) {
                  var selectedId = selectedItems[j] && selectedItems[j].id;
                  if (typeof selectedId === 'string' && selectedId) {
                    selectedShapeIds.push(selectedId);
                  }
                }

                var selectedText = '';
                if (
                  selectedTextRangeOrNull &&
                  selectedTextRangeOrNull.isNullObject === false &&
                  typeof selectedTextRangeOrNull.text === 'string'
                ) {
                  selectedText = selectedTextRangeOrNull.text.trim();
                } else if (selectedTextRange && typeof selectedTextRange.text === 'string') {
                  selectedText = selectedTextRange.text.trim();
                }
                if (!selectedText && asyncSelectedText) {
                  selectedText = asyncSelectedText;
                }

                var snapshotObjects = [];
                for (var k = 0; k < ranked.length; k += 1) {
                  var rankedShape = ranked[k] || {};
                  var shapeText = '';
                  try {
                    shapeText = rankedShape && rankedShape.textFrame && rankedShape.textFrame.textRange
                      ? String(rankedShape.textFrame.textRange.text || '').trim()
                      : '';
                  } catch (_error) {
                    shapeText = '';
                  }

                  snapshotObjects.push({
                    name: rankedShape && rankedShape.name ? rankedShape.name : '',
                    text: shapeText,
                    top: Number(rankedShape && rankedShape.top),
                    left: Number(rankedShape && rankedShape.left)
                  });
                }

                var signature = buildRecommendationActivitySignature({
                  slideId: activeSlide.id || 'active',
                  selectedShapeIds: selectedShapeIds,
                  selectedText: selectedText,
                  objects: snapshotObjects
                });

                return {
                  slideId: activeSlide.id || 'active',
                  selectedShapeIds: selectedShapeIds,
                  selectedText: selectedText,
                  objects: snapshotObjects,
                  signature: signature,
                  hasSignal: Boolean(selectedShapeIds.length || selectedText || snapshotObjects.length)
                };
              });
            });
          });
        })
        .catch(function () {
          return null;
        });
    });
  }

  function scheduleRecommendationMonitorTick(delayMs) {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.recommendationMonitorTimer) {
      window.clearTimeout(uiState.recommendationMonitorTimer);
    }
    uiState.recommendationMonitorTimer = window.setTimeout(function () {
      pollRecommendationActivity();
    }, Math.max(100, Number(delayMs || AUTO_RECOMMEND_POLL_MS)));
  }

  function pollRecommendationActivity() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.recommendationMonitorTicking) {
      scheduleRecommendationMonitorTick(AUTO_RECOMMEND_POLL_MS);
      return;
    }
    uiState.recommendationMonitorTicking = true;

    collectRecommendationActivitySnapshot()
      .then(function (snapshot) {
        if (!snapshot || !snapshot.signature) {
          return;
        }

        var currentTime = nowMs();
        if (!uiState.lastObservedActivitySignature) {
          uiState.lastObservedActivitySignature = snapshot.signature;
          clearPendingAutoRecommendation();
          return;
        }

        if (!snapshot.hasSignal) {
          if (snapshot.signature !== uiState.lastObservedActivitySignature) {
            uiState.lastObservedActivitySignature = snapshot.signature;
          }
          clearPendingAutoRecommendation();
          return;
        }

        if (snapshot.signature !== uiState.lastObservedActivitySignature) {
          uiState.lastObservedActivitySignature = snapshot.signature;
          uiState.pendingActivitySignature = snapshot.signature;
          uiState.pendingActivityAt = currentTime;
          return;
        }

        if (!uiState.pendingActivitySignature || uiState.pendingActivitySignature !== snapshot.signature) {
          return;
        }

        if (!shouldAllowAutomaticRecommendation()) {
          return;
        }

        if ((currentTime - Number(uiState.pendingActivityAt || 0)) < AUTO_RECOMMEND_IDLE_MS) {
          return;
        }

        if (
          uiState.lastRecommendationSignature === snapshot.signature &&
          (currentTime - Number(uiState.lastRecommendationAt || 0)) < AUTO_RECOMMEND_COOLDOWN_MS
        ) {
          clearPendingAutoRecommendation();
          return;
        }

        clearPendingAutoRecommendation();
        runRecommendationCycle({
          trigger: 'idle',
          activitySignature: snapshot.signature
        });
      })
      .catch(function (_error) {
        // best-effort background polling only
      })
      .then(function () {
        uiState.recommendationMonitorTicking = false;
        scheduleRecommendationMonitorTick(AUTO_RECOMMEND_POLL_MS);
      }, function () {
        uiState.recommendationMonitorTicking = false;
        scheduleRecommendationMonitorTick(AUTO_RECOMMEND_POLL_MS);
      });
  }

  function startRecommendationIdleMonitor() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.recommendationMonitorStarted) {
      return;
    }
    uiState.recommendationMonitorStarted = true;
    scheduleRecommendationMonitorTick(300);
  }

  function isSlidePreviewApiSupported() {
    if (!Office || !Office.context || !Office.context.requirements) {
      return false;
    }
    if (typeof Office.context.requirements.isSetSupported !== 'function') {
      return false;
    }
    return Office.context.requirements.isSetSupported('PowerPointApi', '1.8');
  }

  function slideIdsFromCollection(slides) {
    var ids = [];
    var items = slides && Array.isArray(slides.items) ? slides.items : [];
    for (var i = 0; i < items.length; i += 1) {
      var slideId = items[i] && items[i].id;
      if (typeof slideId === 'string' && slideId) {
        ids.push(slideId);
      }
    }
    return ids;
  }

  function findSlideByIdInCollection(slides, slideId) {
    var items = slides && Array.isArray(slides.items) ? slides.items : [];
    for (var i = 0; i < items.length; i += 1) {
      if (items[i] && items[i].id === slideId) {
        return items[i];
      }
    }
    return null;
  }

  function getSlideIndexById(slides, slideId) {
    var items = slides && Array.isArray(slides.items) ? slides.items : [];
    for (var i = 0; i < items.length; i += 1) {
      if (items[i] && items[i].id === slideId) {
        return i;
      }
    }
    return -1;
  }

  function setButtonDisabled(id, disabled) {
    var button = document.getElementById(id);
    if (button) {
      button.disabled = Boolean(disabled);
    }
  }

  function canUndoAcceptedVersion() {
    var history = window.PPTAutomation.uiState.versionHistory;
    return Boolean(
      history &&
      Array.isArray(history.versions) &&
      history.versions.length > 1 &&
      Number(history.currentIndex) > 0 &&
      typeof history.currentSlideId === 'string' &&
      history.currentSlideId
    );
  }

  function canRedoAcceptedVersion() {
    var history = window.PPTAutomation.uiState.versionHistory;
    return Boolean(
      history &&
      Array.isArray(history.versions) &&
      history.versions.length > 1 &&
      Number(history.currentIndex) >= 0 &&
      Number(history.currentIndex) < history.versions.length - 1 &&
      typeof history.currentSlideId === 'string' &&
      history.currentSlideId
    );
  }

  function updateUndoRedoButtons() {
    var uiState = window.PPTAutomation.uiState;
    var isBusy = uiState.isRecommending || uiState.isAddingReference || uiState.isApplyingPlan || Boolean(uiState.previewSession);
    setButtonDisabled('undoAcceptedBtn', isBusy || !canUndoAcceptedVersion());
    setButtonDisabled('redoAcceptedBtn', isBusy || !canRedoAcceptedVersion());
  }

  function clearPreviewState() {
    var uiState = window.PPTAutomation.uiState;
    uiState.previewSession = null;
    uiState.pendingPlan = null;
    uiState.latestPlan = null;
    hidePreviewOverlay();
    updateUndoRedoButtons();
  }

  function getActiveSlideDescriptor() {
    if (!isSlidePreviewApiSupported()) {
      return Promise.reject(new Error('This PowerPoint host does not support slide preview APIs.'));
    }

    return PowerPoint.run(function (context) {
      var slides = context.presentation.slides;
      slides.load('items/id');

      return context.sync()
        .then(function () {
          return resolveActiveSlide(context, slides);
        })
        .then(function (activeSlide) {
          if (!activeSlide) {
            return null;
          }
          activeSlide.load('id');
          return context.sync().then(function () {
            return {
              slideId: activeSlide.id,
              index: getSlideIndexById(slides, activeSlide.id)
            };
          });
        });
    });
  }

  function exportSlideSnapshot(slideId) {
    if (!slideId) {
      return Promise.reject(new Error('A slide must be selected first.'));
    }

    return PowerPoint.run(function (context) {
      var slides = context.presentation.slides;
      slides.load('items/id');

      return context.sync()
        .then(function () {
          var slide = findSlideByIdInCollection(slides, slideId);
          if (!slide || typeof slide.exportAsBase64 !== 'function') {
            throw new Error('Unable to export the selected slide.');
          }
          var exported = slide.exportAsBase64();
          return context.sync().then(function () {
            var value = exported && typeof exported.value === 'string' ? exported.value : '';
            if (!value) {
              throw new Error('Slide export returned no data.');
            }
            return value;
          });
        });
    });
  }

  function insertSlideSnapshotAfter(snapshotBase64, targetSlideId) {
    if (!snapshotBase64) {
      return Promise.reject(new Error('Slide snapshot data is missing.'));
    }

    return PowerPoint.run(function (context) {
      var slides = context.presentation.slides;
      slides.load('items/id');

      return context.sync()
        .then(function () {
          var beforeIds = slideIdsFromCollection(slides);
          if (targetSlideId) {
            context.presentation.insertSlidesFromBase64(snapshotBase64, {
              targetSlideId: targetSlideId
            });
          } else {
            context.presentation.insertSlidesFromBase64(snapshotBase64);
          }

          return context.sync().then(function () {
            slides.load('items/id');
            return context.sync().then(function () {
              var afterIds = slideIdsFromCollection(slides);
              for (var i = 0; i < afterIds.length; i += 1) {
                if (beforeIds.indexOf(afterIds[i]) < 0) {
                  return afterIds[i];
                }
              }
              throw new Error('Unable to identify the inserted preview slide.');
            });
          });
        });
    });
  }

  function selectSlideById(slideId) {
    if (!slideId) {
      return Promise.resolve(false);
    }

    return PowerPoint.run(function (context) {
      if (!context.presentation || typeof context.presentation.setSelectedSlides !== 'function') {
        throw new Error('This PowerPoint host cannot select preview slides.');
      }
      context.presentation.setSelectedSlides([slideId]);
      return context.sync().then(function () {
        return true;
      });
    });
  }

  function deleteSlideById(slideId) {
    if (!slideId) {
      return Promise.resolve(false);
    }

    return PowerPoint.run(function (context) {
      var slides = context.presentation.slides;
      slides.load('items/id');

      return context.sync()
        .then(function () {
          var slide = findSlideByIdInCollection(slides, slideId);
          if (!slide || typeof slide.delete !== 'function') {
            return false;
          }
          slide.delete();
          return context.sync().then(function () {
            return true;
          });
        });
    });
  }

  function safelyDeleteSlideById(slideId) {
    return deleteSlideById(slideId).catch(function () {
      return false;
    });
  }

  function replaceTrackedSlideWithSnapshot(currentSlideId, snapshotBase64) {
    var insertedSlideId = '';
    return insertSlideSnapshotAfter(snapshotBase64, currentSlideId)
      .then(function (newSlideId) {
        insertedSlideId = newSlideId;
        return deleteSlideById(currentSlideId);
      })
      .then(function (deleted) {
        if (!deleted) {
          return safelyDeleteSlideById(insertedSlideId).then(function () {
            throw new Error('Failed to replace the current slide version.');
          });
        }
        return selectSlideById(insertedSlideId).then(function () {
          return insertedSlideId;
        }, function () {
          return insertedSlideId;
        });
      });
  }

  function recordAcceptedSlideVersion(previewSession, acceptedSlideBase64) {
    var uiState = window.PPTAutomation.uiState;
    var beforeBase64 = previewSession && previewSession.originalSlideBase64
      ? previewSession.originalSlideBase64
      : '';
    var afterBase64 = String(acceptedSlideBase64 || '');
    var acceptedSlideId = previewSession && previewSession.previewSlideId ? previewSession.previewSlideId : '';
    var originalSlideId = previewSession && previewSession.originalSlideId ? previewSession.originalSlideId : '';
    var history = uiState.versionHistory;

    if (!beforeBase64 || !afterBase64 || !acceptedSlideId) {
      uiState.versionHistory = null;
      updateUndoRedoButtons();
      return;
    }

    if (
      history &&
      Array.isArray(history.versions) &&
      history.currentSlideId === originalSlideId &&
      Number(history.currentIndex) >= 0
    ) {
      history.versions = history.versions.slice(0, history.currentIndex + 1);
      history.versions[history.currentIndex] = beforeBase64;
      history.versions.push(afterBase64);
      history.currentIndex = history.versions.length - 1;
      history.currentSlideId = acceptedSlideId;
    } else {
      uiState.versionHistory = {
        versions: [beforeBase64, afterBase64],
        currentIndex: 1,
        currentSlideId: acceptedSlideId
      };
    }

    updateUndoRedoButtons();
  }

  function createPlanPreviewOnDuplicateSlide(plan, slideContext) {
    var uiState = window.PPTAutomation.uiState;
    var originalSlide = null;
    var previewSession = null;

    if (typeof window.PPTAutomation.applyExecutionPlan !== 'function') {
      return Promise.reject(new Error('Plan applier is unavailable.'));
    }
    if (uiState.previewSession) {
      return Promise.reject(new Error('Accept or reject the current preview first.'));
    }
    if (!isSlidePreviewApiSupported()) {
      return Promise.reject(new Error('This PowerPoint host does not support live slide previews.'));
    }

    return Promise.resolve()
      .then(function () {
        return getActiveSlideDescriptor();
      })
      .then(function (activeSlide) {
        if (!activeSlide || !activeSlide.slideId) {
          throw new Error('Select a slide before generating a preview.');
        }
        originalSlide = activeSlide;
        return exportSlideSnapshot(activeSlide.slideId);
      })
      .then(function (originalBase64) {
        previewSession = {
          originalSlideId: originalSlide.slideId,
          originalSlideIndex: originalSlide.index,
          originalSlideBase64: originalBase64,
          previewSlideId: ''
        };
        return insertSlideSnapshotAfter(originalBase64, originalSlide.slideId);
      })
      .then(function (previewSlideId) {
        previewSession.previewSlideId = previewSlideId;
        return selectSlideById(previewSlideId).then(function () {
          return window.PPTAutomation.applyExecutionPlan(plan, slideContext, {
            targetSlideId: previewSlideId
          });
        });
      })
      .then(function (applyResult) {
        uiState.previewSession = previewSession;
        updateUndoRedoButtons();
        return {
          previewSession: previewSession,
          applyResult: applyResult || { appliedCount: 0, warnings: [] }
        };
      })
      .catch(function (error) {
        return Promise.resolve()
          .then(function () {
            if (previewSession && previewSession.previewSlideId) {
              return safelyDeleteSlideById(previewSession.previewSlideId);
            }
            return false;
          })
          .then(function () {
            if (originalSlide && originalSlide.slideId) {
              return selectSlideById(originalSlide.slideId).catch(function () {
                return false;
              });
            }
            return false;
          })
          .then(function () {
            throw error;
          });
      });
  }

  function acceptPreviewSession() {
    var uiState = window.PPTAutomation.uiState;
    var previewSession = uiState.previewSession;
    if (!previewSession || !previewSession.previewSlideId || !previewSession.originalSlideId) {
      return Promise.reject(new Error('No preview slide is ready to accept.'));
    }

    var acceptedBase64 = '';
    return exportSlideSnapshot(previewSession.previewSlideId)
      .then(function (snapshotBase64) {
        acceptedBase64 = snapshotBase64;
        return deleteSlideById(previewSession.originalSlideId);
      })
      .then(function (deleted) {
        if (!deleted) {
          throw new Error('Failed to finalize the preview slide.');
        }
        return selectSlideById(previewSession.previewSlideId).catch(function () {
          return false;
        });
      })
      .then(function () {
        recordAcceptedSlideVersion(previewSession, acceptedBase64);
        clearPreviewState();
      });
  }

  function rejectPreviewSession() {
    var uiState = window.PPTAutomation.uiState;
    var previewSession = uiState.previewSession;
    if (!previewSession || !previewSession.previewSlideId) {
      clearPreviewState();
      return Promise.resolve(false);
    }

    return deleteSlideById(previewSession.previewSlideId)
      .then(function () {
        if (previewSession.originalSlideId) {
          return selectSlideById(previewSession.originalSlideId).catch(function () {
            return false;
          });
        }
        return false;
      })
      .then(function () {
        clearPreviewState();
        return true;
      });
  }

  function restoreAcceptedSlideVersion(direction) {
    var uiState = window.PPTAutomation.uiState;
    var history = uiState.versionHistory;
    var isUndo = direction === 'undo';
    if (!history || !Array.isArray(history.versions) || !history.currentSlideId) {
      return Promise.reject(new Error('No accepted slide version is available.'));
    }

    var targetIndex = isUndo ? (history.currentIndex - 1) : (history.currentIndex + 1);
    if (targetIndex < 0 || targetIndex >= history.versions.length) {
      return Promise.reject(new Error(isUndo ? 'Nothing to undo.' : 'Nothing to redo.'));
    }

    var snapshotBase64 = history.versions[targetIndex];
    return replaceTrackedSlideWithSnapshot(history.currentSlideId, snapshotBase64)
      .then(function (restoredSlideId) {
        history.currentSlideId = restoredSlideId;
        history.currentIndex = targetIndex;
        updateUndoRedoButtons();
        return restoredSlideId;
      })
      .catch(function (error) {
        uiState.versionHistory = null;
        updateUndoRedoButtons();
        throw error;
      });
  }

  function requestPlan(payload) {
    return fetch('/api/plans', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(function (response) {
      if (!response.ok) {
        return response.text().then(function (t) {
          throw new Error('Plan generation failed: ' + t);
        });
      }
      return response.json();
    });
  }

  function requestReferenceSuggestion(payload) {
    return fetch('/api/references', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).then(function (response) {
      if (!response.ok) {
        return response.text().then(function (t) {
          throw new Error('Reference lookup failed: ' + t);
        });
      }
      return response.json();
    });
  }

  function inferSelectedItemLabel(slideContext) {
    var ids = slideContext && slideContext.selection && Array.isArray(slideContext.selection.shapeIds)
      ? slideContext.selection.shapeIds
      : [];
    if (!ids.length) {
      return null;
    }

    var objects = slideContext && Array.isArray(slideContext.objects) ? slideContext.objects : [];
    var selectedId = ids[0];
    var obj = null;
    var i;
    for (i = 0; i < objects.length; i += 1) {
      if (objects[i] && objects[i].id === selectedId) {
        obj = objects[i];
        break;
      }
    }

    if (!obj) {
      return { shapeId: selectedId, label: 'Selected item', fullText: '' };
    }

    var sourceText = typeof obj.text === 'string' ? obj.text.trim() : '';
    var firstLine = sourceText ? sourceText.split(/\r?\n/)[0] : '';
    if (firstLine && firstLine.length > 0) {
      return { shapeId: selectedId, label: firstLine.slice(0, 120), fullText: sourceText };
    }

    var fallbackName = typeof obj.name === 'string' && obj.name.trim() ? obj.name.trim() : 'Selected item';
    return { shapeId: selectedId, label: fallbackName.slice(0, 120), fullText: '' };
  }

  function sanitizeSourceUrl(value) {
    var raw = String(value || '').trim();
    if (!raw) {
      return null;
    }

    if (!/^https?:\/\//i.test(raw)) {
      raw = 'https://' + raw;
    }

    try {
      var parsed = new URL(raw);
      if (!/^https?:$/i.test(parsed.protocol)) {
        return null;
      }
      return parsed.toString();
    } catch (_error) {
      return null;
    }
  }

  function buildFallbackReferenceUrl(text) {
    var query = String(text || '').trim() || 'reference';
    return 'https://en.wikipedia.org/w/index.php?search=' + encodeURIComponent(query);
  }

  function summarizeErrorMessage(error, maxLen) {
    var text = '';
    if (error && typeof error.message === 'string') {
      text = error.message;
    } else {
      text = String(error || '');
    }
    text = text.replace(/\s+/g, ' ').trim();
    return text.slice(0, maxLen || 180);
  }

  function stripBulletPrefix(line) {
    var text = String(line || '').replace(/\r/g, '');
    return text.replace(/^\s*[\u2022\u25CF\u25E6\u25AA\u25AB\-*]+\s+/, '').trim();
  }

  function stripTrailingReferenceMarkers(text) {
    var value = String(text || '');
    return value
      .replace(/\s*(?:[\u2070\u00B9\u00B2\u00B3\u2074\u2075\u2076\u2077\u2078\u2079]+|\[\d+\])+\s*$/g, '')
      .trim();
  }

  function normalizeClaimText(value) {
    return stripTrailingReferenceMarkers(stripBulletPrefix(String(value || '')))
      .replace(/\s+/g, ' ')
      .trim();
  }

  function normalizeLineBreaks(value) {
    return String(value || '')
      .replace(/\r\n/g, '\n')
      .replace(/[\r\v\f\u2028\u2029]/g, '\n');
  }

  function escapeRegExp(text) {
    return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function dedupeClaimTexts(values) {
    var seen = {};
    var output = [];
    for (var i = 0; i < values.length; i += 1) {
      var value = String(values[i] || '').replace(/\s+/g, ' ').trim();
      if (!value) {
        continue;
      }
      var key = value.toLowerCase();
      if (seen[key]) {
        continue;
      }
      seen[key] = true;
      output.push(value);
    }
    return output;
  }

  function extractBulletClaims(cleanText) {
    var lines = normalizeLineBreaks(cleanText).split('\n');
    var bulletPattern = /^\s*[\u2022\u25CF\u25E6\u25AA\u25AB\-*]+\s+/;
    var claims = [];
    var current = '';
    var collecting = false;

    for (var i = 0; i < lines.length; i += 1) {
      var rawLine = String(lines[i] || '').replace(/\r/g, '');
      var isBullet = bulletPattern.test(rawLine);
      var content = isBullet
        ? rawLine.replace(bulletPattern, '').trim()
        : rawLine.trim();

      if (isBullet) {
        if (collecting && current) {
          claims.push(current);
        }
        current = content;
        collecting = true;
        continue;
      }

      if (!collecting) {
        continue;
      }

      if (!content) {
        if (current) {
          claims.push(current);
        }
        current = '';
        collecting = false;
        continue;
      }

      current = current ? (current + ' ' + content) : content;
    }

    if (collecting && current) {
      claims.push(current);
    }

    var output = [];
    for (var j = 0; j < claims.length; j += 1) {
      var claim = normalizeClaimText(claims[j]);
      if (claim.length >= 18) {
        output.push(claim);
      }
    }
    return output;
  }

  function isLikelyContinuationLine(previousLine, nextRawLine) {
    var prev = String(previousLine || '').trim();
    var next = String(nextRawLine || '').trim();
    if (!prev || !next) {
      return false;
    }
    if (/[,;:-]$/.test(prev)) {
      return true;
    }
    if (/^[a-z(]/.test(next)) {
      return true;
    }
    if (/^(and|or|with|to|for|from|of|in|on)\b/i.test(next)) {
      return true;
    }
    return false;
  }

  function extractLineBasedClaims(cleanText) {
    var lines = normalizeLineBreaks(cleanText).split('\n');
    var claims = [];
    var current = '';

    for (var i = 0; i < lines.length; i += 1) {
      var raw = String(lines[i] || '').replace(/\r/g, '');
      var normalized = normalizeClaimText(raw);
      if (!normalized) {
        if (current) {
          claims.push(current);
          current = '';
        }
        continue;
      }

      if (!current) {
        current = normalized;
        continue;
      }

      if (isLikelyContinuationLine(current, raw)) {
        current = current + ' ' + normalized;
      } else {
        claims.push(current);
        current = normalized;
      }
    }

    if (current) {
      claims.push(current);
    }

    var output = [];
    for (var j = 0; j < claims.length; j += 1) {
      var claim = String(claims[j] || '').trim();
      if (claim.length >= 18) {
        output.push(claim);
      }
    }
    return output;
  }

  function extractFactClaims(selectedItem) {
    var fullText = selectedItem && typeof selectedItem.fullText === 'string'
      ? selectedItem.fullText
      : '';
    var cleanText = normalizeLineBreaks(fullText).trim();
    if (!cleanText) {
      return [selectedItem && selectedItem.label ? selectedItem.label : 'Selected item'];
    }

    var bulletClaims = extractBulletClaims(cleanText);
    if (bulletClaims.length >= 1) {
      return dedupeClaimTexts(bulletClaims).slice(0, 6);
    }

    var lineClaims = extractLineBasedClaims(cleanText);
    if (lineClaims.length >= 2) {
      return dedupeClaimTexts(lineClaims).slice(0, 6);
    }

    var sentenceParts = cleanText.match(/[^.!?;]+[.!?;]?/g) || [];
    var sentenceClaims = [];
    for (var i = 0; i < sentenceParts.length; i += 1) {
      var sentence = normalizeClaimText(String(sentenceParts[i] || '').replace(/[.!?;]+$/, '').trim());
      if (sentence.length >= 22) {
        sentenceClaims.push(sentence);
      }
    }
    if (sentenceClaims.length >= 2) {
      return dedupeClaimTexts(sentenceClaims).slice(0, 6);
    }

    if (cleanText.length > 130 && cleanText.indexOf(',') >= 0) {
      var clauses = cleanText.split(/\s*,\s*/);
      var clauseClaims = [];
      for (var j = 0; j < clauses.length; j += 1) {
        var clause = normalizeClaimText(clauses[j]);
        if (clause.length >= 22) {
          clauseClaims.push(clause);
        }
      }
      if (clauseClaims.length >= 2) {
        return dedupeClaimTexts(clauseClaims).slice(0, 6);
      }
    }

    var fallbackLabel = stripTrailingReferenceMarkers(
      selectedItem && selectedItem.label ? selectedItem.label : cleanText.slice(0, 120)
    );
    return [fallbackLabel || 'Selected item'];
  }

  function inferFocusedClaimFromSelection(selectedItem, slideContext, extractedClaims) {
    var rawSelection = slideContext && slideContext.selection && typeof slideContext.selection.text === 'string'
      ? slideContext.selection.text
      : '';
    var normalizedSelection = normalizeLineBreaks(rawSelection).trim();
    if (!normalizedSelection) {
      return null;
    }

    var lines = normalizedSelection.split(/\n+/);
    var parts = [];
    for (var i = 0; i < lines.length; i += 1) {
      var normalized = normalizeClaimText(lines[i]);
      if (normalized && normalized.length >= 12) {
        parts.push(normalized);
      }
    }
    var focusedClaim = dedupeClaimTexts(parts).join(' ').trim();
    if (!focusedClaim || focusedClaim.length < 12) {
      return null;
    }

    var fullText = selectedItem && typeof selectedItem.fullText === 'string'
      ? normalizeLineBreaks(selectedItem.fullText).replace(/\s+/g, ' ').trim()
      : '';
    if (fullText && focusedClaim.length >= Math.max(32, Math.floor(fullText.length * 0.8))) {
      return null;
    }

    var focusedKey = normalizedClaimKey(focusedClaim);
    var fullKey = normalizedClaimKey(fullText);
    if (fullKey && focusedKey && fullKey.indexOf(focusedKey) < 0) {
      return null;
    }

    var claims = Array.isArray(extractedClaims) ? extractedClaims : [];
    var bestIndex = -1;
    for (var j = 0; j < claims.length; j += 1) {
      var claimKey = normalizedClaimKey(claims[j]);
      if (!claimKey) {
        continue;
      }
      if (
        claimKey === focusedKey ||
        claimKey.indexOf(focusedKey) >= 0 ||
        focusedKey.indexOf(claimKey) >= 0
      ) {
        bestIndex = j;
        break;
      }
    }

    return {
      claimText: focusedClaim,
      claimIndex: bestIndex
    };
  }

  function hasReferenceMarkerNearPosition(text, position) {
    var full = normalizeLineBreaks(text);
    var start = Number(position);
    if (!Number.isFinite(start) || start < 0) {
      return false;
    }
    if (start > full.length) {
      return false;
    }
    var tail = full.slice(start, start + 24);
    if (/^\s*(?:[.,;:!?)]\s*)*(?:[\u2070\u00B9\u00B2\u00B3\u2074\u2075\u2076\u2077\u2078\u2079]+|\[\d{1,3}\])/.test(tail)) {
      return true;
    }

    var headStart = Math.max(0, start - 14);
    var head = full.slice(headStart, start);
    return /(?:[\u2070\u00B9\u00B2\u00B3\u2074\u2075\u2076\u2077\u2078\u2079]+|\[\d{1,3}\])\s*$/.test(head);
  }

  function filterClaimsNeedingReference(selectedItem, claimsWithIndex) {
    var list = Array.isArray(claimsWithIndex) ? claimsWithIndex : [];
    var fullText = selectedItem && typeof selectedItem.fullText === 'string'
      ? normalizeLineBreaks(selectedItem.fullText)
      : '';
    if (!fullText) {
      return list.slice(0, 6);
    }

    var pending = [];
    for (var i = 0; i < list.length; i += 1) {
      var claim = list[i] || {};
      var claimText = String(claim.claimText || '').trim();
      var claimIndex = Number(claim.claimIndex);
      if (!claimText) {
        continue;
      }
      if (!Number.isFinite(claimIndex)) {
        claimIndex = i;
      }

      var insertion = findClaimInsertionIndex(fullText, claimText, claimIndex);
      if (insertion < 0 || !hasReferenceMarkerNearPosition(fullText, insertion)) {
        pending.push({
          claimText: claimText,
          claimIndex: claimIndex
        });
      }
      if (pending.length >= 6) {
        break;
      }
    }
    return pending;
  }

  function requestReferencesForClaims(claims, summarizedContext) {
    var list = Array.isArray(claims) ? claims.slice(0, 6) : [];
    return list.reduce(function (chain, claimItem, listIndex) {
      return chain.then(function (results) {
        var claim = '';
        var claimIndex = listIndex;
        if (typeof claimItem === 'string') {
          claim = claimItem;
        } else if (claimItem && typeof claimItem === 'object') {
          claim = String(claimItem.claimText || claimItem.text || '');
          var explicitIndex = Number(claimItem.claimIndex);
          if (Number.isFinite(explicitIndex)) {
            claimIndex = Math.floor(explicitIndex);
          }
        }
        claim = claim.trim();
        if (!claim) {
          return results;
        }
        return requestReferenceSuggestion({
          itemText: claim,
          slideContext: summarizedContext
        }).then(function (payload) {
          var reference = payload && payload.reference ? payload.reference : null;
          var url = reference && typeof reference.url === 'string' ? reference.url : '';
          if (!url) {
            throw new Error('No URL returned');
          }
          var title = reference && typeof reference.title === 'string' && reference.title.trim()
            ? reference.title.trim()
            : claim;
          results.push({
            claimText: claim,
            claimIndex: claimIndex,
            linkText: title,
            sourceUrl: url,
            reachable: reference && reference.reachable !== false,
            usedFallback: false
          });
          return results;
        }).catch(function (_error) {
          results.push({
            claimText: claim,
            claimIndex: claimIndex,
            linkText: claim,
            sourceUrl: buildFallbackReferenceUrl(claim),
            reachable: false,
            usedFallback: true
          });
          return results;
        });
      });
    }, Promise.resolve([]));
  }

  function isHyperlinkApiSupported() {
    if (!Office || !Office.context || !Office.context.requirements) {
      return false;
    }
    if (typeof Office.context.requirements.isSetSupported !== 'function') {
      return false;
    }
    return Office.context.requirements.isSetSupported('PowerPointApi', '1.10');
  }

  function resolveSlideSize(slideContext, shapes) {
    var fallback = { w: 960, h: 540 };
    var known = slideContext && slideContext.slide && slideContext.slide.size ? slideContext.slide.size : null;
    var width = Number(known && known.w);
    var height = Number(known && known.h);
    if (Number.isFinite(width) && width > 100 && Number.isFinite(height) && height > 100) {
      return { w: width, h: height };
    }

    var maxRight = fallback.w;
    var maxBottom = fallback.h;
    var shapeItems = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    for (var i = 0; i < shapeItems.length; i += 1) {
      var shape = shapeItems[i];
      var right = Number(shape.left) + Number(shape.width);
      var bottom = Number(shape.top) + Number(shape.height);
      if (Number.isFinite(right)) maxRight = Math.max(maxRight, right);
      if (Number.isFinite(bottom)) maxBottom = Math.max(maxBottom, bottom);
    }
    return { w: maxRight, h: maxBottom };
  }

  function countReferenceShapes(shapes) {
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    var count = 0;
    for (var i = 0; i < items.length; i += 1) {
      var name = String(items[i] && items[i].name ? items[i].name : '');
      var isOldRef = name.indexOf(REFERENCE_SHAPE_PREFIX) === 0;
      var isListRef = name.indexOf(REFERENCE_LIST_SHAPE_PREFIX) === 0;
      if (isOldRef || isListRef) {
        count += 1;
      }
    }
    return count;
  }

  function parseNumberAfterPrefix(value, prefix) {
    var text = String(value || '');
    if (text.indexOf(prefix) !== 0) {
      return null;
    }
    var raw = text.slice(prefix.length);
    var num = Number(raw);
    if (!Number.isFinite(num) || num <= 0) {
      return null;
    }
    return Math.floor(num);
  }

  function getExistingReferenceNumbers(shapes) {
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    var numbers = [];
    for (var i = 0; i < items.length; i += 1) {
      var name = String(items[i] && items[i].name ? items[i].name : '');
      var n = parseNumberAfterPrefix(name, REFERENCE_LIST_SHAPE_PREFIX);
      if (n !== null) {
        numbers.push(n);
      }
    }
    return numbers;
  }

  function extractReferenceNumbersFromText(text) {
    var content = String(text || '');
    var regex = /(?:^|\s)(\d+)\.\s/g;
    var match;
    var numbers = [];
    while ((match = regex.exec(content)) !== null) {
      var n = Number(match[1]);
      if (Number.isFinite(n) && n > 0) {
        numbers.push(Math.floor(n));
      }
    }
    return numbers;
  }

  function getSourceBoxText(shape) {
    if (!shape || !shape.textFrame || !shape.textFrame.textRange) {
      return '';
    }
    return normalizeLineBreaks(shape.textFrame.textRange.text || '');
  }

  function stripSourcesPrefix(text) {
    var value = String(text || '').trim();
    if (!value) {
      return '';
    }
    return value.replace(/^sources:\s*/i, '').trim();
  }

  function squeezeInlineText(text) {
    return String(text || '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function estimateSourceBoxHeight(text, width, fontSize) {
    var safeText = String(text || '');
    var w = Math.max(180, Number(width || 180));
    var fs = Math.max(8, Number(fontSize || 9));
    var charsPerLine = Math.max(26, Math.floor((w - 12) / (fs * 0.52)));
    var lines = Math.max(1, Math.ceil(safeText.length / charsPerLine));
    var lineHeight = Math.max(12, Math.ceil(fs * 1.3));
    return Math.max(14, lines * lineHeight + 6);
  }

  function getSourceHostLabel(url) {
    try {
      var hostname = new URL(String(url || '')).hostname || '';
      return hostname.replace(/^www\./i, '');
    } catch (_error) {
      return '';
    }
  }

  function buildSourceEntryText(referenceNumber, linkText, sourceUrl) {
    var label = squeezeInlineText(String(linkText || 'Reference')).slice(0, 64);
    var host = getSourceHostLabel(sourceUrl);
    if (host) {
      return String(referenceNumber) + '. ' + label + ' [' + host + ']';
    }
    return String(referenceNumber) + '. ' + label;
  }

  function parseSourceEntries(text) {
    var body = stripSourcesPrefix(text);
    if (!body) {
      return [];
    }
    var parts = body.split(/\s+\|\s+/);
    var entries = [];
    for (var i = 0; i < parts.length; i += 1) {
      var raw = String(parts[i] || '').trim();
      if (!raw) {
        continue;
      }
      var match = raw.match(/^(\d+)\.\s*(.+)$/);
      if (!match) {
        continue;
      }
      var number = Number(match[1]);
      if (!Number.isFinite(number) || number <= 0) {
        continue;
      }
      entries.push({
        number: Math.floor(number),
        label: squeezeInlineText(match[2]).slice(0, 90)
      });
    }
    entries.sort(function (a, b) {
      return a.number - b.number;
    });
    return entries;
  }

  function composeSourcesText(entries) {
    var items = Array.isArray(entries) ? entries : [];
    if (!items.length) {
      return 'Sources:';
    }
    var parts = [];
    for (var i = 0; i < items.length; i += 1) {
      var entry = items[i] || {};
      if (!entry.number || !entry.label) {
        continue;
      }
      parts.push(String(entry.number) + '. ' + String(entry.label));
    }
    if (!parts.length) {
      return 'Sources:';
    }
    return 'Sources: ' + parts.join(' | ');
  }

  function getSlideSourceLinkMap(slideKey) {
    var uiState = window.PPTAutomation && window.PPTAutomation.uiState
      ? window.PPTAutomation.uiState
      : null;
    if (!uiState) {
      return {};
    }
    if (!uiState.sourceLinksBySlide || typeof uiState.sourceLinksBySlide !== 'object') {
      uiState.sourceLinksBySlide = {};
    }
    if (!uiState.sourceLinksBySlide[slideKey] || typeof uiState.sourceLinksBySlide[slideKey] !== 'object') {
      uiState.sourceLinksBySlide[slideKey] = {};
    }
    return uiState.sourceLinksBySlide[slideKey];
  }

  function applyEntryHyperlinks(shape, sourcesText, entries, sourceLinkMap) {
    if (
      !shape ||
      !shape.textFrame ||
      !shape.textFrame.textRange ||
      typeof shape.textFrame.textRange.getSubstring !== 'function'
    ) {
      return false;
    }

    var linked = false;
    var cursor = 0;
    for (var i = 0; i < entries.length; i += 1) {
      var entry = entries[i] || {};
      var numberKey = String(entry.number || '');
      var url = sourceLinkMap && typeof sourceLinkMap[numberKey] === 'string'
        ? sourceLinkMap[numberKey]
        : '';
      if (!url) {
        continue;
      }
      var segment = String(entry.number) + '. ' + String(entry.label || '');
      var start = sourcesText.indexOf(segment, cursor);
      if (start < 0) {
        continue;
      }
      cursor = start + segment.length;
      var label = String(entry.label || '');
      var labelStart = start + String(entry.number).length + 2;
      var labelLength = label.length;
      if (labelLength <= 0) {
        continue;
      }
      try {
        var range = shape.textFrame.textRange.getSubstring(labelStart, labelLength);
        if (range && typeof range.setHyperlink === 'function') {
          range.setHyperlink({
            address: url,
            screenTip: 'Open source'
          });
          linked = true;
        }
      } catch (_error) {
        // ignore per-entry hyperlink failure
      }
    }
    return linked;
  }

  function colorSourceEntryLabels(shape, sourcesText, entries, markerColor) {
    if (
      !shape ||
      !shape.textFrame ||
      !shape.textFrame.textRange ||
      typeof shape.textFrame.textRange.getSubstring !== 'function'
    ) {
      return;
    }

    var safeColor = normalizeHexColor(markerColor) || '#C62828';
    var cursor = 0;
    for (var i = 0; i < entries.length; i += 1) {
      var entry = entries[i] || {};
      if (!entry.number || !entry.label) {
        continue;
      }
      var prefix = String(entry.number) + '.';
      var segment = prefix + ' ' + String(entry.label);
      var start = sourcesText.indexOf(segment, cursor);
      if (start < 0) {
        continue;
      }
      cursor = start + segment.length;
      try {
        var prefixRange = shape.textFrame.textRange.getSubstring(start, prefix.length);
        if (prefixRange && prefixRange.font) {
          prefixRange.font.color = safeColor;
        }
      } catch (_error) {
        // ignore per-entry styling failure
      }
    }
  }

  function cleanLegacyReferenceShapes(shapes) {
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    for (var i = 0; i < items.length; i += 1) {
      var shape = items[i];
      var name = String(shape && shape.name ? shape.name : '');
      var isLegacyList = name.indexOf(REFERENCE_LIST_SHAPE_PREFIX) === 0;
      var isLegacyHeader = name === REFERENCE_HEADER_SHAPE_NAME;
      var isLegacyOld = name.indexOf(REFERENCE_SHAPE_PREFIX) === 0;
      if (isLegacyList || isLegacyHeader || isLegacyOld) {
        try {
          if (shape && typeof shape.delete === 'function') {
            shape.delete();
          }
        } catch (_error) {
          // ignore cleanup failures
        }
      }
    }
  }

  function collectLegacyReferenceEntries(shapes) {
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    var map = {};
    for (var i = 0; i < items.length; i += 1) {
      var shape = items[i];
      var name = String(shape && shape.name ? shape.name : '');
      var number = parseNumberAfterPrefix(name, REFERENCE_LIST_SHAPE_PREFIX);
      if (number === null) {
        number = parseNumberAfterPrefix(name, REFERENCE_SHAPE_PREFIX);
      }
      if (number === null) {
        continue;
      }

      var rawText = '';
      if (shape && shape.textFrame && shape.textFrame.textRange) {
        rawText = normalizeLineBreaks(shape.textFrame.textRange.text || '');
      }
      var compact = squeezeInlineText(rawText);
      if (!compact) {
        continue;
      }
      var match = compact.match(/^\d+\.\s*(.+)$/);
      var label = match ? match[1] : compact;
      if (!label) {
        continue;
      }
      map[String(number)] = label.slice(0, 90);
    }

    var entries = [];
    var keys = Object.keys(map);
    for (var j = 0; j < keys.length; j += 1) {
      var n = Number(keys[j]);
      if (!Number.isFinite(n) || n <= 0) {
        continue;
      }
      entries.push({
        number: Math.floor(n),
        label: map[keys[j]]
      });
    }
    entries.sort(function (a, b) {
      return a.number - b.number;
    });
    return entries;
  }

  function estimateSourceItemWidth(text, fontSize) {
    var size = Number(fontSize || 9);
    var length = String(text || '').length;
    return Math.max(56, Math.ceil(length * size * 0.5 + 16));
  }

  function collectReferenceListShapes(shapes, extraShape) {
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    var collected = [];
    for (var i = 0; i < items.length; i += 1) {
      var shape = items[i];
      var name = String(shape && shape.name ? shape.name : '');
      var n = parseNumberAfterPrefix(name, REFERENCE_LIST_SHAPE_PREFIX);
      if (n !== null) {
        collected.push({ shape: shape, number: n });
      }
    }
    if (extraShape) {
      var extraName = String(extraShape.name || '');
      var extraNumber = parseNumberAfterPrefix(extraName, REFERENCE_LIST_SHAPE_PREFIX);
      if (extraNumber !== null) {
        collected.push({ shape: extraShape, number: extraNumber });
      } else {
        collected.push({ shape: extraShape, number: collected.length + 1 });
      }
    }
    collected.sort(function (a, b) {
      return a.number - b.number;
    });
    return collected;
  }

  function layoutReferenceListInline(entries, startX, baseY, maxRight, rowHeight, rowGap, itemGap) {
    var x = startX;
    var y = baseY;
    for (var i = 0; i < entries.length; i += 1) {
      var shape = entries[i] && entries[i].shape ? entries[i].shape : null;
      if (!shape) {
        continue;
      }
      var width = Math.max(56, Math.min(220, Number(shape.width || 90)));
      shape.width = width;
      if (x + width > maxRight) {
        x = startX;
        y = y - (rowHeight + rowGap);
      }
      if (y < 2) {
        y = 2;
      }
      shape.left = x;
      shape.top = y;
      x += width + itemGap;
    }
  }

  function findShapeById(shapes, shapeId) {
    if (!shapeId) {
      return null;
    }
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    for (var i = 0; i < items.length; i += 1) {
      if (items[i] && items[i].id === shapeId) {
        return items[i];
      }
    }
    return null;
  }

  function findShapeByName(shapes, shapeName) {
    if (!shapeName) {
      return null;
    }
    var items = shapes && Array.isArray(shapes.items) ? shapes.items : [];
    for (var i = 0; i < items.length; i += 1) {
      if (items[i] && items[i].name === shapeName) {
        return items[i];
      }
    }
    return null;
  }

  function toSuperscriptNumber(value) {
    var digits = String(Math.max(1, Number(value || 1)));
    var map = {
      '0': '\u2070',
      '1': '\u00B9',
      '2': '\u00B2',
      '3': '\u00B3',
      '4': '\u2074',
      '5': '\u2075',
      '6': '\u2076',
      '7': '\u2077',
      '8': '\u2078',
      '9': '\u2079'
    };
    var out = '';
    for (var i = 0; i < digits.length; i += 1) {
      var ch = digits.charAt(i);
      if (!map[ch]) {
        return '[' + digits + ']';
      }
      out += map[ch];
    }
    return out || ('[' + digits + ']');
  }

  function normalizeHexColor(value) {
    var raw = String(value || '').trim();
    if (!raw) {
      return null;
    }
    var hex = raw.charAt(0) === '#' ? raw.slice(1) : raw;
    if (!/^[0-9a-fA-F]{6}$/.test(hex)) {
      return null;
    }
    return '#' + hex.toUpperCase();
  }

  function isRedLikeColor(value) {
    var hex = normalizeHexColor(value);
    if (!hex) {
      return false;
    }
    var r = parseInt(hex.slice(1, 3), 16);
    var g = parseInt(hex.slice(3, 5), 16);
    var b = parseInt(hex.slice(5, 7), 16);
    return r >= 150 && g <= 105 && b <= 105 && r >= (g + 35) && r >= (b + 35);
  }

  function pickReferenceMarkerColor(mainTextColor) {
    if (isRedLikeColor(mainTextColor)) {
      return '#1F4E79';
    }
    return '#C62828';
  }

  function resolveReferenceMarkerColorForShape(context, shape) {
    if (!shape || !shape.textFrame || !shape.textFrame.textRange) {
      return Promise.resolve('#C62828');
    }

    try {
      shape.textFrame.textRange.load('font/color');
    } catch (_error) {
      return Promise.resolve('#C62828');
    }

    return context.sync()
      .then(function () {
        var mainTextColor = shape.textFrame.textRange && shape.textFrame.textRange.font
          ? shape.textFrame.textRange.font.color
          : null;
        return pickReferenceMarkerColor(mainTextColor);
      })
      .catch(function () {
        return '#C62828';
      });
  }

  function colorAllReferenceMarkers(shape, text, markerColor) {
    if (
      !shape ||
      !shape.textFrame ||
      !shape.textFrame.textRange ||
      typeof shape.textFrame.textRange.getSubstring !== 'function'
    ) {
      return;
    }

    try {
      var regex = /[\u2070\u00B9\u00B2\u00B3\u2074\u2075\u2076\u2077\u2078\u2079]+|\[\d+\]/g;
      var sourceText = String(text || '');
      var match;
      while ((match = regex.exec(sourceText)) !== null) {
        var start = match.index;
        var len = match[0].length;
        if (len <= 0) {
          continue;
        }
        var markerRange = shape.textFrame.textRange.getSubstring(start, len);
        if (markerRange && markerRange.font) {
          markerRange.font.color = markerColor;
        }
      }
    } catch (_error) {
      // Best effort only.
    }
  }

  function normalizedClaimKey(value) {
    return stripBulletPrefix(value)
      .toLowerCase()
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function getFactLineEndPositions(text) {
    var full = normalizeLineBreaks(text);
    var lines = full.split('\n');
    var positions = [];
    var offset = 0;
    for (var i = 0; i < lines.length; i += 1) {
      var line = lines[i];
      var cleaned = stripBulletPrefix(line);
      if (cleaned && cleaned.length >= 18) {
        var end = offset + line.length;
        while (end > offset && /\s/.test(full.charAt(end - 1))) {
          end -= 1;
        }
        positions.push(end);
      }
      offset += line.length;
      if (i < lines.length - 1) {
        offset += 1;
      }
    }
    return positions;
  }

  function getBulletLineEndPositions(text) {
    var full = normalizeLineBreaks(text);
    var lines = full.split('\n');
    var positions = [];
    var offset = 0;
    var bulletPattern = /^\s*[\u2022\u25CF\u25E6\u25AA\u25AB\-*]+\s+/;
    for (var i = 0; i < lines.length; i += 1) {
      var line = lines[i];
      if (bulletPattern.test(line)) {
        var end = offset + line.length;
        while (end > offset && /\s/.test(full.charAt(end - 1))) {
          end -= 1;
        }
        positions.push(end);
      }
      offset += line.length;
      if (i < lines.length - 1) {
        offset += 1;
      }
    }
    return positions;
  }

  function findClaimInsertionIndex(text, claimText, claimIndex) {
    var full = normalizeLineBreaks(text);
    var claim = String(claimText || '').replace(/\s+/g, ' ').trim();
    if (!full || !claim) {
      return -1;
    }

    try {
      var spacedPattern = escapeRegExp(claim).replace(/\s+/g, '\\s+');
      var directRegex = new RegExp(spacedPattern, 'i');
      var regexMatch = directRegex.exec(full);
      if (regexMatch && typeof regexMatch.index === 'number') {
        return regexMatch.index + regexMatch[0].length;
      }
    } catch (_error) {
      // fall through to simple matching heuristics
    }

    var lowerFull = full.toLowerCase();
    var lowerClaim = claim.toLowerCase();
    var direct = lowerFull.indexOf(lowerClaim);
    if (direct >= 0) {
      return direct + claim.length;
    }

    var lines = full.split('\n');
    var offset = 0;
    var claimKey = normalizedClaimKey(claim);

    var ordinal = Number(claimIndex);
    if (Number.isFinite(ordinal) && ordinal >= 0) {
      var bulletEnds = getBulletLineEndPositions(full);
      if (bulletEnds.length > 0) {
        if (bulletEnds.length > ordinal) {
          return bulletEnds[ordinal];
        }
        return bulletEnds[bulletEnds.length - 1];
      }
    }

    for (var i = 0; i < lines.length; i += 1) {
      var line = lines[i];
      var cleanedLine = stripBulletPrefix(line).toLowerCase();
      var lineKey = normalizedClaimKey(line);
      if (
        cleanedLine && (
          cleanedLine.indexOf(lowerClaim) >= 0 ||
          lowerClaim.indexOf(cleanedLine) >= 0 ||
          (claimKey && lineKey && (lineKey.indexOf(claimKey) >= 0 || claimKey.indexOf(lineKey) >= 0))
        )
      ) {
        var lineEnd = offset + line.length;
        while (lineEnd > offset && /\s/.test(full.charAt(lineEnd - 1))) {
          lineEnd -= 1;
        }
        return lineEnd;
      }
      offset += line.length;
      if (i < lines.length - 1) {
        offset += 1;
      }
    }

    if (Number.isFinite(ordinal) && ordinal >= 0) {
      var ends = getFactLineEndPositions(full);
      if (ends.length > ordinal) {
        return ends[ordinal];
      }
      if (ends.length > 0) {
        return ends[ends.length - 1];
      }
    }

    return -1;
  }

  function appendInlineReferenceMarkerToShape(context, shape, referenceNumber, claimText, claimIndex, markerColorHint) {
    if (!shape || !shape.textFrame || !shape.textFrame.textRange) {
      return Promise.resolve(false);
    }

    var marker = toSuperscriptNumber(referenceNumber);
    var fallbackMarker = '[' + String(referenceNumber) + ']';
    var markerColor = normalizeHexColor(markerColorHint) || '#C62828';

    try {
      shape.textFrame.textRange.load('text,font/color');
    } catch (_error) {
      return Promise.resolve(false);
    }

    return context.sync()
      .then(function () {
        var currentRaw = String(shape.textFrame.textRange.text || '');
        var current = normalizeLineBreaks(currentRaw);
        if (!current.trim()) {
          return false;
        }
        var trimmed = current.replace(/\s+$/, '');
        if (!trimmed) {
          return false;
        }
        if (!normalizeHexColor(markerColorHint)) {
          var mainTextColor = shape.textFrame.textRange && shape.textFrame.textRange.font
            ? shape.textFrame.textRange.font.color
            : null;
          markerColor = pickReferenceMarkerColor(mainTextColor);
        }
        var insertionIndex = findClaimInsertionIndex(current, claimText, claimIndex);
        if (insertionIndex < 0 || insertionIndex > current.length) {
          insertionIndex = trimmed.length;
        }
        if (insertionIndex > 0 && current.charAt(insertionIndex - 1) === ' ' && marker.charAt(0) === '[') {
          insertionIndex -= 1;
        }

        if (hasReferenceMarkerNearPosition(current, insertionIndex)) {
          colorAllReferenceMarkers(shape, current, markerColor);
          return false;
        }

        var before = current.slice(0, insertionIndex).replace(/\s+$/, '');
        var after = current.slice(insertionIndex);
        var markerText = marker.charAt(0) === '[' ? (' ' + marker) : marker;

        if (before.slice(-marker.length) === marker || before.slice(-fallbackMarker.length) === fallbackMarker) {
          colorAllReferenceMarkers(shape, current, markerColor);
          return false;
        }
        var afterTrimmed = after.replace(/^\s+/, '');
        if (afterTrimmed.indexOf(marker) === 0 || afterTrimmed.indexOf(fallbackMarker) === 0) {
          colorAllReferenceMarkers(shape, current, markerColor);
          return false;
        }

        var nextText = before + markerText + after;
        shape.textFrame.textRange.text = nextText;
        colorAllReferenceMarkers(shape, nextText, markerColor);

        return context.sync().then(function () {
          return true;
        }, function () {
          return true;
        });
      })
      .catch(function () {
        return false;
      });
  }

  function resolveActiveSlide(context, slides) {
    if (!context || !context.presentation || !slides) {
      return Promise.resolve(null);
    }

    return Promise.resolve()
      .then(function () {
        if (typeof context.presentation.getSelectedSlides !== 'function') {
          return null;
        }
        var selectedSlides = context.presentation.getSelectedSlides();
        selectedSlides.load('items/id');
        return context.sync().then(function () {
          if (selectedSlides.items && selectedSlides.items.length > 0) {
            return selectedSlides.items[0];
          }
          return null;
        }, function () {
          return null;
        });
      })
      .then(function (selected) {
        if (selected) {
          return selected;
        }
        return slides.items && slides.items.length ? slides.items[0] : null;
      });
  }

  function insertReferenceAtBottom(payload) {
    var linkText = payload && payload.linkText ? payload.linkText : 'Source';
    var sourceUrl = payload && payload.sourceUrl ? payload.sourceUrl : '';
    var slideContext = payload && payload.slideContext ? payload.slideContext : {};
    var selectedShapeId = payload && payload.selectedShapeId ? payload.selectedShapeId : null;
    var claimText = payload && payload.claimText ? payload.claimText : '';
    var claimIndex = payload && typeof payload.claimIndex === 'number' ? payload.claimIndex : -1;

    return PowerPoint.run(function (context) {
      var slides = context.presentation.slides;
      slides.load('items');

      return context.sync()
        .then(function () {
          return resolveActiveSlide(context, slides);
        })
        .then(function (activeSlide) {
          if (!activeSlide) {
            throw new Error('Unable to resolve active slide.');
          }

          var shapes = activeSlide.shapes;
          shapes.load('items/id,items/name,items/left,items/top,items/width,items/height,items/textFrame/textRange/text');

          return context.sync().then(function () {
            var targetShape = findShapeById(shapes, selectedShapeId);
            return resolveReferenceMarkerColorForShape(context, targetShape).then(function (referenceMarkerColor) {
            var slideSize = resolveSlideSize(slideContext, shapes);
            var marginX = 20;
            var marginBottom = 8;
            var fontSize = 9;
            var slideKey = String(
              (slideContext && slideContext.slide && slideContext.slide.id) ||
              activeSlide.id ||
              'active'
            );
            var sourceLinkMap = getSlideSourceLinkMap(slideKey);
            var sourceBox = findShapeByName(shapes, REFERENCE_BOX_SHAPE_NAME);
            var existingText = getSourceBoxText(sourceBox);
            var existingEntries = parseSourceEntries(existingText);
            var legacyEntries = collectLegacyReferenceEntries(shapes);
            var legacyNumbers = getExistingReferenceNumbers(shapes);
            var knownNumbers = [];
            var entryMap = {};

            for (var i = 0; i < existingEntries.length; i += 1) {
              var existing = existingEntries[i] || {};
              if (!existing.number || !existing.label) {
                continue;
              }
              knownNumbers.push(existing.number);
              entryMap[String(existing.number)] = existing.label;
            }

            for (var j = 0; j < legacyEntries.length; j += 1) {
              var legacy = legacyEntries[j] || {};
              if (!legacy.number || !legacy.label) {
                continue;
              }
              knownNumbers.push(legacy.number);
              if (!entryMap[String(legacy.number)]) {
                entryMap[String(legacy.number)] = legacy.label;
              }
            }

            var fromTextNumbers = extractReferenceNumbersFromText(existingText);
            for (var k = 0; k < fromTextNumbers.length; k += 1) {
              knownNumbers.push(fromTextNumbers[k]);
            }
            for (var n = 0; n < legacyNumbers.length; n += 1) {
              knownNumbers.push(legacyNumbers[n]);
            }
            var mapKeys = Object.keys(sourceLinkMap);
            for (var m = 0; m < mapKeys.length; m += 1) {
              var mapNumber = Number(mapKeys[m]);
              if (Number.isFinite(mapNumber) && mapNumber > 0) {
                knownNumbers.push(Math.floor(mapNumber));
              }
            }

            var referenceNumber = knownNumbers.length ? (Math.max.apply(null, knownNumbers) + 1) : 1;
            var safeSourceUrl = sanitizeSourceUrl(sourceUrl) || sourceUrl;
            var entryText = buildSourceEntryText(referenceNumber, linkText, safeSourceUrl);
            var entryMatch = entryText.match(/^\d+\.\s*(.+)$/);
            var entryLabel = entryMatch ? entryMatch[1] : squeezeInlineText(String(linkText || 'Reference'));
            entryMap[String(referenceNumber)] = entryLabel.slice(0, 90);

            if (safeSourceUrl) {
              sourceLinkMap[String(referenceNumber)] = safeSourceUrl;
            }

            var entries = [];
            var entryNumbers = Object.keys(entryMap);
            for (var p = 0; p < entryNumbers.length; p += 1) {
              var number = Number(entryNumbers[p]);
              if (!Number.isFinite(number) || number <= 0) {
                continue;
              }
              entries.push({
                number: Math.floor(number),
                label: String(entryMap[entryNumbers[p]] || '').trim().slice(0, 90)
              });
            }
            entries.sort(function (a, b) {
              return a.number - b.number;
            });

            var sourcesText = composeSourcesText(entries);
            var sourceWidth = Math.max(240, slideSize.w - (marginX * 2));
            var sourceHeight = estimateSourceBoxHeight(sourcesText, sourceWidth, fontSize);
            var minTop = 2;
            var top = slideSize.h - marginBottom - sourceHeight;
            if (top < minTop) {
              top = minTop;
              sourceHeight = Math.max(14, slideSize.h - marginBottom - minTop);
            }

            if (!sourceBox) {
              sourceBox = activeSlide.shapes.addTextBox(sourcesText);
              try {
                sourceBox.name = REFERENCE_BOX_SHAPE_NAME;
              } catch (_error) {
                // ignore name assignment limitations
              }
            } else {
              sourceBox.textFrame.textRange.text = sourcesText;
            }

            sourceBox.left = marginX;
            sourceBox.top = top;
            sourceBox.width = sourceWidth;
            sourceBox.height = sourceHeight;
            if (sourceBox.textFrame && sourceBox.textFrame.textRange && sourceBox.textFrame.textRange.font) {
              sourceBox.textFrame.textRange.font.size = fontSize;
            }

            colorSourceEntryLabels(sourceBox, sourcesText, entries, referenceMarkerColor);

            var linked = false;
            if (isHyperlinkApiSupported()) {
              linked = applyEntryHyperlinks(sourceBox, sourcesText, entries, sourceLinkMap);
            }

            cleanLegacyReferenceShapes(shapes);

            return appendInlineReferenceMarkerToShape(
              context,
              targetShape,
              referenceNumber,
              claimText,
              claimIndex,
              referenceMarkerColor
            ).then(function (markerAdded) {
              return context.sync().then(function () {
                return { linked: linked, referenceNumber: referenceNumber, markerAdded: markerAdded };
              });
            });
            });
          });
        });
    });
  }

  function handleConfirmReferenceInput() {
    var labelInput = document.getElementById('referenceLabelInput');
    var urlInput = document.getElementById('referenceUrlInput');
    var label = labelInput ? String(labelInput.value || '').trim() : '';
    var rawUrl = urlInput ? String(urlInput.value || '').trim() : '';
    if (!label) {
      setReferenceOverlayError('Enter a short reference label.');
      return;
    }

    var sourceUrl = sanitizeSourceUrl(rawUrl);
    if (!sourceUrl) {
      setReferenceOverlayError('Enter a valid http(s) source URL.');
      return;
    }

    resolveReferenceDialog({
      label: label,
      sourceUrl: sourceUrl
    });
  }

  function handleCancelReferenceInput() {
    resolveReferenceDialog(null);
  }

  function handleApplyPlan() {
    var uiState = window.PPTAutomation.uiState;
    var plan = uiState.pendingPlan;

    if (!plan) {
      setStatus('Generate a preview first.');
      return;
    }

    if (!uiState.previewSession) {
      setStatus('Preview slide is not ready yet.');
      return;
    }

    setStatus('Accepting preview slide...');
    uiState.isApplyingPlan = true;
    suppressAutoRecommendations(AUTO_RECOMMEND_SUPPRESSION_MS);
    acceptPreviewSession()
      .then(function () {
        setStatus('Preview accepted. Undo is available.');
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Failed to accept preview');
      })
      .then(function () {
        uiState.isApplyingPlan = false;
        updateUndoRedoButtons();
      });
  }

  function handleRejectPlan() {
    var uiState = window.PPTAutomation.uiState;
    if (!uiState.previewSession) {
      clearPreviewState();
      setStatus('Preview rejected. No changes were applied.');
      updateUndoRedoButtons();
      return;
    }

    setStatus('Rejecting preview slide...');
    uiState.isApplyingPlan = true;
    suppressAutoRecommendations(AUTO_RECOMMEND_SUPPRESSION_MS);
    rejectPreviewSession()
      .then(function () {
        setStatus('Preview rejected. Original slide restored.');
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Failed to reject preview');
      })
      .then(function () {
        uiState.isApplyingPlan = false;
        updateUndoRedoButtons();
      });
  }

  function handleUndoAccepted() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.previewSession) {
      setStatus('Accept or reject the current preview first.');
      return;
    }
    if (!canUndoAcceptedVersion()) {
      setStatus('Nothing to undo.');
      return;
    }

    setStatus('Restoring previous accepted slide version...');
    uiState.isApplyingPlan = true;
    suppressAutoRecommendations(AUTO_RECOMMEND_SUPPRESSION_MS);
    restoreAcceptedSlideVersion('undo')
      .then(function () {
        setStatus('Restored the previous accepted slide version.');
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Undo failed.');
      })
      .then(function () {
        uiState.isApplyingPlan = false;
        updateUndoRedoButtons();
      });
  }

  function handleRedoAccepted() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.previewSession) {
      setStatus('Accept or reject the current preview first.');
      return;
    }
    if (!canRedoAcceptedVersion()) {
      setStatus('Nothing to redo.');
      return;
    }

    setStatus('Restoring the newer accepted slide version...');
    uiState.isApplyingPlan = true;
    suppressAutoRecommendations(AUTO_RECOMMEND_SUPPRESSION_MS);
    restoreAcceptedSlideVersion('redo')
      .then(function () {
        setStatus('Restored the newer accepted slide version.');
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Redo failed.');
      })
      .then(function () {
        uiState.isApplyingPlan = false;
        updateUndoRedoButtons();
      });
  }

  function handleAddReference() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.isAddingReference) return;
    if (uiState.previewSession) {
      setStatus('Accept or reject the current preview first.');
      return;
    }

    var addReferenceBtn = document.getElementById('addReferenceBtn');
    if (typeof window.PPTAutomation.collectSlideContext !== 'function') {
      setStatus('Slide collector is unavailable.');
      return;
    }

    uiState.isAddingReference = true;
    if (addReferenceBtn) addReferenceBtn.disabled = true;
    suppressAutoRecommendations(AUTO_RECOMMEND_SUPPRESSION_MS);
    updateUndoRedoButtons();

    function cleanup() {
      uiState.isAddingReference = false;
      if (addReferenceBtn) addReferenceBtn.disabled = false;
      updateUndoRedoButtons();
    }

    setStatus('Reading selected item...');
    Promise.resolve()
      .then(function () {
        return window.PPTAutomation.collectSlideContext();
      })
      .then(function (slideContext) {
        uiState.latestSlideContext = slideContext;

        var selected = inferSelectedItemLabel(slideContext);
        if (!selected) {
          throw new Error('Select a slide item first, then add a reference.');
        }

        var claims = extractFactClaims(selected);
        var claimItems = [];
        var focusedClaim = inferFocusedClaimFromSelection(selected, slideContext, claims);
        if (focusedClaim && focusedClaim.claimText) {
          claimItems.push({
            claimText: focusedClaim.claimText,
            claimIndex: focusedClaim.claimIndex
          });
        } else {
          for (var i = 0; i < claims.length; i += 1) {
            claimItems.push({
              claimText: claims[i],
              claimIndex: i
            });
          }
        }
        var pendingClaims = filterClaimsNeedingReference(selected, claimItems);
        if (!pendingClaims.length) {
          if (focusedClaim && focusedClaim.claimText) {
            throw new Error('Selected statement already has a reference label.');
          }
          throw new Error('All detected statements already have reference labels.');
        }
        var claimCount = pendingClaims.length;
        setStatus('AI is finding ' + String(claimCount) + ' reference(s)...');
        var summarized = buildSummaryContextWithImage(slideContext);
        return requestReferencesForClaims(pendingClaims, summarized).then(function (referenceItems) {
          if (!referenceItems || !referenceItems.length) {
            throw new Error('No references were generated for the selected item.');
          }

          return referenceItems.reduce(function (chain, item, index) {
            return chain.then(function (state) {
              setStatus('Adding reference ' + String(index + 1) + ' of ' + String(referenceItems.length) + '...');
              return insertReferenceAtBottom({
                linkText: item.linkText,
                sourceUrl: item.sourceUrl,
                slideContext: slideContext,
                selectedShapeId: selected.shapeId,
                claimText: item.claimText,
                claimIndex: item.claimIndex
              }).then(function (result) {
                state.push({ item: item, result: result || {} });
                return state;
              });
            });
          }, Promise.resolve([])).then(function (applied) {
            var addedCount = applied.length;
            var fallbackCount = 0;
            var unverifiedCount = 0;
            var hyperlinkUnsupported = false;
            for (var i = 0; i < applied.length; i += 1) {
              if (applied[i].item && applied[i].item.usedFallback) {
                fallbackCount += 1;
              }
              if (applied[i].item && applied[i].item.reachable === false && !applied[i].item.usedFallback) {
                unverifiedCount += 1;
              }
              if (applied[i].result && applied[i].result.linked === false) {
                hyperlinkUnsupported = true;
              }
            }

            var message = 'Added ' + String(addedCount) + ' reference(s).';
            if (fallbackCount > 0) {
              message += ' Fallback used for ' + String(fallbackCount) + '.';
            }
            if (unverifiedCount > 0) {
              message += ' ' + String(unverifiedCount) + ' URL(s) were not reachability-verified.';
            }
            if (hyperlinkUnsupported) {
              message += ' Host does not support clickable link API.';
            }
            setStatus(message);
          });
        });
      })
      .catch(function (error) {
        var message = (error && error.message) || 'Failed to add reference.';
        setStatus(message);
      })
      .then(function () {
        cleanup();
      }, function () {
        cleanup();
      });
  }

  function runRecommendationCycle(options) {
    var opts = options || {};
    var trigger = opts.trigger === 'idle' ? 'idle' : 'manual';
    var uiState = window.PPTAutomation.uiState;
    if (uiState.isRecommending) return Promise.resolve(false);
    if (uiState.previewSession) {
      setStatus('Accept or reject the current preview first.');
      return Promise.resolve(false);
    }
    if (trigger === 'idle' && !shouldAllowAutomaticRecommendation()) {
      return Promise.resolve(false);
    }

    var recEl = document.getElementById('recommendations');
    var previewEl = document.getElementById('planPreview');
    var recommendBtn = document.getElementById('recommendBtn');

    if (typeof window.PPTAutomation.collectSlideContext !== 'function') {
      setStatus('Slide collector is unavailable.');
      return Promise.resolve(false);
    }

    uiState.isRecommending = true;
    if (recommendBtn) recommendBtn.disabled = true;
    clearPendingAutoRecommendation();
    updateUndoRedoButtons();

    setStatus(trigger === 'idle' ? 'Typing paused. Reading slide...' : 'Reading slide...');
    if (recEl) recEl.innerHTML = '';
    if (previewEl) previewEl.textContent = '';
    hidePreviewOverlay();
    uiState.latestPlan = null;
    uiState.pendingPlan = null;
    uiState.latestSlideContext = null;

    function cleanupRecommendState() {
      uiState.isRecommending = false;
      if (recommendBtn) recommendBtn.disabled = false;
      updateUndoRedoButtons();
    }

    return Promise.resolve()
      .then(function () {
        return window.PPTAutomation.collectSlideContext();
      })
      .then(function (slideContext) {
        uiState.latestSlideContext = slideContext;
        uiState.lastRecommendationSignature = opts.activitySignature || buildRecommendationActivitySignature(slideContext);
        uiState.lastRecommendationAt = nowMs();
        var summarized = buildSummaryContextWithImage(slideContext);
        var userPrompt = buildAutoPromptFromContext(slideContext);
        setStatus(trigger === 'idle' ? 'Typing paused. Generating recommendations...' : 'Generating recommendations...');
        return fetch('/api/recommendations', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ userPrompt: userPrompt, slideContext: summarized })
        }).then(function (response) {
          if (!response.ok) {
            return response.text().then(function (t) {
              throw new Error('Recommendations failed: ' + t);
            });
          }
          return response.json().then(function (payload) {
            return { payload: payload, summarized: summarized, userPrompt: userPrompt };
          });
        });
      })
      .then(function (data) {
        var payload = data.payload || {};
        var summarized = data.summarized || {};
        var userPrompt = data.userPrompt || '';
        var recommendations = Array.isArray(payload.recommendations) ? payload.recommendations : [];

        renderRecommendations(recommendations, function (recommendation) {
          setStatus('Generating execution plan...');
          requestPlan({
            selectedRecommendation: recommendation,
            userPrompt: userPrompt,
            slideContext: summarized
          })
            .then(function (planPayload) {
              var plan = planPayload && planPayload.plan ? planPayload.plan : null;
              if (previewEl) previewEl.textContent = JSON.stringify(plan, null, 2);
              if (!plan) {
                throw new Error('Plan generation returned no plan.');
              }
              setStatus('Creating live slide preview...');
              uiState.isApplyingPlan = true;
              suppressAutoRecommendations(AUTO_RECOMMEND_SUPPRESSION_MS);
              return createPlanPreviewOnDuplicateSlide(plan, summarized)
                .then(function (previewResult) {
                  var applyResult = previewResult && previewResult.applyResult ? previewResult.applyResult : {};
                  uiState.latestPlan = plan;
                  uiState.pendingPlan = plan;
                  showPreviewOverlay(plan, {
                    previewMessage: 'A duplicated preview slide is now selected in PowerPoint. Review it, then accept or reject here.',
                    runtimeWarnings: Array.isArray(applyResult.warnings) ? applyResult.warnings : []
                  });
                  var appliedCount = typeof applyResult.appliedCount === 'number' ? applyResult.appliedCount : 0;
                  setStatus('Preview ready on a duplicated slide. ' + String(appliedCount) + ' operation(s) applied.');
                })
                .then(function () {
                  uiState.isApplyingPlan = false;
                  updateUndoRedoButtons();
                }, function (error) {
                  uiState.isApplyingPlan = false;
                  updateUndoRedoButtons();
                  throw error;
                });
            })
            .catch(function (error) {
              setStatus((error && error.message) || 'Plan generation failed');
            });
        });

        setStatus((trigger === 'idle' ? 'Auto-loaded ' : 'Loaded ') + recommendations.length + ' recommendation(s).');
        return true;
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Unexpected error');
        return false;
      })
      .then(function () {
        cleanupRecommendState();
      }, function () {
        cleanupRecommendState();
      });
  }

  function handleRecommend() {
    return runRecommendationCycle({ trigger: 'manual' });
  }

  function buildAutoPromptFromContext(slideContext) {
    var objects = slideContext && Array.isArray(slideContext.objects) ? slideContext.objects : [];
    var title = inferSlideTitle(objects);
    var snippets = [];
    var i;
    for (i = 0; i < objects.length && snippets.length < 8; i += 1) {
      var text = objects[i] && typeof objects[i].text === 'string' ? objects[i].text.trim() : '';
      if (text) snippets.push(text);
    }
    var contentSummary = snippets.length ? snippets.join(' | ').slice(0, 800) : 'No detailed body content found.';
    var selected = slideContext && slideContext.selection && Array.isArray(slideContext.selection.shapeIds)
      ? slideContext.selection.shapeIds
      : [];
    var selectionSummary = selected.length ? ('Selected shape IDs: ' + selected.join(', ')) : 'No shapes selected.';
    var fonts = slideContext && slideContext.themeHints && Array.isArray(slideContext.themeHints.fonts)
      ? slideContext.themeHints.fonts.slice(0, 4)
      : [];
    var colors = slideContext && slideContext.themeHints && Array.isArray(slideContext.themeHints.colors)
      ? slideContext.themeHints.colors.slice(0, 6)
      : [];
    var styleSummary = 'Fonts: ' + (fonts.length ? fonts.join(', ') : 'unknown') +
      '; Colors: ' + (colors.length ? colors.join(', ') : 'unknown');

    return [
      'You are helping improve one PowerPoint slide.',
      '',
      'Goal:',
      'Predict the likely user intent and complete the slide toward that intended final state.',
      '',
      'Context:',
      '- Slide title: ' + title,
      '- Current content summary: ' + contentSummary,
      '- Selected object(s): ' + selectionSummary,
      '- Theme/style hints: ' + styleSummary,
      '',
      'Instructions:',
      '1. Infer the primary intent and desired final slide outcome.',
      '2. Act as a completion engine: fill missing sections/placeholders with ready-to-use content.',
      '3. Prioritize presentation-level quality: visual hierarchy, spacing, contrast, and readability.',
      '4. Propose the best output format (list, table, chart, image, smartart, or layout-improvement).',
      '5. Include at least one formatting/layout recommendation when readability or hierarchy can improve.',
      '6. Preserve existing design system and theme style (fonts/colors/spacing).',
      '7. Keep recommendations concise, high-confidence, and presentation-ready.'
    ].join('\\n');
  }

  function inferSlideTitle(objects) {
    var i;
    var candidates = [];
    for (i = 0; i < objects.length; i += 1) {
      var obj = objects[i] || {};
      var text = typeof obj.text === 'string' ? obj.text.trim() : '';
      if (!text) continue;
      var name = String(obj.name || '').toLowerCase();
      var top = Number.POSITIVE_INFINITY;
      if (Array.isArray(obj.bbox) && obj.bbox.length > 1) {
        top = Number(obj.bbox[1]);
      }
      candidates.push({ text: text, name: name, top: top });
    }

    for (i = 0; i < candidates.length; i += 1) {
      if (candidates[i].name.indexOf('title') >= 0) {
        return candidates[i].text.slice(0, 180);
      }
    }

    candidates.sort(function (a, b) {
      return a.top - b.top;
    });
    return candidates.length ? candidates[0].text.slice(0, 180) : 'Untitled slide';
  }

  function wireUiHandlers() {
    if (window.PPTAutomation.uiState.handlersWired) return;
    window.PPTAutomation.uiState.handlersWired = true;

    var recommendBtn = document.getElementById('recommendBtn');
    var addReferenceBtn = document.getElementById('addReferenceBtn');
    var undoAcceptedBtn = document.getElementById('undoAcceptedBtn');
    var redoAcceptedBtn = document.getElementById('redoAcceptedBtn');
    var confirmApplyBtn = document.getElementById('confirmApplyBtn');
    var rejectApplyBtn = document.getElementById('rejectApplyBtn');
    var confirmReferenceBtn = document.getElementById('confirmReferenceBtn');
    var cancelReferenceBtn = document.getElementById('cancelReferenceBtn');
    var referenceLabelInput = document.getElementById('referenceLabelInput');
    var referenceUrlInput = document.getElementById('referenceUrlInput');

    if (recommendBtn) recommendBtn.addEventListener('click', handleRecommend);
    if (addReferenceBtn) addReferenceBtn.addEventListener('click', handleAddReference);
    if (undoAcceptedBtn) undoAcceptedBtn.addEventListener('click', handleUndoAccepted);
    if (redoAcceptedBtn) redoAcceptedBtn.addEventListener('click', handleRedoAccepted);
    if (confirmApplyBtn) confirmApplyBtn.addEventListener('click', handleApplyPlan);
    if (rejectApplyBtn) rejectApplyBtn.addEventListener('click', handleRejectPlan);
    if (confirmReferenceBtn) confirmReferenceBtn.addEventListener('click', handleConfirmReferenceInput);
    if (cancelReferenceBtn) cancelReferenceBtn.addEventListener('click', handleCancelReferenceInput);
    if (referenceLabelInput) {
      referenceLabelInput.addEventListener('keydown', function (event) {
        if (event && event.key === 'Enter') {
          event.preventDefault();
          if (referenceUrlInput && typeof referenceUrlInput.focus === 'function') {
            referenceUrlInput.focus();
          }
        }
      });
    }
    if (referenceUrlInput) {
      referenceUrlInput.addEventListener('keydown', function (event) {
        if (event && event.key === 'Enter') {
          event.preventDefault();
          handleConfirmReferenceInput();
        }
      });
    }

    hideReferenceOverlay();
    startRecommendationIdleMonitor();
    updateUndoRedoButtons();

    setStatus('Ready.');
    fetchBackendHealthSummary()
      .then(function (summary) {
        setStatus('Ready. ' + summary);
      })
      .catch(function () {
        setStatus('Ready. Backend health unavailable.');
      });
  }

  window.PPTAutomation.handleRecommend = handleRecommend;
  window.PPTAutomation.handleAddReference = handleAddReference;
  window.PPTAutomation.handleApplyPlan = handleApplyPlan;
  window.PPTAutomation.handleRejectPlan = handleRejectPlan;
  window.PPTAutomation.handleUndoAccepted = handleUndoAccepted;
  window.PPTAutomation.handleRedoAccepted = handleRedoAccepted;

  window.addEventListener('error', function (event) {
    var err = event && event.error ? event.error : null;
    // eslint-disable-next-line no-console
    console.error('taskpane window.error', err || event);
    setStatus('Script error: check console');
  });

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireUiHandlers);
  } else {
    wireUiHandlers();
  }

  if (typeof Office !== 'undefined' && typeof Office.onReady === 'function') {
    Office.onReady(function () {
      wireUiHandlers();
    });
  }
})();
