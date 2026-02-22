/* global Office, PowerPoint */
(function () {
  var REFERENCE_SHAPE_PREFIX = 'PPTAutomationReference_';
  var REFERENCE_LIST_SHAPE_PREFIX = 'PPTAutomationReferenceList_';
  var REFERENCE_HEADER_SHAPE_NAME = 'PPTAutomationReferenceHeader';
  var REFERENCE_BOX_SHAPE_NAME = 'PPTAutomationSourcesBox';

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
      referenceDialogResolver: null,
      sourceLinksBySlide: {}
    };
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

  function hidePreviewOverlay() {
    var overlay = document.getElementById('previewOverlay');
    if (!overlay) return;
    overlay.classList.add('hidden');
    overlay.setAttribute('aria-hidden', 'true');
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

  function showPreviewOverlay(plan) {
    var overlay = document.getElementById('previewOverlay');
    var summary = document.getElementById('overlaySummary');
    var warnings = document.getElementById('overlayWarnings');
    var operations = document.getElementById('overlayOperations');
    if (!overlay || !summary || !warnings || !operations) return;

    summary.textContent = (plan && plan.summary) ? plan.summary : 'Review the generated plan before applying.';
    warnings.innerHTML = '';
    operations.innerHTML = '';

    var planWarnings = plan && Array.isArray(plan.warnings) ? plan.warnings : [];
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
        var txt = op.content && typeof op.content.text === 'string' ? op.content.text : '(non-text payload)';
        item.textContent = (j + 1) + '. ' + (op.type || 'unknown') + ' -> ' + (op.target || 'auto-target') + ' | ' + txt.slice(0, 120);
        operations.appendChild(item);
      }
    }

    overlay.classList.remove('hidden');
    overlay.setAttribute('aria-hidden', 'false');
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

        var title = document.createElement('h3');
        title.textContent = rec.title || 'Recommendation';

        var desc = document.createElement('p');
        desc.textContent = rec.description || '';

        var meta = document.createElement('p');
        meta.className = 'meta';
        var conf = typeof rec.confidence === 'number' ? rec.confidence.toFixed(2) : '0.00';
        meta.textContent = 'Type: ' + (rec.outputType || 'other') + ' | Confidence: ' + conf;

        var button = document.createElement('button');
        button.type = 'button';
        button.textContent = 'Generate Plan';
        button.addEventListener('click', function () {
          onSelect(rec);
        });

        item.appendChild(title);
        item.appendChild(desc);
        item.appendChild(meta);
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
    var latestContext = uiState.latestSlideContext || {};

    if (!plan) {
      setStatus('Generate a plan first.');
      return;
    }

    if (typeof window.PPTAutomation.applyExecutionPlan !== 'function') {
      setStatus('Plan applier is unavailable.');
      return;
    }

    setStatus('Applying plan to slide...');
    window.PPTAutomation.applyExecutionPlan(plan, latestContext)
      .then(function (result) {
        var warnings = result && Array.isArray(result.warnings) ? result.warnings : [];
        var applied = result && typeof result.appliedCount === 'number' ? result.appliedCount : 0;
        var warningSummary = warnings.length ? ' Warnings: ' + warnings.join(' | ') : '';
        setStatus('Applied ' + applied + ' operation(s).' + warningSummary);
        uiState.pendingPlan = null;
        hidePreviewOverlay();
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Failed to apply plan');
      });
  }

  function handleRejectPlan() {
    window.PPTAutomation.uiState.pendingPlan = null;
    hidePreviewOverlay();
    setStatus('Plan rejected. No changes were applied.');
  }

  function handleAddReference() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.isAddingReference) return;

    var addReferenceBtn = document.getElementById('addReferenceBtn');
    if (typeof window.PPTAutomation.collectSlideContext !== 'function') {
      setStatus('Slide collector is unavailable.');
      return;
    }

    uiState.isAddingReference = true;
    if (addReferenceBtn) addReferenceBtn.disabled = true;

    function cleanup() {
      uiState.isAddingReference = false;
      if (addReferenceBtn) addReferenceBtn.disabled = false;
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

  function handleRecommend() {
    var uiState = window.PPTAutomation.uiState;
    if (uiState.isRecommending) return;

    var recEl = document.getElementById('recommendations');
    var previewEl = document.getElementById('planPreview');
    var recommendBtn = document.getElementById('recommendBtn');

    if (typeof window.PPTAutomation.collectSlideContext !== 'function') {
      setStatus('Slide collector is unavailable.');
      return;
    }

    uiState.isRecommending = true;
    if (recommendBtn) recommendBtn.disabled = true;

    setStatus('Reading slide...');
    if (recEl) recEl.innerHTML = '';
    if (previewEl) previewEl.textContent = '';
    hidePreviewOverlay();
    uiState.latestPlan = null;
    uiState.pendingPlan = null;
    uiState.latestSlideContext = null;

    function cleanupRecommendState() {
      uiState.isRecommending = false;
      if (recommendBtn) recommendBtn.disabled = false;
    }

    Promise.resolve()
      .then(function () {
        return window.PPTAutomation.collectSlideContext();
      })
      .then(function (slideContext) {
        uiState.latestSlideContext = slideContext;
        var summarized = buildSummaryContextWithImage(slideContext);
        var userPrompt = buildAutoPromptFromContext(slideContext);
        setStatus('Generating recommendations...');
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
              uiState.latestPlan = plan;
              uiState.pendingPlan = plan;
              showPreviewOverlay(plan || {});
              setStatus('Plan generated. Confirm or reject in preview overlay.');
            })
            .catch(function (error) {
              setStatus((error && error.message) || 'Plan generation failed');
            });
        });

        setStatus('Loaded ' + recommendations.length + ' recommendation(s).');
      })
      .catch(function (error) {
        setStatus((error && error.message) || 'Unexpected error');
      })
      .then(function () {
        cleanupRecommendState();
      }, function () {
        cleanupRecommendState();
      });
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
      '3. Propose the best output format (list, table, chart, image, or layout-improvement).',
      '4. Include at least one formatting/layout recommendation when readability or hierarchy can improve.',
      '5. Preserve existing design and placeholders where possible.',
      '6. Keep recommendations concise, high-confidence, and presentation-ready.'
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
    var confirmApplyBtn = document.getElementById('confirmApplyBtn');
    var rejectApplyBtn = document.getElementById('rejectApplyBtn');
    var confirmReferenceBtn = document.getElementById('confirmReferenceBtn');
    var cancelReferenceBtn = document.getElementById('cancelReferenceBtn');
    var referenceLabelInput = document.getElementById('referenceLabelInput');
    var referenceUrlInput = document.getElementById('referenceUrlInput');

    if (recommendBtn) recommendBtn.addEventListener('click', handleRecommend);
    if (addReferenceBtn) addReferenceBtn.addEventListener('click', handleAddReference);
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
