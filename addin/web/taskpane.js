/* global Office */
(function () {
  if (!window.PPTAutomation) {
    window.PPTAutomation = {};
  }
  if (!window.PPTAutomation.uiState) {
    window.PPTAutomation.uiState = {
      latestPlan: null,
      latestSlideContext: null,
      pendingPlan: null,
      isRecommending: false
    };
  }

  function setStatus(message) {
    var node = document.getElementById('status');
    if (node) {
      node.textContent = message;
    }
  }

  function hidePreviewOverlay() {
    var overlay = document.getElementById('previewOverlay');
    if (!overlay) return;
    overlay.classList.add('hidden');
    overlay.setAttribute('aria-hidden', 'true');
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
    var confirmApplyBtn = document.getElementById('confirmApplyBtn');
    var rejectApplyBtn = document.getElementById('rejectApplyBtn');

    if (recommendBtn) recommendBtn.addEventListener('click', handleRecommend);
    if (confirmApplyBtn) confirmApplyBtn.addEventListener('click', handleApplyPlan);
    if (rejectApplyBtn) rejectApplyBtn.addEventListener('click', handleRejectPlan);

    setStatus('Ready.');
  }

  window.PPTAutomation.handleRecommend = handleRecommend;
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
