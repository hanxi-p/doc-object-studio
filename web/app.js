(() => {
  const token = new URLSearchParams(location.hash.slice(1)).get('token') || '';
  const $ = (id) => document.getElementById(id);
  const state = {
    opened: false, opening: false, page: 0, pageCount: 0, zoom: 100, fit: true,
    objects: [], documentObjects: [], selected: new Set(), objectMap: new Map(), pageObjectMap: new Map(),
    activeUid: '', pageState: null,
    indexingObjects: false,
    previewUrl: '', thumbnailUrls: new Map(), thumbnailObserver: null, thumbnailGeneration: 0,
    loadSequence: 0, hoveredUid: ''
  };
  const typeLabels = { text: '文字', image: '图片', shape: '形状', group: '组合', vector: '矢量', table: '表格', ole: '公式', document: '文档' };
  const typeIcons = { text: 'T', image: '▧', shape: '◇', group: '◈', vector: '⌁', table: '▦', ole: '∑' };

  async function request(path, options = {}, responseType = 'json') {
    const headers = { 'X-DocStudio-Token': token, ...(options.headers || {}) };
    const response = await fetch(path, { ...options, headers });
    if (!response.ok) {
      let message = '请求失败（' + response.status + '）';
      try {
        const body = await response.json();
        message = body.error || message;
      } catch (_) {}
      throw new Error(message);
    }
    return responseType === 'blob' ? response.blob() : response.json();
  }

  function uploadFile(file, onProgress) {
    return new Promise((resolve, reject) => {
      const form = new FormData();
      form.append('file', file, file.name);
      const xhr = new XMLHttpRequest();
      xhr.open('POST', '/api/open');
      xhr.timeout = 10 * 60 * 1000;
      xhr.setRequestHeader('X-DocStudio-Token', token);
      xhr.setRequestHeader('X-File-Name', encodeURIComponent(file.name));
      xhr.upload.addEventListener('progress', (event) => {
        if (event.lengthComputable) onProgress(event.loaded, event.total, false);
      });
      xhr.upload.addEventListener('load', () => onProgress(file.size, file.size, true));
      xhr.addEventListener('load', () => {
        let payload = {};
        try { payload = JSON.parse(xhr.responseText || '{}'); } catch (_) {}
        if (xhr.status >= 200 && xhr.status < 300) resolve(payload);
        else reject(new Error(payload.error || '打开失败（HTTP ' + xhr.status + '）'));
      });
      xhr.addEventListener('error', () => reject(new Error('上传失败，请检查本机服务是否仍在运行。')));
      xhr.addEventListener('timeout', () => reject(new Error('打开文档超时。')));
      xhr.addEventListener('abort', () => reject(new Error('打开文档已取消。')));
      xhr.send(form);
    });
  }

  function toast(message, error = false) {
    const node = $('toast');
    node.textContent = message;
    node.style.background = error ? '#943e4a' : '#20283a';
    node.classList.add('visible');
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => node.classList.remove('visible'), 2800);
  }

  function setStatus(message, busy = false) {
    const node = $('statusMessage');
    node.replaceChildren();
    const indicator = document.createElement('span');
    indicator.className = 'status-indicator';
    indicator.style.background = busy ? '#dfaa54' : '#53b69b';
    node.append(indicator, document.createTextNode(message));
  }

  function showLoading(percent, detail, title) {
    const panel = $('loadingPanel');
    const value = Math.max(0, Math.min(100, Math.round(percent)));
    panel.hidden = false;
    $('loadingTitle').textContent = title || '正在打开文档';
    $('loadingDetail').textContent = detail || '正在准备文件…';
    $('loadingPercent').textContent = value + '%';
    $('loadingFill').style.width = value + '%';
  }

  function hideLoading() {
    $('loadingPanel').hidden = true;
  }

  function flatten(items, depth = 0, result = []) {
    for (const item of items || []) {
      result.push({ item, depth });
      flatten(item.children || [], depth + 1, result);
    }
    return result;
  }

  function clearThumbnailCache() {
    state.thumbnailGeneration += 1;
    if (state.thumbnailObserver) state.thumbnailObserver.disconnect();
    state.thumbnailObserver = null;
    for (const url of state.thumbnailUrls.values()) URL.revokeObjectURL(url);
    state.thumbnailUrls.clear();
  }

  function clearDocument() {
    state.loadSequence += 1;
    state.opened = false;
    state.page = 0;
    state.pageCount = 0;
    state.objects = [];
    state.documentObjects = [];
    state.pageState = null;
    state.selected.clear();
    state.activeUid = '';
    state.objectMap.clear();
    state.pageObjectMap.clear();
    state.hoveredUid = '';
    $('pageCanvas').hidden = true;
    $('emptyState').hidden = false;
    $('pageList').innerHTML = '<div class="rail-empty">打开文档后<br>页面会显示在这里</div>';
    $('objectList').innerHTML = '<div class="object-list-empty">打开文档后<br>可选取整篇中的图片与表格</div>';
    $('objectDetails').textContent = '仅复制图片与表格；勾选项可跨页面保留。';
    $('objectCount').textContent = '0';
    $('pageCount').textContent = '—';
    $('pagePosition').textContent = '— / —';
    $('fileName').textContent = '尚未打开文档';
    $('fileBadge').classList.add('is-empty');
    $('fileBadge').classList.remove('is-open');
    $('previousPage').disabled = true;
    $('nextPage').disabled = true;
    $('copyButton').disabled = true;
    $('copyPanelButton').disabled = true;
    if (state.previewUrl) URL.revokeObjectURL(state.previewUrl);
    state.previewUrl = '';
    $('pageImage').removeAttribute('src');
    $('objectOverlay').replaceChildren();
    clearThumbnailCache();
    hideLoading();
  }

  async function openFile(file) {
    if (!file || state.opening) return;
    const extension = file.name.split('.').pop().toLowerCase();
    if (!['pdf', 'pptx', 'docx'].includes(extension)) {
      toast('请选择 PDF、PPTX 或 DOCX 文件。', true);
      return;
    }

    state.opening = true;
    clearDocument();
    $('emptyState').hidden = true;
    $('openButton').disabled = true;
    setStatus('正在上传并解析文档…', true);
    showLoading(2, '正在上传文件…', '正在打开文档');
    let uploadFinished = false;
    let displayedProgress = 2;
    const progressTimer = setInterval(async () => {
      if (!uploadFinished) return;
      try {
        const progress = await request('/api/progress');
        if (!state.opening) return;
        const backendPercent = Number(progress.percent || 0);
        if (state.indexingObjects) {
          displayedProgress = Math.max(displayedProgress, 97 + backendPercent * 0.025);
          showLoading(displayedProgress, progress.message || '正在检查全文图片与表格…', '整理可复制对象');
          return;
        }
        const mapped = Math.min(96, 20 + backendPercent * 0.76);
        displayedProgress = Math.max(displayedProgress, mapped);
        showLoading(displayedProgress, progress.message || '正在解析和渲染页面…');
      } catch (_) {}
    }, 350);

    try {
      const result = await uploadFile(file, (loaded, total, complete) => {
        if (complete) {
          uploadFinished = true;
          displayedProgress = Math.max(displayedProgress, 20);
          showLoading(displayedProgress, '文件上传完成，正在解析文档…');
        } else if (total > 0) {
          const percent = 2 + (loaded / total) * 18;
          displayedProgress = Math.max(displayedProgress, percent);
          showLoading(displayedProgress, '正在上传文件 ' + Math.round((loaded / total) * 100) + '%…');
        }
      });
      uploadFinished = true;
      if (!result.ok) throw new Error('服务未能打开此文档。');
      state.opened = true;
      state.page = 0;
      state.pageCount = result.pages;
      $('fileName').textContent = result.file_name + ' · ' + result.kind.toUpperCase();
      $('fileBadge').classList.remove('is-empty');
      $('fileBadge').classList.add('is-open');
      $('emptyState').hidden = true;
      $('pageCanvas').hidden = false;
      $('pageCount').textContent = String(state.pageCount);
      renderPageList();
      showLoading(97, '正在载入高清首屏和对象坐标…', '页面已渲染');
      await loadPage(0, { quiet: true });
      if (state.fit) fitPageToStage();
      state.indexingObjects = true;
      showLoading(97, '正在索引整份文档的图片与表格…', '整理可复制对象');
      const documentState = await request('/api/objects');
      state.indexingObjects = false;
      state.documentObjects = documentState.objects || [];
      state.objectMap.clear();
      state.documentObjects.forEach((item) => state.objectMap.set(item.uid, item));
      renderObjectList();
      showLoading(100, '页面和全篇图片、表格列表已就绪', '打开完成');
      setStatus(result.kind.toUpperCase() + ' · 共 ' + state.pageCount + ' 页');
      toast('已打开 ' + result.file_name);
      setTimeout(hideLoading, 300);
    } catch (error) {
      state.opened = false;
      $('pageCanvas').hidden = true;
      $('emptyState').hidden = false;
      hideLoading();
      setStatus('打开失败');
      toast(error.message, true);
    } finally {
      clearInterval(progressTimer);
      state.indexingObjects = false;
      state.opening = false;
      $('openButton').disabled = false;
      $('fileInput').value = '';
    }
  }

  async function loadPage(index, options = {}) {
    if (!state.opened) return;
    const pageIndex = Math.max(0, Math.min(index, state.pageCount - 1));
    state.page = pageIndex;
    state.hoveredUid = '';
    const sequence = ++state.loadSequence;
    $('pagePosition').textContent = String(pageIndex + 1).padStart(2, '0') + ' / ' + String(state.pageCount).padStart(2, '0');
    $('previousPage').disabled = pageIndex <= 0;
    $('nextPage').disabled = pageIndex >= state.pageCount - 1;
    $('pageList').querySelectorAll('.page-button').forEach((button, i) => {
      button.classList.toggle('active', i === pageIndex);
      if (i === pageIndex) button.setAttribute('aria-current', 'page');
      else button.removeAttribute('aria-current');
    });
    setStatus('正在载入页面…', true);
    try {
      const pageState = await request('/api/state?page=' + pageIndex);
      if (sequence !== state.loadSequence) return;
      state.pageState = pageState;
      state.objects = pageState.objects || [];
      state.pageObjectMap.clear();
      flatten(state.objects).forEach((entry) => {
        if ((entry.item.kind === 'image' || entry.item.kind === 'table') &&
          (entry.item.display_bbox || entry.item.bbox)) {
          state.pageObjectMap.set(entry.item.uid, entry.item);
          if (entry.item.native_table_uid) state.pageObjectMap.set(entry.item.native_table_uid, entry.item);
        }
      });
      const imageBlob = await request('/api/preview/' + pageIndex, {}, 'blob');
      const nextUrl = URL.createObjectURL(imageBlob);
      if (sequence !== state.loadSequence) {
        URL.revokeObjectURL(nextUrl);
        return;
      }
      const image = $('pageImage');
      if (state.previewUrl) URL.revokeObjectURL(state.previewUrl);
      state.previewUrl = nextUrl;
      image.src = nextUrl;
      if (typeof image.decode === 'function') await image.decode();
      else await new Promise((resolve, reject) => {
        image.onload = resolve;
        image.onerror = () => reject(new Error('页面图片载入失败。'));
      });
      if (sequence !== state.loadSequence) return;
      if (state.fit) fitPageToStage();
      renderSelection();
      drawOverlay();
      setStatus(pageState.kind.toUpperCase() + ' · 第 ' + (pageIndex + 1) + ' 页');
    } catch (error) {
      if (sequence !== state.loadSequence) return;
      if (options.quiet) throw error;
      setStatus('页面载入失败');
      toast(error.message, true);
    }
  }

  function renderPageList() {
    const list = $('pageList');
    list.replaceChildren();
    clearThumbnailCache();
    const canObserve = 'IntersectionObserver' in window;
    if (canObserve) {
      const generation = state.thumbnailGeneration;
      state.thumbnailObserver = new IntersectionObserver((entries) => {
        entries.forEach((entry) => {
          if (!entry.isIntersecting) return;
          const image = entry.target;
          state.thumbnailObserver.unobserve(image);
          loadThumbnail(Number(image.dataset.page), image, generation);
        });
      }, { root: list, rootMargin: '100px' });
    }

    for (let index = 0; index < state.pageCount; index += 1) {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = 'page-button' + (index === state.page ? ' active' : '');
      button.dataset.page = String(index);
      button.setAttribute('aria-label', '打开第 ' + (index + 1) + ' 页');
      button.setAttribute('aria-keyshortcuts', 'ArrowUp ArrowDown ArrowLeft ArrowRight');
      if (index === state.page) button.setAttribute('aria-current', 'page');
      const frame = document.createElement('span');
      frame.className = 'page-thumb-frame';
      const image = document.createElement('img');
      image.className = 'page-thumbnail';
      image.alt = '';
      image.dataset.page = String(index);
      image.loading = index < 3 ? 'eager' : 'lazy';
      frame.append(image);
      const caption = document.createElement('span');
      caption.className = 'page-caption';
      const label = document.createElement('span');
      label.className = 'page-label';
      label.textContent = '第 ' + (index + 1) + ' 页';
      const number = document.createElement('span');
      number.className = 'page-number';
      number.textContent = String(index + 1).padStart(2, '0');
      caption.append(label, number);
      button.append(frame, caption);
      button.addEventListener('click', () => loadPage(index));
      list.append(button);
      if (state.thumbnailObserver) state.thumbnailObserver.observe(image);
      else loadThumbnail(index, image, state.thumbnailGeneration);
    }
  }

  async function loadThumbnail(index, image, generation = state.thumbnailGeneration, attempt = 0) {
    if (!image.isConnected || generation !== state.thumbnailGeneration) return;
    const showFailureOrRetry = () => {
      if (!image.isConnected || generation !== state.thumbnailGeneration) return;
      if (attempt < 2) {
        setTimeout(() => loadThumbnail(index, image, generation, attempt + 1), 250 * (attempt + 1));
      } else {
        image.parentElement.classList.add('thumbnail-error');
      }
    };
    try {
      let url = state.thumbnailUrls.get(index);
      if (!url) {
        const blob = await request('/api/thumbnail/' + index, {}, 'blob');
        if (!image.isConnected || generation !== state.thumbnailGeneration) return;
        url = URL.createObjectURL(blob);
        state.thumbnailUrls.set(index, url);
      }
      image.onload = () => {
        if (generation === state.thumbnailGeneration) image.parentElement.classList.remove('thumbnail-error');
      };
      image.onerror = () => {
        if (state.thumbnailUrls.get(index) === url) {
          state.thumbnailUrls.delete(index);
          URL.revokeObjectURL(url);
        }
        showFailureOrRetry();
      };
      image.src = url;
    } catch (_) {
      showFailureOrRetry();
    }
  }

  function handlePageRailKeys(event) {
    if (!state.opened || event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) return;
    const direction = ({
      ArrowUp: -1,
      ArrowLeft: -1,
      ArrowDown: 1,
      ArrowRight: 1,
    })[event.key];
    const currentButton = event.target.closest && event.target.closest('.page-button');
    if (!direction || !currentButton) return;

    event.preventDefault();
    const currentIndex = Number(currentButton.dataset.page);
    const nextIndex = Math.max(0, Math.min(state.pageCount - 1, currentIndex + direction));
    if (nextIndex === currentIndex) return;
    const nextButton = $('pageList').querySelector('.page-button[data-page="' + nextIndex + '"]');
    if (nextButton) {
      nextButton.focus({ preventScroll: true });
      nextButton.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    loadPage(nextIndex);
  }

  function renderObjectList() {
    const list = $('objectList');
    list.replaceChildren();
    const objects = state.documentObjects;
    $('objectCount').textContent = String(objects.length);
    $('objectTitle').textContent = '全篇图片与表格';
    if (!objects.length) {
      list.innerHTML = '<div class="object-list-empty">整份文档中没有可复制的图片或表格</div>';
      $('selectionHint').textContent = '只列出整篇文档中的图片与表格';
      updateCopyButtons();
      return;
    }

    for (const item of objects) {
      const row = document.createElement('label');
      row.className = 'object-row';
      row.dataset.uid = item.uid;
      row.classList.toggle('selected', state.selected.has(item.uid));
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.className = 'object-check';
      checkbox.checked = state.selected.has(item.uid);
      checkbox.setAttribute('aria-label', '勾选' + (typeLabels[item.kind] || '对象') + '：' + item.label);
      checkbox.addEventListener('change', () => setObjectChecked(item.uid, checkbox.checked));
      const icon = document.createElement('span');
      icon.className = 'object-icon';
      icon.textContent = typeIcons[item.kind] || '•';
      icon.setAttribute('aria-hidden', 'true');
      const label = document.createElement('span');
      label.className = 'object-label';
      label.textContent = item.label;
      const type = document.createElement('span');
      type.className = 'object-kind';
      type.textContent = item.page === null || item.page === undefined ? '全文档' : '第' + (item.page + 1) + '页';
      row.append(checkbox, icon, label, type);
      row.addEventListener('mouseenter', () => setHovered(item.uid));
      row.addEventListener('mouseleave', () => setHovered(''));
      checkbox.addEventListener('focus', () => setHovered(item.uid));
      checkbox.addEventListener('blur', () => setHovered(''));
      list.append(row);
    }
    applySearch();
    renderSelection();
  }

  function setObjectChecked(uid, checked, navigate = true) {
    const item = state.objectMap.get(uid);
    if (!item) return;
    if (checked) {
      state.selected.add(uid);
      state.activeUid = uid;
    } else {
      state.selected.delete(uid);
      if (state.activeUid === uid) state.activeUid = [...state.selected].pop() || '';
    }
    renderSelection();
    if (navigate && checked && Number.isInteger(item.page) && item.page !== state.page) loadPage(item.page);
  }

  function renderSelection() {
    document.querySelectorAll('.object-row').forEach((row) => {
      const checked = state.selected.has(row.dataset.uid);
      row.classList.toggle('selected', checked);
      const checkbox = row.querySelector('.object-check');
      if (checkbox) checkbox.checked = checked;
    });
    renderDetails();
    updateCopyButtons();
    $('selectionHint').textContent = state.selected.size
      ? '整份文档已勾选 ' + state.selected.size + ' 项'
      : '整份文档 · 勾选图片或表格，可跨页多选';
    drawOverlay();
  }

  function renderDetails() {
    const details = $('objectDetails');
    if (!state.selected.size) {
      details.textContent = '仅复制图片与表格；勾选项会跨页面保留。';
      return;
    }
    if (state.selected.size > 1) {
      details.replaceChildren();
      const strong = document.createElement('strong');
      strong.textContent = '已勾选 ' + state.selected.size + ' 个图表对象';
      const hint = document.createElement('span');
      hint.className = 'detail-meta';
      hint.textContent = '跨页内容会保留原图数据，并以 HTML 形式合并复制。';
      details.append(strong, hint);
      return;
    }
    const obj = state.objectMap.get(state.activeUid) || state.objectMap.get([...state.selected][0]);
    if (!obj) return;
    details.replaceChildren();
    const title = document.createElement('strong');
    title.textContent = obj.label;
    const type = document.createElement('span');
    type.className = 'detail-meta';
    const pageLabel = obj.page === null || obj.page === undefined ? ' · 文档级对象' : ' · 第 ' + (obj.page + 1) + ' 页';
    const imageSize = Array.isArray(obj.image_size_px) && obj.image_size_px.length === 2
      ? ' · 原图 ' + obj.image_size_px[0] + '×' + obj.image_size_px[1] + ' px'
      : '';
    const bytes = obj.data_size ? ' · ' + formatBytes(obj.data_size) : '';
    type.textContent = (typeLabels[obj.kind] || obj.kind) + pageLabel + imageSize + bytes;
    details.append(title, type);
  }

  function setHovered(uid) {
    if (state.hoveredUid === uid) return;
    state.hoveredUid = uid;
    drawOverlay();
  }

  function displayBox(item) {
    return item.display_bbox || item.bbox;
  }

  function drawOverlay() {
    const overlay = $('objectOverlay');
    overlay.replaceChildren();
    if (!state.pageState) return;
    const visible = new Set([...state.selected].filter((uid) => state.pageObjectMap.has(uid)));
    if (state.hoveredUid) visible.add(state.hoveredUid);
    const coordinateWidth = Number(state.pageState.coordinate_width || 0);
    const coordinateHeight = Number(state.pageState.coordinate_height || 0);
    if (coordinateWidth <= 0 || coordinateHeight <= 0) return;

    for (const uid of visible) {
      const item = state.pageObjectMap.get(uid);
      const bounds = item && displayBox(item);
      if (!bounds || bounds.length !== 4) continue;
      const x0 = Number(bounds[0]);
      const y0 = Number(bounds[1]);
      const x1 = Number(bounds[2]);
      const y1 = Number(bounds[3]);
      if (![x0, y0, x1, y1].every(Number.isFinite)) continue;
      const box = document.createElement('div');
      box.className = 'object-box' +
        (state.selected.has(uid) ? ' selected' : '') +
        (state.hoveredUid === uid ? ' hovered' : '');
      box.dataset.uid = uid;
      box.title = item.label;
      box.style.left = (x0 / coordinateWidth * 100) + '%';
      box.style.top = (y0 / coordinateHeight * 100) + '%';
      box.style.width = (Math.max(0, x1 - x0) / coordinateWidth * 100) + '%';
      box.style.height = (Math.max(0, y1 - y0) / coordinateHeight * 100) + '%';
      overlay.append(box);
    }
  }

  function hitTest(event) {
    if (!state.pageState) return null;
    const image = $('pageImage');
    const rect = image.getBoundingClientRect();
    if (!rect.width || !rect.height) return null;
    const x = (event.clientX - rect.left) / rect.width * Number(state.pageState.coordinate_width);
    const y = (event.clientY - rect.top) / rect.height * Number(state.pageState.coordinate_height);
    const candidates = flatten(state.objects).map((entry) => entry.item).filter((item) => {
      if (item.kind !== 'image' && item.kind !== 'table') return false;
      if (!state.objectMap.has(item.native_table_uid || item.uid)) return false;
      const box = displayBox(item);
      return box && box.length === 4 && x >= Number(box[0]) && x <= Number(box[2]) &&
        y >= Number(box[1]) && y <= Number(box[3]);
    });
    if (!candidates.length) return null;

    const area = (item) => {
      const box = displayBox(item);
      return Math.max(0, Number(box[2]) - Number(box[0])) *
        Math.max(0, Number(box[3]) - Number(box[1]));
    };
    if (state.pageState.kind === 'pptx') {
      candidates.sort((a, b) => {
        const ap = a.z_path || [];
        const bp = b.z_path || [];
        if (ap.length !== bp.length) return bp.length - ap.length;
        for (let index = 0; index < Math.max(ap.length, bp.length); index += 1) {
          const diff = Number(bp[index] || 0) - Number(ap[index] || 0);
          if (diff) return diff;
        }
        return area(a) - area(b);
      });
    } else {
      candidates.sort((a, b) => area(a) - area(b));
    }
    return candidates[0];
  }

  function updateCopyButtons() {
    const disabled = !state.selected.size;
    $('copyButton').disabled = disabled;
    $('copyPanelButton').disabled = disabled;
  }

  async function copySelected() {
    if (!state.selected.size) return;
    try {
      setStatus('正在复制到 Windows 剪贴板…', true);
      const result = await request('/api/copy', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ uids: [...state.selected] })
      });
      setStatus(result.message || '已复制到剪贴板');
      toast(result.message || '已复制到 Windows 剪贴板');
    } catch (error) {
      setStatus('复制失败');
      toast(error.message, true);
    }
  }

  function formatBytes(value) {
    if (value < 1024) return value + ' B';
    if (value < 1024 * 1024) return (value / 1024).toFixed(1) + ' KB';
    return (value / (1024 * 1024)).toFixed(1) + ' MB';
  }

  function applySearch() {
    const query = $('objectSearch').value.trim().toLocaleLowerCase();
    document.querySelectorAll('.object-row').forEach((row) => {
      row.classList.toggle('hidden', !!query && !row.textContent.toLocaleLowerCase().includes(query));
    });
  }

  function confineColumnWheel(columnSelector, scrollerId) {
    const column = document.querySelector(columnSelector);
    const scroller = $(scrollerId);
    if (!column || !scroller) return;
    column.addEventListener('wheel', (event) => {
      // The narrow layout stacks panels vertically and must keep page scrolling.
      if (window.matchMedia('(max-width: 930px)').matches) return;
      // Keep the browser's standard Ctrl+wheel page zoom shortcut intact.
      if (event.ctrlKey || event.metaKey) return;
      const lineHeight = Number.parseFloat(getComputedStyle(scroller).lineHeight) || 16;
      const multiplier = event.deltaMode === WheelEvent.DOM_DELTA_LINE
        ? lineHeight
        : event.deltaMode === WheelEvent.DOM_DELTA_PAGE ? scroller.clientHeight : 1;
      const dx = event.deltaX * multiplier;
      const dy = event.deltaY * multiplier;
      if (!dx && !dy) return;

      event.preventDefault();
      const horizontalOnly = scroller.scrollHeight <= scroller.clientHeight &&
        scroller.scrollWidth > scroller.clientWidth;
      if (event.shiftKey || horizontalOnly) {
        scroller.scrollLeft += dx || dy;
        if (!horizontalOnly) scroller.scrollTop += dy;
      } else {
        scroller.scrollTop += dy;
        scroller.scrollLeft += dx;
      }
    }, { passive: false });
  }

  function applyZoom() {
    $('pageCanvas').style.setProperty('--zoom', state.zoom);
    $('zoomValue').textContent = Math.round(state.zoom) + '%';
  }

  function fitPageToStage() {
    const image = $('pageImage');
    const stage = $('dropTarget');
    if (!image.naturalWidth || !image.naturalHeight) return;
    const style = getComputedStyle(stage);
    const width = Math.max(100, stage.clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight));
    const height = Math.max(100, stage.clientHeight - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom));
    const pageHeightAtFullWidth = width * image.naturalHeight / image.naturalWidth;
    state.zoom = Math.max(30, Math.min(100, height / pageHeightAtFullWidth * 100));
    state.fit = true;
    applyZoom();
  }

  function setManualZoom(delta) {
    state.fit = false;
    state.zoom = Math.max(40, Math.min(240, state.zoom + delta));
    applyZoom();
  }

  function pageClick(event) {
    const item = hitTest(event);
    if (item) setObjectChecked(item.native_table_uid || item.uid, true, false);
    else if (!event.ctrlKey && !event.metaKey) {
      state.selected.clear();
      state.activeUid = '';
      renderSelection();
    }
  }

  $('openButton').addEventListener('click', () => $('fileInput').click());
  $('emptyOpenButton').addEventListener('click', () => $('fileInput').click());
  $('fileInput').addEventListener('change', (event) => openFile(event.target.files[0]));
  $('copyButton').addEventListener('click', copySelected);
  $('copyPanelButton').addEventListener('click', copySelected);
  $('previousPage').addEventListener('click', () => loadPage(state.page - 1));
  $('nextPage').addEventListener('click', () => loadPage(state.page + 1));
  $('zoomOut').addEventListener('click', () => setManualZoom(-10));
  $('zoomIn').addEventListener('click', () => setManualZoom(10));
  $('fitPage').addEventListener('click', () => fitPageToStage());
  $('objectSearch').addEventListener('input', applySearch);
  $('pageList').addEventListener('keydown', handlePageRailKeys);
  confineColumnWheel('.page-rail', 'pageList');
  confineColumnWheel('.preview-column', 'dropTarget');
  confineColumnWheel('.object-panel', 'objectList');
  $('dropTarget').addEventListener('dragover', (event) => {
    event.preventDefault();
    $('dropTarget').classList.add('is-dragging');
  });
  $('dropTarget').addEventListener('dragleave', (event) => {
    if (!event.relatedTarget || !$('dropTarget').contains(event.relatedTarget)) $('dropTarget').classList.remove('is-dragging');
  });
  $('dropTarget').addEventListener('drop', (event) => {
    event.preventDefault();
    $('dropTarget').classList.remove('is-dragging');
    openFile(event.dataTransfer.files[0]);
  });
  $('pageCanvas').addEventListener('pointermove', (event) => {
    const item = hitTest(event);
    setHovered(item ? (item.native_table_uid || item.uid) : '');
  });
  $('pageCanvas').addEventListener('pointerleave', () => setHovered(''));
  $('pageCanvas').addEventListener('click', pageClick);
  $('quitButton').addEventListener('click', async () => {
    try { await request('/api/quit', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' }); }
    finally { document.body.innerHTML = '<main style="font:15px Segoe UI,sans-serif;padding:48px;color:#344054">文档对象工坊已退出，可以关闭此标签页。</main>'; }
  });
  window.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      state.selected.clear();
      state.activeUid = '';
      renderSelection();
    }
    if (event.key === '/' && document.activeElement !== $('objectSearch')) {
      event.preventDefault();
      $('objectSearch').focus();
    }
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'c' &&
      state.selected.size && document.activeElement !== $('objectSearch')) {
      event.preventDefault();
      copySelected();
    }
  });
  window.addEventListener('resize', () => {
    if (state.fit) fitPageToStage();
  });

  if (!token) {
    clearDocument();
    toast('缺少本地会话凭据，请重新启动文档对象工坊。', true);
  } else {
    clearDocument();
    request('/api/state?page=0').catch((error) => toast(error.message, true));
  }
})();
