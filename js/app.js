// js/app.js
// ツリービュー + プレビューのメインロジック
(function () {
  "use strict";

  // ── DOM 参照 ──────────────────────────────────
  let signinScreen, signinBtn, signinError, appHeader, appBody;
  let loginBtn, logoutBtn, userNameEl;
  let fileListStatus, treeContainer;
  let previewFrame, previewPlaceholder, previewSpinner;
  let searchInput, searchClearBtn, searchResultsContainer;

  // ── 状態 ──────────────────────────────────────
  let rootNode     = null;   // ツリーのルートノード
  let activeNodeId = null;   // 現在選択中の PPTX ノード ID
  let searchDebounceTimer = null;
  let searchActiveId      = null;  // ツリーの activeNodeId と独立

  // ── 起動 ──────────────────────────────────────
  async function boot() {
    signinScreen  = document.getElementById("signin-screen");
    signinBtn     = document.getElementById("signin-btn");
    signinError   = document.getElementById("signin-error");
    appHeader     = document.getElementById("app-header");
    appBody       = document.getElementById("app-body");
    loginBtn      = document.getElementById("login-btn");
    logoutBtn     = document.getElementById("logout-btn");
    userNameEl    = document.getElementById("user-name");
    fileListStatus    = document.getElementById("file-list-status");
    treeContainer     = document.getElementById("tree-container");
    previewFrame      = document.getElementById("preview-frame");
    previewPlaceholder = document.getElementById("preview-placeholder");
    previewSpinner    = document.getElementById("preview-spinner");
    searchInput            = document.getElementById("search-input");
    searchClearBtn         = document.getElementById("search-clear-btn");
    searchResultsContainer = document.getElementById("search-results-container");

    signinBtn.addEventListener("click", handleLogin);
    logoutBtn.addEventListener("click", handleLogout);
    if (searchInput)    searchInput.addEventListener("input",  handleSearchInput);
    if (searchClearBtn) searchClearBtn.addEventListener("click", clearSearch);

    setUiState("loading");

    try {
      const account = await window.Auth.initAuth();
      if (account) {
        setUiState("authenticated", account);
        await loadTree();
      } else {
        setUiState("unauthenticated");
      }
    } catch (err) {
      console.error("起動エラー:", err);
      setUiState("unauthenticated");
    }
  }

  // ── 認証 ──────────────────────────────────────
  async function handleLogin() {
    signinBtn.disabled    = true;
    signinBtn.textContent = "サインイン中...";
    signinError.style.display = "none";
    try {
      const account = await window.Auth.login();
      setUiState("authenticated", account);
      await loadTree();
    } catch (err) {
      console.error("ログイン失敗:", err);
      signinError.textContent   = "サインインに失敗しました: " + err.message;
      signinError.style.display = "block";
      signinBtn.disabled    = false;
      signinBtn.innerHTML   = `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M11.5 2C6.25 2 2 6.25 2 11.5S6.25 21 11.5 21c2.91 0 5.52-1.25 7.35-3.24l-1.46-1.46A7.44 7.44 0 0 1 11.5 19 7.5 7.5 0 0 1 4 11.5 7.5 7.5 0 0 1 11.5 4c3.36 0 6.21 2.22 7.13 5.27H16v2h6V5h-2v2.46A9.46 9.46 0 0 0 11.5 2z" fill="currentColor"/></svg> Microsoft アカウントでサインイン`;
    }
  }

  async function handleLogout() {
    try { await window.Auth.logout(); } finally {
      rootNode = null;
      treeContainer.innerHTML = "";
      clearSearch();
      clearPreview();
      setUiState("unauthenticated");
    }
  }

  // ── ツリー読み込み ─────────────────────────────
  async function loadTree() {
    setStatus("loading", "フォルダを読み込み中...");
    treeContainer.innerHTML = "";

    try {
      // 起点フォルダのメタデータ取得
      rootNode = await window.Graph.loadRootFolder();

      // 起点フォルダの子アイテムを即時読み込み（最初は展開済み）
      rootNode.loading  = true;
      rootNode.children = await window.Graph.loadFolderChildren(rootNode.driveId, rootNode.id);
      rootNode.loading  = false;
      rootNode.expanded = true;

      renderTree();
      searchInput.disabled = false;

      const pptxCount = countPptx(rootNode);
      setStatus("done", `${pptxCount} 件の PPTX ファイル`);
    } catch (err) {
      console.error("ツリー読み込みエラー:", err);
      setStatus("error", "フォルダの読み込みに失敗しました: " + err.message);
    }
  }

  // ツリー全体の PPTX 件数を再帰的にカウント（読み込み済み分のみ）
  function countPptx(node) {
    if (!node.children) return 0;
    let count = 0;
    for (const child of node.children) {
      if (child.isPptx) count++;
      else if (child.isFolder) count += countPptx(child);
    }
    return count;
  }

  // ── ツリー描画 ────────────────────────────────
  function renderTree() {
    treeContainer.innerHTML = "";
    if (!rootNode?.children) return;
    // ルートフォルダ自体はヘッダーに表示するため、子から描画
    renderChildren(treeContainer, rootNode.children, 0);
  }

  function renderChildren(container, children, depth) {
    for (const node of children) {
      container.appendChild(createNodeElement(node, depth));
    }
  }

  function createNodeElement(node, depth) {
    const INDENT = 18; // px per depth level

    // ── ラッパー ──
    const wrapper = document.createElement("div");
    wrapper.className = "tree-item";

    // ── 行 ──
    const row = document.createElement("div");
    row.className = "tree-row" + (node.isPptx ? " tree-row--file" : " tree-row--folder");
    row.style.paddingLeft = `${depth * INDENT + 10}px`;
    row.dataset.depth = depth;

    if (node.isPptx && node.id === activeNodeId) {
      row.classList.add("tree-row--active");
    }

    // 矢印（フォルダのみ）
    const arrow = document.createElement("span");
    arrow.className = "tree-arrow";
    arrow.setAttribute("aria-hidden", "true");
    if (node.isFolder) {
      arrow.textContent = node.expanded ? "▼" : "▶";
    }
    row.appendChild(arrow);

    // アイコン
    const icon = document.createElement("span");
    icon.className = "tree-icon";
    icon.setAttribute("aria-hidden", "true");
    if (node.isFolder)    icon.textContent = node.expanded ? "📂" : "📁";
    else if (node.isPptx) icon.textContent = "📊";
    row.appendChild(icon);

    // ラベル
    const label = document.createElement("span");
    label.className = "tree-label";
    label.textContent = node.name;
    label.title = node.name;
    row.appendChild(label);

    // ロード中スピナー（フォルダ展開中）
    if (node.loading) {
      const sp = document.createElement("span");
      sp.className = "tree-inline-spinner";
      row.appendChild(sp);
    }

    wrapper.appendChild(row);

    // ── 子コンテナ（フォルダのみ） ──
    if (node.isFolder) {
      const childrenEl = document.createElement("div");
      childrenEl.className = "tree-children";
      childrenEl.style.display = node.expanded ? "block" : "none";

      if (node.children?.length > 0) {
        renderChildren(childrenEl, node.children, depth + 1);
      } else if (node.children?.length === 0) {
        childrenEl.appendChild(makeEmptyLabel(depth + 1, INDENT));
      }

      wrapper.appendChild(childrenEl);

      // フォルダクリックで展開・折りたたみ
      row.setAttribute("role", "button");
      row.setAttribute("tabindex", "0");
      row.setAttribute("aria-label", `${node.name} フォルダを開閉`);
      row.addEventListener("click",   () => toggleFolder(node, wrapper));
      row.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); toggleFolder(node, wrapper); }
      });
    }

    // ── PPTX クリックでプレビュー ──
    if (node.isPptx) {
      row.setAttribute("role", "button");
      row.setAttribute("tabindex", "0");
      row.setAttribute("aria-label", `${node.name} のプレビューを表示`);
      row.addEventListener("click",   () => handleFileClick(node, row));
      row.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handleFileClick(node, row); }
      });
    }

    return wrapper;
  }

  function makeEmptyLabel(depth, INDENT) {
    const el = document.createElement("div");
    el.className = "tree-empty";
    el.style.paddingLeft = `${depth * INDENT + 10}px`;
    el.textContent = "（ファイルなし）";
    return el;
  }

  // ── フォルダ展開・折りたたみ ──────────────────
  async function toggleFolder(node, wrapperEl) {
    if (node.loading) return;

    const childrenEl = wrapperEl.querySelector(":scope > .tree-children");
    const arrowEl    = wrapperEl.querySelector(":scope > .tree-row .tree-arrow");
    const iconEl     = wrapperEl.querySelector(":scope > .tree-row .tree-icon");
    const depth      = parseInt(wrapperEl.querySelector(":scope > .tree-row").dataset.depth || 0);
    const INDENT     = 18;

    if (node.children === null) {
      // 未読み込み → API から取得
      node.loading = true;
      arrowEl.textContent = "◌";

      try {
        node.children = await window.Graph.loadFolderChildren(node.driveId, node.id);
        node.expanded = true;
        node.loading  = false;

        arrowEl.textContent = "▼";
        iconEl.textContent  = "📂";
        childrenEl.innerHTML = "";

        if (node.children.length > 0) {
          renderChildren(childrenEl, node.children, depth + 1);
        } else {
          childrenEl.appendChild(makeEmptyLabel(depth + 1, INDENT));
        }
        childrenEl.style.display = "block";

        // 読み込んだ PPTX 数をステータスに反映
        const total = countPptx(rootNode);
        setStatus("done", `${total} 件の PPTX ファイル`);
      } catch (err) {
        node.loading = false;
        arrowEl.textContent = "▶";
        console.error("フォルダ読み込みエラー:", err);
        setStatus("error", "フォルダの読み込みに失敗しました: " + err.message);
      }
    } else {
      // 読み込み済み → 表示切り替え
      node.expanded = !node.expanded;
      arrowEl.textContent = node.expanded ? "▼" : "▶";
      iconEl.textContent  = node.expanded ? "📂" : "📁";
      childrenEl.style.display = node.expanded ? "block" : "none";
    }
  }

  // ── プレビュー ────────────────────────────────
  async function handleFileClick(node, rowEl) {
    if (activeNodeId === node.id) return;

    // アクティブ強調を更新
    document.querySelectorAll(".tree-row--active")
      .forEach((el) => el.classList.remove("tree-row--active"));
    rowEl.classList.add("tree-row--active");
    activeNodeId = node.id;

    showPreviewSpinner();

    try {
      const url = await window.Graph.getPreviewUrl(node.driveId, node.id);
      previewFrame.src           = url;
      previewFrame.style.display = "block";
      previewPlaceholder.style.display = "none";
    } catch (err) {
      console.error("プレビューエラー:", err);
      clearPreview();
      setStatus("error", "プレビューを表示できませんでした: " + err.message);
    } finally {
      hidePreviewSpinner();
    }
  }

  // ── UI ヘルパー ───────────────────────────────
  function setUiState(state, account) {
    switch (state) {
      case "unauthenticated":
        signinScreen.style.display = "flex";
        appHeader.style.display    = "none";
        appBody.style.display      = "none";
        signinBtn.disabled         = false;
        if (searchInput) { searchInput.disabled = true; searchInput.value = ""; }
        if (searchClearBtn) searchClearBtn.style.display = "none";
        showTreeView();
        break;
      case "authenticated":
        signinScreen.style.display = "none";
        appHeader.style.display    = "flex";
        appBody.style.display      = "flex";
        logoutBtn.style.display    = "inline-block";
        userNameEl.textContent     = account?.name || account?.username || "";
        break;
      case "loading":
        signinScreen.style.display = "flex";
        appHeader.style.display    = "none";
        appBody.style.display      = "none";
        signinBtn.disabled         = true;
        if (searchInput) searchInput.disabled = true;
        break;
    }
  }

  function setStatus(type, message) {
    fileListStatus.textContent = message;
    fileListStatus.className   = type ? `status status-${type}` : "status";
  }

  function showPreviewSpinner() {
    previewSpinner.style.display     = "flex";
    previewFrame.style.display       = "none";
    previewPlaceholder.style.display = "none";
  }

  function hidePreviewSpinner() {
    previewSpinner.style.display = "none";
  }

  function clearPreview() {
    previewFrame.src                 = "about:blank";
    previewFrame.style.display       = "none";
    previewPlaceholder.style.display = "flex";
    activeNodeId = null;
  }

  // ── 検索 ─────────────────────────────────────
  function handleSearchInput() {
    const raw = searchInput.value;
    searchClearBtn.style.display = raw.length > 0 ? "inline-block" : "none";
    clearTimeout(searchDebounceTimer);
    if (raw.trim().length < 2) {
      showTreeView();
      return;
    }
    searchDebounceTimer = setTimeout(() => executeSearch(raw.trim()), 400);
  }

  async function executeSearch(query) {
    showSearchView();
    setStatus("loading", "検索中...");
    searchResultsContainer.innerHTML =
      `<div class="search-results-loading"><span class="search-spinner-inline"></span>検索中...</div>`;

    try {
      const results = await window.Graph.searchFiles(rootNode.driveId, query);

      if (searchInput.value.trim() !== query) return;

      renderSearchResults(results, query);

      if (results.length === 0) {
        setStatus("empty", `「${query}」に一致するファイルはありません`);
      } else {
        setStatus("done", `${results.length}${results.length === 25 ? " 件以上" : " 件"}見つかりました`);
      }
    } catch (err) {
      console.error("検索エラー:", err);
      searchResultsContainer.innerHTML =
        `<div class="search-results-error">検索エラー: ${escapeHtml(err.message)}</div>`;
      setStatus("error", "検索に失敗しました");
    }
  }

  function renderSearchResults(results, query) {
    searchResultsContainer.innerHTML = "";
    if (results.length === 0) {
      searchResultsContainer.innerHTML =
        `<div class="search-results-empty">「${escapeHtml(query)}」に一致するファイルはありません</div>`;
      return;
    }
    for (const result of results) {
      searchResultsContainer.appendChild(createSearchResultElement(result));
    }
  }

  function createSearchResultElement(result) {
    const row = document.createElement("div");
    row.className = "search-result-row";
    if (result.id === searchActiveId) row.classList.add("search-result-row--active");
    row.setAttribute("role", "listitem");
    row.setAttribute("tabindex", "0");
    row.setAttribute("aria-label", `${result.name}、フォルダ: ${result.folderPath}`);

    const icon = document.createElement("span");
    icon.className   = "search-result-icon";
    icon.textContent = "📊";
    icon.setAttribute("aria-hidden", "true");
    row.appendChild(icon);

    const textBlock = document.createElement("div");
    textBlock.className = "search-result-text";

    const nameEl = document.createElement("div");
    nameEl.className   = "search-result-name";
    nameEl.textContent = result.name;
    nameEl.title       = result.name;
    textBlock.appendChild(nameEl);

    if (result.folderPath) {
      const pathEl = document.createElement("div");
      pathEl.className   = "search-result-path";
      pathEl.textContent = result.folderPath;
      pathEl.title       = result.folderPath;
      textBlock.appendChild(pathEl);
    }

    row.appendChild(textBlock);

    row.addEventListener("click", () => handleSearchResultClick(result, row));
    row.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") { e.preventDefault(); handleSearchResultClick(result, row); }
    });

    return row;
  }

  async function handleSearchResultClick(result, rowEl) {
    if (searchActiveId === result.id) return;

    searchResultsContainer.querySelectorAll(".search-result-row--active")
      .forEach((el) => el.classList.remove("search-result-row--active"));
    rowEl.classList.add("search-result-row--active");
    searchActiveId = result.id;

    showPreviewSpinner();
    try {
      const url = await window.Graph.getPreviewUrl(result.driveId, result.id);
      previewFrame.src                 = url;
      previewFrame.style.display       = "block";
      previewPlaceholder.style.display = "none";
    } catch (err) {
      console.error("プレビューエラー (検索結果):", err);
      clearPreview();
      setStatus("error", "プレビューを表示できませんでした: " + err.message);
    } finally {
      hidePreviewSpinner();
    }
  }

  function clearSearch() {
    searchInput.value            = "";
    searchClearBtn.style.display = "none";
    searchActiveId               = null;
    clearTimeout(searchDebounceTimer);
    showTreeView();
    if (rootNode) {
      const total = countPptx(rootNode);
      setStatus("done", `${total} 件の PPTX ファイル`);
    }
  }

  function showTreeView() {
    treeContainer.style.display          = "";
    searchResultsContainer.style.display = "none";
    searchResultsContainer.innerHTML     = "";
  }

  function showSearchView() {
    treeContainer.style.display          = "none";
    searchResultsContainer.style.display = "";
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ── エントリーポイント ─────────────────────────
  document.addEventListener("DOMContentLoaded", boot);
})();
