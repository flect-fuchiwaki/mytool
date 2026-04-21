// js/graph.js
// Microsoft Graph API 呼び出しモジュール
(function () {
  "use strict";

  const BASE = window.APP_CONFIG.graphBaseUrl;

  // 共通 fetch ラッパー（トークン付与・エラーハンドリング）
  async function graphFetch(url, options = {}) {
    const token = await window.Auth.getToken();
    const res = await fetch(url, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
        ...options.headers,
      },
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error(`Graph ${res.status}: ${err?.error?.message || res.statusText}`);
    }
    return res.json();
  }

  // ──────────────────────────────────────────────
  // ツリービュー用：起点フォルダのメタデータを取得
  // ──────────────────────────────────────────────
  async function loadRootFolder() {
    const { searchSite, searchFolderPath } = window.APP_CONFIG;

    // フォルダパスの各セグメントを個別にエンコード
    const encodedPath = searchFolderPath
      .split("/")
      .map(encodeURIComponent)
      .join("/");

    const url = `${BASE}/sites/${searchSite}:/drive/root:/${encodedPath}`;
    const data = await graphFetch(url);

    return {
      id:       data.id,
      name:     data.name,
      driveId:  data.parentReference.driveId,
      isFolder: true,
      isPptx:   false,
      children: null,   // null = 未読み込み
      expanded: false,
      loading:  false,
    };
  }

  // ──────────────────────────────────────────────
  // フォルダの子アイテムを取得（フォルダと .pptx のみ返す）
  // ──────────────────────────────────────────────
  async function loadFolderChildren(driveId, itemId) {
    let url = `${BASE}/drives/${driveId}/items/${itemId}/children`
            + `?$orderby=name asc&$top=200`
            + `&$select=id,name,folder,file,lastModifiedDateTime,size,webUrl,parentReference`;

    const items = [];
    // nextLink によるページング対応
    while (url) {
      const data = await graphFetch(url);
      items.push(...(data.value || []));
      url = data["@odata.nextLink"] || null;
    }

    return items
      // フォルダ または .pptx のみ対象
      .filter((item) => item.folder || item.name.toLowerCase().endsWith(".pptx"))
      .map((item) => ({
        id:          item.id,
        name:        item.name,
        driveId:     driveId,
        isFolder:    !!item.folder,
        isPptx:      !item.folder && item.name.toLowerCase().endsWith(".pptx"),
        lastModified: item.lastModifiedDateTime,
        webUrl:      item.webUrl,
        children:    item.folder ? null : undefined, // フォルダは未読み込み状態
        expanded:    false,
        loading:     false,
      }));
  }

  // ──────────────────────────────────────────────
  // Office Online 埋め込みプレビュー URL を取得
  // ──────────────────────────────────────────────
  async function getPreviewUrl(driveId, itemId) {
    const url = `${BASE}/drives/${driveId}/items/${itemId}/preview`;
    const data = await graphFetch(url, {
      method: "POST",
      body: JSON.stringify({
        viewer:     "office",
        chromeless: true,
        allowEdit:  false,
      }),
    });

    let embedUrl = data.getUrl;
    if (!embedUrl) throw new Error("プレビュー URL が取得できませんでした");

    // Office Online のバナーを非表示にする
    if (!embedUrl.includes("nb=true")) {
      embedUrl += (embedUrl.includes("?") ? "&" : "?") + "nb=true";
    }
    return embedUrl;
  }

  // ──────────────────────────────────────────────
  // ドライブ内を全文検索（ファイル名 + コンテンツ）
  // ──────────────────────────────────────────────
  async function searchFiles(driveId, query) {
    const encodedQ = encodeURIComponent(query);
    const url = `${BASE}/drives/${driveId}/search(q='${encodedQ}')`
      + `?$select=id,name,parentReference,webUrl,lastModifiedDateTime,file`
      + `&$top=25`;
    const data = await graphFetch(url);
    return (data.value || [])
      .filter((item) => item.name.toLowerCase().endsWith(".pptx"))
      .map((item) => ({
        id:           item.id,
        name:         item.name,
        driveId:      driveId,
        webUrl:       item.webUrl || "",
        folderPath:   extractFolderPath(item.parentReference?.path || ""),
        lastModified: item.lastModifiedDateTime || "",
      }));
  }

  function extractFolderPath(rawPath) {
    const match = rawPath.match(/\/root:\/?(.*)$/);
    return match ? decodeURIComponent(match[1]) : rawPath;
  }

  window.Graph = { loadRootFolder, loadFolderChildren, getPreviewUrl, searchFiles };
})();
