(function () {
  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("fileInput");
  const dropFileName = document.getElementById("dropFileName");
  const panelFile = document.getElementById("panel-file");
  const panelPaste = document.getElementById("panel-paste");
  const switchBtn = document.getElementById("switchMode");
  const subtext = document.getElementById("subtext");
  const statusEl = document.getElementById("status");
  const convertBtn = document.getElementById("convertBtn");
  const uploadLabel = panelFile.querySelector(".upload-btn");
  const uploadLabelText = uploadLabel.querySelector("span:last-child");
  const defaultUploadLabel = uploadLabelText.textContent;

  let mode = "file";

  function setMode(next) {
    mode = next;
    if (mode === "file") {
      panelFile.classList.remove("hidden");
      panelPaste.classList.add("hidden");
      subtext.textContent = "HTMLファイルをアップロードしてください";
      switchBtn.textContent = "HTMLを直接貼り付ける";
    } else {
      panelFile.classList.add("hidden");
      panelPaste.classList.remove("hidden");
      subtext.textContent = "HTMLソースを貼り付けて変換します";
      switchBtn.textContent = "ファイルをアップロードする";
    }
    setStatus("", null);
  }

  switchBtn.addEventListener("click", () => setMode(mode === "file" ? "paste" : "file"));

  function showSelectedFile(file) {
    if (!file) {
      dropFileName.textContent = "";
      dropFileName.classList.remove("show");
      return;
    }
    dropFileName.textContent = `📄 ${file.name}`;
    dropFileName.classList.add("show");
  }

  function setStatus(msg, kind) {
    statusEl.textContent = msg;
    statusEl.className = "status" + (kind ? ` ${kind}` : "");
  }

  function setLoading(isLoading) {
    uploadLabel.style.pointerEvents = isLoading ? "none" : "";
    uploadLabel.style.opacity = isLoading ? "0.6" : "";
    uploadLabelText.textContent = isLoading ? "変換中..." : defaultUploadLabel;
    if (convertBtn) {
      convertBtn.disabled = isLoading;
      const spinner = convertBtn.querySelector(".btn-spinner");
      const label = convertBtn.querySelector(".btn-label");
      if (spinner) spinner.hidden = !isLoading;
      if (label) label.textContent = isLoading ? "変換中..." : "PPTXに変換してダウンロード";
    }
  }

  async function runConvert({ file, html }) {
    let body;
    const headers = {};

    if (file) {
      const form = new FormData();
      form.append("file", file);
      body = form;
    } else {
      if (!html || !html.trim()) return setStatus("HTMLを貼り付けてください。", "error");
      body = JSON.stringify({ html });
      headers["Content-Type"] = "application/json";
    }

    setLoading(true);
    setStatus("変換中です。初回アクセス直後はサーバーの起動に時間がかかることがあります。", "info");

    try {
      const res = await fetch("/api/convert", { method: "POST", headers, body });
      if (!res.ok) {
        const err = await res.json().catch(() => ({ error: res.statusText }));
        throw new Error(err.error || "変換に失敗しました。");
      }
      const route = res.headers.get("X-Conversion-Route");
      const warnings = res.headers.get("X-Conversion-Warnings");
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "converted.pptx";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
      setStatus(
        `✅ 変換が完了しました (${route}${warnings && warnings !== "0" ? ` / warnings: ${warnings}` : ""})`,
        "success"
      );
    } catch (e) {
      setStatus(`⚠️ エラー: ${e.message}`, "error");
    } finally {
      setLoading(false);
    }
  }

  fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    if (!file) return;
    showSelectedFile(file);
    runConvert({ file });
  });

  ["dragenter", "dragover"].forEach((evt) =>
    dropzone.addEventListener(evt, (e) => {
      e.preventDefault();
      dropzone.classList.add("dragover");
    })
  );
  ["dragleave", "dragend"].forEach((evt) =>
    dropzone.addEventListener(evt, () => dropzone.classList.remove("dragover"))
  );
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    const file = e.dataTransfer.files[0];
    if (!file) return;
    fileInput.files = e.dataTransfer.files;
    showSelectedFile(file);
    runConvert({ file });
  });

  if (convertBtn) {
    convertBtn.addEventListener("click", () => {
      const html = document.getElementById("htmlInput").value;
      runConvert({ html });
    });
  }
})();
