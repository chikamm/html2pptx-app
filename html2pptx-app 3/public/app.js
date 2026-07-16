(function () {
  const tabs = document.querySelectorAll(".tab");
  const panels = { file: document.getElementById("panel-file"), paste: document.getElementById("panel-paste") };
  tabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      tabs.forEach((t) => t.classList.remove("active"));
      tab.classList.add("active");
      Object.entries(panels).forEach(([key, el]) => el.classList.toggle("hidden", key !== tab.dataset.tab));
    });
  });

  const statusEl = document.getElementById("status");
  const btn = document.getElementById("convertBtn");

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  btn.addEventListener("click", async () => {
    const activeTab = document.querySelector(".tab.active").dataset.tab;
    let body;
    let headers = {};

    if (activeTab === "file") {
      const file = document.getElementById("fileInput").files[0];
      if (!file) return setStatus("HTMLファイルを選択してください。");
      const form = new FormData();
      form.append("file", file);
      body = form;
    } else {
      const html = document.getElementById("htmlInput").value;
      if (!html.trim()) return setStatus("HTMLを貼り付けてください。");
      body = JSON.stringify({ html });
      headers["Content-Type"] = "application/json";
    }

    btn.disabled = true;
    setStatus("変換中... (初回はChromiumの起動に少し時間がかかります)");

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
      setStatus(`完了 (${route}${warnings && warnings !== "0" ? ` / warnings: ${warnings}` : ""})`);
    } catch (e) {
      setStatus(`エラー: ${e.message}`);
    } finally {
      btn.disabled = false;
    }
  });
})();
