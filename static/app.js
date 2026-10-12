/* 环洋 HYE 内部运价查询 */
(function () {
  // Expected SHA-256 hex of "username:password" (built offline; no plaintext secret here)
  const AUTH_HASH =
    "0d05643b48a2f23df7a0d19ced29f1c320467afc1199238a683d742aa3a02b58";
  const AUTH_KEY = "hye_rate_auth_v1";

  async function sha256Hex(text) {
    const data = new TextEncoder().encode(text);
    const buf = await crypto.subtle.digest("SHA-256", data);
    return [...new Uint8Array(buf)].map((b) => b.toString(16).padStart(2, "0")).join("");
  }

  function isAuthed() {
    try {
      return localStorage.getItem(AUTH_KEY) === "1";
    } catch (_) {
      return false;
    }
  }

  function setAuthed(ok) {
    try {
      if (ok) localStorage.setItem(AUTH_KEY, "1");
      else localStorage.removeItem(AUTH_KEY);
    } catch (_) {}
  }

  function showApp(show) {
    const gate = document.getElementById("loginGate");
    const shell = document.getElementById("appShell");
    if (gate) gate.classList.toggle("hidden", show);
    if (shell) shell.classList.toggle("hidden", !show);
  }

  async function tryLogin(user, pass) {
    const digest = await sha256Hex(`${user.trim()}:${pass}`);
    return digest === AUTH_HASH;
  }

  function wireAuth(onReady) {
    const form = document.getElementById("loginForm");
    const err = document.getElementById("loginError");
    const btnLogout = document.getElementById("btnLogout");

    if (isAuthed()) {
      showApp(true);
      onReady();
    } else {
      showApp(false);
    }

    if (form) {
      form.addEventListener("submit", async (e) => {
        e.preventDefault();
        const user = document.getElementById("loginUser").value;
        const pass = document.getElementById("loginPass").value;
        const ok = await tryLogin(user, pass);
        if (!ok) {
          if (err) err.classList.remove("hidden");
          return;
        }
        if (err) err.classList.add("hidden");
        setAuthed(true);
        showApp(true);
        onReady();
      });
    }

    if (btnLogout) {
      btnLogout.addEventListener("click", () => {
        setAuthed(false);
        showApp(false);
        const u = document.getElementById("loginUser");
        const p = document.getElementById("loginPass");
        if (u) u.value = "";
        if (p) p.value = "";
        if (err) err.classList.add("hidden");
      });
    }
  }

  const state = {
    index: null,
    meta: null,
    fbaCodes: [],
    cache: {}, // category -> rates[]
  };

  const $ = (id) => document.getElementById(id);
  const els = {
    category: $("category"),
    product: $("product"),
    warehouse: $("warehouse"),
    taxMode: $("taxMode"),
    zip: $("zip"),
    zone: $("zone"),
    fba: $("fba"),
    weight: $("weight"),
    cbm: $("cbm"),
    form: $("queryForm"),
    resultArea: $("resultArea"),
    resultCount: $("resultCount"),
    zoneHint: $("zoneHint"),
    metaLine: $("metaLine"),
    disclaimer: $("disclaimer"),
    extraCard: $("extraCard"),
    extraTitle: $("extraTitle"),
    extraArea: $("extraArea"),
    fbaList: $("fbaList"),
  };

  function zipToZone(zip) {
    const z = parseInt(String(zip).trim().slice(0, 5), 10);
    if (Number.isNaN(z)) return null;
    if (z >= 0 && z <= 39999) return "美东";
    if (z >= 40000 && z <= 79999) return "美中";
    if (z >= 80000 && z <= 99999) return "美西";
    return null;
  }

  function matchesWeightBracket(row, weight, cbm) {
    // If user didn't enter weight/cbm, keep all
    const hasW = weight != null && weight !== "" && !Number.isNaN(weight);
    const hasC = cbm != null && cbm !== "" && !Number.isNaN(cbm);

    if (row.unit === "元/CBM" || (row.cbm_min != null && row.weight_min_kg == null)) {
      if (hasC) {
        if (row.cbm_min != null && cbm < row.cbm_min) return false;
      }
      return true;
    }
    if (hasW) {
      if (row.weight_min_kg != null && weight < row.weight_min_kg) return false;
      if (row.weight_max_kg != null && weight > row.weight_max_kg) return false;
    }
    return true;
  }

  function estimateTotal(row, weight, cbm) {
    const price = row.unit_price;
    if (price == null) return null;
    if (row.flat_rate || row.unit === "元/票") return price;
    if (row.unit === "元/CBM") {
      if (cbm != null && cbm !== "" && !Number.isNaN(cbm)) return +(price * cbm).toFixed(2);
      return null;
    }
    // 元/KG
    if (weight != null && weight !== "" && !Number.isNaN(weight)) {
      return +(price * weight).toFixed(2);
    }
    return null;
  }

  function codeMatch(row, fba) {
    if (!fba) return true;
    const q = fba.trim().toUpperCase();
    const codes = row.warehouse_codes || [];
    return codes.some((c) => String(c).toUpperCase() === q || String(c).toUpperCase().includes(q));
  }

  function zoneMatch(row, zone) {
    if (!zone) return true;
    if (!row.zone) return true; // 海卡等无分区时不因分区过滤掉
    return row.zone === zone;
  }

  async function loadJSON(path) {
    const res = await fetch(path);
    if (!res.ok) throw new Error("无法加载 " + path);
    return res.json();
  }

  async function ensureCategory(cat) {
    if (state.cache[cat]) return state.cache[cat];
    const path = `data/parsed/rates_${cat}.json`;
    els.resultArea.innerHTML = `<div class="empty">正在加载「${cat}」价格数据…</div>`;
    const rows = await loadJSON(path);
    state.cache[cat] = rows;
    return rows;
  }

  function fillCategories() {
    const cats = state.index.categories || [];
    els.category.innerHTML =
      `<option value="">请选择</option>` +
      cats.map((c) => `<option value="${c}">${c}</option>`).join("");
  }

  function fillProducts() {
    const cat = els.category.value;
    const list = (state.index.products_by_category || {})[cat] || [];
    els.product.innerHTML =
      `<option value="">全部</option>` +
      list.map((p) => `<option value="${escapeAttr(p)}">${p}</option>`).join("");
  }

  function escapeAttr(s) {
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;");
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;");
  }

  function updateZoneHint() {
    const zip = els.zip.value.trim();
    if (zip.length >= 5) {
      const z = zipToZone(zip);
      if (z) {
        els.zoneHint.textContent = `邮编 ${zip} → 分区：${z}（${(state.index.zip_zones || {})[z] || ""}）`;
        if (!els.zone.value) {
          // soft suggest — don't overwrite manual choice unless empty
        }
        return;
      }
    }
    els.zoneHint.textContent = "";
  }

  function dedupeKey(r) {
    return [
      r.product,
      r.channel_name,
      r.warehouse,
      r.tax_mode,
      r.zone,
      r.weight_bracket,
      r.unit,
      r.unit_price,
      (r.warehouse_codes || []).join("/"),
      r.region_label || "",
    ].join("|");
  }

  function query() {
    const cat = els.category.value;
    if (!cat) {
      alert("请先选择渠道大类");
      return;
    }
    const product = els.product.value;
    const warehouse = els.warehouse.value;
    const taxMode = els.taxMode.value;
    const zip = els.zip.value.trim();
    let zone = els.zone.value;
    const fba = els.fba.value.trim();
    const weight = els.weight.value === "" ? null : parseFloat(els.weight.value);
    const cbm = els.cbm.value === "" ? null : parseFloat(els.cbm.value);

    if (!zone && zip.length >= 5) {
      zone = zipToZone(zip) || "";
    }

    ensureCategory(cat)
      .then((rows) => {
        let filtered = rows.filter((r) => {
          if (product && r.product !== product && r.channel_name !== product) return false;
          if (warehouse && r.warehouse !== warehouse) return false;
          if (taxMode && r.tax_mode && r.tax_mode !== taxMode) return false;
          if (taxMode && !r.tax_mode && cat !== "小包裹") return false;
          if (!zoneMatch(r, zone)) return false;
          if (!codeMatch(r, fba)) return false;
          if (!matchesWeightBracket(r, weight, cbm)) return false;
          return true;
        });

        // Prefer 价格汇总 rows when FBA code search for 海卡 (cleaner per-code)
        if (fba && cat === "海卡") {
          const fromSummary = filtered.filter((r) => r.source_sheet === "价格汇总");
          if (fromSummary.length) filtered = fromSummary;
        }

        // Dedupe
        const seen = new Set();
        filtered = filtered.filter((r) => {
          const k = dedupeKey(r);
          if (seen.has(k)) return false;
          seen.add(k);
          return true;
        });

        // Sort by unit price asc
        filtered.sort((a, b) => (a.unit_price || 0) - (b.unit_price || 0));

        // Cap display
        const MAX = 200;
        const shown = filtered.slice(0, MAX);
        renderResults(shown, filtered.length, { weight, cbm, zone, zip, fba });
      })
      .catch((err) => {
        els.resultArea.innerHTML = `<div class="empty">加载失败：${escapeHtml(err.message)}<br>请用本地服务器打开（见 README）。</div>`;
      });
  }

  function renderResults(rows, total, ctx) {
    els.resultCount.textContent = total ? `共 ${total} 条` + (total > rows.length ? `（显示前 ${rows.length}）` : "") : "无匹配";
    if (!rows.length) {
      els.resultArea.innerHTML = `<div class="empty">没有匹配的价格行。可尝试放宽条件（如清空邮编/仓库代码/重量档）。${
        ctx.zone ? `<br>当前分区：${escapeHtml(ctx.zone)}` : ""
      }</div>`;
      return;
    }

    const head = `
      <div class="table-wrap"><table>
        <thead><tr>
          <th>产品/渠道</th>
          <th>交货仓</th>
          <th>税别</th>
          <th>分区/代码</th>
          <th>计费档</th>
          <th>单价</th>
          <th>估算总价</th>
          <th>时效</th>
          <th>备注</th>
        </tr></thead><tbody>
    `;

    const body = rows
      .map((r) => {
        const totalEst = estimateTotal(r, ctx.weight, ctx.cbm);
        const taxClass = r.tax_mode === "包税" ? "tax-bao" : r.tax_mode === "自税" ? "tax-zi" : "";
        const codes = (r.warehouse_codes || []).slice(0, 8).join(" / ");
        const zoneBits = [r.zone, r.region_label, codes || r.zip_range].filter(Boolean).join(" · ");
        return `<tr>
          <td>
            <div>${escapeHtml(r.product || r.channel_name || "")}</div>
            <div class="muted">${escapeHtml(r.source_sheet || "")}</div>
          </td>
          <td>${escapeHtml(r.warehouse || "—")}</td>
          <td>${r.tax_mode ? `<span class="tag ${taxClass}">${escapeHtml(r.tax_mode)}</span>` : "—"}</td>
          <td>${escapeHtml(zoneBits || "—")}</td>
          <td>${escapeHtml(r.weight_bracket || "—")}</td>
          <td class="price">${r.unit_price != null ? r.unit_price + " " + escapeHtml(r.unit || "") : "—"}</td>
          <td class="total">${totalEst != null ? "¥ " + totalEst : "<span class='muted'>—</span>"}</td>
          <td class="muted">${escapeHtml((r.transit_time || "—").split("\n")[0])}</td>
          <td class="notes">${escapeHtml((r.notes || "").slice(0, 180))}</td>
        </tr>`;
      })
      .join("");

    els.resultArea.innerHTML =
      head +
      body +
      `</tbody></table></div>
       <p class="muted" style="margin-top:10px">⚠ 以上为参考单价；实际计费可能含附加费、比重优惠、最低计费重等，请以业务确认为准。</p>`;
  }

  async function showSchedule() {
    try {
      if (!state.meta) state.meta = await loadJSON("data/parsed/meta.json");
      const rows = state.meta.schedule || [];
      els.extraTitle.textContent = "船期表（参考）";
      els.extraCard.classList.remove("hidden");
      if (!rows.length) {
        els.extraArea.innerHTML = "<div class='empty'>暂无船期数据</div>";
        return;
      }
      els.extraArea.innerHTML = `<div class="table-wrap"><table>
        <thead><tr><th>区域</th><th>渠道</th><th>起运港</th><th>ETD</th><th>目的港</th><th>ETA</th><th>船司</th><th>航次</th><th>截单</th></tr></thead>
        <tbody>${rows
          .map(
            (r) => `<tr>
          <td>${escapeHtml(r.region)}</td>
          <td>${escapeHtml(r.channel)}</td>
          <td>${escapeHtml(r.pol)}</td>
          <td>${escapeHtml(r.etd)}</td>
          <td>${escapeHtml(r.pod)}</td>
          <td>${escapeHtml(r.eta)}</td>
          <td>${escapeHtml(r.carrier)}</td>
          <td>${escapeHtml(r.vessel)}</td>
          <td class="muted">${escapeHtml(r.cutoff_shanghai || r.cutoff_yiwu || r.cutoff_bantan || "")}</td>
        </tr>`
          )
          .join("")}</tbody></table></div>`;
    } catch (e) {
      alert(e.message);
    }
  }

  async function showSurcharge() {
    try {
      if (!state.meta) state.meta = await loadJSON("data/parsed/meta.json");
      const rows = state.meta.surcharges || [];
      els.extraTitle.textContent = "海运品名附加费";
      els.extraCard.classList.remove("hidden");
      els.extraArea.innerHTML = `<div class="table-wrap"><table>
        <thead><tr><th>品名</th><th>费用</th><th>说明</th></tr></thead>
        <tbody>${rows
          .map(
            (r) => `<tr>
          <td>${escapeHtml(r.name)}</td>
          <td class="price">${escapeHtml(r.fee)}</td>
          <td class="notes">${escapeHtml(r.description)}</td>
        </tr>`
          )
          .join("")}</tbody></table></div>`;
    } catch (e) {
      alert(e.message);
    }
  }

  async function init() {
    try {
      state.index = await loadJSON("data/parsed/index.json");
      fillCategories();
      els.metaLine.textContent = `${state.index.company} · 已解析 ${state.index.rate_count} 条价格 · 更新于 ${state.index.parsed_at}`;
      els.disclaimer.textContent = state.index.disclaimer || "";

      // FBA datalist (optional)
      try {
        state.fbaCodes = await loadJSON("data/parsed/fba_codes.json");
        els.fbaList.innerHTML = state.fbaCodes
          .slice(0, 500)
          .map((c) => `<option value="${c}">`)
          .join("");
      } catch (_) {}
    } catch (e) {
      els.metaLine.textContent = "数据加载失败，请用本地服务器打开（见 README）";
      els.resultArea.innerHTML = `<div class="empty">${escapeHtml(e.message)}<br><br>
        请在项目目录执行：<code>python3 -m http.server 8765</code><br>
        然后浏览器打开 <code>http://127.0.0.1:8765/</code></div>`;
    }
  }

  els.category.addEventListener("change", fillProducts);
  els.zip.addEventListener("input", updateZoneHint);
  els.form.addEventListener("submit", (e) => {
    e.preventDefault();
    query();
  });
  $("btnReset").addEventListener("click", () => {
    els.form.reset();
    fillProducts();
    els.zoneHint.textContent = "";
    els.resultCount.textContent = "尚未查询";
    els.resultArea.innerHTML = "请选择渠道大类并点击「查询运价」。价格仅供参考。";
  });
  $("btnSchedule").addEventListener("click", showSchedule);
  $("btnSurcharge").addEventListener("click", showSurcharge);
  $("btnCloseExtra").addEventListener("click", () => els.extraCard.classList.add("hidden"));

  // Auto-switch to 海卡 when typing FBA-like code
  els.fba.addEventListener("change", () => {
    if (els.fba.value.trim() && !els.category.value) {
      const opt = [...els.category.options].find((o) => o.value === "海卡");
      if (opt) {
        els.category.value = "海卡";
        fillProducts();
      }
    }
  });

  let appStarted = false;
  function startAppOnce() {
    if (appStarted) return;
    appStarted = true;
    init();
  }

  wireAuth(startAppOnce);
})();
