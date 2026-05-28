(function () {
  const cfg = window.DASHBOARD_CONFIG;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Sites.Read.All"];

  const msalConfig = {
    auth: {
      clientId: cfg.clientId,
      authority: "https://login.microsoftonline.com/" + cfg.tenantId,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
  };
  const msalApp = new msal.PublicClientApplication(msalConfig);
  let account = null;
  let siteId = null;
  let listId = null;
  let chart = null;
  let rotationTimer = null;
  let currentPage = 0;

  document.addEventListener("DOMContentLoaded", async () => {
    startClock();
    document.getElementById("signin-btn").addEventListener("click", interactiveSignIn);
    try {
      if (typeof msalApp.initialize === "function") {
        await msalApp.initialize();
      }
      await handleRedirect();
      await ensureSignedIn();
      await refresh();
      setInterval(refresh, cfg.refreshMinutes * 60 * 1000);
    } catch (err) {
      console.error("Bootstrap failed", err);
      showError(err.message || String(err));
    }
  });

  async function handleRedirect() {
    const response = await msalApp.handleRedirectPromise();
    if (response && response.account) {
      msalApp.setActiveAccount(response.account);
    }
  }

  async function ensureSignedIn() {
    const accounts = msalApp.getAllAccounts();
    if (accounts.length > 0) {
      account = accounts[0];
      msalApp.setActiveAccount(account);
      return;
    }
    try {
      const result = await msalApp.ssoSilent({ scopes: SCOPES });
      account = result.account;
      msalApp.setActiveAccount(account);
    } catch (e) {
      document.getElementById("signin-overlay").classList.add("show");
      throw new Error("Sign-in required");
    }
  }

  async function interactiveSignIn() {
    try {
      await msalApp.loginRedirect({ scopes: SCOPES });
    } catch (e) {
      console.error(e);
      showError("Sign-in failed: " + (e.message || e));
    }
  }

  async function getToken() {
    const req = { scopes: SCOPES, account };
    try {
      const result = await msalApp.acquireTokenSilent(req);
      return result.accessToken;
    } catch (e) {
      await msalApp.acquireTokenRedirect(req);
      return null;
    }
  }

  async function graphGet(path) {
    const token = await getToken();
    if (!token) throw new Error("No token");
    const res = await fetch(GRAPH + path, {
      headers: { Authorization: "Bearer " + token, Accept: "application/json" }
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error("Graph " + res.status + ": " + text);
    }
    return res.json();
  }

  async function resolveSiteAndList() {
    if (siteId && listId) return;
    const hostname = cfg.sharePointHostname + ".sharepoint.com";
    const sitePath = cfg.sitePath && cfg.sitePath !== "/" ? ":" + cfg.sitePath : "";
    const site = await graphGet("/sites/" + hostname + sitePath);
    siteId = site.id;
    const listFilter = "displayName eq '" + encodeURIComponent(cfg.listName).replace(/'/g, "''") + "'";
    const lists = await graphGet("/sites/" + siteId + "/lists?$filter=" + listFilter);
    if (!lists.value || lists.value.length === 0) {
      throw new Error('SharePoint list "' + cfg.listName + '" not found on site.');
    }
    listId = lists.value[0].id;
  }

  async function fetchItems() {
    await resolveSiteAndList();
    const items = [];
    let url = "/sites/" + siteId + "/lists/" + listId + "/items?$expand=fields&$top=999";
    while (url) {
      const page = await graphGet(url);
      for (const item of page.value || []) items.push(item.fields || {});
      url = page["@odata.nextLink"]
        ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
        : null;
    }
    return items;
  }

  async function refresh() {
    try {
      setStatus(true, "Refreshing…");
      const rows = await fetchItems();
      const jobs = rows.map(normaliseJob);
      renderKPIs(jobs);
      renderFunnel(jobs);
      renderTable(jobs);
      setStatus(true, "Connected");
      document.getElementById("last-updated").textContent =
        "Last updated: " + new Date().toLocaleTimeString(cfg.locale);
      clearError();
    } catch (err) {
      console.error("Refresh failed", err);
      setStatus(false, "Error");
      showError(err.message || String(err));
    }
  }

  function normaliseJob(fields) {
    const c = cfg.columns;
    return {
      customer:       fields[c.customer]       || "",
      jobNo:          fields[c.jobNo]          || "",
      jobName:        fields[c.jobName]        || "",
      projectStatus:  (fields[c.projectStatus] || "").toString(),
      projectType:    (fields[c.projectType]   || "").toString(),
      enquiryDate:    parseDate(fields[c.enquiryDate]),
      quoteDate:      parseDate(fields[c.quoteDate]),
      orderDate:      parseDate(fields[c.orderDate]),
      orderDueDate:   parseDate(fields[c.orderDueDate]),
      productionCompleteDate: parseDate(fields[c.productionCompleteDate]),
      declinedDate:   parseDate(fields[c.declinedDate]),
      declinedReason: fields[c.declinedReason] || "",
      netValue:       parseMoney(fields[c.netValue]),
      liveProject:    isYes(fields[c.liveProject]),
      pricedBy:       fields[c.pricedBy] || ""
    };
  }

  function parseDate(v) {
    if (!v) return null;
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  function parseMoney(v) {
    if (v == null || v === "") return 0;
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g, ""));
    return isNaN(n) ? 0 : n;
  }
  function isYes(v) {
    if (v === true || v === 1) return true;
    const s = String(v == null ? "" : v).toLowerCase().trim();
    return s === "yes" || s === "true" || s === "1";
  }

  function classify(j) {
    if (j.declinedDate) return "declined";
    if (j.productionCompleteDate) return "complete";
    if (j.orderDate)              return "ordered";
    if (j.quoteDate)              return "quoted";
    if (j.enquiryDate)            return "enquiry";
    return "none";
  }

  function renderKPIs(jobs) {
    const today = startOfDay(new Date());
    const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

    const openEnquiries = jobs.filter(function (j) { return classify(j) === "enquiry"; }).length;
    const quotesOut     = jobs.filter(function (j) { return classify(j) === "quoted"; });
    const wonThisMonth  = jobs.filter(function (j) { return j.orderDate && j.orderDate >= startOfMonth; });

    const pipelineValue = quotesOut.reduce(function (s, j) { return s + (j.netValue || 0); }, 0);
    const wonValue      = wonThisMonth.reduce(function (s, j) { return s + (j.netValue || 0); }, 0);

    setText("kpi-enquiries", openEnquiries);
    setText("kpi-quotes",    quotesOut.length);
    setText("kpi-won",       wonThisMonth.length);
    setText("kpi-pipeline",  formatCurrency(pipelineValue));

    setText("kpi-quotes-sub",   pipelineValue ? "worth " + formatCurrency(pipelineValue) : "awaiting decision");
    setText("kpi-won-sub",      wonValue ? "worth " + formatCurrency(wonValue) : "this month");
    setText("kpi-pipeline-sub", quotesOut.length + " open " + (quotesOut.length === 1 ? "quote" : "quotes"));
  }

  function formatCurrency(n) {
    const sym = cfg.currencySymbol || "£";
    const rounded = Math.round(n);
    if (Math.abs(rounded) >= 1000000) {
      return sym + (rounded / 1000000).toFixed(1).replace(/\.0$/, "") + "m";
    }
    if (Math.abs(rounded) >= 10000) {
      return sym + (rounded / 1000).toFixed(0) + "k";
    }
    return sym + rounded.toLocaleString(cfg.locale || "en-GB");
  }

  function renderFunnel(jobs) {
    const buckets = { enquiry: 0, quoted: 0, ordered: 0, complete30: 0 };
    const today = startOfDay(new Date());
    const thirtyDaysAgo = new Date(today); thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

    for (const j of jobs) {
      const s = classify(j);
      if (s === "enquiry") buckets.enquiry++;
      else if (s === "quoted") buckets.quoted++;
      else if (s === "ordered") buckets.ordered++;
      else if (s === "complete" && j.productionCompleteDate >= thirtyDaysAgo) buckets.complete30++;
    }

    const labels = ["Open Enquiries", "Quotes Out", "Active Orders", "Completed (30d)"];
    const data   = [buckets.enquiry, buckets.quoted, buckets.ordered, buckets.complete30];
    const colors = ["#60a5fa", "#fbbf24", "#a78bfa", "#4ade80"];

    const ctx = document.getElementById("funnel-chart");
    if (chart) chart.destroy();
    chart = new Chart(ctx, {
      type: "bar",
      data: {
        labels: labels,
        datasets: [{ data: data, backgroundColor: colors, borderRadius: 8, borderSkipped: false }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false }, tooltip: { enabled: false } },
        scales: {
          x: { beginAtZero: true, ticks: { color: "#93a0bd", font: { size: 14 } }, grid: { color: "#1d2540" } },
          y: { ticks: { color: "#e8ecf3", font: { size: 16, weight: "600" } }, grid: { display: false } }
        },
        animation: { duration: 600 }
      },
      plugins: [{
        id: "valueLabels",
        afterDatasetsDraw: function (c) {
          const ctx2 = c.ctx;
          c.data.datasets[0].data.forEach(function (v, i) {
            const bar = c.getDatasetMeta(0).data[i];
            if (!bar) return;
            ctx2.save();
            ctx2.fillStyle = "#e8ecf3";
            ctx2.font = "bold 18px sans-serif";
            ctx2.textBaseline = "middle";
            ctx2.fillText(v, bar.x + 8, bar.y);
            ctx2.restore();
          });
        }
      }]
    });
  }

  function renderTable(jobs) {
    const today = startOfDay(new Date());
    const openQuotes = jobs
      .filter(function (j) { return classify(j) === "quoted"; })
      .sort(function (a, b) {
        if (!a.quoteDate && !b.quoteDate) return 0;
        if (!a.quoteDate) return 1;
        if (!b.quoteDate) return -1;
        return a.quoteDate - b.quoteDate;
      });

    const totalValue = openQuotes.reduce(function (s, j) { return s + (j.netValue || 0); }, 0);
    setText("table-footer",
      openQuotes.length
        ? openQuotes.length + " open " + (openQuotes.length === 1 ? "quote" : "quotes") + " · total " + formatCurrency(totalValue)
        : "No open quotes awaiting decision."
    );

    const pageSize = cfg.rowsInTable;
    const pages = [];
    for (let i = 0; i < openQuotes.length; i += pageSize) {
      pages.push(openQuotes.slice(i, i + pageSize));
    }
    if (pages.length === 0) pages.push([]);

    const drawPage = function (idx) {
      const tbody = document.getElementById("jobs-rows");
      const rows = pages[idx % pages.length];
      if (rows.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" style="color:#6b7896;padding:24px 8px;">No open quotes awaiting decision.</td></tr>';
        return;
      }
      tbody.innerHTML = rows.map(function (j) {
        const daysOut = j.quoteDate ? Math.floor((today - startOfDay(j.quoteDate)) / 86400000) : null;
        const ageClass = daysOut == null ? ""
                      : daysOut >= 21 ? "overdue"
                      : daysOut >= 10 ? "due-soon"
                      : "";
        const ageText = j.quoteDate
          ? j.quoteDate.toLocaleDateString(cfg.locale) + ' <span class="age">(' + daysOut + 'd)</span>'
          : "—";
        return '<tr>'
          + '<td class="mono">' + escapeHtml(String(j.jobNo || "")) + '</td>'
          + '<td>' + escapeHtml(j.customer) + '</td>'
          + '<td>' + escapeHtml(j.jobName) + '</td>'
          + '<td class="num">' + (j.netValue ? formatCurrency(j.netValue) : "—") + '</td>'
          + '<td class="' + ageClass + '">' + ageText + '</td>'
          + '</tr>';
      }).join("");
    };

    if (rotationTimer) clearInterval(rotationTimer);
    currentPage = 0;
    drawPage(currentPage);
    if (cfg.rotateTable && pages.length > 1) {
      rotationTimer = setInterval(function () {
        currentPage = (currentPage + 1) % pages.length;
        drawPage(currentPage);
      }, cfg.pageRotateSeconds * 1000);
    }
  }

  function startOfDay(d) {
    const x = new Date(d);
    x.setHours(0, 0, 0, 0);
    return x;
  }
  function setText(id, v) {
    const el = document.getElementById(id);
    if (el) el.textContent = v;
  }
  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function startClock() {
    const tick = function () {
      const now = new Date();
      setText("clock", now.toLocaleTimeString(cfg.locale, { hour: "2-digit", minute: "2-digit" }));
      setText("date",  now.toLocaleDateString(cfg.locale, { weekday: "long", day: "numeric", month: "long", year: "numeric" }));
    };
    tick();
    setInterval(tick, 1000 * 30);
  }

  function setStatus(ok, text) {
    const el = document.getElementById("conn-status");
    el.classList.toggle("error", !ok);
    setText("conn-status-text", text);
  }
  function showError(msg) {
    const el = document.getElementById("error-banner");
    el.textContent = msg;
    el.classList.add("show");
  }
  function clearError() {
    const el = document.getElementById("error-banner");
    el.textContent = "";
    el.classList.remove("show");
  }
})();
