/* ============================================================
   Sales Pipeline Dashboard
   - Fetches SharePoint list items via Microsoft Graph
   - KPIs: Open Enquiries, Quotes Out, Won This Month, Pipeline £
   - Chart: pipeline funnel (horizontal bar)
   - Table: open quotes awaiting decision, rotating pages
   ============================================================ */

(function () {
  const cfg = window.DASHBOARD_CONFIG;
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const SCOPES = ["Sites.Read.All"];

  // ---------- MSAL setup -----------------------------------
  const msalConfig = {
    auth: {
      clientId: cfg.clientId,
      authority: `https://login.microsoftonline.com/${cfg.tenantId}`,
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false
    }
  };
  const msalApp = new msal.PublicClientApplication(msalConfig);
  let account = null;
  let siteId = null;
  let listId = null;
  let chart = null;
  let rotationTimer = null;
  let currentPage = 0;

  // ---------- Bootstrap ------------------------------------
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

  // ---------- Graph helpers --------------------------------
  async function graphGet(path) {
    const token = await getToken();
    if (!token) throw new Error("No token");
    const res = await fetch(GRAPH + path, {
      headers: { Authorization: "Bearer " + token, Accept: "application/json" }
    });
    if (!res.ok) {
      const text = await res.text();
      throw new Error(`Graph ${res.status}: ${text}`);
    }
    return res.json();
  }

  async function resolveSiteAndList() {
    if (siteId && listId) return;
    const hostname = `${cfg.sharePointHostname}.sharepoint.com`;
    const sitePath = cfg.sitePath && cfg.sitePath !== "/" ? `:${cfg.sitePath}` : "";
    const site = await graphGet(`/sites/${hostname}${sitePath}`);
    siteId = site.id;

    const lists = await graphGet(
      `/sites/${siteId}/lists?$filter=displayName eq '${encodeURIComponent(cfg.listName).replace(/'/g, "''")}'`
    );
    if (!lists.value || lists.value.length === 0) {
      throw new Error(`SharePoint list "${cfg.listName}" not found on site.`);
    }
    listId = lists.value[0].id;
  }

  async function fetchItems() {
    await resolveSiteAndList();
    const items = [];
    let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=999`;
    while (url) {
      const page = await graphGet(url);
      for (const item of page.value || []) items.push(item.fields || {});
      url = page["@odata.nextLink"]
        ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
        : null;
    }
    return items;
  }

  // ---------- Refresh + normalise --------------------------
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
      customer:       fields[c.customer]       ?? "",
      jobNo:          fields[c.jobNo]          ?? "",
      jobName:        fields[c.jobName]        ?? "",
      projectStatus:  (fields[c.projectStatus] ?? "").toString(),
      projectType:    (fields[c.projectType]   ?? "").toString(),
      enquiryDate:    parseDate(fields[c.enquiryDate]),
      quoteDate:      parseDate(fields[c.quoteDate]),
      orderDate:      parseDate(fields[c.orderDate]),
      orderDueDate:   parseDate(fields[c.orderDueDate]),
      productionCompleteDate: parseDate(fields[c.productionCompleteDate]),
      declinedDate:   parseDate(fields[c.declinedDate]),
      declinedReason: fields[c.declinedReason] ?? "",
      netValue:       parseMoney(fields[c.netValue]),
      liveProject:    isYes(fields[c.liveProject]),
      pricedBy:       fields[c.pricedBy] ?? ""
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
    const s = String(v ?? "").toLowerCase().trim();
    return s === "yes" || s === "true" || s === "1";
  }

  // ---------- Pipeline-stage classification ----------------
  // A row is in exactly one stage at a time, in this order of precedence:
  //   declined  -> not shown anywhere except as exclusion
  //   complete  -> production complete date set
  //   ordered   -> order placed, not yet complete
  //   quoted    -> quote issued, not yet ordered/declined
  //   enquiry   -> enquiry received, not yet quoted/declined
  //   none      -> no enquiry yet (shouldn't normally happen)
  function classify(j) {
    if (j.declinedDate) return "declined";
    if (j.productionCompleteDate) return "complete";
    if (j.orderDate)              return "ordered";
    if (j.quoteDate)              return "quoted";
    if (j.enquiryDate)            return "enquiry";
    return "none";
  }

  // ---------- KPI tiles ------------------------------------
  function renderKPIs(jobs) {
    const today = startOfDay(new Date());
    const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);

    const openEnquiries = jobs.filter(j => classify(j) === "enquiry").length;
    const quotesOut     = jobs.filter(j => classify(j) === "quoted");
    const wonThisMonth  = jobs.filter(j =>
      j.orderDate && j.orderDate >= startOfMonth
    );

    const pipelineValue = quotesOut.reduce((s, j) => s + (j.netValue || 0), 0);
    const wonValue      = wonThisMonth.reduce((s, j) => s + (j.netValue || 0), 0);

    setText("kpi-enquiries", openEnquiries);
    setText("kpi-quotes",    quotesOut.length);
    setText("kpi-won",       wonThisMonth.length);
    setText("kpi-pipeline",  formatCurrency(pipelineValue));

    setText("kpi-quotes-sub",   pipelineValue ? `worth ${formatCurrency(pipelineValue)}` : "awaiting decision");
    setText("kpi-won-sub",      wonValue ? `worth ${formatCurrency(wonValue)}` : "this month");
    setText("kpi-pipeline-sub", `${quotesOut.length} open ${quotesOut.length === 1 ? "quote" : "quotes"}`);
  }

  function formatCurrency(n) {
    const sym = cfg.currencySymbol || "£";
    const rounded = Math.round(n);
    if (Math.abs(rounded) >= 1_000_000) {
      return sym + (rounded / 1_000_000).toFixed(1).replace(/\.0$/, "") + "m";
    }
    if (Math.abs(rounded) >= 10_000) {
      return sym + (rounded / 1000).toFixed(0) + "k";
    }
    return sym + rounded.toLocaleString(cfg.locale || "en-GB");
  }

  // ---------- Funnel chart ---------------------------------
  function renderFunnel(jobs) {
    // Counts of currently-live items at each stage
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
        labels,
        datasets: [{
          data,
          backgroundColor: colors,
          borderRadius: 8,
          borderSkipped: false
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: { enabled: false },
          datalabels: { display: false }
        },
        scales: {
          x: {
            beginAtZero: true,
            ticks: { color: "#93a0bd", font: { size: 14 } },
            grid:  { color: "#1d2540" }
          },
          y: {
            ticks: { color: "#e8ecf3", font: { size: 16, weight: "600" } },
            grid:  { display: false }
          }
        },
        animation: { duration: 600 }
      },
      plugins: [{
        id: "valueLabels",
        afterDatasetsDraw(c) {
          const { ctx } = c;
          c.data.datasets[0].data.forEach((v, i) => {
            const bar = c.getDatasetMeta(0).data[i];
            if (!bar) return;
            ctx.save();
            ctx.fillStyle = "#e8ecf3";
            ctx.font = "bold 18px sans-serif";
            ctx.textBaseline = "middle";
            ctx.fillText(v, bar.x + 8, bar.y);
            ctx.restore();
          });
        }
      }]
    });
  }

  // ---------- Table: quotes awaiting decis