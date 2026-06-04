(function () {
  const loginBtn = document.getElementById("loginBtn");
  const logoutBtn = document.getElementById("logoutBtn");
  const loadBtn = document.getElementById("loadBtn");
  const expandAllBtn = document.getElementById("expandAllBtn");
  const collapseAllBtn = document.getElementById("collapseAllBtn");
  const rootUpnInput = document.getElementById("rootUpn");
  const maxDepthInput = document.getElementById("maxDepth");
  const includeManagerCheckbox = document.getElementById("includeManager");
  const searchInput = document.getElementById("searchInput");
  const statusText = document.getElementById("statusText");
  const signedInText = document.getElementById("signedInText");
  const chartContainer = document.getElementById("chartContainer");

  const personCache = new Map();
  const expanded = new Set();

  let msalApp = null;
  let activeAccount = null;
  let graphToken = null;
  let orgTree = null;
  let searchTerm = "";

  const cfg = window.APP_CONFIG;

  if (!cfg || !cfg.auth || !cfg.graph) {
    setStatus("Missing config. Open config.js and set APP_CONFIG values.", true);
    return;
  }

  if (!cfg.auth.clientId || cfg.auth.clientId === "PUT_YOUR_CLIENT_ID_HERE") {
    setStatus("Set your Microsoft Entra clientId in config.js before signing in.", true);
  }

  initialize().catch(function (err) {
    console.error(err);
    setStatus("Initialization failed: " + safeMessage(err), true);
  });

  async function initialize() {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: cfg.auth.clientId,
        authority: "https://login.microsoftonline.com/" + cfg.auth.tenantId,
        redirectUri: cfg.auth.redirectUri
      },
      cache: {
        cacheLocation: "localStorage"
      }
    });

    await msalApp.initialize();

    const redirectResult = await msalApp.handleRedirectPromise();
    if (redirectResult && redirectResult.account) {
      activeAccount = redirectResult.account;
      msalApp.setActiveAccount(activeAccount);
    }

    const accounts = msalApp.getAllAccounts();
    if (!activeAccount && accounts.length > 0) {
      activeAccount = accounts[0];
      msalApp.setActiveAccount(activeAccount);
    }

    wireEvents();
    refreshAuthUI();
  }

  function wireEvents() {
    loginBtn.addEventListener("click", onLogin);
    logoutBtn.addEventListener("click", onLogout);
    loadBtn.addEventListener("click", onLoadChart);
    expandAllBtn.addEventListener("click", function () {
      expandCollapseAll(true);
      renderTree();
    });
    collapseAllBtn.addEventListener("click", function () {
      expandCollapseAll(false);
      renderTree();
    });
    searchInput.addEventListener("input", function (e) {
      searchTerm = (e.target.value || "").trim().toLowerCase();
      autoExpandForSearch();
      renderTree();
    });
  }

  async function onLogin() {
    if (!isConfigured()) {
      setStatus("Set your clientId in config.js first.", true);
      return;
    }

    try {
      await msalApp.loginPopup({ scopes: cfg.graph.scopes });
      const accounts = msalApp.getAllAccounts();
      activeAccount = accounts[0] || null;
      if (activeAccount) {
        msalApp.setActiveAccount(activeAccount);
      }
      await ensureToken();
      setStatus("Signed in. You can load your org chart.", false);
      refreshAuthUI();
    } catch (err) {
      setStatus("Sign-in failed: " + safeMessage(err), true);
    }
  }

  async function onLogout() {
    try {
      await msalApp.logoutPopup({ account: activeAccount || undefined });
    } catch (err) {
      console.warn("Logout warning", err);
    }

    activeAccount = null;
    graphToken = null;
    orgTree = null;
    expanded.clear();
    personCache.clear();
    searchTerm = "";
    searchInput.value = "";

    refreshAuthUI();
    clearChart("Signed out. Sign in to load your org chart.");
  }

  async function onLoadChart() {
    if (!activeAccount) {
      setStatus("Sign in first.", true);
      return;
    }

    try {
      setStatus("Loading organization data...", false);
      await ensureToken();

      const upn = rootUpnInput.value.trim();
      const depth = clampDepth(maxDepthInput.value);
      maxDepthInput.value = String(depth);

      let rootNode = upn
        ? await getUserByUpn(upn)
        : await graphGet("/me?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation");

      if (!rootNode || !rootNode.id) {
        throw new Error("Could not resolve root user. Check the UPN or permissions.");
      }

      rootNode = normalizePerson(rootNode);

      const includeManagerChain = includeManagerCheckbox.checked;
      let managerChain = [];
      if (includeManagerChain) {
        managerChain = await fetchManagerChain(rootNode.id);
      }

      const subtree = await buildSubtree(rootNode.id, 0, depth);

      orgTree = attachManagerChain(managerChain, subtree);
      expanded.clear();
      expandDefaults(orgTree, 0);
      autoExpandForSearch();
      renderTree();

      const total = countNodes(orgTree);
      setStatus("Loaded " + total + " people.", false);
    } catch (err) {
      console.error(err);
      setStatus("Failed to load org chart: " + safeMessage(err), true);
      clearChart("Unable to load data. Confirm Graph permissions and admin consent.");
    }
  }

  function refreshAuthUI() {
    const signedIn = !!activeAccount;
    loginBtn.disabled = signedIn;
    logoutBtn.disabled = !signedIn;
    loadBtn.disabled = !signedIn;
    expandAllBtn.disabled = !signedIn;
    collapseAllBtn.disabled = !signedIn;
    searchInput.disabled = !signedIn;

    if (signedIn) {
      signedInText.textContent = "Signed in as " + (activeAccount.username || activeAccount.name || "account");
    } else {
      signedInText.textContent = "Not signed in";
    }
  }

  async function ensureToken() {
    const account = activeAccount || msalApp.getActiveAccount();
    if (!account) {
      throw new Error("No signed-in account.");
    }

    try {
      const result = await msalApp.acquireTokenSilent({
        account: account,
        scopes: cfg.graph.scopes
      });
      graphToken = result.accessToken;
      return graphToken;
    } catch (err) {
      if (err instanceof msal.InteractionRequiredAuthError) {
        const result = await msalApp.acquireTokenPopup({ scopes: cfg.graph.scopes });
        graphToken = result.accessToken;
        return graphToken;
      }
      throw err;
    }
  }

  async function graphGet(path) {
    if (!graphToken) {
      await ensureToken();
    }

    const response = await fetch("https://graph.microsoft.com/v1.0" + path, {
      headers: {
        Authorization: "Bearer " + graphToken
      }
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error("Graph " + response.status + ": " + text);
    }

    return response.json();
  }

  async function getUserById(id) {
    if (personCache.has(id)) {
      return personCache.get(id);
    }

    const user = await graphGet("/users/" + encodeURIComponent(id) + "?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation");
    const normalized = normalizePerson(user);
    personCache.set(normalized.id, normalized);
    return normalized;
  }

  async function getUserByUpn(upn) {
    const user = await graphGet("/users/" + encodeURIComponent(upn) + "?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation");
    const normalized = normalizePerson(user);
    personCache.set(normalized.id, normalized);
    return normalized;
  }

  async function getDirectReports(userId) {
    const path =
      "/users/" +
      encodeURIComponent(userId) +
      "/directReports/microsoft.graph.user?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation&$top=999";

    const result = await graphGet(path);
    const users = (result.value || []).map(normalizePerson);

    users.forEach(function (u) {
      personCache.set(u.id, u);
    });

    users.sort(function (a, b) {
      return (a.displayName || "").localeCompare(b.displayName || "");
    });

    return users;
  }

  async function fetchManagerChain(userId) {
    const chain = [];
    let currentId = userId;

    while (true) {
      let manager;
      try {
        manager = await graphGet(
          "/users/" +
            encodeURIComponent(currentId) +
            "/manager/microsoft.graph.user?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation"
        );
      } catch (err) {
        const msg = safeMessage(err).toLowerCase();
        if (msg.includes("404") || msg.includes("resource not found")) {
          break;
        }
        throw err;
      }

      if (!manager || !manager.id) {
        break;
      }

      const normalized = normalizePerson(manager);
      chain.push(normalized);
      personCache.set(normalized.id, normalized);
      currentId = normalized.id;
    }

    return chain.reverse();
  }

  async function buildSubtree(userId, depth, maxDepth) {
    const person = await getUserById(userId);
    const node = cloneNode(person);

    if (depth >= maxDepth) {
      return node;
    }

    const reports = await getDirectReports(userId);
    const children = [];
    for (let i = 0; i < reports.length; i += 1) {
      const child = reports[i];
      const childNode = await buildSubtree(child.id, depth + 1, maxDepth);
      childNode.parentId = node.id;
      children.push(childNode);
    }

    node.children = children;
    return node;
  }

  function attachManagerChain(chain, root) {
    if (!chain || chain.length === 0) {
      return root;
    }

    let current = root;
    for (let i = chain.length - 1; i >= 0; i -= 1) {
      const manager = cloneNode(chain[i]);
      manager.children = [current];
      current.parentId = manager.id;
      current = manager;
    }

    return current;
  }

  function renderTree() {
    if (!orgTree) {
      clearChart("Org chart will appear here after you sign in and load data.");
      return;
    }

    chartContainer.classList.remove("empty-state");
    chartContainer.innerHTML = "";

    const rootWrap = document.createElement("div");
    rootWrap.className = "tree-root";
    rootWrap.appendChild(renderNode(orgTree));
    chartContainer.appendChild(rootWrap);
  }

  function renderNode(node) {
    const wrapper = document.createElement("div");
    wrapper.className = "node";

    const card = document.createElement("div");
    card.className = "card";

    if (matchesSearch(node)) {
      card.classList.add("highlight");
    }

    const header = document.createElement("div");
    header.className = "card-header";

    const avatar = document.createElement("span");
    avatar.className = "avatar";
    avatar.textContent = initials(node.displayName);

    const text = document.createElement("div");

    const name = document.createElement("div");
    name.className = "name";
    name.textContent = node.displayName || "Unknown";

    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = [node.jobTitle, node.department].filter(Boolean).join(" | ") || "No title";

    text.appendChild(name);
    text.appendChild(meta);

    header.appendChild(avatar);
    header.appendChild(text);

    card.appendChild(header);

    if (node.children.length > 0) {
      const controls = document.createElement("div");
      controls.className = "card-controls";

      const count = document.createElement("span");
      count.className = "meta";
      count.textContent = node.children.length + " reports";

      const toggle = document.createElement("button");
      toggle.className = "toggle";
      const isExpanded = expanded.has(node.id);
      toggle.textContent = isExpanded ? "Collapse" : "Expand";
      toggle.addEventListener("click", function () {
        if (expanded.has(node.id)) {
          expanded.delete(node.id);
        } else {
          expanded.add(node.id);
        }
        renderTree();
      });

      controls.appendChild(count);
      controls.appendChild(toggle);
      card.appendChild(controls);
    }

    wrapper.appendChild(card);

    if (node.children.length > 0 && expanded.has(node.id)) {
      const children = document.createElement("div");
      children.className = "children";

      node.children.forEach(function (child) {
        const childItem = document.createElement("div");
        childItem.className = "child-item";
        childItem.appendChild(renderNode(child));
        children.appendChild(childItem);
      });

      wrapper.appendChild(children);
    }

    return wrapper;
  }

  function expandCollapseAll(shouldExpand) {
    if (!orgTree) {
      return;
    }

    if (!shouldExpand) {
      expanded.clear();
      expanded.add(orgTree.id);
      return;
    }

    walkTree(orgTree, function (node) {
      if (node.children.length > 0) {
        expanded.add(node.id);
      }
    });
  }

  function expandDefaults(node, depth) {
    if (!node) {
      return;
    }

    if (depth < 2 && node.children.length > 0) {
      expanded.add(node.id);
    }

    node.children.forEach(function (child) {
      expandDefaults(child, depth + 1);
    });
  }

  function autoExpandForSearch() {
    if (!orgTree) {
      return;
    }

    if (!searchTerm) {
      return;
    }

    const parentMap = new Map();
    walkTree(orgTree, function (node) {
      node.children.forEach(function (child) {
        parentMap.set(child.id, node.id);
      });
    });

    walkTree(orgTree, function (node) {
      if (matchesSearch(node)) {
        let current = node.id;
        while (parentMap.has(current)) {
          const parentId = parentMap.get(current);
          expanded.add(parentId);
          current = parentId;
        }
      }
    });
  }

  function walkTree(node, callback) {
    callback(node);
    node.children.forEach(function (child) {
      walkTree(child, callback);
    });
  }

  function matchesSearch(node) {
    if (!searchTerm) {
      return false;
    }

    return [node.displayName, node.jobTitle, node.department]
      .filter(Boolean)
      .join(" ")
      .toLowerCase()
      .includes(searchTerm);
  }

  function countNodes(root) {
    let total = 0;
    walkTree(root, function () {
      total += 1;
    });
    return total;
  }

  function clearChart(message) {
    chartContainer.innerHTML = "<p>" + escapeHtml(message) + "</p>";
    chartContainer.classList.add("empty-state");
  }

  function normalizePerson(raw) {
    return {
      id: raw.id,
      displayName: raw.displayName || "Unknown",
      jobTitle: raw.jobTitle || "",
      department: raw.department || "",
      mail: raw.mail || "",
      userPrincipalName: raw.userPrincipalName || "",
      officeLocation: raw.officeLocation || "",
      parentId: null,
      children: []
    };
  }

  function cloneNode(person) {
    return {
      id: person.id,
      displayName: person.displayName,
      jobTitle: person.jobTitle,
      department: person.department,
      mail: person.mail,
      userPrincipalName: person.userPrincipalName,
      officeLocation: person.officeLocation,
      parentId: null,
      children: []
    };
  }

  function initials(name) {
    return (name || "?")
      .split(" ")
      .filter(Boolean)
      .slice(0, 2)
      .map(function (part) {
        return part.charAt(0).toUpperCase();
      })
      .join("");
  }

  function clampDepth(value) {
    const parsed = Number(value);
    if (!Number.isFinite(parsed)) {
      return 4;
    }
    return Math.max(1, Math.min(8, Math.floor(parsed)));
  }

  function setStatus(text, isError) {
    statusText.textContent = text;
    statusText.style.color = isError ? "var(--danger)" : "var(--ink-soft)";
  }

  function safeMessage(err) {
    if (!err) {
      return "Unknown error";
    }
    if (typeof err === "string") {
      return err;
    }
    return err.message || String(err);
  }

  function escapeHtml(str) {
    return String(str)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#39;");
  }

  function isConfigured() {
    return cfg && cfg.auth && cfg.auth.clientId && cfg.auth.clientId !== "PUT_YOUR_CLIENT_ID_HERE";
  }
})();
