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

  scrubLegacyMsalCache();

  initialize().catch(function (err) {
    console.error(err);
    setStatus("Initialization failed: " + safeMessage(err), true);
  });

  function scrubLegacyMsalCache() {
    // Remove stale request cache entries from older builds that requested invalid scopes.
    try {
      const toRemove = [];
      for (let i = 0; i < localStorage.length; i += 1) {
        const key = localStorage.key(i);
        if (!key || !key.toLowerCase().includes("msal")) {
          continue;
        }

        const value = localStorage.getItem(key) || "";
        if (value.includes("Org.Read.All")) {
          toRemove.push(key);
        }
      }

      toRemove.forEach(function (key) {
        localStorage.removeItem(key);
      });
    } catch (err) {
      console.warn("Could not scrub stale MSAL cache entries.", err);
    }
  }

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

      if (upn) {
        let rootNode = await getUserByUpn(upn);

        if (!rootNode || !rootNode.id) {
          throw new Error("Could not resolve root user. Check the UPN or permissions.");
        }

        rootNode = normalizePerson(rootNode);

        if (!rootNode.accountEnabled) {
          throw new Error("The selected root user account is disabled.");
        }

        const includeManagerChain = includeManagerCheckbox.checked;
        let managerChain = [];
        if (includeManagerChain) {
          managerChain = await fetchManagerChain(rootNode.id);
        }

        const subtree = await buildSubtree(rootNode.id, 0, depth);
        orgTree = attachManagerChain(managerChain, subtree);
      } else {
        orgTree = await buildOrgForest(depth);
      }

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

    const user = await graphGet("/users/" + encodeURIComponent(id) + "?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation,accountEnabled,userType");
    const normalized = normalizePerson(user);
    personCache.set(normalized.id, normalized);
    return normalized;
  }

  async function getUserByUpn(upn) {
    const user = await graphGet("/users/" + encodeURIComponent(upn) + "?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation,accountEnabled,userType");
    const normalized = normalizePerson(user);
    personCache.set(normalized.id, normalized);
    return normalized;
  }

  async function getDirectReports(userId) {
    const path =
      "/users/" +
      encodeURIComponent(userId) +
      "/directReports/microsoft.graph.user?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation,accountEnabled,userType&$top=999";

    const result = await graphGet(path);
    const rawUsers = (result.value || []).map(normalizePerson);
    const users = [];

    for (let i = 0; i < rawUsers.length; i += 1) {
      const validatedUser = await resolveEnabledPerson(rawUsers[i]);
      if (validatedUser) {
        users.push(validatedUser);
        personCache.set(validatedUser.id, validatedUser);
      }
    }

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
          "/users/" + encodeURIComponent(currentId) + "/manager/microsoft.graph.user?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation,accountEnabled,userType"
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
      const validatedManager = await resolveEnabledPerson(normalized);
      if (validatedManager) {
        personCache.set(validatedManager.id, validatedManager);
        chain.push(validatedManager);
      }
      currentId = normalized.id;
    }

    return chain.reverse();
  }

  async function buildSubtree(userId, depth, maxDepth) {
    const person = await getUserById(userId);
    if (!person || !person.accountEnabled) {
      return null;
    }

    const node = cloneNode(person);

    if (depth >= maxDepth) {
      return node;
    }

    const reports = await getDirectReports(userId);
    const children = [];
    for (let i = 0; i < reports.length; i += 1) {
      const child = reports[i];
      const childNode = await buildSubtree(child.id, depth + 1, maxDepth);
      if (!childNode) {
        continue;
      }
      childNode.parentId = node.id;
      children.push(childNode);
    }

    node.children = children;
    return node;
  }

  async function buildOrgForest(maxDepth) {
    const users = await getAllActiveUsers();

    if (users.length === 0) {
      throw new Error("No top-level users were found in Microsoft 365.");
    }

    const nodeMap = new Map();
    users.forEach(function (user) {
      nodeMap.set(user.id, cloneNode(user));
    });

    const roots = [];
    users.forEach(function (user) {
      const node = nodeMap.get(user.id);
      const managerId = user.managerId || null;
      if (managerId && nodeMap.has(managerId)) {
        nodeMap.get(managerId).children.push(node);
      } else {
        roots.push(node);
      }
    });

    nodeMap.forEach(function (node) {
      node.children.sort(function (a, b) {
        return (a.displayName || "").localeCompare(b.displayName || "");
      });
    });

    roots.sort(function (a, b) {
      return (a.displayName || "").localeCompare(b.displayName || "");
    });

    if (roots.length === 1) {
      return roots[0];
    }

    return {
      id: "org-root",
      displayName: "Executive Leadership",
      accountEnabled: true,
      jobTitle: "Multiple top-level leaders",
      department: "",
      mail: "",
      userPrincipalName: "",
      officeLocation: "",
      parentId: null,
      children: roots
    };
  }

  async function getAllActiveUsers() {
    let path = "/users?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation,accountEnabled,userType&$filter=accountEnabled eq true and userType eq 'Member'&$top=999";
    const users = [];

    while (path) {
      const result = await graphGet(path);
      const batch = (result.value || []).map(normalizePerson).filter(function (person) {
        return isActiveMemberPerson(person);
      });

      for (let i = 0; i < batch.length; i += 1) {
        const validatedUser = await resolveActiveUserWithManager(batch[i]);
        if (validatedUser) {
          users.push(validatedUser);
          personCache.set(validatedUser.id, validatedUser);
        }
      }

      path = result["@odata.nextLink"] ? result["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "") : null;
    }

    users.sort(function (a, b) {
      return (a.displayName || "").localeCompare(b.displayName || "");
    });

    return users;
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
      accountEnabled: typeof raw.accountEnabled === "boolean" ? raw.accountEnabled : null,
      userType: raw.userType || null,
      managerId: raw.managerId || raw.manager?.id || null,
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
      accountEnabled: person.accountEnabled,
      userType: person.userType || null,
      managerId: person.managerId || null,
      jobTitle: person.jobTitle,
      department: person.department,
      mail: person.mail,
      userPrincipalName: person.userPrincipalName,
      officeLocation: person.officeLocation,
      parentId: null,
      children: []
    };
  }

  function isEnabledPerson(person) {
    return !!person && person.accountEnabled === true;
  }

  function isGuestPerson(person) {
    if (!person) {
      return false;
    }

    const identity = (person.userPrincipalName || person.mail || "").toLowerCase();
    return person.userType === "Guest" || identity.includes("#ext#");
  }

  function isActiveMemberPerson(person) {
    return isEnabledPerson(person) && !isGuestPerson(person);
  }

  async function resolveActiveUserWithManager(person) {
    if (!person || !person.id) {
      return null;
    }

    if (typeof person.accountEnabled === "boolean") {
      if (!person.accountEnabled) {
        return null;
      }
      if (person.managerId) {
        return person;
      }
    }

    const fullPerson = await getUserById(person.id);
    if (!isActiveMemberPerson(fullPerson)) {
      return null;
    }

    if (!fullPerson.managerId) {
      try {
        const manager = await graphGet(
          "/users/" +
            encodeURIComponent(fullPerson.id) +
            "/manager/microsoft.graph.user?$select=id,displayName,jobTitle,department,mail,userPrincipalName,officeLocation,accountEnabled,userType"
        );
        fullPerson.managerId = manager?.id || null;
      } catch (err) {
        const msg = safeMessage(err).toLowerCase();
        if (!msg.includes("404") && !msg.includes("resource not found")) {
          throw err;
        }
        fullPerson.managerId = null;
      }
    }

    return fullPerson;
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
