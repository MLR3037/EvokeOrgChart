window.APP_CONFIG = {
  auth: {
    // Microsoft Entra app registration client ID (GUID)
    clientId: "38155c8a-9831-4ff7-b683-89ea2254e56d",
    // Use your tenant ID for single-tenant or 'common' for multi-tenant
        tenantId: "a4adcc38-7b4e-485c-80f9-7d9ca4e83d64",
    // For GitHub Pages use: https://<org-or-user>.github.io/<repo>/
        redirectUri: "https://mlr3037.github.io/EvokeOrgChart/github-webapp/"
  },
  graph: {
    // These scopes are requested during sign in.
    scopes: ["User.Read", "User.Read.All", "Organization.Read.All"]
  }
};
