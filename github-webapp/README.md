# GitHub-Hosted Microsoft 365 Org Chart Web App

This folder contains a static web app you can host on GitHub Pages.

It uses:
- Microsoft Entra ID (MSAL) for browser sign-in
- Microsoft Graph for org data
- Plain HTML/CSS/JavaScript (no build step required)

## 1) Register an app in Microsoft Entra

1. Go to Entra admin center -> App registrations -> New registration.
2. Name it: GitHub Org Chart SPA.
3. Supported account types:
   - Single tenant if internal only, or
   - Multitenant if needed.
4. Redirect URI:
   - Platform: Single-page application
   - URI: your future GitHub Pages URL, for example:
     - https://your-user.github.io/Evoke-org-chart/github-webapp/
5. After creation, copy the Application (client) ID.

## 2) Configure API permissions

In the app registration:
1. API permissions -> Add a permission -> Microsoft Graph -> Delegated permissions.
2. Add these:
   - User.Read
   - User.Read.All
   - Organization.Read.All
3. Click Grant admin consent for your tenant.

If your tenant policy blocks broad delegated permissions, ask your Entra admin for approved alternatives.

## 3) Update app config

Edit config.js:
1. Set auth.clientId to your app client ID.
2. Set auth.tenantId:
   - Your tenant GUID for single-tenant, or
   - common for multi-tenant.
3. Set auth.redirectUri to your published GitHub Pages URL path.

Important: redirectUri must exactly match the URI in the app registration.

## 4) Publish on GitHub Pages

Recommended approach:
1. Commit this folder to your repository.
2. Push to GitHub.
3. In repository Settings -> Pages:
   - Build and deployment -> Source: Deploy from a branch
   - Branch: main
   - Folder: / (root)
4. Access your app at:
   - https://your-user.github.io/Evoke-org-chart/github-webapp/

Alternative:
- If your Pages setup only publishes /docs, copy github-webapp contents into a docs folder.

## 5) Use the app

1. Open the published URL.
2. Click Sign in.
3. Optional: enter a root UPN/email (example: ceo@contoso.com).
4. Set max depth.
5. Click Load org chart.
6. Expand/collapse nodes and use search to highlight people.

## Troubleshooting

- AADSTS50011 (redirect mismatch):
  - Fix redirect URI in both Entra app and config.js.
- Insufficient privileges:
  - Confirm delegated permissions and admin consent.
- Missing manager chain:
  - Some users may not have a manager value in Entra.
- Large tenant performance:
  - Lower max depth to reduce Graph calls.

## Security notes

- This is a client-side SPA using delegated user permissions.
- Do not put secrets in this app.
- clientId is public and safe to expose.
