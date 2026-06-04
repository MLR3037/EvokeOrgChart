var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import styles from './OrgChart.module.scss';
// ─── Helpers ─────────────────────────────────────────────────────────────────
/** Fetch all users in one call with manager expanded, then build tree client-side */
function fetchOrgTree(client, rootEmail) {
    return __awaiter(this, void 0, void 0, function () {
        var allUsers, nextLink, response, batch, userMap, roots, pinned;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    allUsers = [];
                    nextLink = "/users?$select=id,displayName,jobTitle,department,mail,userPrincipalName&$expand=manager($select=id)&$top=999&$filter=accountEnabled eq true";
                    _a.label = 1;
                case 1:
                    if (!nextLink) return [3 /*break*/, 3];
                    return [4 /*yield*/, client.api(nextLink).get()];
                case 2:
                    response = _a.sent();
                    batch = (response.value || []).map(function (u) {
                        var _a;
                        return ({
                            id: u.id,
                            displayName: u.displayName || '(No Name)',
                            jobTitle: u.jobTitle || '',
                            department: u.department || '',
                            mail: u.mail || u.userPrincipalName || '',
                            userPrincipalName: u.userPrincipalName || '',
                            managerId: ((_a = u.manager) === null || _a === void 0 ? void 0 : _a.id) || null,
                            children: [],
                        });
                    });
                    allUsers = allUsers.concat(batch);
                    nextLink = response['@odata.nextLink'];
                    return [3 /*break*/, 1];
                case 3:
                    userMap = new Map();
                    allUsers.forEach(function (u) { return userMap.set(u.id, u); });
                    roots = [];
                    allUsers.forEach(function (u) {
                        if (u.managerId && userMap.has(u.managerId)) {
                            userMap.get(u.managerId).children.push(u);
                        }
                        else {
                            roots.push(u);
                        }
                    });
                    // 3. Determine the single root to display
                    if (rootEmail) {
                        pinned = allUsers.find(function (u) {
                            var _a, _b;
                            return ((_a = u.mail) === null || _a === void 0 ? void 0 : _a.toLowerCase()) === rootEmail.toLowerCase() ||
                                ((_b = u.userPrincipalName) === null || _b === void 0 ? void 0 : _b.toLowerCase()) === rootEmail.toLowerCase();
                        });
                        return [2 /*return*/, pinned || roots[0] || null];
                    }
                    // If multiple roots exist (e.g. shared mailboxes, guests), prefer the one with most reports
                    if (roots.length === 0)
                        return [2 /*return*/, null];
                    roots.sort(function (a, b) { return countDescendants(b) - countDescendants(a); });
                    return [2 /*return*/, roots[0]];
            }
        });
    });
}
function countDescendants(user) {
    return user.children.reduce(function (sum, c) { return sum + 1 + countDescendants(c); }, 0);
}
/** Fetch a user's photo as a data URL */
function fetchPhoto(client, userId) {
    return __awaiter(this, void 0, void 0, function () {
        var blob, _a;
        return __generator(this, function (_b) {
            switch (_b.label) {
                case 0:
                    _b.trys.push([0, 2, , 3]);
                    return [4 /*yield*/, client.api("/users/".concat(userId, "/photo/$value")).getStream()];
                case 1:
                    blob = _b.sent();
                    return [2 /*return*/, URL.createObjectURL(blob)];
                case 2:
                    _a = _b.sent();
                    return [2 /*return*/, null];
                case 3: return [2 /*return*/];
            }
        });
    });
}
var PersonCard = function (_a) {
    var user = _a.user, isRoot = _a.isRoot, graphClient = _a.graphClient, depth = _a.depth;
    var _b = useState(null), photoUrl = _b[0], setPhotoUrl = _b[1];
    var _c = useState(depth < 2), expanded = _c[0], setExpanded = _c[1]; // auto-expand first 2 levels
    useEffect(function () {
        fetchPhoto(graphClient, user.id).then(function (url) { return setPhotoUrl(url); });
    }, [user.id]);
    var initials = user.displayName
        .split(' ')
        .map(function (p) { return p[0]; })
        .slice(0, 2)
        .join('')
        .toUpperCase();
    var hasChildren = user.children.length > 0;
    return (React.createElement("div", { className: "".concat(styles.nodeWrapper, " ").concat(isRoot ? styles.root : '') },
        React.createElement("div", { className: "".concat(styles.card, " ").concat(isRoot ? styles.rootCard : '') },
            React.createElement("div", { className: styles.avatar }, photoUrl
                ? React.createElement("img", { src: photoUrl, alt: user.displayName })
                : React.createElement("span", { className: styles.initials }, initials)),
            React.createElement("div", { className: styles.info },
                React.createElement("div", { className: styles.name }, user.displayName),
                user.jobTitle && React.createElement("div", { className: styles.title }, user.jobTitle),
                user.department && React.createElement("div", { className: styles.dept }, user.department)),
            hasChildren && (React.createElement("button", { className: "".concat(styles.toggle, " ").concat(expanded ? styles.expanded : ''), onClick: function () { return setExpanded(!expanded); }, title: expanded ? 'Collapse' : "Expand (".concat(user.children.length, ")") },
                React.createElement("span", null, expanded ? '▲' : "\u25BC ".concat(user.children.length))))),
        hasChildren && expanded && (React.createElement("div", { className: styles.children },
            React.createElement("div", { className: styles.connector }),
            React.createElement("div", { className: styles.childRow }, user.children
                .slice()
                .sort(function (a, b) { return a.displayName.localeCompare(b.displayName); })
                .map(function (child) { return (React.createElement(PersonCard, { key: child.id, user: child, graphClient: graphClient, depth: depth + 1 })); }))))));
};
// ─── Main Component ───────────────────────────────────────────────────────────
var OrgChart = function (_a) {
    var graphClient = _a.graphClient, rootUserEmail = _a.rootUserEmail;
    var _b = useState(null), root = _b[0], setRoot = _b[1];
    var _c = useState(true), loading = _c[0], setLoading = _c[1];
    var _d = useState(null), error = _d[0], setError = _d[1];
    var _e = useState(''), search = _e[0], setSearch = _e[1];
    var _f = useState(1), zoom = _f[0], setZoom = _f[1];
    useEffect(function () {
        setLoading(true);
        fetchOrgTree(graphClient, rootUserEmail)
            .then(function (tree) {
            setRoot(tree);
            setLoading(false);
        })
            .catch(function (err) {
            setError("Failed to load org chart: ".concat(err.message || err));
            setLoading(false);
        });
    }, [graphClient, rootUserEmail]);
    // Simple search: find user in tree and highlight (future: scroll to)
    var handleSearch = useCallback(function (e) {
        setSearch(e.target.value);
    }, []);
    if (loading) {
        return (React.createElement("div", { className: styles.state },
            React.createElement("div", { className: styles.spinner }),
            React.createElement("p", null, "Loading organizational chart\u2026")));
    }
    if (error) {
        return (React.createElement("div", { className: styles.state },
            React.createElement("div", { className: styles.errorIcon }, "\u26A0\uFE0F"),
            React.createElement("p", null, error),
            React.createElement("p", { className: styles.hint },
                "Ensure ",
                React.createElement("strong", null, "User.Read.All"),
                " Graph permission is approved in SharePoint Admin \u2192 API Access.")));
    }
    if (!root) {
        return React.createElement("div", { className: styles.state },
            React.createElement("p", null, "No organizational data found."));
    }
    return (React.createElement("div", { className: styles.container },
        React.createElement("div", { className: styles.toolbar },
            React.createElement("div", { className: styles.titleBlock },
                React.createElement("h2", { className: styles.chartTitle }, "Organizational Chart")),
            React.createElement("div", { className: styles.controls },
                React.createElement("input", { className: styles.search, type: "search", placeholder: "Search by name\u2026", value: search, onChange: handleSearch }),
                React.createElement("div", { className: styles.zoomControls },
                    React.createElement("button", { onClick: function () { return setZoom(function (z) { return Math.max(0.4, z - 0.1); }); } }, "\u2212"),
                    React.createElement("span", null,
                        Math.round(zoom * 100),
                        "%"),
                    React.createElement("button", { onClick: function () { return setZoom(function (z) { return Math.min(2, z + 0.1); }); } }, "+"),
                    React.createElement("button", { onClick: function () { return setZoom(1); } }, "Reset")))),
        React.createElement("div", { className: styles.scrollArea },
            React.createElement("div", { className: styles.tree, style: { transform: "scale(".concat(zoom, ")"), transformOrigin: 'top center' } },
                React.createElement(PersonCard, { user: root, isRoot: true, graphClient: graphClient, depth: 0 })))));
};
export default OrgChart;
//# sourceMappingURL=OrgChart.js.map