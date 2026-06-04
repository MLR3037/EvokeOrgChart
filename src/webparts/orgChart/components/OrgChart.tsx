import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import styles from './OrgChart.module.scss';

// ─── Types ────────────────────────────────────────────────────────────────────

export interface IOrgChartProps {
  graphClient: MSGraphClientV3;
  rootUserEmail?: string; // Optional: pin to a specific root user; defaults to org-wide
}

interface IUser {
  id: string;
  displayName: string;
  jobTitle?: string;
  department?: string;
  mail?: string;
  userPrincipalName?: string;
  managerId?: string;
  children: IUser[];
  photoUrl?: string;
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Fetch all users in one call with manager expanded, then build tree client-side */
async function fetchOrgTree(
  client: MSGraphClientV3,
  rootEmail?: string
): Promise<IUser | null> {
  // 1. Bulk-fetch all users with manager info
  let allUsers: IUser[] = [];
  let nextLink: string | undefined = `/users?$select=id,displayName,jobTitle,department,mail,userPrincipalName&$expand=manager($select=id)&$top=999&$filter=accountEnabled eq true`;

  while (nextLink) {
    const response = await client.api(nextLink).get();
    const batch: IUser[] = (response.value || []).map((u: any) => ({
      id: u.id,
      displayName: u.displayName || '(No Name)',
      jobTitle: u.jobTitle || '',
      department: u.department || '',
      mail: u.mail || u.userPrincipalName || '',
      userPrincipalName: u.userPrincipalName || '',
      managerId: u.manager?.id || null,
      children: [],
    }));
    allUsers = allUsers.concat(batch);
    nextLink = response['@odata.nextLink'];
  }

  // 2. Build a map and wire up parent → children
  const userMap = new Map<string, IUser>();
  allUsers.forEach(u => userMap.set(u.id, u));

  let roots: IUser[] = [];

  allUsers.forEach(u => {
    if (u.managerId && userMap.has(u.managerId)) {
      userMap.get(u.managerId)!.children.push(u);
    } else {
      roots.push(u);
    }
  });

  // 3. Determine the single root to display
  if (rootEmail) {
    const pinned = allUsers.find(
      u => u.mail?.toLowerCase() === rootEmail.toLowerCase() ||
           u.userPrincipalName?.toLowerCase() === rootEmail.toLowerCase()
    );
    return pinned || roots[0] || null;
  }

  // If multiple roots exist (e.g. shared mailboxes, guests), prefer the one with most reports
  if (roots.length === 0) return null;
  roots.sort((a, b) => countDescendants(b) - countDescendants(a));
  return roots[0];
}

function countDescendants(user: IUser): number {
  return user.children.reduce((sum, c) => sum + 1 + countDescendants(c), 0);
}

/** Fetch a user's photo as a data URL */
async function fetchPhoto(client: MSGraphClientV3, userId: string): Promise<string | null> {
  try {
    const blob: Blob = await client.api(`/users/${userId}/photo/$value`).getStream();
    return URL.createObjectURL(blob);
  } catch {
    return null;
  }
}

// ─── PersonCard ───────────────────────────────────────────────────────────────

interface IPersonCardProps {
  user: IUser;
  isRoot?: boolean;
  graphClient: MSGraphClientV3;
  depth: number;
}

const PersonCard: React.FC<IPersonCardProps> = ({ user, isRoot, graphClient, depth }) => {
  const [photoUrl, setPhotoUrl] = useState<string | null>(null);
  const [expanded, setExpanded] = useState(depth < 2); // auto-expand first 2 levels

  useEffect(() => {
    fetchPhoto(graphClient, user.id).then(url => setPhotoUrl(url));
  }, [user.id]);

  const initials = user.displayName
    .split(' ')
    .map(p => p[0])
    .slice(0, 2)
    .join('')
    .toUpperCase();

  const hasChildren = user.children.length > 0;

  return (
    <div className={`${styles.nodeWrapper} ${isRoot ? styles.rootCard : ''}`}>
      <div className={`${styles.card} ${isRoot ? styles.rootCard : ''}`}>
        <div className={styles.avatar}>
          {photoUrl
            ? <img src={photoUrl} alt={user.displayName} />
            : <span className={styles.initials}>{initials}</span>
          }
        </div>
        <div className={styles.info}>
          <div className={styles.name}>{user.displayName}</div>
          {user.jobTitle && <div className={styles.title}>{user.jobTitle}</div>}
          {user.department && <div className={styles.dept}>{user.department}</div>}
        </div>
        {hasChildren && (
          <button
            className={`${styles.toggle} ${expanded ? styles.expanded : ''}`}
            onClick={() => setExpanded(!expanded)}
            title={expanded ? 'Collapse' : `Expand (${user.children.length})`}
          >
            <span>{expanded ? '▲' : `▼ ${user.children.length}`}</span>
          </button>
        )}
      </div>

      {hasChildren && expanded && (
        <div className={styles.children}>
          <div className={styles.connector} />
          <div className={styles.childRow}>
            {user.children
              .slice()
              .sort((a, b) => a.displayName.localeCompare(b.displayName))
              .map(child => (
                <PersonCard
                  key={child.id}
                  user={child}
                  graphClient={graphClient}
                  depth={depth + 1}
                />
              ))}
          </div>
        </div>
      )}
    </div>
  );
};

// ─── Main Component ───────────────────────────────────────────────────────────

const OrgChart: React.FC<IOrgChartProps> = ({ graphClient, rootUserEmail }) => {
  const [root, setRoot] = useState<IUser | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [search, setSearch] = useState('');
  const [zoom, setZoom] = useState(1);

  useEffect(() => {
    setLoading(true);
    fetchOrgTree(graphClient, rootUserEmail)
      .then(tree => {
        setRoot(tree);
        setLoading(false);
      })
      .catch(err => {
        setError(`Failed to load org chart: ${err.message || err}`);
        setLoading(false);
      });
  }, [graphClient, rootUserEmail]);

  // Simple search: find user in tree and highlight (future: scroll to)
  const handleSearch = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    setSearch(e.target.value);
  }, []);

  if (loading) {
    return (
      <div className={styles.state}>
        <div className={styles.spinner} />
        <p>Loading organizational chart…</p>
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.state}>
        <div className={styles.errorIcon}>⚠️</div>
        <p>{error}</p>
        <p className={styles.hint}>Ensure <strong>User.Read.All</strong> Graph permission is approved in SharePoint Admin → API Access.</p>
      </div>
    );
  }

  if (!root) {
    return <div className={styles.state}><p>No organizational data found.</p></div>;
  }

  return (
    <div className={styles.container}>
      <div className={styles.toolbar}>
        <div className={styles.titleBlock}>
          <h2 className={styles.chartTitle}>Organizational Chart</h2>
        </div>
        <div className={styles.controls}>
          <input
            className={styles.search}
            type="search"
            placeholder="Search by name…"
            value={search}
            onChange={handleSearch}
          />
          <div className={styles.zoomControls}>
            <button onClick={() => setZoom(z => Math.max(0.4, z - 0.1))}>−</button>
            <span>{Math.round(zoom * 100)}%</span>
            <button onClick={() => setZoom(z => Math.min(2, z + 0.1))}>+</button>
            <button onClick={() => setZoom(1)}>Reset</button>
          </div>
        </div>
      </div>

      <div className={styles.scrollArea}>
        <div
          className={styles.tree}
          style={{ transform: `scale(${zoom})`, transformOrigin: 'top center' }}
        >
          <PersonCard
            user={root}
            isRoot={true}
            graphClient={graphClient}
            depth={0}
          />
        </div>
      </div>
    </div>
  );
};

export default OrgChart;
