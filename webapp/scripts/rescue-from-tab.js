// SecretaryBench storyline rescue (2026-06-10 incident).
//
// If you have a SecretaryBench tab that was opened BEFORE the storylines were
// deleted and you have NOT reloaded it, that tab still holds every storyline in
// memory. This script pulls them out of the page's React state, downloads a
// backup JSON to your computer, and re-uploads any storyline the server no
// longer has. It never overwrites a storyline that still exists on the server.
//
// How to run:
//   1. DO NOT reload or close the tab.
//   2. Open DevTools (Cmd+Option+J on Mac, Ctrl+Shift+J on Windows).
//   3. Paste this entire file into the Console and press Enter.
//   4. Read the output. It should list the storylines it found, download a
//      backup file, and print "restored" for each missing one.
(async () => {
  const containers = [...document.querySelectorAll('body, body *')].slice(0, 2000);
  let fiber = null;
  for (const el of containers) {
    const key = Object.keys(el).find((k) => k.startsWith('__reactFiber$') || k.startsWith('__reactContainer$'));
    if (key) { fiber = el[key]; break; }
  }
  if (!fiber) { console.error('No React internals found. Is this the SecretaryBench tab?'); return; }

  // Walk the fiber tree looking for a useState hook whose value is the storyline array.
  const looksLikeNodes = (v) => Array.isArray(v) && v.length > 0 && v.every((x) => x && typeof x === 'object' && typeof x.id === 'string' && Array.isArray(x.emails));
  const seen = new Set(); const stack = [fiber]; let nodes = null;
  while (stack.length && !nodes) {
    const f = stack.pop();
    if (!f || seen.has(f)) continue; seen.add(f);
    let h = f.memoizedState, guard = 0;
    while (h && typeof h === 'object' && guard++ < 200) {
      if (looksLikeNodes(h.memoizedState)) { nodes = h.memoizedState; break; }
      h = h.next;
    }
    if (f.child) stack.push(f.child);
    if (f.sibling) stack.push(f.sibling);
    if (f.alternate && !seen.has(f.alternate)) stack.push(f.alternate);
  }
  if (!nodes) { console.error('No storyline state found in this tab. It may have loaded after the deletion.'); return; }
  console.log(`Found ${nodes.length} storylines in this tab:`, nodes.map((n) => n.id).join(', '));

  // Always download a local backup first, in case anything below fails.
  const blob = new Blob([JSON.stringify(nodes, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `secretarybench-rescue-${nodes.length}-storylines.json`;
  a.click();
  console.log('Backup file downloaded. Now re-uploading missing storylines...');

  // Only upload storylines the server no longer has — never clobber live ones.
  const existing = new Set((await (await fetch('/api/nodes', { cache: 'no-store' })).json()).map((n) => n.id));
  for (const n of nodes) {
    if (existing.has(n.id)) { console.log(`  ${n.id}: skipped (still on server)`); continue; }
    const r = await fetch('/api/nodes', { method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify(n) });
    console.log(`  ${n.id}: ${r.ok ? 'RESTORED' : 'FAILED ' + r.status}`);
  }
  console.log('Done. Reload the app to verify, and send Miguel the downloaded backup file.');
})();
