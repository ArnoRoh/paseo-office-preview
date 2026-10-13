export function isOfficePath(value: string): boolean {
  return /\.(?:docx|xlsx|xls)$/i.test(value);
}

// Only local file references qualify. HTTP links stay with Paseo's normal renderer.
export function documentLinks(text: string): string[] {
  const paths: string[] = [];
  const candidates = [
    ...text.matchAll(/\[[^\]\n]*\]\(\s*(<[^>\n]+>|[^\n]*?)\s*\)/g),
    ...text.matchAll(/`([^`\n]+)`/g),
  ];
  for (const match of candidates) {
    let candidate = match[1]!.replace(/^<|>$/g, '').trim();
    if (/^file:\/\//i.test(candidate)) {
      if (!/^file:\/\/\//i.test(candidate)) continue;
      candidate = candidate.slice('file://'.length);
    } else if (/^[a-z][a-z\d+.-]*:/i.test(candidate) && !/^[a-z]:[\\/]/i.test(candidate)) continue;
    try { candidate = decodeURIComponent(candidate); } catch { continue; }
    candidate = candidate.replace(/:\d+(?::\d+)?$/, '');
    if (isOfficePath(candidate) && candidate.length <= 4096 && !candidate.includes('\0') && !paths.includes(candidate)) paths.push(candidate);
    if (paths.length === 20) break;
  }
  return paths;
}
