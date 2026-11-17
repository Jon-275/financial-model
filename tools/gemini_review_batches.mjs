/** Prepare complete, bounded review inputs from the GitHub changed-file response. */
export const patchLimit = 120_000;

function isReviewable(path) {
  return /\.(?:pyi?|[cm]?js|sh|bas|cls|frm|fragment)$/i.test(path) ||
    /^\.github\/(workflows\/[^/]+|actions\/.+\/action)\.ya?ml$/i.test(path) ||
    path.startsWith('.githooks/') || path === '.gitignore' ||
    /(?:^|\/)requirements(?:[-.][^/]*)?\.txt$/.test(path);
}

function inspectPatch(file) {
  const patch = file.patch ?? '';
  const anchors = { LEFT: [], RIGHT: [] };
  const positions = [];
  let oldLine = 0, newLine = 0, oldLeft = 0, newLeft = 0, offset = 0;
  let hunk = '', added = 0, removed = 0;
  const invalid = () => { throw new Error('GitHub returned an incomplete patch for ' + file.filename); };
  if (typeof patch !== 'string' || !Number.isSafeInteger(file.additions) ||
      !Number.isSafeInteger(file.deletions) || file.additions < 0 || file.deletions < 0) invalid();
  for (const line of patch.split('\n')) {
    if (offset === patch.length && line === '') break;
    const match = /^@@ -(\d+)(?:,(\d+))? \+(\d+)(?:,(\d+))? @@(?: |$)/.exec(line);
    if (match) {
      if (oldLeft || newLeft) invalid();
      oldLine = Number(match[1]);
      newLine = Number(match[3]);
      oldLeft = Number(match[2] ?? 1);
      newLeft = Number(match[4] ?? 1);
      if (![oldLine, newLine, oldLeft, newLeft].every(Number.isSafeInteger)) invalid();
      hunk = match[0].trimEnd();
    }
    positions.push({ offset, oldLine, newLine, hunk, kind: line[0] });
    if (!match) {
      if (!hunk) invalid();
      switch (line[0]) {
        case '-':
          if (--oldLeft < 0) invalid();
          anchors.LEFT.push(oldLine++);
          removed++;
          break;
        case '+':
          if (--newLeft < 0) invalid();
          anchors.RIGHT.push(newLine++);
          added++;
          break;
        case ' ':
          if (--oldLeft < 0 || --newLeft < 0) invalid();
          oldLine++;
          anchors.RIGHT.push(newLine++);
          break;
        default:
          if (line !== '\\ No newline at end of file') invalid();
      }
    }
    offset += line.length + 1;
  }
  if (oldLeft || newLeft || added !== file.additions || removed !== file.deletions) invalid();
  return { patch, positions, target: { path: file.filename, anchors } };
}

function* splitPatch(file, patch, positions, limit) {
  const previous = file.previous_filename ? ' from ' + JSON.stringify(file.previous_filename) : '';
  const header = 'File: ' + JSON.stringify(file.filename) + ' (' + file.status + previous + ')\n';
  let start = 0, index = 0;
  do {
    // Only copied headers and continuation labels are outside the verbatim patch slices.
    while (positions[index + 1]?.offset <= start) index++;
    const position = positions[index];
    let prefix = header;
    if (start > 0 && (position.offset !== start || position.kind !== '@')) {
      prefix += `Continuing ${position.hunk} at LEFT ${position.oldLine}, RIGHT ${position.newLine}`;
      if (position.offset !== start) {
        prefix += ` inside diff line ${JSON.stringify(position.kind)}, character ${start - position.offset + 1}`;
      }
      prefix += '.\n';
    }
    const available = limit - prefix.length;
    if (available < 2) throw new Error('The file header exceeds the Gemini review input limit.');
    let end = Math.min(start + available, patch.length);
    if (end < patch.length) {
      const window = patch.slice(start, end);
      const hunkBreak = window.lastIndexOf('\n@@ ');
      const newline = hunkBreak >= 0 ? hunkBreak : window.lastIndexOf('\n');
      if (newline >= 0) end = start + newline + 1;
      else if (/[\uD800-\uDBFF]/.test(patch[end - 1]) && /[\uDC00-\uDFFF]/.test(patch[end])) end--;
    }
    const content = patch.slice(start, end);
    let scope;
    if (start > 0 || end < patch.length) {
      const anchors = { LEFT: new Set(), RIGHT: new Set() };
      for (let i = index; i < positions.length && positions[i].offset < end; i++) {
        const { kind, oldLine, newLine } = positions[i];
        if (kind === '-') anchors.LEFT.add(oldLine);
        if (kind === '+' || kind === ' ') anchors.RIGHT.add(newLine);
      }
      scope = { path: file.filename, partial: true, anchors };
    }
    yield { prefix, content, scope };
    start = end;
  } while (start < patch.length);
}

export function collectBatches(files, changedFiles, limit = patchLimit) {
  if (files.length !== changedFiles || new Set(files.map(f => f.filename)).size !== changedFiles) {
    throw new Error('GitHub returned an incomplete changed-file list.');
  }
  const codeFiles = files.filter(f => isReviewable(f.filename) ||
    (f.previous_filename && isReviewable(f.previous_filename)));
  codeFiles.sort((a,b) => a.filename < b.filename ? -1 : a.filename > b.filename ? 1 : 0);
  const batches = [], paths = new Map();
  const canonicalNames = new Set(codeFiles.map(file => file.filename));
  let current;
  for (const file of codeFiles) {
    const { patch, positions, target } = inspectPatch(file);
    const aliases = [file.filename];
    if (file.previous_filename && !canonicalNames.has(file.previous_filename)) aliases.push(file.previous_filename);
    for (const name of aliases) paths.set(name, target);
    for (const { prefix, content, scope } of splitPatch(file, patch, positions, limit)) {
      const text = prefix + content;
      if (!current || current.paths.has(file.filename) || current.text.length + 2 + text.length > limit) {
        current = { text: '', paths: new Map() };
        batches.push(current);
      }
      current.text += (current.text ? '\n\n' : '') + text;
      for (const name of aliases) current.paths.set(name, scope ?? target);
    }
  }
  return { batches, paths };
}
