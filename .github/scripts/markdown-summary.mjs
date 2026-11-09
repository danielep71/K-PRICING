// Markdown helpers shared by the label scripts' job summaries.
//
// Every value that comes from the manifest, the policy or the GitHub API is
// data: it must not open a code span, end a table cell or start a new line.
// A code span gets a backtick fence longer than any backtick run inside the
// value; carriage returns and line feeds become spaces; a table cell escapes
// its pipes.

function singleLine(value) {
  return String(value).replace(/\r\n|\r|\n/g, " ");
}

function fenced(text) {
  const longest = Math.max(0, ...(text.match(/`+/g) ?? []).map(run => run.length));
  const fence = "`".repeat(longest + 1);
  if (text === "") return "` `";
  const pad = text.startsWith("`") || text.endsWith("`")
    || (text.startsWith(" ") && text.endsWith(" ")) ? " " : "";
  return `${fence}${pad}${text}${pad}${fence}`;
}

// A code span for a list line or a paragraph.
export function codeSpan(value) {
  return fenced(singleLine(value));
}

// A code span inside a table cell: its pipes are escaped for the table.
export function tableCode(value) {
  return fenced(singleLine(value).replaceAll("|", "\\|"));
}

// Plain text inside a table cell: backslashes, pipes and backticks escaped.
export function tableText(value) {
  return singleLine(value)
    .replaceAll("\\", "\\\\")
    .replaceAll("|", "\\|")
    .replaceAll("`", "\\`");
}

export function selfTestMarkdown(assert) {
  assert.equal(codeSpan("plain"), "`plain`");
  assert.equal(codeSpan("a`b"), "``a`b``");
  assert.equal(codeSpan("``x"), "``` ``x ```");
  assert.equal(codeSpan("line\r\nnext\rlast\nend"), "`line next last end`");
  assert.equal(codeSpan(""), "` `");
  assert.equal(tableCode("a|b"), "`a\\|b`");
  assert.equal(tableCode("x`|`y"), "``x`\\|`y``");
  assert.equal(tableText("a\\b|c`d\ne"), "a\\\\b\\|c\\`d e");
}
