"""Automated checks for documentation links, glossary acronym coverage and prose hygiene."""

from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]
USER_DOCS = (
    'README.md',
    'docs/README.md',
    'docs/glossary.md',
    'docs/getting-started.md',
    'docs/model-guide.md',
    'docs/model-methods.md',
    'docs/risk-analysis.md',
    'docs/input-reference.md',
    'docs/portfolio-management.md',
    'docs/workbook-schedules.md',
)
ALL_DOCS = USER_DOCS + (
    'tests/README.md',
    'docs/annual-tax-review.md',
    'docs/macro-runner.md',
    'docs/publishing.md',
)
# Spreadsheet column letters and hardware size units that are not financial acronyms.
NON_ACRONYM_TOKENS = {
    'AE', 'AH', 'AI', 'AM', 'AR', 'AW', 'BD', 'CE', 'CF',
    'GB', 'RAM', 'README', 'LICENSE',
}
RETIRED_PHRASES = (
    'choose your path',
    'guided safari',
    'executive verdict',
    'decision takeaway',
    '[top 10]',
    'zero macros',
    'benchmark website',
)


def _slugify_heading(text: str) -> str:
    clean = re.sub(r'`([^`]*)`', r'\1', text.strip().lower())
    clean = re.sub(r'[^\w\s-]', '', clean)
    return re.sub(r'\s+', '-', clean).strip('-')


def _anchors_in_markdown(path: Path) -> set[str]:
    text = path.read_text(encoding='utf-8')
    anchors = set(re.findall(r'<a\s+id=["\']([^"\']+)["\']', text))
    seen: dict[str, int] = {}
    for line in text.splitlines():
        match = re.match(r'^#{1,6}\s+(.+?)\s*$', line)
        if not match:
            continue
        base = _slugify_heading(match.group(1))
        count = seen.get(base, 0)
        seen[base] = count + 1
        anchors.add(base if count == 0 else f'{base}-{count}')
    return anchors


class TestDocsAndGlossary(unittest.TestCase):
    def test_relative_links_and_anchors_resolve(self):
        link_re = re.compile(r'\[[^\]]+\]\(([^)]+)\)')
        for rel in ALL_DOCS:
            source = ROOT / rel
            text = re.sub(r'```.*?```', '', source.read_text(encoding='utf-8'), flags=re.S)
            for raw_target in link_re.findall(text):
                target = raw_target.strip()
                if target.startswith(('http://', 'https://', 'mailto:')):
                    continue
                file_part, _, fragment = target.partition('#')
                resolved = (source.parent / file_part).resolve() if file_part else source.resolve()
                self.assertTrue(
                    resolved.exists(),
                    f'{rel}: broken relative link {raw_target!r} -> {resolved}',
                )
                if fragment and resolved.suffix.lower() == '.md':
                    anchors = _anchors_in_markdown(resolved)
                    self.assertIn(
                        fragment,
                        anchors,
                        f'{rel}: missing #{fragment} in {resolved.relative_to(ROOT)}',
                    )

    def test_user_docs_acronyms_are_defined_in_glossary(self):
        glossary_text = (ROOT / 'docs/glossary.md').read_text(encoding='utf-8')
        defined = set(re.findall(r'\*\*([A-Z]{2,6})\*\*', glossary_text))
        missing: dict[str, list[str]] = {}
        for rel in USER_DOCS:
            text = (ROOT / rel).read_text(encoding='utf-8')
            text = re.sub(r'```.*?```', '', text, flags=re.S)
            text = re.sub(r'`[^`]+`', '', text)
            text = re.sub(r'\]\([^)]+\)', ']()', text)
            for token in sorted(set(re.findall(r'\b[A-Z]{2,6}\b', text))):
                if token in NON_ACRONYM_TOKENS or token in defined:
                    continue
                missing.setdefault(token, []).append(rel)
        self.assertEqual(
            missing,
            {},
            f'Define missing acronyms in docs/glossary.md: {missing}',
        )

    def test_user_hub_keeps_maintainer_roadmaps_in_tests_readme(self):
        hub = (ROOT / 'docs/README.md').read_text(encoding='utf-8')
        self.assertIn('(glossary.md)', hub)
        self.assertNotIn('enhancement-plan.md', hub)
        self.assertNotIn('macro-runner.md', hub)

    def test_no_retired_phrases_or_tool_attribution(self):
        banned_tools = re.compile(
            r'\b(claude|chatgpt|copilot|cursor|hermes|codex|openai|anthropic|llm)\b',
            re.I,
        )
        for rel in USER_DOCS:
            text = (ROOT / rel).read_text(encoding='utf-8')
            lower = text.lower()
            for phrase in RETIRED_PHRASES:
                self.assertNotIn(phrase, lower, f'{rel} contains retired phrase {phrase!r}')
            self.assertIsNone(
                banned_tools.search(text),
                f'{rel} contains coding-tool attribution',
            )


if __name__ == '__main__':
    unittest.main()
