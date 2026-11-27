---
name: devloop-setup
description: Set up or repair this project's devloop layout (AGENTS.md managed blocks, skills and per-tool adapters) with devloop init, sync and doctor. Use when doctor reports drift or after editing devloop.toml or a skill; not for writing project code.
---

# devloop setup

The rules and skills every coding agent reads come from one source. Never edit a rendered
copy: change the source and re-render.

## Sources

- `devloop.toml`: the only project-specific configuration.
- `AGENTS.md`: text between `devloop:begin` and `devloop:end` markers is rendered from
  `devloop.toml`; everything else is the project's own and is edited by hand.
- Project skills: the folder named by `skills.project` (default `devloop/skills/`).

## Steps

0. Install the devloop release whose minor version `framework.version` names, for
   example `uv tool install git+https://github.com/Jon-275/dev-loop@v0.11.0`, or pin it in
   the project's dev dependencies.
1. Run `devloop doctor`. It lists every problem and the file it concerns.
2. If the configuration is wrong, fix `devloop.toml` and run `devloop doctor` again.
3. Run `devloop sync` to rewrite the managed blocks, rendered skills and adapters.
4. Run `devloop doctor --check`; it must exit 0. CI runs the same check.
5. Commit the source change and the rendered files together.

## When not to use

- To change what a skill says: edit its source folder, then run `devloop sync`.
- To silence a doctor finding: fix the cause; CI runs the same check.
