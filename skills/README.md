# AI Skill Packages

This directory contains installable `SKILL.md` packages for AI coding agents.

## Available Skills

- `webpro-ppt-film-maker`: Instructions for working on this PowerPoint-to-video project, including CLI/GUI workflows, MiniMax TTS/Hailuo video generation, OpenAI image generation, ffmpeg composition, and safe verification commands.

## Install

Copy the entire skill folder, not only `SKILL.md`.

Claude Code:

```sh
mkdir -p ~/.claude/skills
cp -R skills/webpro-ppt-film-maker ~/.claude/skills/
```

opencode project skill:

```sh
mkdir -p .opencode/skills
cp -R skills/webpro-ppt-film-maker .opencode/skills/
```

opencode global skill:

```sh
mkdir -p ~/.config/opencode/skills
cp -R skills/webpro-ppt-film-maker ~/.config/opencode/skills/
```

Generic `SKILL.md` agents, including Codex-like or OpenClaw-style clients:

```sh
mkdir -p ~/.agents/skills
cp -R skills/webpro-ppt-film-maker ~/.agents/skills/
```

If a client does not auto-load skills, paste or reference `skills/webpro-ppt-film-maker/SKILL.md` from its project instructions file.

After installing into a running agent client, restart the client so it reloads skills.
