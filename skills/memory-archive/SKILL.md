---
name: memory-archive
description: "Use when you need facts evicted from memory."
category: productivity
version: 1.0.0
author: Hermes Agent
license: MIT
metadata:
  hermes:
    tags: [memory, archive, eviction, context]
---

# Memory Archive

## When to Use

Use this skill when you need to recall a fact, lesson, or environment
detail that was evicted from Coding Sensei's persistent memory because
of its character limit. It is the long-term repository for everything
that no longer fits in active memory.

## How this differs from active memory

Active memory is injected into every conversation turn and is strictly
size-capped (~2,200 characters). This archive lives in a skill file and
is loaded only on demand — it has no size cap.

Repository for facts, lessons, and context that were evicted from the
Coding Sensei persistent memory store because of its character limit.
Nothing here is deleted — it is simply moved out of the per-turn memory
budget so it can be reloaded when relevant.

## How entries get here

When `memory` is full and a new fact must be saved, I batch-remove stale
entries and add them here instead of discarding them. Each entry keeps
its original wording as closely as possible.

## Entries

### 2026-11-16 — evicted to make room for timesheet automation facts

- **Discord connector slug**: `'discordbot'` (not `'discord'`). Provides only
  CREATE_MESSAGE (send-only), no listening/webhook capability. Telegram and
  WhatsApp connectors available but not yet connected.
- **Docker**: NOT installed on this Windows host; `docker` command not
  recognized. Use bare Python instead.
- **computer_use capture action**: fails with 'Connection closed' on this
  Windows host; PIL ImageGrab (Python) works as a reliable screenshot fallback.
- **MEDIA:/ delivery**: MEDIA:/ lines in chat produce no delivery confirmation
  — cannot verify whether images actually reach the user on Telegram.
- **Desktop screenshot path**: C:/Users/Administrator/Desktop/ is a more
  visible screenshot delivery location than temp scratch on this host.

---

## How to use

Load with `skill_view(name='memory-archive')` to see everything archived so
far. If you need a specific archived fact back in active memory, tell me and
I can swap it in (evicting something else if needed).

---

*Maintained by Coding Sensei (profile `coding-sensei`).*