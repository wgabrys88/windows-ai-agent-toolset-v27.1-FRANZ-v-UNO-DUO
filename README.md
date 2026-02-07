# FRANZ (single-file) — Changes vs original `main_updated_crlf.py`

This README lists only the changes introduced compared to the original file you shared at the start.

## 1) Per-turn execution log (new file)

- Added `dump/<run>/execution-log.txt`.
- Appends one block per turn immediately after saving the screenshot.
- Each block includes:
  - Timestamp (ms precision)
  - Current process PID
  - Screenshot filename (`stepNNN.png`) so logs can be aligned with images
- Logs only windows owned by the current Python process (PID-based filtering) and only those intersecting the primary screen.

## 2) Programmatic window text capture (non-visual)

- Uses Win32 enumeration and messaging (no OCR, no image parsing):
  - `EnumWindows` to enumerate top-level windows
  - `GetWindowThreadProcessId` to filter to this process
  - `WM_GETTEXT` via `SendMessageTimeoutW` to safely read window text across threads without hanging
  - `EnumChildWindows` to also capture visible child controls’ text
- Captures for each window:
  - hwnd, class name, rect, title
  - top-level `WM_GETTEXT` content (when present)
  - child control `WM_GETTEXT` content (when present)

## 3) Added Win32 API bindings

- Added `kernel32.GetCurrentProcessId`.
- Added `user32.GetWindowThreadProcessId`.
- Added helpers used by the execution log:
  - `_get_class_name`
  - `_safe_sendmessage_wm_gettext`
  - `_format_multiline`
  - `append_execution_log`

## 4) LM Studio + Qwen3-VL request alignment

- Kept `/v1/chat/completions` OpenAI-compatible request shape and explicit `tool_choice: "required"` with `tools` enabled.
- Added an explicit user text instruction alongside the screenshot to improve reliable single-tool calling and enforce inclusion of `story`.

## 5) Generation parameter changes (expanded + retuned)

- Retuned defaults to Qwen3-VL recommended VL settings:
  - `temperature=0.7`, `top_p=0.8`, `top_k=20`, `presence_penalty=1.5`, `repeat_penalty=1.0`
- Expanded payload to always include OpenAI/LM Studio compatible fields:
  - `stream`, `stop`, `frequency_penalty`, `logit_bias`, `seed`, `max_tokens` (and the above parameters)

## 6) Tool-call handling simplification (removed legacy salvage)

Removed the legacy “tool-call repair / coercion” layer. The system now assumes:
- LM Studio returns `message.tool_calls` as per the OpenAI-compatible schema.
- Tool arguments are valid JSON (string or object).
- The tool name must be one of the declared tools (strict validation).

Removed from the original:
- `_strip_code_fences`
- `_coerce_int`
- `_extract_from_values`
- `normalize_tool_call`
- Regex-based coordinate recovery / fallback tool substitution logic

## 7) Main loop change around screenshots

- Screenshot name is stored as `img_name` and reused consistently for:
  - image write
  - execution-log block header

## 8) Cleanup

- Removed unused `re` import after removing the normalization/salvage layer.
- Reduced branching in the tool-call path (strict behavior, fewer heuristics).
