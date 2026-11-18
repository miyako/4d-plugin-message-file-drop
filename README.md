# 4d-plugin-message-file-drop

Accept drag and drop of messages from Microsoft Outlook (and of files from Explorer) into a 4D window. When an email is dropped directly from Outlook, the plugin drives Outlook to export the selected messages to a temporary folder as `.mht` and `.msg` files, then runs a 4D callback method for each message. Standard Explorer file drops are delivered to the same callback.

![version](https://img.shields.io/badge/version-17%2B-3E8B93)
![platform](https://img.shields.io/static/v1?label=platform&message=win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-message-file-drop)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-message-file-drop/total)

**A similar solution for Mac is available**: [apple-file-promises](https://github.com/miyako/4d-plugin-apple-file-promises)

| Command | Returns | Purpose |
|---|---|---|
| [ACCEPT MESSAGE FILES](#accept-message-files) | — | Install a custom OLE drag-and-drop handler on a window and run a callback when a message or file is dropped |

**Platforms:** Windows 64-bit only. This plugin has no macOS implementation; on macOS the command is a no-op.

---

## Requirements & platform notes

- **Windows 64-bit only.** The drop target, the Outlook COM automation, and the export all live under `#if VERSIONWIN`. On macOS `ACCEPT MESSAGE FILES` reads its parameters and does nothing else.
- **Microsoft Outlook must be installed and running** for Outlook message drops to be exported. If Outlook cannot be reached, the drop is silently ignored — no 4D error is raised.
- **Installing the handler disables the window's standard OLE drop events.** Once you call `ACCEPT MESSAGE FILES` with a window, that window no longer receives normal `On Drop`/external drag-and-drop events. Passing `0` turns the custom feature off again, but it cannot restore the standard drop handling for that window.
- **The callback runs in a new process.** It is spawned by the plugin, so it cannot read the local variables of the method that called `ACCEPT MESSAGE FILES`, and it should not be modal (avoid blocking `DIALOG`/`ALERT` in it, or you will hold the monitor process).
- **Do not abort the callback method.** If you abort its execution context, the monitor process keeps running but the callback is never called again until the structure is reopened.
- **The command takes exactly 3 parameters; the first two are mandatory.** There is no optional form.
- **Failure is silent.** The plugin does not raise a 4D error on any of the failure paths below; it either falls back to a different interpretation of the drop or simply does nothing.

---

## ACCEPT MESSAGE FILES

### Syntax

```4d
ACCEPT MESSAGE FILES ( window ; method { ; context } ) → *
```

### Parameters

| Parameter | Type | Description |
|---|---|---|
| `window` | Longint | Window reference of the target window, or `0` / `-1` to disable the handler |
| `method` | Text | Name of the callback method to run for each dropped message/file |
| `context` | Text | Optional arbitrary string that is passed back to the callback as `$4` |
| Result | — | The command returns nothing |

### Description

`ACCEPT MESSAGE FILES` replaces the window's standard OLE drag-and-drop handling with the plugin's own drop target (via `RevokeDragDrop`/`RegisterDragDrop`) and, for the callback side, starts a monitor process.

`window` selects the action:

- **A valid window reference** — registers the drop target on that window and starts the monitor process that dispatches the callback. The window is identified by its HWND (the 4D window reference returned by `Current form window`).
- **`0`** — unregisters the drop target and finishes the monitor process. The custom feature is disabled.
- **`-1`** — same as `0`, but lets the monitor process finish gracefully.

The callback `method` must already exist in the structure when the drop arrives (the plugin looks it up by name with `PA_GetMethodID`). It is called once per dropped item, with four Text parameters:

| Callback parameter | Contents |
|---|---|
| `$1` | The message content. For a direct Outlook drop this is the email's **HTML body**. For a file/folder drop it is the **dropped path**. |
| `$2` | The exported `.mht` path (direct Outlook drop only; empty for file drops). |
| `$3` | The exported `.msg` path (direct Outlook drop only; empty for file drops). |
| `$4` | The `context` string you passed to the command. |

Three drop sources are recognised, in this order:

1. **Direct Outlook drop.** The plugin queries Outlook's running instance and, for each selected mail item, saves an `.mht` (format `10`, `olMHTML`) and a `.msg` (format `9`, `olMSGUnicode`) copy into a temporary folder. `$1` receives the HTML body, `$2` the `.mht` path, `$3` the `.msg` path.
2. **Explorer file/folder drop.** Standard `CF_HDROP` / file-promise data. `$1` receives the dropped file or folder path; `$2` and `$3` are empty. A folder is reported as a single path (its contents are not enumerated).
3. **Chromium web-content drop.** When the drop carries Chromium "web custom MIME" data (`maillistrow` / `multimaillistconversationrows`), `$1` receives the dropped value; `$2`/`$3` are empty.

The monitor watches the temporary folder for the `rename to *.mht` event before firing the callback (a single Outlook export generates several intermediate file events). Temporary `.mht`/`.msg` files are written to the system temp folder and are **not cleaned up** by the plugin.

### Example

Install the handler on the current form's window, passing a context so one callback can serve several call sites:

```4d
// Form method / button
ACCEPT MESSAGE FILES (Current form window ; "On message dropped" ; "mail")
```

```4d
// Callback method "On message dropped"
C_TEXT($1)   // message HTML body, or dropped file/folder path
C_TEXT($2)   // exported .mht path (Outlook drop) or ""
C_TEXT($3)   // exported .msg path (Outlook drop) or ""
C_TEXT($4)   // the context string passed to ACCEPT MESSAGE FILES

If ($4 = "mail")
    // An Outlook message was dropped: the body is in $1.
    // Optionally, $2 and $3 point at saved .mht / .msg copies.
    SET WINDOW TITLE (Current form window ; "Dropped: " + $2)
End if
```

Disable the custom handler (for example on the form's `On Unload`):

```4d
ACCEPT MESSAGE FILES (0 ; "" ; "")
```

> **Note:** the exact commands you use to process `$1`/`$2`/`$3` (HTML parsing, reading the `.msg`, displaying the message) are ordinary 4D operations — combine them with your usual `TEXT TO DOCUMENT`/`Open document`/HTML commands. Check your 4D Language Reference for the exact syntax on your version.

---

## Error handling & troubleshooting

- **Drop does nothing and no error appears.** The most likely cause is that Outlook is not running (for direct Outlook drops). Outlook must be installed and running for the export to succeed; otherwise the drop is silently ignored. Check that Outlook is open before dropping.
- **Callback is not called after it was aborted once.** If you abort the callback method's execution context, the monitor process keeps running but the callback will not fire again until the structure file is reopened.
- **Standard drag-and-drop is gone on the window.** This is by design: installing the handler overrides `On Drop`. Passing `0` disables the custom handler but does not restore the standard drop behavior for that window.
- **`$2` / `$3` are empty for file drops.** Only direct Outlook drops produce `.mht`/`.msg` files. For Explorer drops the content is in `$1` only.
- **Temporary files accumulate.** The plugin writes `.mht`/`.msg` copies to the system temp folder and does not remove them. On long-running, drop-heavy servers you may want a scheduled cleanup of those files.
- **Platform parity.** This plugin is Windows-only. On macOS the command is a no-op — use [apple-file-promises](https://github.com/miyako/4d-plugin-apple-file-promises) there.

---

## Quick reference

```4d
// Enable on the current window
ACCEPT MESSAGE FILES (Current form window ; "On message dropped" ; "mail")

// Callback "On message dropped"
C_TEXT($1)   // HTML body (Outlook) or dropped file/folder path
C_TEXT($2)   // .mht path or ""
C_TEXT($3)   // .msg path or ""
C_TEXT($4)   // context
// ...process $1..$4...

// Disable the handler
ACCEPT MESSAGE FILES (0 ; "" ; "")
```