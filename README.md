# 4d-plugin-message-file-drop
Accept drag and drop of messages from Microsoft Outlook

![version](https://img.shields.io/badge/version-17%2B-3E8B93)
![platform](https://img.shields.io/static/v1?label=platform&message=win-64&color=blue)
[![license](https://img.shields.io/github/license/miyako/4d-plugin-message-file-drop)](LICENSE)
![downloads](https://img.shields.io/github/downloads/miyako/4d-plugin-message-file-drop/total)

**Note**: for v17 and earlier, move `manifest.json` to `Contents`

**A similar solution for Mac is available**: [apple-file-promises](https://github.com/miyako/4d-plugin-apple-file-promises)

## Syntax

```
ACCEPT MESSAGE FILES (window;method{;context})
```

Parameter|Type|Description
------------|------------|----
window|LONGINT|windowRef or ``0``
method|TEXT|callback method(``$1:text;$2:text``)
context|TEXT|any string (``$2`` passed to ``method``)

Configures the specified window to implement a custom [OLE drag and drop](https://msdn.microsoft.com/en-us/library/96826a87.aspx) feature, over-riding standard form events. 

**Note**: You can disable the custom "OLE drag and drop" feature by passing ``0`` to the command. However, it is not possible to restore standard drop events (from external source) for that window. 

When an email is dropped directly from Outlook, a VBA script is executed (``csript.exe``) which asks Outlook to export the selected messages to a temporary folder in ``.mht`` format. 

The plugin is monitoring the temporary folder in a background thread. The specified callback method is executed when an ``.mht`` file is created.

**Do not abort the callback method**. If you abort the execution context, the process will keep running but the method will not longer be called from the plugin until you reopen the structure file.

---

#### Technical Details

When Outlook exports an email in  ``.mht`` format, several file system events are generated:

1. create "name.mh_"
2. modify "name.mh_"
3. rename from "name.mh_"
4. rename to "name.mht" 
5. create "~$name.mht"
6. modify "~$name.mht"
7. remove "~$name.mht"

In addition, the script creates a new subfolder for each drop event.

The callback is called when the event is "rename to", not a folder, has the  extension "mht" and does not have the temporary file prefix (``~$``).

The callback method is also executed for standard file drop events. No name checks are performed in this instance. Multiple items call multiple events. However, when a folder with contents is dropped, on the path of the folder (not its sub contents) are reported.
