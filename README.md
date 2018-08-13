# 4d-plugin-message-file-drop
Accept drag and drop of messages from Microsoft Outlook

### Platform

| carbon | cocoa | win32 | win64 |
|:------:|:-----:|:---------:|:---------:|
|||<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|<img src="https://cloud.githubusercontent.com/assets/1725068/22371562/1b091f0a-e4db-11e6-8458-8653954a7cce.png" width="24" height="24" />|

**A similar solution for Mac is available**: [apple-file-promises](https://github.com/miyako/4d-plugin-apple-file-promises)

### Version

<img src="https://cloud.githubusercontent.com/assets/1725068/18940649/21945000-8645-11e6-86ed-4a0f800e5a73.png" width="32" height="32" /> <img src="https://cloud.githubusercontent.com/assets/1725068/18940648/2192ddba-8645-11e6-864d-6d5692d55717.png" width="32" height="32" /> <img src="https://user-images.githubusercontent.com/1725068/41266195-ddf767b2-6e30-11e8-9d6b-2adf6a9f57a5.png" width="32" height="32" />

### Releases

[2.1](https://github.com/miyako/4d-plugin-message-file-drop/releases/tag/2.1)

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
