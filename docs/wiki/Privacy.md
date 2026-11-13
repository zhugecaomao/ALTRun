**English** · [中文](Privacy-zh)

# Privacy and network access

ALTRun collects and uploads nothing: no statistics, telemetry or ads. Searches, clipboard history and learned ranking are processed and stored only on your PC.

## When ALTRun goes online
ALTRun itself only goes online to check for updates and install one you choose, to download exchange rates when currency conversion is turned on (off by default), and to download a web search's icon when you click "Download Site Icon". When each happens, where to, and how to turn it off: [Security policy → Network connections](https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md#network-connections) (the one list, kept with the privacy policy).

Apart from that, web searches, bookmarks, web address commands and help links are handed to your browser (the default one, or the one in Preferences → General → Browser); ALTRun itself does not contact those sites. Talking to Everything and Total Commander happens locally (Windows messages), not over the network.

## What is stored on your PC
Everything lives in the Data folder and a few local folders; what each file is, whether it can be deleted and how to back it up is in [Data files](Data-Files). Points related to privacy:
- **Clipboard history**: content that password managers and other programs mark as "do not record" (`ExcludeClipboardContentFromMonitorProcessing`, `CanIncludeInClipboardHistory = 0`) is never recorded, and neither is content copied in the programs listed under Preferences → Clipboard → "Never record from" (a few password managers by default). To leave no history on disk, set `Features.Clipboard.Persist` to 0: the history stays in memory and is gone when ALTRun quits; or set Preferences → Clipboard → "Save in" to keep everything on this PC only, so it doesn't follow a Data folder in a sync folder. See [Clipboard history](Clipboard)
- **Browser history** (`h words`): read from the browsers' own files on this PC only when you type the keyword; ALTRun keeps no copy of it. Turn it off in Preferences → Features → "Browser history". See [Browser history](Bookmarks#browser-history)
- **Usage statistics** (`Usage.json`): only how many times each feature (and each system command) was used per day; never what you typed or opened
- **Learned ranking** (`Knowledge.json`): "what was typed → which item was picked", used to put frequent picks first; only the 3000 items picked most recently are kept; Preferences → Advanced → "Reset Learned Ranking" clears it
- **Snippet expansion** never triggers in Windows password boxes, and Preferences → Snippets → "[Don't expand in](Snippets#dont-expand-in)" can exclude Remote Desktop, password managers and so on
- **Debug log** is off by default; Preferences → Advanced → "Write a debug log" writes `%Temp%\ALTRun.log` for troubleshooting, with timings of startup and background tasks and error messages (never what you type in the search box). Look it over before attaching it to an issue, see [How do I send a debug log?](FAQ#how-do-i-send-a-debug-log)

## Permissions
ALTRun runs with normal user rights: no administrator rights, no system-wide registry entries, no background service. The one exception: ticking Preferences → Keyboard → "Also in windows running as administrator" creates a scheduled task that runs a small Caps Lock-only helper as administrator (one UAC confirmation when it is created and one when it is removed), see [Caps Lock switches the input method](Hotkeys#caps-lock-switches-the-input-method).

## Removing everything
See [Uninstalling](Installation#uninstalling).

For more on security (reporting security issues, verifying release files) see the [security policy](https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md).
