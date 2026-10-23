**English** · [中文](Privacy-zh)

# Privacy and network access

ALTRun collects and uploads nothing: no statistics, telemetry or ads. Searches, clipboard history and learned ranking are processed and stored only on your PC.

## When ALTRun goes online
ALTRun itself only accesses the network in these cases:

| When | Where | What for | How to turn it off |
|---|---|---|---|
| Automatic update check: 1 minute after startup (if the last check was over 1 hour ago), then every 6 hours while it keeps running | `api.github.com` | Reads the latest release's version number; sends nothing about your PC | Preferences → General → "Check for updates automatically" (`General.CheckForUpdates`) |
| One-key update: when you press "Install update" | `github.com` (the release download) | Downloads the new version's zip and checks GitHub's SHA256 before replacing anything | Nothing is downloaded unless you press it |
| Currency conversion (off by default): 1.5 minutes after startup, when the last download is over 12 hours old | `api.frankfurter.dev` | Downloads reference exchange rates published by the ECB and other central banks, saved in `Data\Currency.json` | Preferences → Calculator → "Currency conversion" (`Features.Calculator.Currency`) |
| Download Site Icon: when you click "Download Site Icon" in the web search editor | The website in the URL | Downloads that site's icon and saves it in `Data\Icons` | Never accessed unless you click it |

Apart from that, web searches, bookmarks, web address commands and help links are handed to your default browser; ALTRun itself does not contact those sites. Talking to Everything and Total Commander happens locally (Windows messages), not over the network.

## What is stored on your PC
Everything lives in the Data folder and a few local folders; what each file is, whether it can be deleted and how to back it up is in [Data files](Data-Files). Points related to privacy:
- **Clipboard history**: content that password managers and other programs mark as "do not record" (`ExcludeClipboardContentFromMonitorProcessing`, `CanIncludeInClipboardHistory = 0`) is never recorded, and neither is content copied in the programs listed under Preferences → Clipboard → "Never record from" (a few password managers by default). To leave no history on disk, set `Features.Clipboard.Persist` to 0: the history stays in memory and is gone when ALTRun quits; or set Preferences → Clipboard → "Save in" to keep everything on this PC only, so it doesn't follow a Data folder in a sync folder. See [Clipboard history](Clipboard)
- **Usage statistics** (`Usage.json`): only how many times each feature was used per day; never what you typed or opened
- **Learned ranking** (`Knowledge.json`): "what was typed → which item was picked", used to put frequent picks first; Preferences → Advanced → "Reset Learned Ranking" clears it
- **Snippet expansion** never triggers in Windows password boxes, and Preferences → Snippets → "[Don't expand in](Snippets#dont-expand-in)" can exclude Remote Desktop, password managers and so on
- **Debug log** is off by default; Preferences → Advanced → "Write a debug log" writes `%Temp%\ALTRun.log` for troubleshooting, with timings of startup and background tasks and error messages (never what you type in the search box). Look it over before attaching it to an issue, see [How do I send a debug log?](FAQ#how-do-i-send-a-debug-log)

## Permissions
ALTRun runs with normal user rights: no administrator rights, no system-wide registry entries, no background service. The one exception: ticking Preferences → Keyboard → "Also in windows running as administrator" creates a scheduled task that runs a small Caps Lock-only helper as administrator (one UAC confirmation when it is created and one when it is removed), see [Caps Lock switches the input method](Hotkeys#caps-lock-switches-the-input-method).

## Removing everything
See [Uninstalling](Installation#uninstalling).

For more on security (reporting security issues, verifying release files) see the [security policy](https://github.com/zhugecaomao/ALTRun/blob/main/SECURITY.md).
