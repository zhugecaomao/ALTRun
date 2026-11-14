**English** · [中文](Recent-zh)

# Pinned and recent items

Before you type anything, the search window lists (pinned items have a pin at the bottom right of their icon):

![Empty search box: pinned and recent items](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/empty.png)
- **Pinned items** (shown by default): on any result, press `→` to open the action panel and choose **Pin to Empty Search Box**. Pinned items are saved in `ALTRun.json` (`Features.Recent.Pinned`), so they travel with your settings to another PC. To change their order or unpin several at once: Preferences → Search Window → Empty search box → "Show pinned items" → **Edit...** (Move Up / Move Down / Unpin)
- **Recently opened items**: not shown by default; check Preferences → Search Window → Empty search box → "Show the last [5] recently used items" (the number can be changed). **Clear...** next to it forgets them all (pinned items stay)
- **Gray text in the search box**: "ALTRun Search" by default; change it in Preferences → Search Window → Empty search box → "Placeholder text" (leave blank for the default, or type a single space to show nothing)

![Preferences → Search Window → Edit...: the pinned items](https://raw.githubusercontent.com/zhugecaomao/ALTRun/main/docs/images/screenshots/pinned-editor.png)

Only results you can open again are remembered: apps, files, folders, URLs, custom commands and system commands; clipboard history, calculations, snippets, web searches and windows are not. Select an item and press `Ctrl+Del` to unpin it or remove it from the recent items; Preferences → Advanced → "Reset Learned Ranking" clears the recent items too. Uncheck both boxes to show neither.
