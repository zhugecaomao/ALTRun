## Folder Sync: automatic search preview and window state

- Enter on a task result from the `sync <task>` keyword now starts its Preview immediately. Ordinary name matches and Preferences → Folder Sync → **Sync...** keep opening the Ready window; a task already in progress still opens its progress window.
- Folder Sync windows remember their normal size and position plus whether they were maximized, independently for each task. The saved geometry is restored on the next open and clamped to the current monitor work area.

Window state is stored in `Data/SyncPreviewWindows.json`.
