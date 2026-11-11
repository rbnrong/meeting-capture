# Contributing

Please open an issue describing the problem or proposed change before a substantial redesign. Keep fixes focused and include reproduction steps and relevant validation.

## Build locally

Use Windows 10/11 x64 with its .NET Framework desktop components enabled. Run `build.cmd`; the output is `dist\MeetingCapture.exe`. The source uses C# 5 so the compiler bundled with Windows can build it without installing an SDK.

Run `build-release.ps1` to build the portable ZIP and checksums. Generated outputs belong in `dist/` and must not be committed.

## Manual verification

Use a disposable output folder and a synthetic slide or test window.

- Paste a folder containing spaces with and without surrounding `"` or smart quotes. Confirm it is cleaned and used for saving.
- Select a region, cancel a second selection, and confirm the earlier selection remains usable.
- Capture repeatedly; check PNG dimensions, ordered filenames, and a counter that changes only after successful saves.
- Put the toolbar inside the region and check that it is absent from saved images.
- Test fullscreen PowerPoint capture and a quick double-click. Neither should change the active slide.
- Test a missing or unwritable folder. The error must be visible and the app must not switch folders silently.
- On multiple monitors, test different scaling settings, selection on a secondary monitor, and moving the toolbar without changing the selected region.
- Check that ending a session clears the region, and display changes pause capture until a new region is selected.

Remove disposable images, logs, and temporary build files when done. Do not commit personal slide content, local machine paths, settings, or credentials.

## Pull requests

Describe the problem, resulting behavior, and checks actually performed. Distinguish compilation from interactive verification. Keep the single-EXE distribution and dependency-free Windows build unless a change in requirements is agreed first.
