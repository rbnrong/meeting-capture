# Project guidance

Meeting Capture is a Windows x64 desktop utility for repeated fixed-region PNG capture. It is not a scientific analysis pipeline and does not read or modify presentation files.

- Check the repository root and existing changes before editing. Preserve changes outside the current task.
- Build entry: `build.cmd`. Release packaging: `build-release.ps1`. The application source and embedded manifest/icon are in `app/`.
- Keep compatibility with the built-in .NET Framework compiler and C# 5. Avoid new runtime or installer requirements without an explicit product decision.
- Preserve quoted-path paste handling, collision-safe PNG publication, successful-save-only counts, physical-pixel coordinates, and foreground preservation.
- The temporary mouse hook must remain limited to capture time and the toolbar's former position. It must be released on success, failure, and close.
- Use disposable fixtures for checks and remove generated test images and logs afterwards. Never commit `dist/`, settings, credentials, or personal screenshots.
- The minimum acceptance check is a clean Windows build plus targeted interactive checks for changed behavior. CI compilation is not proof of PowerPoint, mixed-DPI, or clean-machine behavior.
- Record only durable user-facing changes and validation limits in existing documentation. Do not add transient test logs or machine-specific audit reports.
