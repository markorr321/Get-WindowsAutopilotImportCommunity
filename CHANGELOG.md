# Changelog

All notable changes to this project are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/); versioning is [SemVer](https://semver.org/).

## [1.2.2] - 2026-07-27

### Fixed

- **Autopilot diagnostics results vanished the moment the run finished.** The report streamed
  into the Logs pane correctly, then disappeared and left a pane with nothing in it — no
  content, and therefore no scrollbar — so the results could never actually be read.

  `Update-ApLogsPage` replaces that pane with the GUI's own session log, and it was called
  unconditionally whenever *any* run completed. For a diagnostics run, whose output box **is**
  the Logs pane, that overwrote a 40,000-character report with a few dozen lines of session
  log. Navigating away from Logs and back did it again, via the nav handler.

  The pane now tracks which of its two owners is on screen. A run that streams there keeps it
  until the operator clicks **Refresh**, which is an explicit request for the session log.
  Reaching the end of a run no longer discards what the run produced.

- A finished diagnostics report now scrolls back to the **top** rather than staying pinned to
  the tail where streaming left it, and the run's log-file path is appended to the output, so
  the full report is findable on disk regardless.

## [1.2.1] - 2026-07-27

### Fixed

- **Text typed into the registration fields was invisible.** Assigned user, computer name, add
  to Entra group, the batch import CSV path and the Advanced product key all accepted input,
  kept it, and passed it to the engine — but painted nothing on screen. Most visible in
  Autopilot v1, because v2 hides that whole card.

  The theme hand-templates `TextBox` and bound the control's `Padding` to the content host's
  `Margin`. WPF's text host applies `TextBox.Padding` itself, as the stock template relies on,
  so the inset landed twice: a 38px-high field lost 16px to the margin and another 16px
  internally, leaving a 4px line box that an 18.6px glyph run could not render into. The caret
  still blinked, which is why the fields looked merely empty rather than broken. The group tag
  was unaffected only because `DarkComboBox` uses `Padding="10,0"` and so lost nothing
  vertically. Present since 1.0.0.

  Fixed by dropping the `Margin` binding, matching the stock WPF template. The distribution
  harness now measures the line box of every editable `TextBox` on every page and fails if one
  is shorter than its own font size — the condition is invisible to a parse check, and this
  guard reproduces the bug on the pre-fix build.

- **The Logs and Live output panes would not scroll sideways.** The `ScrollBar` style set
  `Width`/`MinWidth` unconditionally, which is right for a vertical bar and wrong for a
  horizontal one: instead of spanning the viewport it collapsed to a 12×12 stub in the corner.
  Log lines are deliberately unwrapped, so anything past the right edge — most engine
  narration — was simply unreachable. Released via an `Orientation` trigger, and the thumb is
  now inset evenly on all four sides so it reads correctly either way.

- **Scrolling up during a run was undone 250ms later.** Every batch of engine output called
  `ScrollToEnd()`, so reading back through a live run was impossible. Output now follows the
  tail only when the view is already at the bottom; scroll up and the position holds, return to
  the bottom and it resumes following.

- **The vertical scrollbar was effectively invisible.** It was there and it worked, but a 6px
  `#3A3A3A` thumb on the `#121212` output pane, over a transparent trough, is near-black on
  black — indistinguishable from no scrollbar. Bars are now 14px with a visible `#242424`
  track, the thumb is `#6A6A6A` (brightening on hover, accent while dragging), and log panes
  keep the vertical bar on screen permanently instead of `Auto`, which had it appear and vanish
  as output arrived.

- **The junction of the two scrollbars was a white square.** `ScrollViewer` is not
  hand-templated, and the stock template fills that corner with `SystemColors.ControlBrush`
  (#F0F0F0). The theme now overrides the brush key.

  All four are covered by new harness checks: both bars visible and correctly sized on a pane
  overflowing each way, and no pale corner rectangle.

## [1.2.0] - 2026-07-27

### Added

- **`-Online` mode for Autopilot diagnostics.** The Advanced page ran
  `Get-AutopilotDiagnosticsCommunity.ps1` with no parameters, so every app, policy and
  platform script in the ESP output appeared as a bare GUID. A new **Resolve app and policy
  names from Intune** checkbox passes `-Online`, which signs in read-only
  (`DeviceManagementApps.Read.All`, `DeviceManagementConfiguration.Read.All`) and resolves
  those GUIDs to display names. Persisted as `diagnosticsOnline` in `config.json`.

  Off by default, because the local read is the case that has to work in OOBE: without an
  `Online` key in the parameter hashtable the launcher skips the sign-in prep block entirely,
  so a plain diagnostics run still touches neither the PowerShell Gallery nor a browser. When
  the switch is on, the run reuses the same prerequisite handling as a registration — and warns
  up front if the Graph module is installed but unloadable, which would otherwise abort the run
  before any sign-in prompt appeared.

## [1.1.0] - 2026-07-26

### Added

- **Restart after import for Device Preparation (v2).** The engine cannot do this: its
  `-Reboot` switch sits inside the assignment-wait block that only the Autopilot v1 path
  reaches, so `-identifier` runs ignore it entirely. The restart is therefore performed by the
  GUI once the import reports success, via `Restart-Computer -Force` with a `shutdown.exe /r`
  fallback. No new switch is sent to the engine; v2 runs still pass only `-identifier -Online`.

  It only fires on a clean run — never after a failure or a cancel, because rebooting then
  would throw away the log and leave the device unregistered at OOBE. Ticking it prompts for
  confirmation first, since Device Preparation targets devices through the Entra group on the
  policy: restarting before this device is a member of that group returns it to OOBE before the
  policy can apply. Persisted as `rebootAfterV2Import` in `config.json`, off by default.

- **A "Restart now" button** in v2 mode, in the pinned actions row beside Register. This is
  usually the one you actually want: import the identifier, add the device to the policy's
  Entra group, then come back and restart. The automatic checkbox cannot wait for that group
  membership, so on its own it restarts too early for most workflows.

  It refuses while a run is in progress, and confirms before restarting. It lives in the
  actions row rather than the Options card deliberately — placed in the card it laid out at
  y=498 inside a form area that stops scrolling around y=470, so it was clipped below the fold
  and effectively invisible. The harness now asserts it is not inside a `ScrollViewer`.

### Changed

- The Options card now swaps its contents by mode instead of dimming: v1 shows the assignment
  wait, reboot and existing-device policy; v2 shows only the restart option.
- The whole **Registration details** card is hidden in v2, not just the group tag. Every field
  in it (group tag, assigned user, computer name, Entra group) is ignored by the identifier
  path, and leaving them on screen dimmed pushed the one live v2 option below the fold.

## [1.0.0] - 2026-07-26

Published to the PowerShell Gallery as
[Get-WindowsAutopilotImportGUICommunity](https://www.powershellgallery.com/packages/Get-WindowsAutopilotImportGUICommunity).
Verified after publishing: the downloaded package is byte-identical to `dist\`, and the engine
it extracts still reports a Valid Authenticode signature.

First release. A ground-up rewrite of [ugurkocde/AutoPilot_Import_GUI](https://github.com/ugurkocde/AutoPilot_Import_GUI) built around
[Andrew S Taylor's Windows Autopilot Community script](https://github.com/andrew-s-taylor/WindowsAutopilotInfo) instead of Michael Niehaus'
original, which is what makes Autopilot v2 possible.

### Added

- **Autopilot v2 (Device Preparation) support** via the engine's `-identifier` path, importing the
  `Manufacturer,Model,Serial` identifier to `deviceManagement/importedDeviceIdentities`. The
  make and model are normalised exactly as the engine does (trim, strip `.` and `,`) so the
  identifier previewed in the UI is the one that gets imported.
- Resizable dark window with sidebar navigation across six pages: Register, Device, Batch
  import, Network check, Advanced, Logs. Palette, typography and the
  `PrimaryButton` / `FieldLabel` / `Segment` control styles are shared with VM-Pilot so the
  tools read as one family.
- **Staged progress** driven by parsing the engine's own output: Connect → Collect → Import →
  Sync → Assign → Complete, with a working cancel button that terminates the whole process
  tree (the engine can spawn sysprep, changepk and the pre-provisioning helper).
- **Preview command** dry-run that renders the exact engine invocation without executing it.
- Group tag as an editable dropdown backed by remembered history in `config.json`.
- Explicit handling for an already-registered serial: update tag, delete and re-add, or
  assume-new (`-newdevice`, which skips the engine's full-tenant enumeration).
- Assigned user, computer name, and Entra group membership (comma or semicolon separated).
- Batch CSV import (`-InputFile`) for v1 and v2, with the expected column layouts documented
  in the UI.
- Offline export with no tenant connection: hardware hash CSV, device identifier CSV, partner
  CSV format, append mode, and clipboard copies of both the hash and the identifier.
- Advanced post-assignment actions: pre-provisioning, sysprep, Intune wipe, product key
  change. Destructive combinations require confirmation that shows the command to be run.
- Autopilot diagnostics via the vendored `Get-AutopilotDiagnosticsCommunity.ps1`.
- Engine integrity verification: SHA256 against `vendor\VERSION.json` plus Authenticode status.
- Session logging to `%ProgramData%\AutopilotImportGUI\Logs\` with an in-app Logs page,
  replacing the original's write to the root of `C:\`.
- Single-file build (`build.ps1`) that inlines both XAML documents and embeds the vendored
  scripts as base64 of their exact bytes, so the Authenticode signatures survive. Verified at
  build time and again at runtime.
- 75 Pester unit tests over the argument builder, progress parser and configuration layer,
  plus a 29-check distribution smoke test. `build.ps1` will not produce a build if any fail.
- Branding and attribution: a persistent credit in the sidebar footer, an **About** card on the
  Advanced page with clickable links, `PSScriptInfo` metadata on both the development entry
  point and the generated build (so the version and author survive a PowerShell Gallery
  publish), and the author line in every session log header. Links open through `Start-ApUrl`,
  which accepts only `http` and `https` — handing an arbitrary string to `Start-Process` would
  otherwise let a UI link launch a local executable.

### Fixed

Carried over from the original project:

- **The network connectivity check never worked.** Every one of its ~29 probes called
  `Write-Output -NoNewline -ForegroundColor ...`; `Write-Output` has neither parameter, so each
  probe threw a `ParameterBindingException` that was then swallowed by
  `$ErrorActionPreference = 'SilentlyContinue'`. The technician saw an empty window and
  concluded the network was fine. Rewritten with raw `TcpClient` probes on a runspace pool:
  26 endpoints in roughly 0.4 s, with latency, required-vs-optional weighting and a
  colour-coded grid.
- The endpoint list probed the bare host `azure.net`, which does not resolve; replaced with the
  documented TPM attestation endpoint. `delivery.mp.microsoft.com` was likewise a
  wildcard-only base that never resolves, so real hosts are probed instead.
- Reachability was tested by pinging `8.8.8.8`, which fails on any network that blocks ICMP or
  public DNS while Microsoft endpoints work fine. Now a TCP 443 connect to
  `login.microsoftonline.com`.
- Free space summed every logical disk, overstating it on multi-disk machines. Now reports the
  system drive.
- Deprecated `Get-WmiObject` replaced with CIM throughout.
- `Set-Location "$env:ProgramFiles\WindowsPowerShell\Scripts"` and the assumption that the
  engine is installed there are both gone; the engine is embedded and extracted per session.
- Elevation happened per button press via `-Verb RunAs`. Now the GUI elevates once at launch and
  child runs inherit the token: one UAC prompt per session.
- Windows Update no longer needs `PSWindowsUpdate` from PSGallery. It uses the in-box Windows
  Update agent COM API, which matters on a restricted OOBE network — and avoids generating a
  script that chains `Install-PackageProvider`, `Set-PSRepository -Trusted` and
  `Install-Module -Force`, a shape Defender's AMSI scanner blocks on sight.

### Engine constraints encoded in the argument builder

These three engine behaviours fail *silently* if the caller gets them wrong, so they are
enforced in one tested place rather than trusted to the UI:

- `-Reboot`, `-Wipe`, `-Sysprep`, `-preprov` and `-ChangePK` are all nested inside
  `if ($Assign)`. Without `-Assign` they are accepted and never executed. The builder forces
  `-Assign` on and says so.
- With neither `-delete` nor `-updatetag`, an already-registered serial reaches
  `Read-Host "Do you want to delete or update?"` and blocks forever behind a hidden console.
  Every online run now carries an explicit policy plus `-Force`.
- `-identifier` takes an entirely separate branch that ignores `GroupTag`, `AssignedUser`,
  `AssignedComputerName`, `AddToGroup`, `Assign` and `Reboot`. The UI disables those controls in
  v2 mode and warns instead of silently dropping values.

### Performance

- Device information is read on a background runspace triggered by the window's `Loaded`
  event, not synchronously during construction. When the process is not elevated,
  `root/cimv2/security/microsofttpm` and `root/cimv2/mdm/dmmap` are inaccessible and each
  query sits on a ~5 second DCOM timeout, so building the window synchronously took
  **10,479 ms** before anything appeared. It is now **408 ms**, with the panes showing
  "reading..." until the data lands.

### Notes for maintainers

- The engine writes almost exclusively with `Write-Host`. Capturing it requires `*>&1` in the
  launcher; without the `*` the run log is empty.
- The run log is written through a single `StreamWriter` held open for the whole run.
  `Add-Content` per line was tried first and intermittently *lost lines*: it opens and closes
  the file on each call, so it collides with the GUI's tail reader, and the failure is
  non-terminating — the line vanishes and the run continues. Reproduced by diffing streamed
  output against the log file.
- The tail reader buffers an incomplete trailing line between polls. Without that, a read
  landing mid-line surfaces one log line as two fragments and the progress regexes stop
  matching.
- Do not rename `Start-ApCsvExport` to the more obvious `Invoke-Ap` + `Export` spelling. That
  identifier lowercases into a string that collides with a Microsoft Defender HackTool
  signature family, and Defender then refuses to load the entire file with
  `ScriptContainedMaliciousContent`. Confirmed by elimination: a file containing only that one
  function was blocked while all 22 other functions in the same file loaded cleanly.
- **An exception escaping a `DispatcherTimer` tick propagates out of `Window.ShowDialog` and
  terminates the application.** It is not merely logged. Every tick handler here is wrapped in
  try/catch for that reason; without it the window vanished a second after opening.
- Inside a `GetNewClosure()` scriptblock, `$script:` binds to the closure's own captured
  session state, *not* the defining script's scope. `$script:ApBackgroundJobs.Remove(...)` in
  the background tick was therefore a method call on `$null`; it threw from a `finally` block,
  the continuation never ran, and the UI sat on "reading..." with nothing logged. Capture such
  state into a local before creating the closure.
- A WPF event handler receives `(sender, eventArgs)` positionally. `$_` is not the event args:
  `$_.Cancel = $true` in the `Closing` handler threw "Argument types do not match" out of
  `Window.OnClosing`, so the cancel-the-close path never worked.
- Local variable names are case-insensitive and will silently overwrite a parameter of the same
  name. A local `$border` inside a function taking a `$Border` parameter replaced the control
  with a colour string.
- XML comments cannot contain `--`, so em-dash-style double hyphens break the XAML parse.
- `RenderTargetBitmap` will not capture a page switched from `Collapsed` to `Visible` in the
  same tick — it has been arranged but not rendered. Pump the dispatcher to `Loaded` priority
  before capturing screenshots.

[1.0.0]: https://github.com/markorr321/Get-WindowsAutopilotImportCommunity/releases/tag/v1.0.0
