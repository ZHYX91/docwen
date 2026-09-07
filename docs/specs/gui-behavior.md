# GUI behavior / GUI 行为

The GUI renders one application/runtime state through view models. Widgets own presentation and interaction wiring; they do not become configuration or task-truth sources.

GUI 通过 ViewModel 呈现唯一的 application/runtime 状态。Widget 负责展示和交互接线，不成为配置或任务真相源。

## Main behavior / 主要行为

- Single and batch input modes support file dialog, drag/drop, filtering and stable ordering.
- Available conversion panels and actions derive from selected files and route capabilities.
- Batch rows expose truthful pending, processing, completed, partial, failed and cancelled states.
- Admitted input files cannot be removed until their worker finishes. Completed operation records own failure details and retry intent independently of editable batch rows; retry restores removed inputs through normal admission.
- Mixed batches use the intersection of routes for every admitted source format. XLSX protection analysis runs in the background with bounded file-identity caching; only the latest selection can publish results. Pending or unknown protection blocks ODS delivery. Empty protection elements are not enabled locks, and execution still revalidates inputs.
- Cancellation becomes reachable immediately, is idempotent and does not publish incomplete output.
- Settings use persisted, draft and preview layers; Apply/OK/Cancel and reset operations preserve their ownership boundaries.
- Light, dark and system themes, semantic typography presets and high-DPI geometry remain supported.
- Form rows reflow after font, style, text and geometry changes. System theme preference survives font changes and responds to Qt system color-scheme events.
- Keyboard shortcuts are suppressed while text-editing controls own focus.
- A second launch forwards files to the existing instance through IPC.

## Accessibility and feedback / 可访问性与反馈

Visible controls require localized labels or accessible names. Errors, warnings and confirmations use the shared feedback layer. Terminal summaries, history and retained artifacts must agree with runtime truth.

## Visual system / 视觉规范

- Shared cards use a left-aligned, theme-aware header band and 16 logical-pixel body padding. Main workflow and settings cards share this primitive.
- Regular controls have a 36 logical-pixel minimum height; execution buttons have a 40 logical-pixel minimum height and fill their card's content width. Short text buttons have an 80 logical-pixel minimum width. Font metrics may increase these dimensions.
- Related controls use an 8-pixel horizontal gap, functional groups and form rows use 12 pixels, and cards use 16 pixels. Qt handles system DPI scaling; these tokens are not scaled a second time.
- Checkbox labels follow the indicator on its right. Form fields align within a settings page; narrow rows and long localized text reflow without clipping. Proofreading choices use one column when two do not fit.
- Markdown target selection occupies a separate form row above the full-width generation action. The action names its target format. Batch conversion buttons show the same category count used by request dispatch; merge buttons show their own matching-input count.
- The workspace previews committed output-directory settings. Template selection has a persistent check marker and a visible selected name. Idle feedback is borderless; task, notification and result feedback retain their card.
- Light and dark themes use identical geometry, with readable hover, pressed, disabled and keyboard-focus states.
- Theme-independent design tokens and shared control geometry own dimensions and spacing. Theme colour changes reuse the same complete button box model; new themes do not redeclare padding or minimum sizes.

## Regression / 回归

Widget/view-model tests cover deterministic behavior. GUI smoke, screenshots and physical desktop interaction cover the final host-dependent surface. Current reference screenshots are stored under `docs/assets/screenshots/`.
