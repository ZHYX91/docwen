# GUI behavior / GUI 行为

The GUI renders one application/runtime state through view models. Widgets own presentation and interaction wiring; they do not become configuration or task-truth sources.

GUI 通过 ViewModel 呈现唯一的 application/runtime 状态。Widget 负责展示和交互接线，不成为配置或任务真相源。

## Main behavior / 主要行为

- Single and batch input modes support file dialog, drag/drop, filtering and stable ordering.
- Available conversion panels and actions derive from selected files and route capabilities.
- Batch rows expose truthful pending, processing, completed, partial, failed and cancelled states.
- Cancellation becomes reachable immediately, is idempotent and does not publish incomplete output.
- Settings use persisted, draft and preview layers; Apply/OK/Cancel and reset operations preserve their ownership boundaries.
- Light, dark and system themes, semantic typography presets and high-DPI geometry remain supported.
- Keyboard shortcuts are suppressed while text-editing controls own focus.
- A second launch forwards files to the existing instance through IPC.

## Accessibility and feedback / 可访问性与反馈

Visible controls require localized labels or accessible names. Errors, warnings and confirmations use the shared feedback layer. Terminal summaries, history and retained artifacts must agree with runtime truth.

## Regression / 回归

Widget/view-model tests cover deterministic behavior. GUI smoke, screenshots and physical desktop interaction cover the final host-dependent surface. Current reference screenshots are stored under `docs/assets/screenshots/`.
