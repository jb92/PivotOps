/**
 * PivotOps — desktop (WinUI + WebView2) entry point.
 *
 * Installs the Office.js shim before the task pane bundle is evaluated, so the
 * existing add-in code runs unmodified against the native workbook host.
 */

import "./office-shim";
import "../taskpane/taskpane";
