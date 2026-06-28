/**
 * PivotOps — Ribbon Commands
 * Functions invoked directly from the ribbon (outside the task pane).
 */

Office.onReady(() => {
  // Register command functions
});

function openDashboard(event: Office.AddinCommands.Event): void {
  // The dashboard button opens the task pane via ShowTaskpane action
  // This is a fallback for ExecuteFunction
  event.completed();
}

// Expose to Office runtime
(globalThis as Record<string, unknown>).openDashboard = openDashboard;
