/* eslint-disable no-console */

// Ribbon / function commands (lightweight – no Vue needed)
Office.onReady(() => {
  // Register command handlers
  Office.actions.associate("ShowTaskpane", showTaskpane);
});

function showTaskpane(_event: Office.AddinCommands.Event) {
  // Surface the task pane
  Office.addin.showAsTaskpane();
  _event.completed();
}
