// Clicking the toolbar icon opens the sender as Chrome's right-hand side panel,
// next to the Gmail tab you're already signed in to.
chrome.sidePanel
  .setPanelBehavior({ openPanelOnActionClick: true })
  .catch((err) => console.error(err));
