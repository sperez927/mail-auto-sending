// Timers inside a dedicated worker are not throttled when the dashboard tab
// is in the background, so the random delays stay accurate.
self.onmessage = (e) => {
  const { id, ms } = e.data;
  setTimeout(() => self.postMessage({ id }), ms);
};
