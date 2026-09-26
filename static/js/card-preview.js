// Independent preview lifecycle: debounce edits, cancel stale work, release images.
export function createCardPreview({ image, status, getRequest, fetcher = fetch }) {
  let timer, controller, objectUrl, version = 0;
  function update(message, error = false) {
    status.textContent = message;
    status.style.color = error ? 'var(--danger, #ff6b6b)' : 'var(--muted)';
  }
  async function render(current) {
    controller = new AbortController();
    image.style.opacity = '0.5';
    status.setAttribute('aria-busy', 'true');
    update('Preparing your card…');
    try {
      const { url, body } = getRequest();
      const response = await fetcher(url, {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body), signal: controller.signal, cache: 'no-store'
      });
      if (!response.ok || !response.headers.get('content-type')?.includes('image/png')) {
        const data = await response.json().catch(() => ({}));
        throw new Error(data.message || 'Preview unavailable. Check your connection or sign in again.');
      }
      const blob = await response.blob();
      if (current !== version) return;
      const next = URL.createObjectURL(blob);
      image.src = next;
      image.hidden = false;
      if (objectUrl) URL.revokeObjectURL(objectUrl);
      objectUrl = next;
      update('Print preview ready. Save the design to use it for downloads.');
    } catch (error) {
      if (current === version && error.name !== 'AbortError') {
        image.hidden = true;
        update(error.message, true);
      }
    } finally {
      if (current === version) {
        image.style.opacity = '1';
        status.setAttribute('aria-busy', 'false');
      }
    }
  }
  return {
    refresh() {
      const current = ++version;
      clearTimeout(timer);
      controller?.abort();
      update('Updating preview…');
      timer = setTimeout(() => render(current), 250);
    },
    dispose() {
      ++version;
      clearTimeout(timer);
      controller?.abort();
      if (objectUrl) URL.revokeObjectURL(objectUrl);
    }
  };
}
