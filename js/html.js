// HTML, markdown, and text helpers.
export function installHtml(app) {
  app.renderMarkdown = function renderMarkdown(text, keyOverride) {
    const keyToUse = keyOverride || app.currentAiCacheKey;

    // Use marked library
    // breaks: true allows newlines to be rendered as <br>
    const htmlContent = marked.parse(text, { breaks: true });

    const div = document.createElement('div');
    div.innerHTML = htmlContent;

    // Enable checkboxes and restore state
    div.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
      checkbox.removeAttribute('disabled');
      const li = checkbox.closest('li');
      if (li) {
        li.style.listStyleType = 'none'; // Hide bullet for task items
        const label = li.textContent.trim();
        const storageKey = keyToUse ? (keyToUse + '_chk_' + encodeURIComponent(label)) : null;

        if (storageKey) {
          checkbox.dataset.key = storageKey;
          const stored = app.db.get(storageKey);
          if (stored !== null) {
            checkbox.checked = stored === 'true';
          }
        }
      }
    });

    // Enforce secure links
    div.querySelectorAll('a').forEach(a => {
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
    });

    return div.innerHTML;
  };

  app.escapeHtml = function escapeHtml(s) {
    const div = document.createElement('div');
    div.textContent = s;
    return div.innerHTML;
  };

  app.getTextContent = function getTextContent(htmlStr) {
    const div = document.createElement('div');
    div.innerHTML = htmlStr;
    return div.textContent || '';
  };

  app.linkifyAndSanitize = function linkifyAndSanitize(str) {
    const div = document.createElement('div');
    div.innerHTML = str;
    // Convert newlines to <br> to ensure they are preserved as breaks
    div.innerHTML = (str || '').replace(/\n/g, '<br>');

    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) {
        return app.escapeHtml(node.textContent);
      }
      if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'BR') return '<br>';
        if (node.tagName === 'A' && node.hasAttribute('href')) {
          const a = document.createElement('a');
          a.href = node.getAttribute('href');
          a.target = '_blank';
          a.rel = 'noopener noreferrer';
          a.innerHTML = Array.from(node.childNodes).map(walk).join('');
          return a.outerHTML;
        }
        return Array.from(node.childNodes).map(walk).join('');
      }
      return '';
    }
    return Array.from(div.childNodes).map(walk).join('');
  };

  app.simplifyHtml = function simplifyHtml(html) {
    const div = document.createElement('div');
    div.innerHTML = html;
    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) return node.textContent;
      if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'A' && node.hasAttribute('href')) {
          const content = Array.from(node.childNodes).map(walk).join('');
          return `<a href="${node.getAttribute('href')}">${content}</a>`;
        }
        if (node.tagName === 'BR') return '\n';
        if (node.tagName === 'P' || node.tagName === 'DIV') return '\n' + Array.from(node.childNodes).map(walk).join('');
        return Array.from(node.childNodes).map(walk).join('');
      }
      return '';
    }
    return Array.from(div.childNodes).map(walk).join('').trim();
  };
}
