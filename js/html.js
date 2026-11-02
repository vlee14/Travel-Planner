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

  app.safeHttpUrl = function safeHttpUrl(href) {
    const raw = (href || '').trim();
    if (!raw) return '';
    const withScheme = /^www\./i.test(raw) ? `https://${raw}` : raw;
    try {
      const url = new URL(withScheme);
      if (url.protocol !== 'http:' && url.protocol !== 'https:') return '';
      return url.href;
    } catch (err) {
      return '';
    }
  };

  app.getTextContent = function getTextContent(htmlStr) {
    const div = document.createElement('div');
    div.innerHTML = htmlStr || '';
    const text = div.textContent || '';
    return text.replace(/\[([^\]\n]+)\]\((https?:\/\/[^\s)]+)\)/gi, '$1 ($2)');
  };

  app.activityEditorText = function activityEditorText(html) {
    const div = document.createElement('div');
    div.innerHTML = html || '';
    function walk(node) {
      if (node.nodeType === Node.TEXT_NODE) return node.textContent;
      if (node.nodeType !== Node.ELEMENT_NODE) return '';
      if (node.tagName === 'BR') return '\n';
      const inner = Array.from(node.childNodes).map(walk).join('');
      if (node.tagName === 'A') {
        const href = app.safeHttpUrl(node.getAttribute('href') || '');
        const label = inner.replace(/\s+/g, ' ').trim().replace(/[\[\]]/g, '');
        if (!href) return label;
        return `[${label || href}](${href.replace(/[()]/g, match => (match === '(' ? '%28' : '%29'))})`;
      }
      if (node.tagName === 'P' || node.tagName === 'DIV' || node.tagName === 'LI') {
        return `${inner.replace(/[ \t]+$/g, '')}\n`;
      }
      return inner;
    }
    return Array.from(div.childNodes).map(walk).join('')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n{3,}/g, '\n\n')
      .trim();
  };

  app.linkifyPlainText = function linkifyPlainText(text) {
    const str = text || '';
    const re = /((?:https?:\/\/|www\.)[^\s<]+)/gi;
    let out = '';
    let last = 0;
    for (const match of str.matchAll(re)) {
      const end = match.index + match[0].length;
      const trail = match[0].match(/[.,);:!?]+$/)?.[0] || '';
      const raw = trail ? match[0].slice(0, -trail.length) : match[0];
      const href = app.safeHttpUrl(raw);
      out += app.escapeHtml(str.slice(last, match.index));
      if (!href) out += app.escapeHtml(match[0]);
      else {
        out += `<a href="${href}" target="_blank" rel="noopener noreferrer">${app.escapeHtml(raw)}</a>`;
        out += app.escapeHtml(trail);
      }
      last = end;
    }
    out += app.escapeHtml(str.slice(last));
    return out;
  };

  app.linkifyAndSanitize = function linkifyAndSanitize(str) {
    const withMarkdown = (str || '').replace(/\[([^\]\n]+)\]\(([^)\s]+)\)/g, (match, label, href) => {
      const url = app.safeHttpUrl(href);
      if (!url) return match;
      return `<a href="${url}">${app.escapeHtml(label)}</a>`;
    });
    const div = document.createElement('div');
    div.innerHTML = withMarkdown.replace(/\n/g, '<br>');

    function walk(node, inLink) {
      if (node.nodeType === Node.TEXT_NODE) {
        return inLink ? app.escapeHtml(node.textContent) : app.linkifyPlainText(node.textContent);
      }
      if (node.nodeType === Node.ELEMENT_NODE) {
        if (node.tagName === 'BR') return '<br>';
        if (node.tagName === 'A' && node.hasAttribute('href')) {
          const href = app.safeHttpUrl(node.getAttribute('href'));
          const inner = Array.from(node.childNodes).map(child => walk(child, true)).join('');
          if (!href) return inner;
          const a = document.createElement('a');
          a.href = href;
          a.target = '_blank';
          a.rel = 'noopener noreferrer';
          a.innerHTML = inner;
          return a.outerHTML;
        }
        return Array.from(node.childNodes).map(child => walk(child, inLink)).join('');
      }
      return '';
    }
    return Array.from(div.childNodes).map(node => walk(node, false)).join('');
  };

  // One outline per action. Stroke matches the header icons.
  app.icon = function icon(name, size = 18) {
    const glyphs = {
      trash: '<path d="M3 6h18"></path><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6"></path><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path><path d="M10 11v6"></path><path d="M14 11v6"></path>',
      file: '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line><polyline points="10 9 9 9 8 9"></polyline>',
      sparkle: '<path d="M12 3l1.6 4.8L18.5 9.4 13.6 11 12 15.8 10.4 11 5.5 9.4l4.9-1.6L12 3z"></path><path d="M18 14l.6 1.8 1.8.6-1.8.6L18 19l-.6-1.8-1.8-.6 1.8-.6L18 14z"></path>'
    };
    const body = glyphs[name];
    if (!body) return '';
    const px = Number.isFinite(size) ? size : 18;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${body}</svg>`;
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
