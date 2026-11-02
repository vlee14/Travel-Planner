// Enough of the DOM for the HTML helpers to run under Node.
const VOID = new Set(['AREA', 'BASE', 'BR', 'COL', 'EMBED', 'HR', 'IMG', 'INPUT', 'LINK', 'META', 'SOURCE', 'TRACK', 'WBR']);

function escapeText(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function escapeAttr(value) {
  return escapeText(value).replace(/"/g, '&quot;');
}

function decodeEntities(value) {
  return value.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (all, body) => {
    const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: '\u00a0' };
    if (body[0] !== '#') return named[body.toLowerCase()] ?? all;
    const code = body[1] === 'x' || body[1] === 'X'
      ? parseInt(body.slice(2), 16)
      : parseInt(body.slice(1), 10);
    return Number.isFinite(code) ? String.fromCodePoint(code) : all;
  });
}

class TextNode {
  constructor(text) {
    this.nodeValue = String(text);
    this.parentNode = null;
    this.childNodes = [];
  }

  get nodeType() { return 3; }
  get textContent() { return this.nodeValue; }
  set textContent(value) { this.nodeValue = String(value); }
}

class Element {
  constructor(tagName) {
    this.tagName = tagName;
    this.childNodes = [];
    this.parentNode = null;
    this.attrs = {};
    this.style = {};
  }

  get nodeType() { return 1; }

  setAttribute(name, value) {
    this.attrs[String(name).toLowerCase()] = String(value);
  }

  getAttribute(name) {
    const key = String(name).toLowerCase();
    return Object.prototype.hasOwnProperty.call(this.attrs, key) ? this.attrs[key] : null;
  }

  hasAttribute(name) {
    return Object.prototype.hasOwnProperty.call(this.attrs, String(name).toLowerCase());
  }

  removeAttribute(name) {
    delete this.attrs[String(name).toLowerCase()];
  }

  get textContent() {
    return this.childNodes.map(node => node.textContent).join('');
  }

  set textContent(value) {
    this.childNodes = [];
    const text = value == null ? '' : String(value);
    if (!text) return;
    const node = new TextNode(text);
    node.parentNode = this;
    this.childNodes.push(node);
  }

  get innerHTML() {
    return this.childNodes.map(serializeNode).join('');
  }

  set innerHTML(value) {
    this.childNodes = parseFragment(value);
    this.childNodes.forEach(node => { node.parentNode = this; });
  }

  get outerHTML() {
    const name = this.tagName.toLowerCase();
    if (VOID.has(this.tagName)) return `<${name}>`;
    const attrs = Object.entries(this.attrs)
      .map(([key, value]) => ` ${key}="${escapeAttr(value)}"`)
      .join('');
    return `<${name}${attrs}>${this.innerHTML}</${name}>`;
  }

  querySelectorAll() { return []; }
  closest() { return null; }

  set href(value) { this.setAttribute('href', value); }
  get href() { return this.getAttribute('href') || ''; }
  set target(value) { this.setAttribute('target', value); }
  get target() { return this.getAttribute('target') || ''; }
  set rel(value) { this.setAttribute('rel', value); }
  get rel() { return this.getAttribute('rel') || ''; }
}

function serializeNode(node) {
  if (node.nodeType === 3) return escapeText(node.textContent);
  return node.outerHTML;
}

function parseAttrs(raw) {
  const attrs = [];
  const re = /([^\s=/]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
  let match;
  while ((match = re.exec(raw))) {
    attrs.push([match[1].toLowerCase(), match[2] ?? match[3] ?? match[4] ?? '']);
  }
  return attrs;
}

function parseFragment(html) {
  const root = { childNodes: [], tagName: '' };
  const stack = [root];
  const source = String(html ?? '');
  let i = 0;

  function append(node) {
    const parent = stack[stack.length - 1];
    node.parentNode = parent.tagName ? parent : null;
    parent.childNodes.push(node);
  }

  function appendText(text) {
    if (!text) return;
    append(new TextNode(decodeEntities(text)));
  }

  while (i < source.length) {
    const lt = source.indexOf('<', i);
    if (lt === -1) {
      appendText(source.slice(i));
      break;
    }
    appendText(source.slice(i, lt));
    if (source.startsWith('<!--', lt)) {
      const end = source.indexOf('-->', lt + 4);
      i = end === -1 ? source.length : end + 3;
      continue;
    }
    const tag = /^<\/?([a-zA-Z][\w:-]*)([^>]*)>/.exec(source.slice(lt));
    if (!tag) {
      appendText('<');
      i = lt + 1;
      continue;
    }
    const closing = source[lt + 1] === '/';
    const name = tag[1].toUpperCase();
    const attrText = tag[2] || '';
    i = lt + tag[0].length;
    if (closing) {
      for (let s = stack.length - 1; s > 0; s--) {
        if (stack[s].tagName === name) {
          stack.length = s;
          break;
        }
      }
      continue;
    }
    const el = new Element(name);
    parseAttrs(attrText).forEach(([key, value]) => el.setAttribute(key, decodeEntities(value)));
    append(el);
    const selfClosing = /\/\s*$/.test(attrText) || VOID.has(name);
    if (!selfClosing) stack.push(el);
  }
  return root.childNodes;
}

globalThis.Node = { ELEMENT_NODE: 1, TEXT_NODE: 3 };
globalThis.document = {
  createElement(tag) {
    return new Element(String(tag || 'div').toUpperCase());
  }
};
