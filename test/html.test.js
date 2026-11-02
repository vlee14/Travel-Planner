import assert from 'node:assert/strict';
import { describe, test } from 'node:test';
import { createApp } from './harness.js';

describe('html', () => {
  test('escapes text that would otherwise be markup', () => {
    const app = createApp();
    assert.equal(app.escapeHtml('<b>a & b</b>'), '&lt;b&gt;a &amp; b&lt;/b&gt;');
    assert.equal(app.escapeHtml(''), '');
  });

  test('allows only http and https urls', () => {
    const app = createApp();
    assert.equal(app.safeHttpUrl('https://example.com/a'), 'https://example.com/a');
    assert.equal(app.safeHttpUrl('http://ok.test/x'), 'http://ok.test/x');
    assert.equal(app.safeHttpUrl('www.example.com/path'), 'https://www.example.com/path');
    assert.equal(app.safeHttpUrl('javascript:alert(1)'), '');
    assert.equal(app.safeHttpUrl('ftp://files.test/a'), '');
    assert.equal(app.safeHttpUrl('   '), '');
  });

  test('turns plain urls into links and leaves the trailing punctuation outside', () => {
    const app = createApp();
    assert.equal(
      app.linkifyPlainText('See https://example.com/a.'),
      'See <a href="https://example.com/a" target="_blank" rel="noopener noreferrer">https://example.com/a</a>.'
    );
    assert.equal(
      app.linkifyPlainText('Visit www.example.com/path'),
      'Visit <a href="https://www.example.com/path" target="_blank" rel="noopener noreferrer">www.example.com/path</a>'
    );
  });

  test('keeps safe links and drops every other tag', () => {
    const app = createApp();
    assert.equal(
      app.linkifyAndSanitize('[Museum](https://example.com/a)'),
      '<a href="https://example.com/a" target="_blank" rel="noopener noreferrer">Museum</a>'
    );
    assert.equal(
      app.linkifyAndSanitize('[A & B](https://example.com/a)'),
      '<a href="https://example.com/a" target="_blank" rel="noopener noreferrer">A &amp; B</a>'
    );
    assert.equal(app.linkifyAndSanitize('[click](javascript:alert(1))'), '[click](javascript:alert(1))');
    assert.equal(app.linkifyAndSanitize('<script>alert(1)</script>'), 'alert(1)');
    assert.equal(app.linkifyAndSanitize('<img src="x" onerror="alert(1)">'), '');
    assert.equal(
      app.linkifyAndSanitize('<a href="https://example.com/a" onclick="alert(1)">Map</a>'),
      '<a href="https://example.com/a" target="_blank" rel="noopener noreferrer">Map</a>'
    );
    assert.equal(app.linkifyAndSanitize('<b>hi</b>'), 'hi');
  });

  test('round-trips an activity link through the editor', () => {
    const app = createApp();
    const editor = app.activityEditorText('<a href="https://example.com/a">Museum</a><br>Next');
    assert.equal(editor, '[Museum](https://example.com/a)\nNext');
    assert.equal(
      app.linkifyAndSanitize(editor),
      '<a href="https://example.com/a" target="_blank" rel="noopener noreferrer">Museum</a><br>Next'
    );
    assert.equal(app.activityEditorText('<p>Hello</p><div>There</div>'), 'Hello\nThere');
    assert.equal(app.activityEditorText('<a href="javascript:alert(1)">x</a>'), 'x');
  });

  test('reads visible text and flattens a sheet cell', () => {
    const app = createApp();
    assert.equal(app.getTextContent('<b>Kyoto</b>'), 'Kyoto');
    assert.equal(app.getTextContent('Temple &amp; garden'), 'Temple & garden');
    assert.equal(app.getTextContent('[Map](https://example.com/m)'), 'Map (https://example.com/m)');
    assert.equal(
      app.simplifyHtml('<div>Hi <a href="https://example.com">there</a><br>next</div>'),
      'Hi <a href="https://example.com">there</a>\nnext'
    );
  });
});
