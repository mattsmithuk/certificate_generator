// PhantomJS: generate PDF to stdout from HTML in stdin.
// Example: echo "<b>test</b>" | phantomjs makepdf.js > test.pdf && open test.pdf

var page = require('webpage').create(),
    fs = require('fs');

page.viewportSize = { width: 600, height: 600 };
page.paperSize = { format: 'Letter', orientation: 'portrait', margin: '1cm' };

page.content = fs.read('/dev/stdin');

window.setTimeout(function() {
    page.render('/dev/stdout', { format: 'pdf' });
    phantom.exit();
}, 1);