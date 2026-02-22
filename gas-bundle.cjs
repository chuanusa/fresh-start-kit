const fs = require('fs');
const path = require('path');

async function bundle() {
    const distPath = path.join(__dirname, 'dist');
    const indexPath = path.join(distPath, 'index.html');

    if (!fs.existsSync(indexPath)) {
        console.error('Error: dist/index.html not found. Run npm run build first.');
        process.exit(1);
    }

    let html = fs.readFileSync(indexPath, 'utf8');

    // 1. Inline JS robustly
    // Match any script tag with an asset src
    html = html.replace(/<script\b[^>]*src="[^"]*\/assets\/([^"]+\.js)"[^>]*><\/script>/gi, (match, fileName) => {
        const assetPath = path.join(distPath, 'assets', fileName);
        if (fs.existsSync(assetPath)) {
            console.log(`Inlining JS: ${fileName}`);
            const content = fs.readFileSync(assetPath, 'utf8');
            // Remove source mapping URLs which can confuse GAS
            const cleanedContent = content.replace(/\/\/# sourceMappingURL=.*/g, '');
            return `<script>\n${cleanedContent}\n</script>`;
        }
        return match;
    });

    // 2. Inline CSS robustly
    html = html.replace(/<link\b[^>]*href="[^"]*\/assets\/([^"]+\.css)"[^>]*>/gi, (match, fileName) => {
        const assetPath = path.join(distPath, 'assets', fileName);
        if (fs.existsSync(assetPath)) {
            console.log(`Inlining CSS: ${fileName}`);
            const content = fs.readFileSync(assetPath, 'utf8');
            return `<style>\n${content}\n</style>`;
        }
        return match;
    });

    // 3. Specific Cleanups for GAS HtmlService
    // GAS rejects HTML if it has certain tags or attributes in the wrong place

    // Remove modulepreload links
    html = html.replace(/<link [^>]*rel="modulepreload"[^>]*>/gi, '');

    // Remove type="module" from any remaining scripts (just in case)
    html = html.replace(/<script\b([^>]*)type="module"([^>]*)>/gi, '<script$1$2>');

    // Remove crossorigin from any remaining tags
    html = html.replace(/<(script|link|img)\b([^>]*)crossorigin(="[^"]*")?([^>]*)>/gi, '<$1$2$4>');

    // 4. Structural Cleanups (Optional but safer)
    // GAS often prefers just the body content or very clean html/head/body
    html = html.replace(/<!doctype html>/gi, '');
    // Keep html/head/body but make them plain
    html = html.replace(/<html[^>]*>/gi, '<html>');
    html = html.replace(/<head[^>]*>/gi, '<head>');
    html = html.replace(/<body[^>]*>/gi, '<body>');

    // Save to AppMain.html
    fs.writeFileSync(path.join(__dirname, 'AppMain.html'), html);
    console.log('Successfully bundled dist/ into AppMain.html (Robust Mode)');
}

bundle().catch(err => {
    console.error(err);
    process.exit(1);
});
