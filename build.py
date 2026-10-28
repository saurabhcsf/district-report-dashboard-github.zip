"""Builds dist/district-report-dashboard.html: one self-contained, offline file."""
import base64, pathlib
root = pathlib.Path(__file__).parent
deps = root.parent / 'deps'
tpl = (root/'src/template.html').read_text(encoding='utf-8')
fontdir = deps/'fontsource-mukta-5.0.17/package/files'
ranges = {
 'devanagari': 'U+0900-097F,U+1CD0-1CF9,U+200C-200D,U+20A8,U+20B9,U+25CC,U+A830-A839,U+A8E0-A8FF',
 'latin': 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0300-0301,U+0303-0304,U+0308-0309,U+0323,U+0329,U+2000-206F,U+2074,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD',
}
faces=[]
for sub, rng in ranges.items():
    for w in (400, 600, 700):
        b64 = base64.b64encode((fontdir/f'mukta-{sub}-{w}-normal.woff2').read_bytes()).decode()
        faces.append(f"@font-face{{font-family:'Mukta';font-style:normal;font-display:swap;font-weight:{w};src:url(data:font/woff2;base64,{b64}) format('woff2');unicode-range:{rng};}}")
xlsx = (deps/'xlsx-0.18.5/package/dist/xlsx.full.min.js').read_text(encoding='utf-8').replace('</script', '<\\/script')
app = (root/'src/core.js').read_text(encoding='utf-8') + '\n' + (root/'src/ui.js').read_text(encoding='utf-8')
out = tpl.replace('/*__FONTS__*/', '\n'.join(faces)).replace('/*__XLSX__*/', xlsx).replace('/*__APP__*/', app)

dst = root/'index.html'
dst.write_text(out, encoding='utf-8')
print(dst, round(dst.stat().st_size/1024), 'KB')
