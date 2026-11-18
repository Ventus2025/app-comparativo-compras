// Genera appscript/*.html a partir de src/*.html.
// HtmlService elimina "comentarios" del JavaScript en línea sin entender template literals
// ni expresiones regulares, y corta cadenas que contienen // o /*. Para evitarlo, el script
// principal de cada página se publica codificado en base64url (sin / * ni comillas) y un cargador
// mínimo lo decodifica y lo inserta como <script> normal (mismo alcance global que el original).
// Uso: node build.js            → genera los archivos
//      node build.js --check    → solo verifica que appscript/ esté al día con src/ (sale con 1 si no)
// Luego: cd appscript && clasp push
const fs = require('fs'), path = require('path');
const SRC = path.join(__dirname, 'src'), OUT = path.join(__dirname, 'appscript');
const check = process.argv.includes('--check');

const loader = b64 => '<script>(function(){var b=' + JSON.stringify(b64) + ".split('-').join('+').split('_').join('/');" +
  "while(b.length%4)b+='=';var s=atob(b),u=new Uint8Array(s.length);for(var i=0;i<s.length;i++)u[i]=s.charCodeAt(i);" +
  "var el=document.createElement('script');el.textContent=new TextDecoder().decode(u);document.body.appendChild(el)})();</script>";

let stale = 0;
for (const f of fs.readdirSync(SRC).filter(f => f.endsWith('.html'))) {
  const html = fs.readFileSync(path.join(SRC, f), 'utf8');
  const m = html.match(/<script>([\s\S]*)<\/script>/);
  let out = html;
  if (m) {
    try { new Function(m[1]); } catch (e) { throw new Error(`${f}: error de sintaxis en el script: ${e.message}`); }
    out = html.slice(0, m.index) + loader(Buffer.from(m[1], 'utf8').toString('base64url')) + html.slice(m.index + m[0].length);
  }
  const dest = path.join(OUT, f), cur = fs.existsSync(dest) ? fs.readFileSync(dest, 'utf8') : null;
  if (check) {
    if (cur !== out) { stale++; console.log('DESACTUALIZADO', f); } else console.log('OK', f);
  } else {
    fs.writeFileSync(dest, out);
    console.log(cur === out ? 'sin cambios' : 'generado  ', path.relative(__dirname, dest), out.length, 'caracteres');
  }
}
if (stale) process.exit(1);
