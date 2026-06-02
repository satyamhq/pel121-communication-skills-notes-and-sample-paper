const fs = require('fs');

const gen1Path = 'd:\\pel121\\gen_pel121 (1).js';
const gen1 = fs.readFileSync(gen1Path, 'utf8');

// We want to extract the bodies of unit1() to unit6()
const units = [];
for (let i = 1; i <= 6; i++) {
  const regex = new RegExp(`function unit${i}\\(\\) \\{\\s*return \\s*\\[([\\s\\S]*?)\\];\\s*\\}`, 'm');
  const match = gen1.match(regex);
  if (match) {
    units.push({ id: `unit${i}`, body: match[1] });
  }
}

let htmlContent = '';

units.forEach(unit => {
  let body = unit.body;
  
  // Basic transformations
  // h1("...")
  body = body.replace(/h1\("(.*?)"\)/g, '<h1>$1</h1>');
  // h2("...")
  body = body.replace(/h2\("(.*?)"\)/g, '<h2>$1</h2>');
  // h3("...")
  body = body.replace(/h3\("(.*?)"\)/g, '<h3>$1</h3>');
  // para("...")
  body = body.replace(/para\("(.*?)"(?:,.*?)?\)/g, '<p>$1</p>');
  // noteBox("...")
  body = body.replace(/noteBox\("(.*?)"\)/g, '<div class="alert note">💡 $1</div>');
  // tipBox("...")
  body = body.replace(/tipBox\("(.*?)"\)/g, '<div class="alert tip">⚡ $1</div>');
  // errorBox("...")
  body = body.replace(/errorBox\("(.*?)"\)/g, '<div class="alert error">❌ $1</div>');
  // space(), divider(), pageBreak()
  body = body.replace(/space\(\)/g, '<br>');
  body = body.replace(/divider\(\)/g, '<hr>');
  body = body.replace(/pageBreak\(\)/g, '');
  
  // twoColTable([...], ... )
  // We need to parse arrays. Let's just do a simpler approach:
  // Instead of full parsing, replace twoColTable with a special syntax or just evaluate it.
  
  htmlContent += `<section id="${unit.id}">\n`;
  
  // A somewhat hacky but effective way to evaluate the docx structure into HTML
  // We will define the helper functions to return HTML strings.
  
  const env = {
    h1: (text) => `<h1>${text}</h1>`,
    h2: (text) => `<h2>${text}</h2>`,
    h3: (text) => `<h3>${text}</h3>`,
    para: (text, opts) => `<p>${text}</p>`,
    bold: (text) => `<strong>${text}</strong>`,
    bullet: (text) => `<li>${text}</li>`,
    numbered: (text) => `<li>${text}</li>`,
    space: () => `<div class="spacer"></div>`,
    divider: () => `<hr>`,
    pageBreak: () => ``,
    twoColTable: (rows) => {
      let html = '<table class="styled-table">';
      rows.forEach(r => {
        html += `<tr><td>${r[0]}</td><td>${r[1]}</td></tr>`;
      });
      html += '</table>';
      return html;
    },
    threeColTable: (headers, rows) => {
      let html = '<table class="styled-table"><thead><tr>';
      headers.forEach(h => html += `<th>${h}</th>`);
      html += '</tr></thead><tbody>';
      rows.forEach(r => {
        html += `<tr><td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td></tr>`;
      });
      html += '</tbody></table>';
      return html;
    },
    noteBox: (text) => `<div class="alert note"><strong>Note:</strong> ${text}</div>`,
    tipBox: (text) => `<div class="alert tip"><strong>Tip:</strong> ${text}</div>`,
    errorBox: (text) => `<div class="alert error"><strong>Error:</strong> ${text}</div>`
  };
  
  try {
    const fn = new Function(...Object.keys(env), `return [${unit.body}].join('');`);
    const generatedHtml = fn(...Object.values(env));
    htmlContent += generatedHtml;
  } catch(e) {
    console.error(`Error parsing ${unit.id}:`, e);
    // fallback to regex replaced
    htmlContent += body;
  }
  
  htmlContent += `</section>\n`;
});

fs.writeFileSync('d:\\pel121\\notes_data.js', 'const notesHTML = `' + htmlContent.replace(/`/g, '\\`') + '`;\n');
console.log('Notes extracted to notes_data.js');
