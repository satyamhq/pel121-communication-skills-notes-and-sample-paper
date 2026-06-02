const fs = require('fs');

const gen1Path = 'd:\\pel121\\gen_pel121 (1).js';
const gen2Path = 'd:\\pel121\\generate_papers.js';

let gen1 = fs.readFileSync(gen1Path, 'utf8');
let gen2 = fs.readFileSync(gen2Path, 'utf8');

// Extract paper 1 questions
const p1Match = gen1.match(/const qs = \[([\s\S]*?)\];/);
let p1QuestionsText = p1Match ? p1Match[1] : '';

// The p1 questions are in format: ["1", "Question text", ["(a) option1", ...], "CO3,L3"]
// I'll extract these using regex

const p1Qs = [];
const p1Regex = /\["(\d+)",\s*"(.*?)",\s*\[(.*?)\](?:,\s*"(.*?)")?\]/g;
let match;
while ((match = p1Regex.exec(p1QuestionsText)) !== null) {
  const num = match[1];
  const text = match[2];
  const optionsRaw = match[3];
  const optsMatch = optionsRaw.match(/"(.*?)"/g);
  const options = optsMatch ? optsMatch.map(o => o.replace(/(^"|"$)/g, '')) : [];
  p1Qs.push({ q: parseInt(num), text, opts: options });
}

// Extract paper 2 and 3 questions
const p2qMatch = gen2.match(/const paper2Questions = (\[[\s\S]*?\]);/);
const p2aMatch = gen2.match(/const paper2Answers = (\[[\s\S]*?\]);/);
const p3qMatch = gen2.match(/const paper3Questions = (\[[\s\S]*?\]);/);
const p3aMatch = gen2.match(/const paper3Answers = (\[[\s\S]*?\]);/);

function evalCode(code) {
  // Simple evaluation since it's just JS objects
  return eval(code);
}

let paper2Questions = p2qMatch ? evalCode(p2qMatch[1]) : [];
let paper2Answers = p2aMatch ? evalCode(p2aMatch[1]) : [];
let paper3Questions = p3qMatch ? evalCode(p3qMatch[1]) : [];
let paper3Answers = p3aMatch ? evalCode(p3aMatch[1]) : [];

// Combine data
const data = {
  paper1: p1Qs, // Note: Paper 1 answers aren't explicitly listed as an answer key object in gen_pel121, but in the options themselves. Wait, let's see. 
  paper2: {
    questions: paper2Questions,
    answers: paper2Answers
  },
  paper3: {
    questions: paper3Questions,
    answers: paper3Answers
  }
};

fs.writeFileSync('d:\\pel121\\data.js', 'const courseData = ' + JSON.stringify(data, null, 2) + ';\n');
console.log('Data extracted to data.js');
