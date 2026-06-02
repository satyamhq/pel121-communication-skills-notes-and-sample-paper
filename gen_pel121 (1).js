const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak, TableOfContents
} = require('docx');
const fs = require('fs');

// ─── Helpers ─────────────────────────────────────────────────────────────────

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
  right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
};

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function para(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...opts,
    children: [new TextRun({ text, ...opts.run })],
  });
}
function bold(text) {
  return new Paragraph({ spacing: { after: 100 }, children: [new TextRun({ text, bold: true })] });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { after: 80 },
    children: [new TextRun(text)],
  });
}
function numbered(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "numbers", level },
    spacing: { after: 80 },
    children: [new TextRun(text)],
  });
}
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}
function space() {
  return new Paragraph({ children: [new TextRun("")], spacing: { after: 80 } });
}
function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E4057", space: 1 } },
    children: [new TextRun("")],
    spacing: { after: 160 },
  });
}

function twoColRow(c1, c2, w1 = 4680, w2 = 4680) {
  return new TableRow({
    children: [
      new TableCell({
        borders, width: { size: w1, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun(c1)] })],
      }),
      new TableCell({
        borders, width: { size: w2, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun(c2)] })],
      }),
    ],
  });
}
function twoColTable(rows, w1 = 4680, w2 = 4680) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [w1, w2],
    rows: rows.map(([a, b]) => twoColRow(a, b, w1, w2)),
  });
}
function threeColRow(c1, c2, c3, header = false) {
  const makeCell = (text, w) => new TableCell({
    borders,
    width: { size: w, type: WidthType.DXA },
    shading: header ? { fill: "2E4057", type: ShadingType.CLEAR } : undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, bold: header, color: header ? "FFFFFF" : "000000" })] })],
  });
  return new TableRow({ children: [makeCell(c1, 2340), makeCell(c2, 3510), makeCell(c3, 3510)] });
}
function threeColTable(headers, rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2340, 3510, 3510],
    rows: [threeColRow(...headers, true), ...rows.map(r => threeColRow(...r))],
  });
}
function noteBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "E8A838" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "E8A838" }, left: { style: BorderStyle.THICK, size: 12, color: "E8A838" }, right: { style: BorderStyle.NONE } },
        shading: { fill: "FFF8E7", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: "💡 " + text, italics: true })] })],
      })],
    })],
  });
}
function tipBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "27AE60" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "27AE60" }, left: { style: BorderStyle.THICK, size: 12, color: "27AE60" }, right: { style: BorderStyle.NONE } },
        shading: { fill: "EAFAF1", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: "SHORTCUT: " + text, bold: true })] })],
      })],
    })],
  });
}
function errorBox(text) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 4, color: "E74C3C" }, bottom: { style: BorderStyle.SINGLE, size: 4, color: "E74C3C" }, left: { style: BorderStyle.THICK, size: 12, color: "E74C3C" }, right: { style: BorderStyle.NONE } },
        shading: { fill: "FDECEA", type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 160, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: "COMMON ERROR: " + text, bold: true, color: "C0392B" })] })],
      })],
    })],
  });
}

// MCQ question paragraph (bold Q number + options)
function mcqQ(num, questionText, opts, co) {
  const children = [];
  return [
    new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: `Q(${num})  ${questionText}`, bold: true })] }),
    ...opts.map(o => new Paragraph({ spacing: { after: 40 }, indent: { left: 360 }, children: [new TextRun(o)] })),
    new Paragraph({ spacing: { after: 140 }, alignment: AlignmentType.RIGHT, children: [new TextRun({ text: co, italics: true, color: "555555", size: 18 })] }),
  ];
}

// ─── NOTES CONTENT ────────────────────────────────────────────────────────────

// UNIT I – Parts of Speech
function unit1() {
  return [
    h1("UNIT I: Parts of Speech"),
    divider(),
    h2("1.1 Types of Nouns"),
    para("A noun is a word used to name a person, place, thing, idea, or quality."),
    threeColTable(
      ["Type", "Definition", "Examples"],
      [
        ["Proper Noun", "Names a specific person/place/thing (always capitalised)", "Delhi, Ravi, Eiffel Tower"],
        ["Common Noun", "Names a general person, place, or thing", "city, boy, book, river"],
        ["Collective Noun", "Names a group of persons/things", "team, fleet, jury, flock, army"],
        ["Abstract Noun", "Names an idea, quality, or state", "love, courage, honesty, freedom"],
        ["Countable Noun", "Can be counted; has singular & plural", "apple / apples, child / children"],
        ["Uncountable Noun", "Cannot be counted; no plural form", "water, advice, furniture, information"],
        ["Compound Noun", "Two or more words forming one noun", "toothpaste, bedroom, swimming pool"],
        ["Concrete Noun", "Can be perceived by senses", "table, music, rose, thunder"],
      ]
    ),
    space(),
    h2("1.2 Apostrophe 's with Nouns"),
    para("Use possessive 's to show ownership."),
    twoColTable([
      ["Rule", "Example"],
      ["Singular noun: add 's", "the girl's bag, Ravi's pen"],
      ["Plural noun ending in s: add ' only", "the teachers' room, the boys' school"],
      ["Plural noun NOT ending in s: add 's", "the children's books, the men's club"],
      ["Two owners (joint possession): 's after last noun", "Ram and Shyam's house"],
      ["Two owners (separate possession): 's after each", "Ram's and Shyam's cars"],
    ], 3120, 6240),
    space(),
    errorBox("WRONG: The students's papers. CORRECT: The students' papers."),
    space(),
    h2("1.3 Order of Adjectives"),
    para("When multiple adjectives precede a noun, they follow a fixed royal order:"),
    noteBox("DOSASMP → Determiner – Opinion – Size – Age – Shape – Material – Purpose/Qualifier + NOUN"),
    space(),
    twoColTable([
      ["Position", "Category / Examples"],
      ["1. Determiner", "a, an, the, my, this, three"],
      ["2. Opinion", "beautiful, wonderful, ugly, delicious"],
      ["3. Size", "large, small, tall, tiny, huge"],
      ["4. Age / Temperature", "old, young, new, hot, cold"],
      ["5. Shape", "round, square, flat, oval"],
      ["6. Colour", "red, blue, golden, dark"],
      ["7. Origin / Nationality", "Indian, French, American"],
      ["8. Material", "wooden, cotton, silver, plastic"],
      ["9. Purpose / Qualifier + Noun", "sleeping (bag), writing (desk)"],
    ], 2340, 7020),
    space(),
    tipBox("'Every intelligent Indian teacher' = Order: Determiner + Opinion + Origin + Noun. (Q.42 in sample paper)"),
    space(),
    h2("1.4 Degrees of Adjectives"),
    threeColTable(
      ["Adjective", "Comparative", "Superlative"],
      [
        ["tall", "taller", "tallest"],
        ["intelligent", "more intelligent", "most intelligent"],
        ["good", "better", "best"],
        ["bad", "worse", "worst"],
        ["far", "farther / further", "farthest / furthest"],
        ["little", "less", "least"],
        ["many / much", "more", "most"],
      ]
    ),
    space(),
    errorBox("WRONG: She is more taller. CORRECT: She is taller. (Never use 'more' with -er forms.)"),
    space(),
    h2("1.5 Prepositions"),
    para("Prepositions show the relationship between a noun/pronoun and other words."),
    twoColTable([
      ["Type", "Examples & Usage"],
      ["Place / Position", "in, on, at, under, over, between, among, beside, behind, in front of"],
      ["Time", "in (months/years), on (days/dates), at (specific time)"],
      ["Direction / Movement", "to, towards, into, onto, from, across, through, along"],
      ["Manner / Means", "by, with, on (foot), in (a car)"],
      ["Cause / Purpose", "for, because of, due to, owing to"],
    ], 2520, 6840),
    space(),
    noteBox("'The professor explained the concept with great clarity.' → 'with' = preposition of manner. (Q.18)"),
    space(),
    h2("UNIT I – Quick Revision Sheet"),
    divider(),
    twoColTable([
      ["Topic", "Key Points"],
      ["Proper Noun", "Specific name; always capital – Delhi, Ravi"],
      ["Collective Noun", "Group – team, jury, fleet, flock"],
      ["Abstract Noun", "Qualities/ideas – love, honesty, freedom"],
      ["Possessive 's", "Singular: girl's | Plural-s: teachers' | Plural non-s: children's"],
      ["Adj. Order (Royal)", "DOSASMP → Opinion → Size → Age → Shape → Colour → Origin → Material → Purpose"],
      ["Degrees", "tall/taller/tallest; good/better/best; bad/worse/worst"],
      ["Key Prepositions", "in/on/at (time & place); with (manner); for (purpose); by (agent in passive)"],
      ["Common Error", "Never: more taller / most tallest. Never: The students's."],
    ], 2520, 6840),
    pageBreak(),
  ];
}

// UNIT II – Building Sentences
function unit2() {
  return [
    h1("UNIT II: Building Sentences"),
    divider(),
    h2("2.1 Conjunctions & Connectors"),
    para("Conjunctions join words, phrases, or clauses. Choose the right one for the right logical relationship."),
    threeColTable(
      ["Connector", "Function / Meaning", "Example"],
      [
        ["and", "Addition", "She sings and dances."],
        ["but", "Contrast / Exception", "He tried but failed."],
        ["both … and", "Equal addition (two items)", "Both Ram and Sita attended."],
        ["either … or", "Choice / Alternative", "Either you or he is wrong."],
        ["neither … nor", "Negative alternative", "Neither she nor he came."],
        ["because / as / since / for", "Reason / Cause", "I left early because it rained."],
        ["in case", "Precaution", "Carry an umbrella in case it rains."],
        ["so / therefore", "Result", "It rained, so we stayed inside."],
        ["so that / in order that", "Purpose", "She studied hard so that she could pass."],
        ["although / though / even though", "Concession (unexpected contrast)", "Although it rained, they played."],
        ["while / whereas", "Simultaneous action OR contrast", "While I cook, you set the table. / She likes tea, whereas he prefers coffee."],
        ["however / nevertheless", "Contrast (formal)", "It was hard; however, she succeeded."],
        ["despite / in spite of", "Concession (followed by noun/V-ing)", "Despite the rain, they played."],
      ]
    ),
    space(),
    h2("2.2 Relative Clauses"),
    para("Relative clauses modify nouns. Introduced by relative pronouns: who, whom, whose, which, that."),
    twoColTable([
      ["Pronoun", "Use & Example"],
      ["who", "Subject of verb → 'people who we went out with' (Q.25: 'people whom we went out with' is correct when object)"],
      ["whom", "Object of verb → 'The people whom we went out with are nice.' (formal/correct)"],
      ["whose", "Possession → 'The man whose car broke down called us.'"],
      ["which", "Things → 'The book which I borrowed is good.'"],
      ["that", "Things or persons (defining clauses) → 'The phone that you bought is new.'"],
    ], 2340, 7020),
    space(),
    tipBox("Q.25: 'The people whom we went out with are very nice.' is grammatically correct. 'who' is used when it is the subject of the verb."),
    space(),
    errorBox("WRONG: The people which we met. CORRECT: The people whom/who we met. (Use 'which' only for things.)"),
    space(),
    h2("2.3 Sentence Types"),
    twoColTable([
      ["Type", "Definition & Example"],
      ["Simple Sentence", "One independent clause: 'She reads books.'"],
      ["Compound Sentence", "Two independent clauses joined by conjunction: 'She reads, but he watches TV.'"],
      ["Complex Sentence", "One independent + one/more dependent clauses: 'Although it rained, she went out.'"],
      ["Compound-Complex", "Two independent + one dependent: 'She studies hard, so she passes, although topics are tough.'"],
    ], 2520, 6840),
    space(),
    h2("UNIT II – Quick Revision Sheet"),
    divider(),
    twoColTable([
      ["Connector", "Use"],
      ["because / as / since / for", "Reason"],
      ["so / therefore", "Result"],
      ["so that / in order to", "Purpose"],
      ["although / despite / however", "Contrast/Concession"],
      ["while / whereas", "Contrast or simultaneous"],
      ["either…or", "Choice; neither…nor for negative"],
      ["who/whom", "People (who = subject; whom = object)"],
      ["which", "Things only"],
      ["whose", "Possession"],
    ], 2520, 6840),
    pageBreak(),
  ];
}

// UNIT III – Tenses
function unit3() {
  return [
    h1("UNIT III: Tenses"),
    divider(),
    h2("3.1 Present Tenses"),
    threeColTable(
      ["Tense", "Formula", "Usage & Signal Words"],
      [
        ["Present Simple", "Subject + V1 (s/es for 3rd person singular)", "Habits, routines, facts. Signal: always, usually, every day, never. e.g., She reads books every day."],
        ["Present Continuous", "Subject + am/is/are + V-ing", "Actions happening now, temporary actions, fixed future plans. Signal: now, at the moment, currently. e.g., They are building a new school."],
        ["Present Perfect", "Subject + have/has + V3", "Past actions with present relevance, experiences, recent events. Signal: just, already, yet, ever, never, recently, since, for. e.g., We have learnt our lessons. The baby has eaten all the chips."],
        ["Present Perfect Continuous", "Subject + have/has + been + V-ing", "Actions that started in past and continue now. Signal: for (duration), since (starting point). e.g., I have been living here for a month."],
      ]
    ),
    space(),
    h2("3.2 Past Tenses"),
    threeColTable(
      ["Tense", "Formula", "Usage & Signal Words"],
      [
        ["Past Simple", "Subject + V2", "Completed actions at a specific past time. Signal: yesterday, ago, last week, in 2005. e.g., Charles Darwin proposed the theory of evolution in 1859."],
        ["Past Continuous", "Subject + was/were + V-ing", "Action in progress at a past time; interrupted action. Signal: while, when, as. e.g., He was working when the crime occurred."],
        ["Past Perfect", "Subject + had + V3", "Action completed before another past action. Signal: before, after, already, by the time, when. e.g., The team had won the match before the storm started."],
        ["Past Perfect Continuous", "Subject + had + been + V-ing", "Ongoing action in past before another past event. Signal: for, since (in past context). e.g., Mr Arnold had been writing his book for four months."],
      ]
    ),
    space(),
    h2("3.3 Future Tenses"),
    threeColTable(
      ["Tense", "Formula", "Usage"],
      [
        ["Future Simple", "Subject + will + V1", "Predictions, promises, spontaneous decisions. e.g., She will call you."],
        ["Future Continuous", "Subject + will + be + V-ing", "Ongoing actions at a future moment. e.g., They will be sleeping at 10 pm."],
        ["Future Perfect", "Subject + will + have + V3", "Actions completed before a future point. e.g., She will have finished by Monday."],
        ["be going to + V1", "Subject + am/is/are + going to + V1", "Planned intentions or evidence-based predictions. e.g., Look at those clouds – it is going to rain."],
      ]
    ),
    space(),
    h2("3.4 Question Tags"),
    para("A question tag is a short question added at the end of a statement to confirm or check information."),
    twoColTable([
      ["Rule", "Example"],
      ["Positive statement → Negative tag", "She is coming, isn't she?"],
      ["Negative statement → Positive tag", "He doesn't know, does he?"],
      ["Auxiliary from statement used in tag", "They can swim, can't they?"],
      ["'I am' → tag is 'aren't I?'", "I am right, aren't I?"],
      ["Imperative → 'will you?' or 'won't you?'", "Close the door, will you?"],
    ], 3600, 5760),
    space(),
    h2("3.5 Negative & Imperative Forms"),
    twoColTable([
      ["Form", "Example"],
      ["Present Simple Negative", "She does not (doesn't) play. / He did not (didn't) come. (Q.17: 'Mary doesn't mind the children to play' – error is 'to play'; should be 'playing')"],
      ["Imperative (command)", "Open your books. / Please sit down."],
      ["Imperative Negative", "Don't open the door. / Never be late."],
    ], 2520, 6840),
    space(),
    tipBox("Identify tense signal words first: 'just/already/yet/ever' → Present Perfect. 'was/were + Ving + when/while' → Past Continuous. 'had + V3 + before/by the time' → Past Perfect."),
    errorBox("WRONG: I have seen him yesterday. CORRECT: I saw him yesterday. (Past simple goes with specific past time.)"),
    space(),
    h2("UNIT III – Quick Revision Sheet"),
    divider(),
    threeColTable(
      ["Tense", "Key Signal Words", "Formula"],
      [
        ["Present Simple", "always, usually, every, never, often, rarely", "S + V1(s/es)"],
        ["Present Continuous", "now, currently, at the moment, still", "S + am/is/are + Ving"],
        ["Present Perfect", "just, already, yet, ever, never, since, for, recently", "S + have/has + V3"],
        ["Present Perf. Cont.", "for + duration (ongoing)", "S + have/has + been + Ving"],
        ["Past Simple", "yesterday, last, ago, in [year]", "S + V2"],
        ["Past Continuous", "while, when, at that time", "S + was/were + Ving"],
        ["Past Perfect", "before, after, by the time, already (in past)", "S + had + V3"],
        ["Past Perf. Cont.", "for/since (before a past moment)", "S + had been + Ving"],
        ["Future Simple", "tomorrow, soon, I think, probably", "S + will + V1"],
      ]
    ),
    pageBreak(),
  ];
}

// UNIT IV – Articles & Indefinites
function unit4() {
  return [
    h1("UNIT IV: Articles and Indefinites"),
    divider(),
    h2("4.1 Articles Overview"),
    threeColTable(
      ["Article", "Use", "Examples"],
      [
        ["a", "Before singular countable noun starting with consonant sound; first mention; one of many", "a book, a university (y-sound), a European"],
        ["an", "Before singular countable noun starting with vowel sound", "an apple, an hour (h is silent), an honest man, an MBA"],
        ["the", "Specific / previously mentioned; unique things; superlatives; with ordinals; geographical features (rivers, mountain ranges, oceans, deserts)", "the sun, the Eiffel Tower, the best student, the first time, the Nile, the Alps"],
        ["No article (∅)", "Plural/uncountable nouns in general sense; proper nouns (most); meals, languages, sports, academic subjects (general)", "∅ Dogs are loyal. ∅ Kashmir is Switzerland of India. ∅ She plays cricket. ∅ I study science."],
      ]
    ),
    space(),
    noteBox("Q.51: '…Kashmir is …. Switzerland of India.' → Answer: no article – no article (a/an/the not needed for Kashmir; 'Switzerland of India' is a metaphor here so no article before 'Switzerland' either). The correct answer is 'no article – no article.'"),
    space(),
    h2("4.2 Definite Article with Special Objects"),
    twoColTable([
      ["Rule", "Example"],
      ["Unique objects / celestial bodies", "the sun, the moon, the earth, the sky"],
      ["Rivers, oceans, seas, mountain ranges, island groups", "the Ganges, the Pacific, the Himalayas, the Andamans"],
      ["With superlatives", "the tallest building, the most intelligent student"],
      ["With ordinals (first, second...)", "the first chapter, the second floor"],
      ["Musical instruments", "She plays the guitar."],
      ["Titles without name", "the President, the Principal (but: President Biden, Principal Sharma – no 'the')"],
    ], 3600, 5760),
    space(),
    h2("4.3 Quantifiers: much, many, (a) few, (a) little, more, every, each, one, another, other, others"),
    threeColTable(
      ["Quantifier", "Used with", "Example"],
      [
        ["much", "Uncountable nouns (negative/question mostly)", "I don't have much time."],
        ["many", "Countable nouns", "How many students are here?"],
        ["a few / few", "Countable – 'a few' = some; 'few' = not many (negative)", "I have a few friends. / Unfortunately I have little talent. (Q.9 → 'little' for uncountable noun 'talent')"],
        ["a little / little", "Uncountable – 'a little' = some; 'little' = not much (negative)", "There is a little water left."],
        ["every / each", "Singular countable (every = all in group; each = one by one)", "Every student must attend. Each child gets a gift."],
        ["another", "Additional/different singular countable", "Can I have another cup?"],
        ["other / others", "'other' + noun; 'others' = pronoun (alone)", "Other students passed. Others failed."],
        ["more", "Comparative quantity (countable & uncountable)", "I need more time / more books."],
      ]
    ),
    space(),
    errorBox("WRONG: I have a little friends. CORRECT: I have a few friends. ('few' for countable, 'little' for uncountable)"),
    space(),
    h2("UNIT IV – Quick Revision Sheet"),
    divider(),
    twoColTable([
      ["Rule / Item", "Quick Tip"],
      ["a vs an", "'a' before consonant sound, 'an' before vowel sound – judge by SOUND not spelling"],
      ["the – unique/specific", "the sun, the moon, the Nile, the Himalayas, the best, the first"],
      ["No article", "General plural (Dogs are…), proper names (Delhi, Ravi), meals (at breakfast), sports, languages"],
      ["much / many", "much = uncountable; many = countable"],
      ["few / little", "a few = some (count); few = hardly any (count); a little = some (uncount); little = hardly any (uncount)"],
      ["every vs each", "every = whole group together; each = one by one"],
      ["another vs other", "another = one more; other + noun; others = alone as pronoun"],
      ["School/uniform article", "'a' uniform for 'the' students → the–the (Q.21 answer: a-the)"],
    ], 3600, 5760),
    pageBreak(),
  ];
}

// UNIT V – Active & Passive Voice
function unit5() {
  return [
    h1("UNIT V: Active and Passive Voice"),
    divider(),
    h2("5.1 Rules of Passive Voice"),
    para("In active voice, the subject performs the action. In passive voice, the subject receives the action."),
    noteBox("Active: Subject + Verb + Object  →  Passive: Object + be (appropriate form) + V3 + by + Subject"),
    space(),
    threeColTable(
      ["Tense (Active)", "Active Formula", "Passive Formula"],
      [
        ["Present Simple", "S + V1(s/es) + O", "O + is/am/are + V3 + by S"],
        ["Present Continuous", "S + is/am/are + Ving + O", "O + is/am/are + being + V3 + by S"],
        ["Present Perfect", "S + has/have + V3 + O", "O + has/have + been + V3 + by S"],
        ["Past Simple", "S + V2 + O", "O + was/were + V3 + by S"],
        ["Past Continuous", "S + was/were + Ving + O", "O + was/were + being + V3 + by S"],
        ["Past Perfect", "S + had + V3 + O", "O + had + been + V3 + by S"],
        ["Future Simple", "S + will + V1 + O", "O + will + be + V3 + by S"],
        ["Modal Verbs", "S + modal + V1 + O", "O + modal + be + V3 + by S"],
      ]
    ),
    space(),
    h2("5.2 Worked Examples"),
    twoColTable([
      ["Active", "Passive"],
      ["The children ate the cake. (Past Simple)", "The cake was eaten by the children. (Q.12 – answer: b)"],
      ["He mailed his application for a new job.", "The application for a new job was mailed by him. (Q.15 – answer: b)"],
      ["The match had been won by the team before the storm started.", "Active: The team had won the match before the storm started. (Q.20 – answer: c)"],
      ["The manager called off the meeting.", "The meeting was called off by the manager. (Q.29 – answer: b)"],
      ["Someone stole my wallet.", "My wallet was stolen by someone. (Q.37 – answer: a)"],
      ["They asked him to leave the room.", "He was asked to leave the room. (Q.57 – answer: a)"],
    ], 4680, 4680),
    space(),
    h2("5.3 Special Cases"),
    twoColTable([
      ["Case", "Rule & Example"],
      ["Phrasal verb passive", "Keep phrasal verb intact: 'call off' → 'was called off'. Do NOT separate. (Q.29)"],
      ["Impersonal passive", "When agent unknown: 'The window had been broken by someone.' (Q.43)"],
      ["Sentences with two objects", "Either object can become subject: 'He gave me a gift' → 'I was given a gift' OR 'A gift was given to me'"],
      ["'By' phrase omission", "Omit 'by + agent' when unknown, obvious, or unimportant"],
    ], 2520, 6840),
    space(),
    tipBox("Steps: (1) Find object → make it subject. (2) Choose correct 'be' form for the tense. (3) Use V3. (4) Add 'by' + original subject if needed."),
    errorBox("WRONG: The cake was ate by them. CORRECT: The cake was eaten by them. (Always use V3, not V2 in passive.)"),
    space(),
    h2("UNIT V – Quick Revision Sheet"),
    divider(),
    threeColTable(
      ["Tense", "Active Example", "Passive Example"],
      [
        ["Present Simple", "She writes a letter.", "A letter is written by her."],
        ["Past Simple", "He bought the car.", "The car was bought by him."],
        ["Future Simple", "They will finish the project.", "The project will be finished by them."],
        ["Present Perfect", "She has eaten the food.", "The food has been eaten by her."],
        ["Past Perfect", "He had locked the door.", "The door had been locked by him."],
        ["Modal", "You should complete the task.", "The task should be completed by you."],
      ]
    ),
    pageBreak(),
  ];
}

// UNIT VI – Phrasal Verbs
function unit6() {
  return [
    h1("UNIT VI: Phrasal Verbs"),
    divider(),
    h2("6.1 What is a Phrasal Verb?"),
    para("A phrasal verb = verb + particle (preposition or adverb). The combination has a distinct meaning different from the individual words."),
    space(),
    h2("6.2 Common Phrasal Verbs (Exam-Focused)"),
    threeColTable(
      ["Phrasal Verb", "Meaning", "Example"],
      [
        ["look after", "take care of", "Our babysitter looks after our kids. (Q.7 – answer: a)"],
        ["look into", "investigate", "The police are looking into the case."],
        ["look up", "search for information", "Look up the word in the dictionary."],
        ["look out", "be careful / watch out", "Look out! The floor is wet! (Q.28 – answer: a)"],
        ["put up", "build / erect / tolerate / stay temporarily", "'They put up a tent near the river.' = To build. (Q.8 – answer: b)"],
        ["put off", "postpone / delay", "Don't put off what you can do today."],
        ["put on", "wear / start wearing", "Put on your jacket."],
        ["get along (with)", "have a good relationship", "I get along really well with my colleagues. (Q.10 – answer: a)"],
        ["find out", "discover / learn information", "I found out the truth."],
        ["break up", "end a relationship / dissolve", "They broke up last year."],
        ["call off", "cancel", "The manager called off the meeting. (Q.24 – answer: c in context)"],
        ["call back", "return a phone call", "I'll call you back later."],
        ["call on", "visit / ask someone to speak", "She called on her old friend."],
        ["call up", "phone someone / summon", "She'll call up her friends tonight."],
        ["give away", "donate / reveal", "She gave them away to charity. (Q.11 – answer: b; pronoun goes between verb and particle for 'give away')"],
        ["set off / set out", "start a journey", "They set off/out early to avoid traffic. (Q.36)"],
        ["make up", "reconcile after a quarrel / invent", "We had a quarrel but could make up. (Q.33 – answer: b)"],
        ["fall out (with)", "quarrel / have a disagreement", "I don't want to fall out with you. (Q.39 – answer: a)"],
        ["get round to", "find time to do something", "I just didn't have time to get round to finishing. (Q.19 – answer: a)"],
        ["make out", "understand / manage / perceive", "I can't make out what he is saying."],
        ["take off", "remove / leave ground (plane) / become successful", "The plane took off at 9 am."],
        ["turn up", "arrive unexpectedly / increase volume", "He turned up an hour late."],
      ]
    ),
    space(),
    h2("6.3 Separable vs Inseparable Phrasal Verbs"),
    twoColTable([
      ["Type", "Rule & Example"],
      ["Separable", "Object (noun or pronoun) can go between verb and particle. With pronoun, MUST be separated: 'She gave them away.' NOT 'She gave away them.' (Q.11)"],
      ["Inseparable", "Object always AFTER the particle: 'She looked after the children.' NOT 'She looked the children after.'"],
    ], 2340, 7020),
    space(),
    tipBox("If the object is a pronoun (him/her/them/it), it MUST go between the verb and particle for separable phrasal verbs."),
    space(),
    h2("6.4 Words Often Confused"),
    threeColTable(
      ["Word Pair", "Word 1", "Word 2"],
      [
        ["affect / effect", "affect (verb) – to influence: 'Pollution affects health.'", "effect (noun) – result: 'The effect of pollution is harmful.'"],
        ["lie / lay", "lie – recline (intransitive): 'I lie down.'", "lay – put/place (transitive): 'Lay the book on the table.'"],
        ["rise / raise", "rise – go up (intransitive): 'The sun rises.'", "raise – lift (transitive): 'Raise your hand.'"],
        ["fewer / less", "fewer – countable: 'fewer students'", "less – uncountable: 'less water'"],
        ["among / between", "among – 3 or more: 'among friends'", "between – 2: 'between you and me'"],
        ["principal / principle", "principal – head/main: 'the school principal'", "principle – rule/belief: 'a matter of principle'"],
        ["complement / compliment", "complement – complete/match: 'The sauce complements the dish.'", "compliment – praise: 'She gave him a compliment.'"],
      ]
    ),
    space(),
    h2("UNIT VI – Quick Revision Sheet"),
    divider(),
    twoColTable([
      ["Phrasal Verb", "Meaning (Memory Tip)"],
      ["look after", "care (after = behind them watching)"],
      ["look out / look out for", "beware / spot"],
      ["put up", "erect / tolerate ('put up with' = tolerate)"],
      ["put off", "postpone (push it off)"],
      ["get along with", "live/work in harmony"],
      ["call off", "cancel (call it off the schedule)"],
      ["set off / set out", "begin a journey"],
      ["give away", "donate; pronoun goes in middle"],
      ["fall out with", "quarrel with"],
      ["make up", "reconcile; also invent a story"],
      ["get round to", "eventually do (circular path)"],
      ["find out", "discover truth / information"],
    ], 2520, 6840),
    pageBreak(),
  ];
}

// ─── SAMPLE PAPER 1 ───────────────────────────────────────────────────────────

function samplePaper1() {
  const qs = [
    // Q1-10
    ["1", "Choose the correct option. She _____ to music while she _____ dinner yesterday.", ["(a) listened / was cooking", "(b) was listening / cooked", "(c) has listened / cooked", "(d) listens / cooks"], "CO3,L3"],
    ["2", "Which of the following nouns is ABSTRACT?", ["(a) Mountain", "(b) River", "(c) Justice", "(d) Telephone"], "CO1,L2"],
    ["3", "Choose the correct article: I need _____ advice about _____ career.", ["(a) an / a", "(b) the / the", "(c) a / a", "(d) no article / a"], "CO1,L3"],
    ["4", "Identify the correct passive voice of: 'The teacher teaches the students every day.'", ["(a) The students were taught by the teacher every day.", "(b) The students had been taught by the teacher every day.", "(c) The students are taught by the teacher every day.", "(d) The students will be taught by the teacher every day."], "CO1,L2"],
    ["5", "Choose the correct phrasal verb: He _____ from the university after failing his exams.", ["(a) dropped in", "(b) dropped off", "(c) dropped out", "(d) dropped by"], "CO2,L3"],
    ["6", "Identify the correct order of adjectives: She wore a(n) _____ dress.", ["(a) beautiful small red silk", "(b) small beautiful silk red", "(c) red silk beautiful small", "(d) silk small red beautiful"], "CO1,L1"],
    ["7", "Choose the correct sentence using the conjunction 'despite':", ["(a) Despite of working hard, she failed.", "(b) Despite working hard, she failed.", "(c) Despite she worked hard, she failed.", "(d) Despite to work hard, she failed."], "CO3,L3"],
    ["8", "Which sentence is in the Present Perfect Continuous tense?", ["(a) She has written three novels.", "(b) She was writing a novel for two hours.", "(c) She has been writing the novel for three hours.", "(d) She wrote a novel last year."], "CO1,L2"],
    ["9", "Select the correct option: 'We _____ dinner when the guests arrived.'", ["(a) had been having", "(b) were having", "(c) have been having", "(d) have had"], "CO3,L3"],
    ["10", "Which of the following uses 'much' correctly?", ["(a) There are much students in the hall.", "(b) How much books do you have?", "(c) There isn't much traffic today.", "(d) Much people attended the event."], "CO1,L5"],
    // Q11-20
    ["11", "Choose the correct possessive form: The report prepared by the _____ committee was submitted.", ["(a) teachers's", "(b) teacher's", "(c) teachers'", "(d) teachers"], "CO1,L2"],
    ["12", "Choose the correct relative pronoun: The scientist _____ discovered penicillin was Alexander Fleming.", ["(a) which", "(b) whom", "(c) who", "(d) whose"], "CO1,L3"],
    ["13", "Identify the sentence with the correct use of 'few' and 'little':", ["(a) I have a little friends here.", "(b) Only a few information is available.", "(c) There is a little hope left, but only a few days remain.", "(d) Few water remains in the bottle."], "CO1,L5"],
    ["14", "Choose the correct passive: 'They have cancelled the concert.'", ["(a) The concert has been cancelled.", "(b) The concert was cancelled.", "(c) The concert had been cancelled.", "(d) The concert is cancelled."], "CO1,L2"],
    ["15", "Which sentence uses a phrasal verb INCORRECTLY?", ["(a) Please look into the matter.", "(b) She looked after her sick mother.", "(c) They called the meeting off.", "(d) He gave away them his old books."], "CO3,L3"],
    ["16", "Choose the correct option: 'He _____ used to live in Jaipur, but now he lives in Mumbai.'", ["(a) has", "(b) had", "(c) would", "(d) Sentence is correct as written (no word needed)"], "CO2,L3"],
    ["17", "Identify the correct comparative form: He is _____ of the two candidates.", ["(a) most experienced", "(b) more experienced", "(c) much more experienced than all", "(d) the more experienced"], "CO1,L5"],
    ["18", "Which preposition correctly completes: 'She succeeded _____ achieving her goal _____ sheer determination.'?", ["(a) in / with", "(b) at / by", "(c) on / with", "(d) in / by"], "CO1,L3"],
    ["19", "Identify the sentence in PAST PERFECT tense:", ["(a) She was studying all night.", "(b) She had submitted the form before the deadline.", "(c) She studied all night.", "(d) She has studied for five hours."], "CO1,L2"],
    ["20", "Select the correct option: 'Neither the manager nor the employees _____ aware of the change.'", ["(a) is", "(b) are", "(c) was", "(d) were"], "CO3,L3"],
    // Q21-30
    ["21", "Choose the correct article: _____ Eiffel Tower is one of _____ most visited monuments in _____ world.", ["(a) The / the / the", "(b) A / the / the", "(c) The / a / a", "(d) A / a / the"], "CO1,L3"],
    ["22", "Which sentence is in the ACTIVE voice?", ["(a) The window was broken by the ball.", "(b) The ball was thrown by him.", "(c) The ball broke the window.", "(d) The window had been broken."], "CO1,L2"],
    ["23", "Choose the correct sentence (tense error free): 'By the time we reach, the train _____.'", ["(a) will leave", "(b) would have left", "(c) will have left", "(d) leaves"], "CO3,L3"],
    ["24", "Identify the NOUN in the underlined group: 'The rapidly flowing river flooded the village.'", ["(a) rapidly", "(b) flowing", "(c) river", "(d) flooded"], "CO1,L2"],
    ["25", "Select the correct option to fill the blank: '_____ student must carry their ID card.'", ["(a) Every", "(b) All", "(c) Many", "(d) Few"], "CO1,L5"],
    ["26", "Which sentence correctly uses 'although'?", ["(a) Although but she was tired, she kept working.", "(b) She kept working although she was tired.", "(c) She kept working although of being tired.", "(d) Although she working, but she was tired."], "CO3,L3"],
    ["27", "Identify the correct question tag: 'You have submitted your assignment, _____?'", ["(a) haven't you", "(b) didn't you", "(c) don't you", "(d) have you"], "CO3,L3"],
    ["28", "Choose the correct passive voice: 'Someone has stolen his bicycle.'", ["(a) His bicycle is stolen by someone.", "(b) His bicycle was stolen by someone.", "(c) His bicycle has been stolen.", "(d) His bicycle had been stolen by someone."], "CO1,L2"],
    ["29", "Choose the most appropriate meaning of 'turn down':", ["(a) To decrease volume / to reject", "(b) To fall down", "(c) To look down upon", "(d) To move in a downward direction"], "CO2,L3"],
    ["30", "Identify the sentence with the correct use of 'each' and 'every':", ["(a) Each of the students have submitted their work.", "(b) Every children in the class is present.", "(c) Each student was given a separate task.", "(d) Every of the boys are playing."], "CO1,L5"],
    // Q31-40
    ["31", "Select the correct option: 'The news _____ shocking to everyone.'", ["(a) were", "(b) was", "(c) are", "(d) have been"], "CO3,L3"],
    ["32", "Choose the correct degree of comparison: 'Gold is _____ than silver.'", ["(a) more precious", "(b) most precious", "(c) preciouser", "(d) much more preciouser"], "CO1,L5"],
    ["33", "Which sentence uses the Past Perfect CORRECTLY?", ["(a) She has finished the work before he arrived.", "(b) She had finished the work before he arrived.", "(c) She finished the work before he had arrived.", "(d) She was finishing the work before he arrived."], "CO3,L3"],
    ["34", "Choose the correct option: 'There is _____ milk left; we need to buy _____.'", ["(a) little / some", "(b) few / any", "(c) a few / some", "(d) little / any"], "CO1,L5"],
    ["35", "Identify the adverb in the sentence: 'The children played happily in the garden.'", ["(a) children", "(b) played", "(c) happily", "(d) garden"], "CO1,L5"],
    ["36", "Which sentence correctly uses 'so that'?", ["(a) She studied hard so that she can pass.", "(b) She studied hard so that she could pass.", "(c) She studied hard so that she passes.", "(d) She studied hard so that passing."], "CO3,L3"],
    ["37", "Choose the correct active voice: 'A letter was being written by him.'", ["(a) He wrote a letter.", "(b) He was writing a letter.", "(c) He has written a letter.", "(d) He had written a letter."], "CO2,L3"],
    ["38", "Identify the part of speech of 'quickly' in: 'She ran quickly to catch the bus.'", ["(a) Adjective", "(b) Verb", "(c) Noun", "(d) Adverb"], "CO1,L2"],
    ["39", "Choose the correct option: 'He is used to _____ early.'", ["(a) wake", "(b) woke", "(c) waking", "(d) have woken"], "CO2,L3"],
    ["40", "Which phrasal verb means 'to investigate'?", ["(a) Look after", "(b) Look into", "(c) Look out", "(d) Look for"], "CO2,L3"],
    // Q41-50
    ["41", "Identify the tense: 'By next month, she will have completed her thesis.'", ["(a) Future Continuous", "(b) Future Perfect", "(c) Present Perfect", "(d) Future Simple"], "CO1,L2"],
    ["42", "Choose the correct article: 'He plays _____ violin in _____ orchestra.'", ["(a) the / an", "(b) a / the", "(c) the / the", "(d) a / an"], "CO1,L3"],
    ["43", "Which sentence correctly uses the conjunction 'whereas'?", ["(a) Whereas she likes coffee, but her sister prefers tea.", "(b) She likes coffee, whereas her sister prefers tea.", "(c) She likes coffee whereas of her sister.", "(d) Whereas she likes coffee and her sister prefers tea."], "CO3,L3"],
    ["44", "Select the correct passive: 'The committee will announce the results tomorrow.'", ["(a) The results were announced by the committee tomorrow.", "(b) The results are announced by the committee tomorrow.", "(c) The results will be announced by the committee tomorrow.", "(d) The results would be announced by the committee tomorrow."], "CO1,L2"],
    ["45", "Choose the correct option: '_____ old woman and _____ honest man sat under _____ oak tree.'", ["(a) An / an / the", "(b) The / an / an", "(c) An / an / an", "(d) An / the / the"], "CO1,L3"],
    ["46", "Identify the COLLECTIVE noun: 'A _____ of wolves howled in the night.'", ["(a) gang", "(b) pack", "(c) flock", "(d) fleet"], "CO1,L2"],
    ["47", "Which sentence uses 'used to' CORRECTLY?", ["(a) She is used to live alone.", "(b) She used to living alone.", "(c) She used to live alone.", "(d) She uses to live alone."], "CO2,L3"],
    ["48", "Choose the correct form: 'The manager made it clear that all employees _____ follow the new policy.'", ["(a) should", "(b) ought", "(c) would", "(d) must"], "CO2,L3"],
    ["49", "Identify the sentence in Present Simple tense:", ["(a) The earth revolves around the sun.", "(b) She has finished her work.", "(c) He is going to school.", "(d) They have been waiting for an hour."], "CO1,L1"],
    ["50", "Choose the correct option: 'I have known him _____ many years.'", ["(a) since", "(b) for", "(c) ago", "(d) before"], "CO2,L5"],
    // Q51-60
    ["51", "Identify the error: 'Neither of the students have submitted their assignments on time.'", ["(a) Neither of", "(b) have submitted", "(c) their assignments", "(d) on time"], "CO3,L6"],
    ["52", "Choose the correct option: '_____ Himalayas separate India from Tibet.'", ["(a) A", "(b) An", "(c) The", "(d) No article"], "CO1,L3"],
    ["53", "Choose the correct passive: 'He was made to apologise by his teacher.'", ["(a) His teacher made him apologise.", "(b) His teacher made him to apologise.", "(c) His teacher was making him apologise.", "(d) His teacher had made him apologise."], "CO1,L2"],
    ["54", "Identify the conjunction type in: 'He worked hard; nevertheless, he failed.'", ["(a) Co-ordinating conjunction", "(b) Subordinating conjunction", "(c) Correlative conjunction", "(d) Conjunctive adverb / transitional connector"], "CO1,L3"],
    ["55", "Which sentence correctly uses the Past Perfect Continuous?", ["(a) She has been working here since 2010.", "(b) She had been working on the project for two years before she quit.", "(c) She was working on the project for two years.", "(d) She worked on the project for two years."], "CO1,L2"],
    ["56", "Select the correct preposition: 'She is very good _____ solving puzzles.'", ["(a) in", "(b) on", "(c) at", "(d) for"], "CO1,L3"],
    ["57", "Choose the sentence where the adjective is used CORRECTLY:", ["(a) He drove careful.", "(b) She looked at him angry.", "(c) The anxious students waited outside.", "(d) She sings beautiful."], "CO1,L5"],
    ["58", "Identify the correct sentence using 'other' or 'others':", ["(a) The others students left early.", "(b) Others students were still studying.", "(c) Some students passed; others failed.", "(d) She is better than the other one's."], "CO1,L5"],
    ["59", "Choose the correct tense: 'When I _____ at the station, the train had already left.'", ["(a) arrived", "(b) arrive", "(c) was arriving", "(d) had arrived"], "CO3,L3"],
    ["60", "Identify the sentence where 'had better' is used CORRECTLY:", ["(a) You had better to call him now.", "(b) We had better leave before it rains.", "(c) She had better to not be late.", "(d) He had better goes now."], "CO2,L4"],
  ];

  const children = [
    h1("SAMPLE PAPER – I"),
    divider(),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Course Code: PEL121    Course Title: COMMUNICATION SKILLS-I", bold: true, size: 24 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: "Time Allowed: 03:00 hrs.                                Max Marks: 60", size: 22 })] }),
    noteBox("This question paper contains 60 questions of 1 mark each. 0.25 marks will be deducted for each wrong answer. Attempt all questions in serial order."),
    space(),
  ];
  for (const [num, q, opts, co] of qs) {
    children.push(...mcqQ(num, q, opts, co));
  }
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240, after: 0 }, children: [new TextRun({ text: "-- End of Sample Paper I --", bold: true, italics: true })] }));
  children.push(pageBreak());
  return children;
}

// ─── SAMPLE PAPER 2 ───────────────────────────────────────────────────────────

function samplePaper2() {
  const qs = [
    ["1", "Identify the part of speech of the underlined word: 'She spoke CONFIDENTLY during the interview.'", ["(a) Adjective", "(b) Noun", "(c) Adverb", "(d) Verb"], "CO1,L2"],
    ["2", "Choose the correct article: 'He is _____ M.B.A. graduate and works for _____ United Nations.'", ["(a) an / the", "(b) a / the", "(c) an / a", "(d) a / a"], "CO1,L3"],
    ["3", "Which sentence is grammatically CORRECT?", ["(a) Despite of his illness, he attended class.", "(b) Despite his illness, he attended class.", "(c) Despite he was ill, he attended class.", "(d) Despite to being ill, he attended class."], "CO3,L3"],
    ["4", "Choose the correct passive voice of: 'The committee is considering the proposal.'", ["(a) The proposal was being considered by the committee.", "(b) The proposal is being considered by the committee.", "(c) The proposal has been considered by the committee.", "(d) The proposal had been considered by the committee."], "CO1,L2"],
    ["5", "Choose the correct option: 'I _____ for this company since 2015.'", ["(a) worked", "(b) have been working", "(c) am working", "(d) was working"], "CO3,L3"],
    ["6", "Identify the COMPOUND noun:", ["(a) beautiful garden", "(b) toothbrush", "(c) quickly done", "(d) very tall"], "CO1,L2"],
    ["7", "Select the correct preposition: 'She insisted _____ attending the meeting personally.'", ["(a) for", "(b) in", "(c) on", "(d) at"], "CO1,L3"],
    ["8", "Which sentence is in the FUTURE PERFECT tense?", ["(a) She will leave for Delhi tomorrow.", "(b) She will be leaving for Delhi at noon.", "(c) She will have left for Delhi before noon.", "(d) She is going to leave for Delhi."], "CO1,L2"],
    ["9", "Choose the correct option: '_____ you work hard, _____ you will succeed.'", ["(a) The more / the more", "(b) More / more", "(c) Most / most", "(d) The most / the most"], "CO3,L3"],
    ["10", "Identify the sentence with INCORRECT use of adjectives:", ["(a) It was a dark, stormy night.", "(b) She wore a beautiful, long, blue gown.", "(c) He is the most tallest player on the team.", "(d) This is the most challenging task I have faced."], "CO1,L6"],
    ["11", "Choose the correct sentence:", ["(a) Neither the principal nor the teachers was present.", "(b) Neither the principal nor the teachers were present.", "(c) Neither the principal nor the teachers is present.", "(d) Neither the principal nor the teachers have been present."], "CO3,L3"],
    ["12", "Identify the tense: 'By 1990, scientists had already mapped much of the human genome.'", ["(a) Past Continuous", "(b) Past Simple", "(c) Past Perfect", "(d) Present Perfect"], "CO1,L2"],
    ["13", "Choose the correct option to fill in the blank: '_____ of the two solutions is correct.'", ["(a) None", "(b) Neither", "(c) Both", "(d) All"], "CO1,L5"],
    ["14", "Which sentence uses 'in case' CORRECTLY?", ["(a) She bought an umbrella in case of it rained.", "(b) She bought an umbrella in case it rained.", "(c) She bought an umbrella in case of raining.", "(d) She bought an umbrella in case to rain."], "CO3,L3"],
    ["15", "Identify the correct passive of: 'The audience laughed at the comedian.'", ["(a) The comedian was laughed at by the audience.", "(b) The comedian was laughed by the audience.", "(c) The comedian is laughed at by the audience.", "(d) The comedian had been laughed by the audience."], "CO1,L2"],
    ["16", "Choose the phrasal verb meaning 'to start a journey':", ["(a) Set aside", "(b) Set back", "(c) Set off", "(d) Set up"], "CO2,L3"],
    ["17", "Identify the ABSTRACT noun in the sentence: 'Her kindness and wisdom inspired everyone.'", ["(a) Her", "(b) kindness and wisdom", "(c) inspired", "(d) everyone"], "CO1,L2"],
    ["18", "Which conjunction correctly fills the blank: 'He was tired; _____, he continued working.'", ["(a) because", "(b) so", "(c) however", "(d) since"], "CO3,L3"],
    ["19", "Choose the correct degree: 'This is _____ book I have ever read.'", ["(a) more interesting", "(b) most interesting", "(c) the more interesting", "(d) the most interesting"], "CO1,L5"],
    ["20", "Choose the correct active form: 'The bridge is being repaired by the workers.'", ["(a) The workers repaired the bridge.", "(b) The workers are repairing the bridge.", "(c) The workers have repaired the bridge.", "(d) The workers were repairing the bridge."], "CO2,L3"],
    ["21", "Identify the sentence with correct use of 'another' and 'other':", ["(a) I need anothers books.", "(b) She took another piece of cake.", "(c) Give me the other ones book.", "(d) Do you have others copies?"], "CO1,L5"],
    ["22", "Choose the correct question tag: 'Let's go for a walk, _____?'", ["(a) don't we", "(b) shall we", "(c) will we", "(d) won't we"], "CO3,L3"],
    ["23", "Which sentence uses the relative pronoun CORRECTLY?", ["(a) The house which John bought it is beautiful.", "(b) The house that John bought is beautiful.", "(c) The house whom John bought is beautiful.", "(d) The house who John bought is beautiful."], "CO3,L3"],
    ["24", "Choose the correct option: 'I wish I _____ harder when I was young.'", ["(a) studied", "(b) had studied", "(c) have studied", "(d) would study"], "CO3,L3"],
    ["25", "Choose the correct passive: 'You must submit the form before Friday.'", ["(a) The form must submit before Friday.", "(b) The form must be submitted before Friday.", "(c) The form has been submitted before Friday.", "(d) The form should have been submitted before Friday."], "CO1,L2"],
    ["26", "Identify the PREPOSITION in: 'The cat jumped onto the table.'", ["(a) cat", "(b) jumped", "(c) onto", "(d) table"], "CO1,L2"],
    ["27", "Choose the correct sentence using 'so … that':", ["(a) She was so tired that she could hardly walk.", "(b) She was so tired so she could hardly walk.", "(c) She was tired so that she could hardly walk.", "(d) She was such tired that she could hardly walk."], "CO3,L3"],
    ["28", "Choose the correct option: 'The police _____ the area thoroughly last night.'", ["(a) had searched", "(b) searched", "(c) searches", "(d) were searching"], "CO3,L3"],
    ["29", "Choose the sentence with the INCORRECT use of the phrasal verb 'put off':", ["(a) Please don't put off your work till tomorrow.", "(b) The meeting was put off till Monday.", "(c) He put off the light before sleeping.", "(d) The event was put off due to rain."], "CO3,L3"],
    ["30", "Identify the correct order of adjectives:", ["(a) a round big old brown wooden table", "(b) a big old round brown wooden table", "(c) a wooden old big round brown table", "(d) a brown big round old wooden table"], "CO1,L1"],
    ["31", "Choose the correct article: 'He was admitted to _____ hospital after _____ accident.'", ["(a) the / the", "(b) a / an", "(c) the / an", "(d) a / the"], "CO1,L3"],
    ["32", "Which sentence is in the PAST PERFECT CONTINUOUS tense?", ["(a) She was painting the wall when I arrived.", "(b) She had painted the wall before I arrived.", "(c) She had been painting the wall for two hours when I arrived.", "(d) She has been painting the wall for two hours."], "CO1,L2"],
    ["33", "Choose the correct option: 'There are _____ people waiting outside.'", ["(a) much", "(b) a little", "(c) a few", "(d) little"], "CO1,L5"],
    ["34", "Identify the INCORRECT sentence:", ["(a) She is interested in learning new languages.", "(b) She is good at playing chess.", "(c) She is capable of handling the situation.", "(d) She is fond to read mystery novels."], "CO3,L6"],
    ["35", "Choose the correct passive voice of: 'They will have completed the project by July.'", ["(a) The project will have been completed by July.", "(b) The project was completed by July.", "(c) The project will be completed by July.", "(d) The project had been completed by July."], "CO1,L2"],
    ["36", "Select the correct sentence using 'while':", ["(a) While she was sleeping, the phone rang.", "(b) While she sleeping, the phone rang.", "(c) While she slept but the phone rang.", "(d) While she sleeps, the phone rang."], "CO3,L3"],
    ["37", "Identify the type of noun: 'The jury has delivered its verdict.'", ["(a) Proper noun", "(b) Abstract noun", "(c) Collective noun", "(d) Countable noun"], "CO1,L2"],
    ["38", "Choose the correct phrasal verb: 'I can't _____ what the professor is saying.'", ["(a) make out", "(b) make off", "(c) make up", "(d) make over"], "CO2,L3"],
    ["39", "Which sentence correctly uses 'used to'?", ["(a) He is used to gets up early.", "(b) He used to get up early.", "(c) He uses to get up early.", "(d) He was used to get up early."], "CO2,L3"],
    ["40", "Choose the correct form: 'This is _____ interesting of the two films.'", ["(a) more", "(b) most", "(c) the more", "(d) the most"], "CO1,L5"],
    ["41", "Identify the error: 'She doesn't knows where he lives.'", ["(a) She", "(b) doesn't knows", "(c) where", "(d) lives"], "CO3,L6"],
    ["42", "Choose the correct option: 'He _____ the answer before anyone else.'", ["(a) had known", "(b) knew", "(c) was knowing", "(d) has known"], "CO3,L3"],
    ["43", "Select the correct preposition: 'She was praised _____ her outstanding performance.'", ["(a) about", "(b) on", "(c) for", "(d) of"], "CO1,L3"],
    ["44", "Which sentence uses 'for' CORRECTLY as a conjunction?", ["(a) She studied for she wanted to succeed.", "(b) She couldn't sleep, for she was very anxious.", "(c) She studied for passing the exam.", "(d) For she was anxious, she couldn't sleep."], "CO3,L3"],
    ["45", "Identify the sentence in the CORRECT future continuous tense:", ["(a) She will finish her homework by evening.", "(b) She will be finishing her homework by evening.", "(c) She will have finished her homework by evening.", "(d) She is going to finish her homework by evening."], "CO1,L2"],
    ["46", "Choose the correct active form: 'The cake had been baked by my mother.'", ["(a) My mother baked the cake.", "(b) My mother was baking the cake.", "(c) My mother had baked the cake.", "(d) My mother has baked the cake."], "CO2,L3"],
    ["47", "Choose the correct option: 'The scientist, _____ research changed the world, won the Nobel Prize.'", ["(a) who", "(b) which", "(c) whose", "(d) whom"], "CO1,L3"],
    ["48", "Identify the CORRECT sentence:", ["(a) She sings beautiful.", "(b) He drives careful.", "(c) They played the game enthusiastically.", "(d) She danced gracious."], "CO1,L5"],
    ["49", "Choose the correct conjunction: '_____ the weather improves, the match will be postponed.'", ["(a) As long as", "(b) In case", "(c) Unless", "(d) When"], "CO3,L3"],
    ["50", "Which sentence correctly uses 'each'?", ["(a) Each of the players have won a medal.", "(b) Each player have won a medal.", "(c) Each player has won a medal.", "(d) Each players has won a medal."], "CO1,L5"],
    ["51", "Identify the correct sentence using the Past Simple tense:", ["(a) She is lived in Delhi for ten years.", "(b) She lived in Delhi for ten years.", "(c) She has lived in Delhi for ten years.", "(d) She had been living in Delhi for ten years."], "CO1,L2"],
    ["52", "Select the sentence with correct subject–verb agreement:", ["(a) The quality of the goods were poor.", "(b) Each of the boys have done their work.", "(c) The committee has decided to postpone the meeting.", "(d) None of the students were present."], "CO3,L3"],
    ["53", "Choose the correct option: 'He ran _____ fast that no one could catch him.'", ["(a) such", "(b) very", "(c) so", "(d) too"], "CO3,L3"],
    ["54", "Identify the sentence with the INCORRECT passive voice:", ["(a) The wall was painted by him.", "(b) The report was submitted on time.", "(c) He was suggested to see a doctor.", "(d) The children were played football by them."], "CO3,L6"],
    ["55", "Choose the correct preposition: 'The professor is known _____ his research on climate change.'", ["(a) by", "(b) for", "(c) about", "(d) of"], "CO1,L3"],
    ["56", "Which sentence uses 'a little' or 'a few' CORRECTLY?", ["(a) She has a few money left.", "(b) He drank a little water.", "(c) I have a few informations.", "(d) There is a few sugar in the bowl."], "CO1,L5"],
    ["57", "Identify the CORRECT sentence:", ["(a) She has been knowing him for years.", "(b) She knew him for years.", "(c) She has known him for years.", "(d) She had been knowing him since 2010."], "CO3,L6"],
    ["58", "Choose the correct passive voice: 'People believe that he is innocent.'", ["(a) It was believed that he was innocent.", "(b) He is believed to be innocent.", "(c) That he is innocent is believed by people.", "(d) Both (a) and (b) are correct."], "CO3,L3"],
    ["59", "Select the phrasal verb that means 'to cancel':", ["(a) Call back", "(b) Call on", "(c) Call off", "(d) Call up"], "CO2,L3"],
    ["60", "Choose the sentence where 'had better' is used INCORRECTLY:", ["(a) You had better finish your work.", "(b) We had better not waste time.", "(c) She had better to take medicine.", "(d) He had better be on time."], "CO2,L4"],
  ];

  const children = [
    h1("SAMPLE PAPER – II"),
    divider(),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Course Code: PEL121    Course Title: COMMUNICATION SKILLS-I", bold: true, size: 24 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: "Time Allowed: 03:00 hrs.                                Max Marks: 60", size: 22 })] }),
    noteBox("This question paper contains 60 questions of 1 mark each. 0.25 marks will be deducted for each wrong answer. Attempt all questions in serial order."),
    space(),
  ];
  for (const [num, q, opts, co] of qs) {
    children.push(...mcqQ(num, q, opts, co));
  }
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 240 }, children: [new TextRun({ text: "-- End of Sample Paper II --", bold: true, italics: true })] }));
  children.push(pageBreak());
  return children;
}

// ─── ANSWER KEY + EXPLANATIONS ───────────────────────────────────────────────

function answerKey1() {
  const answers = [
    ["1","(a)","listened / was cooking","Past Simple + Past Continuous: 'listened' is the action that interrupted the ongoing 'was cooking'."],
    ["2","(c)","Justice","Abstract noun – names a quality/idea that cannot be touched or seen."],
    ["3","(d)","no article / a","'advice' is uncountable (no article); 'career' is singular countable with general reference (a career)."],
    ["4","(c)","The students are taught…","Present Simple active → Present Simple passive uses 'is/are + V3'. 'are taught' is correct."],
    ["5","(c)","dropped out","'drop out' = to leave school/university without completing a course."],
    ["6","(a)","beautiful small red silk","Royal order: Opinion (beautiful) → Size (small) → Colour (red) → Material (silk)."],
    ["7","(b)","Despite working hard, she failed.","'Despite' must be followed by a noun or gerund (V-ing), not a clause."],
    ["8","(c)","She has been writing the novel for three hours.","Formula: has/have + been + V-ing; 'for three hours' confirms ongoing duration."],
    ["9","(b)","were having","Past Continuous: ongoing action (having dinner) interrupted by guests' arrival (arrived = Past Simple)."],
    ["10","(c)","There isn't much traffic today.","'traffic' is uncountable; 'much' is used with uncountable nouns."],
    ["11","(c)","teachers'","Plural noun ending in 's': add only apostrophe after s → teachers'."],
    ["12","(c)","who","'who' is used for persons as the subject of the relative clause."],
    ["13","(c)","There is a little hope left, but only a few days remain.","'a little' with uncountable (hope); 'a few' with countable (days)."],
    ["14","(a)","The concert has been cancelled.","Present Perfect passive: have/has + been + V3. Active: They have cancelled → passive: has been cancelled."],
    ["15","(d)","He gave away them his old books.","Pronoun object must go between verb and particle: 'He gave them away.'"],
    ["16","(d)","Sentence is correct as written","'used to' does not require an additional auxiliary. The sentence needs no insertion."],
    ["17","(d)","the more experienced","Comparing two specific candidates: use 'the + comparative' (the more experienced)."],
    ["18","(a)","in / with","succeed in (doing something); with sheer determination (manner)."],
    ["19","(b)","She had submitted the form before the deadline.","Past Perfect: had + V3; used for an action completed before another past event."],
    ["20","(b)","are","With 'neither … nor', verb agrees with the nearer subject ('employees' – plural) → 'are'."],
    ["21","(a)","The / the / the","Specific famous landmark → the Eiffel Tower; superlative → the most visited; unique entity → the world."],
    ["22","(c)","The ball broke the window.","Active voice: subject (the ball) performs the action (broke)."],
    ["23","(c)","will have left","Future Perfect: will + have + V3; action completed before a future time."],
    ["24","(c)","river","'river' is the noun (concrete, common); 'rapidly' is adverb, 'flowing' is participle/adjective, 'flooded' is verb."],
    ["25","(a)","Every","'Every + singular noun' is used for collective universal statements."],
    ["26","(b)","She kept working although she was tired.","'although' introduces a subordinate clause; no 'but' needed."],
    ["27","(a)","haven't you","Positive statement with Present Perfect (have submitted) → negative tag (haven't you)."],
    ["28","(c)","His bicycle has been stolen.","Present Perfect passive: has + been + V3. 'by someone' omitted as agent is unknown/unimportant."],
    ["29","(a)","To decrease volume / to reject","'turn down' = decrease (volume/heat) OR reject (an offer/application)."],
    ["30","(c)","Each student was given a separate task.","'each + singular noun + singular verb' is correct."],
    ["31","(b)","was","'news' is uncountable and always takes a singular verb → 'was'."],
    ["32","(a)","more precious","Two-syllable+ adjective: use 'more' for comparative. Never add '-er' to 'precious'."],
    ["33","(b)","She had finished the work before he arrived.","The earlier action uses Past Perfect (had finished); the later action uses Past Simple (arrived)."],
    ["34","(a)","little / some","'milk' is uncountable → 'little'; 'some' is used in affirmative statements for uncountable."],
    ["35","(c)","happily","'happily' describes how they played → adverb."],
    ["36","(b)","She studied hard so that she could pass.","'so that' expresses purpose; past tense verb in purpose clause uses 'could' (not 'can')."],
    ["37","(b)","He was writing a letter.","Past Continuous passive (was being written) → active: He was writing a letter."],
    ["38","(d)","Adverb","'quickly' modifies the verb 'ran' → adverb."],
    ["39","(c)","waking","'be used to' is always followed by noun/gerund (V-ing): 'used to waking'."],
    ["40","(b)","Look into","'look into' = to investigate or examine something thoroughly."],
    ["41","(b)","Future Perfect","Formula: will + have + V3 + by [future time]."],
    ["42","(a)","the / an","Musical instruments use 'the'; 'orchestra' begins with vowel sound → 'an'."],
    ["43","(b)","She likes coffee, whereas her sister prefers tea.","'whereas' contrasts two clauses without needing 'but' or any other conjunction."],
    ["44","(c)","The results will be announced by the committee tomorrow.","Future Simple passive: will + be + V3."],
    ["45","(a)","An / an / the","'old' starts with vowel sound → 'an'; 'honest' silent h → 'an'; specific unique oak → 'the'."],
    ["46","(b)","pack","A 'pack' of wolves is the correct collective noun. (flock = birds/sheep; fleet = ships/vehicles; gang = people)"],
    ["47","(c)","She used to live alone.","'used to + V1' describes past habits. Negative/question form: 'She didn't use to…'"],
    ["48","(a)","should","'made it clear that…should' is the correct formal obligation structure."],
    ["49","(a)","The earth revolves around the sun.","Present Simple states facts/universal truths."],
    ["50","(b)","for","'for' is used with a period/duration of time (many years); 'since' is used with a point in time."],
    ["51","(b)","have submitted","'neither of + plural noun' takes a singular verb → 'has submitted', not 'have submitted'."],
    ["52","(c)","The","Mountain ranges use 'the': the Himalayas."],
    ["53","(a)","His teacher made him apologise.","Causative 'make' takes bare infinitive: 'made him apologise' (no 'to')."],
    ["54","(d)","Conjunctive adverb / transitional connector","'nevertheless' is a conjunctive adverb, not a traditional conjunction."],
    ["55","(b)","She had been working on the project for two years before she quit.","Past Perfect Continuous: had + been + V-ing; duration before another past action."],
    ["56","(c)","at","'be good at' is a fixed prepositional phrase."],
    ["57","(c)","The anxious students waited outside.","'anxious' is an adjective correctly modifying 'students'. Others use adjectives where adverbs are needed."],
    ["58","(c)","Some students passed; others failed.","'others' alone acts as a pronoun. 'other' must precede a noun."],
    ["59","(a)","arrived","'When' clause in past with Past Perfect in main clause: 'had already left' + 'arrived' (Past Simple)."],
    ["60","(b)","We had better leave before it rains.","'had better + bare infinitive' (no 'to'): 'had better leave' is correct."],
  ];

  const children = [
    h1("ANSWER KEY – Sample Paper I (with Explanations)"),
    divider(),
  ];

  // Build answer table with header
  const headerRow = new TableRow({
    children: [
      new TableCell({ borders, width: { size: 700, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Q No.", bold: true, color: "FFFFFF" })] })] }),
      new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Answer", bold: true, color: "FFFFFF" })] })] }),
      new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Answer Text", bold: true, color: "FFFFFF" })] })] }),
      new TableCell({ borders, width: { size: 5360, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Explanation", bold: true, color: "FFFFFF" })] })] }),
    ],
  });

  const rows = answers.map(([q, ans, text, exp], i) =>
    new TableRow({
      children: [
        new TableCell({ borders, width: { size: 700, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: q, bold: true })] })] }),
        new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: ans, bold: true, color: "1A6B3A" })] })] }),
        new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun(text)] })] }),
        new TableCell({ borders, width: { size: 5360, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun(exp)] })] }),
      ],
    })
  );

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [700, 800, 2500, 5360],
    rows: [headerRow, ...rows],
  }));
  children.push(pageBreak());
  return children;
}

function answerKey2() {
  const answers = [
    ["1","(c)","Adverb","'confidently' modifies the verb 'spoke' → adverb."],
    ["2","(a)","an / the","'MBA' starts with vowel sound → 'an'; 'United Nations' is a specific organisation → 'the'."],
    ["3","(b)","Despite his illness, he attended class.","'Despite' + noun/gerund, NOT clause. 'Despite of' is always wrong."],
    ["4","(b)","The proposal is being considered…","Present Continuous passive: is/am/are + being + V3."],
    ["5","(b)","have been working","Duration from 2015 to now → Present Perfect Continuous (have been + V-ing)."],
    ["6","(b)","toothbrush","A compound noun is formed by two nouns together. 'toothbrush' = tooth + brush."],
    ["7","(c)","on","'insist on' is a fixed phrasal preposition: insist on doing/attending."],
    ["8","(c)","She will have left for Delhi before noon.","Future Perfect: will + have + V3 + before a future time."],
    ["9","(a)","The more / the more","Parallel comparative structure: 'The more … the more …' (correlative)."],
    ["10","(c)","He is the most tallest player on the team.","'most tallest' is a double superlative – INCORRECT. Use 'the tallest'."],
    ["11","(b)","Neither the principal nor the teachers were present.","With 'neither…nor', verb agrees with nearer subject (teachers – plural) → 'were'."],
    ["12","(c)","Past Perfect","'had already mapped' = had + V3; used for action completed before another past reference point (By 1990)."],
    ["13","(b)","Neither","'neither of the two' = not one of exactly two options."],
    ["14","(b)","She bought an umbrella in case it rained.","'in case' + subject + past tense verb (for hypothetical future in reported speech context)."],
    ["15","(a)","The comedian was laughed at by the audience.","Phrasal verb passive: keep 'at' with the verb: 'laughed at' → 'was laughed at'."],
    ["16","(c)","Set off","'set off' = begin a journey / depart."],
    ["17","(b)","kindness and wisdom","Both are abstract nouns – qualities that cannot be perceived by senses."],
    ["18","(c)","however","'however' shows contrast/concession (but he continued despite being tired)."],
    ["19","(d)","the most interesting","Superlative with 'ever': 'the most interesting book I have ever read'."],
    ["20","(b)","The workers are repairing the bridge.","Present Continuous passive (is being repaired) → Active: are repairing."],
    ["21","(b)","She took another piece of cake.","'another' + singular countable noun is correct. 'anothers', 'other ones book' are wrong."],
    ["22","(b)","shall we","Imperative 'Let's …' always takes tag 'shall we'."],
    ["23","(b)","The house that John bought is beautiful.","No pronoun (it) repeated after relative pronoun. 'that' is correct for things."],
    ["24","(b)","had studied","'I wish + past perfect' expresses regret about the past."],
    ["25","(b)","The form must be submitted before Friday.","Modal passive: modal + be + V3 → 'must be submitted'."],
    ["26","(c)","onto","'onto' is a preposition showing movement/direction (cat jumped onto the table)."],
    ["27","(a)","She was so tired that she could hardly walk.","'so + adjective + that' is the correct structure."],
    ["28","(b)","searched","Completed past action with specific time (last night) → Past Simple."],
    ["29","(c)","He put off the light before sleeping.","'put off' means postpone. To turn off a light, the correct phrasal verb is 'put out' or 'turn off', not 'put off'."],
    ["30","(b)","a big old round brown wooden table","Royal order: Size (big) → Age (old) → Shape (round) → Colour (brown) → Material (wooden)."],
    ["31","(c)","the / an","Specific hospital (previously referenced) → 'the hospital'; 'accident' starts with vowel sound → 'an accident'."],
    ["32","(c)","She had been painting the wall for two hours when I arrived.","Past Perfect Continuous: had + been + V-ing + for [duration] + when [past simple]."],
    ["33","(c)","a few","'people' is countable plural → 'a few'."],
    ["34","(d)","She is fond to read mystery novels.","'fond of + V-ing', not 'fond to': INCORRECT. Correct: 'She is fond of reading mystery novels.'"],
    ["35","(a)","The project will have been completed by July.","Future Perfect Passive: will + have + been + V3."],
    ["36","(a)","While she was sleeping, the phone rang.","'while' + past continuous + past simple for interrupted action."],
    ["37","(c)","Collective noun","'jury' names a group of people acting as one unit → collective noun."],
    ["38","(a)","make out","'make out' = to understand / comprehend something with difficulty."],
    ["39","(b)","He used to get up early.","'used to + bare infinitive' for past habits. 'be used to' is different (followed by gerund)."],
    ["40","(c)","the more","Comparing two definite items: 'the + comparative' → 'the more interesting of the two'."],
    ["41","(b)","doesn't knows","After 'doesn't', use bare infinitive: 'doesn't know' (not 'knows')."],
    ["42","(b)","knew","Stative verb 'know' is not used in continuous tense; Past Simple for a past state."],
    ["43","(c)","for","'be praised for' is a fixed expression: praised for + noun/gerund."],
    ["44","(b)","She couldn't sleep, for she was very anxious.","'for' as a conjunction (formal) = 'because'; must connect two main clauses."],
    ["45","(b)","She will be finishing her homework by evening.","Future Continuous: will + be + V-ing."],
    ["46","(c)","My mother had baked the cake.","Past Perfect active (had + V3) corresponds to Past Perfect passive (had been baked)."],
    ["47","(c)","whose","'whose' shows possession/relationship: the scientist whose research (= his/her research)."],
    ["48","(c)","They played the game enthusiastically.","'enthusiastically' is an adverb correctly modifying 'played'. Others use adjectives where adverbs are needed."],
    ["49","(c)","Unless","'unless' = 'if … not': Unless the weather improves (= if the weather does not improve)."],
    ["50","(c)","Each player has won a medal.","'each + singular noun + singular verb (has)' is correct."],
    ["51","(b)","She lived in Delhi for ten years.","Past Simple for completed past action with a duration that is finished."],
    ["52","(c)","The committee has decided to postpone the meeting.","Collective noun 'committee' takes singular verb. The other sentences have agreement errors."],
    ["53","(c)","so","'so + adjective/adverb + that' structure for result. 'such' is used before nouns: 'such a fast runner that…'"],
    ["54","(d)","The children were played football by them.","'play' is intransitive in this sense (no object being acted upon); cannot be passivised this way. Correct: 'Football was played by the children.'"],
    ["55","(b)","for","'be known for' = to be recognised because of something."],
    ["56","(b)","He drank a little water.","'water' is uncountable → 'a little'. 'a few' is for countable nouns."],
    ["57","(c)","She has known him for years.","'know' is a stative verb – NOT used in Perfect Continuous form ('has been knowing' is wrong). Use Present Perfect simple."],
    ["58","(d)","Both (a) and (b) are correct.","Two standard impersonal passive constructions: 'It is believed that…' and '[Subject] is believed to be…'"],
    ["59","(c)","Call off","'call off' = to cancel an event or activity."],
    ["60","(c)","She had better to take medicine.","'had better + bare infinitive' (NO 'to'): Correct is 'She had better take medicine.'"],
  ];

  const children = [
    h1("ANSWER KEY – Sample Paper II (with Explanations)"),
    divider(),
  ];

  const headerRow = new TableRow({
    children: [
      new TableCell({ borders, width: { size: 700, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Q No.", bold: true, color: "FFFFFF" })] })] }),
      new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Answer", bold: true, color: "FFFFFF" })] })] }),
      new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Answer Text", bold: true, color: "FFFFFF" })] })] }),
      new TableCell({ borders, width: { size: 5360, type: WidthType.DXA }, shading: { fill: "2E4057", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: [new Paragraph({ children: [new TextRun({ text: "Explanation", bold: true, color: "FFFFFF" })] })] }),
    ],
  });

  const rows = answers.map(([q, ans, text, exp], i) =>
    new TableRow({
      children: [
        new TableCell({ borders, width: { size: 700, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: q, bold: true })] })] }),
        new TableCell({ borders, width: { size: 800, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun({ text: ans, bold: true, color: "1A6B3A" })] })] }),
        new TableCell({ borders, width: { size: 2500, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun(text)] })] }),
        new TableCell({ borders, width: { size: 5360, type: WidthType.DXA }, shading: i % 2 === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR } : undefined, margins: { top: 60, bottom: 60, left: 100, right: 60 }, children: [new Paragraph({ children: [new TextRun(exp)] })] }),
      ],
    })
  );

  children.push(new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [700, 800, 2500, 5360],
    rows: [headerRow, ...rows],
  }));
  return children;
}

// ─── TITLE PAGE ───────────────────────────────────────────────────────────────
function titlePage() {
  return [
    new Paragraph({ spacing: { before: 2880, after: 120 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PEL121 COMMUNICATION SKILLS-I", bold: true, size: 52, color: "2E4057" })] }),
    new Paragraph({ spacing: { after: 120 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Complete Notes & 2 Sample Papers with Solutions", bold: true, size: 36, color: "E8A838" })] }),
    divider(),
    space(), space(),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "Prepared for B.Tech / B.Sc. Students", size: 28, italics: true })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "Based on OXFORD PRACTICE GRAMMAR", size: 26 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "by Norman Coe, Mark Harrison, Ken Peterson", size: 24 })] }),
    space(), space(), space(),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "CONTENTS", bold: true, size: 32, color: "2E4057" })] }),
    divider(),
    twoColTable([
      ["Section", "Description"],
      ["Units I – VI", "Comprehensive Exam-Oriented Notes with Rules, Examples & Common Errors"],
      ["Quick Revision Sheets", "One-page summary for each of the 6 units"],
      ["Sample Paper I", "60 MCQs – University Style (Paper Code A)"],
      ["Sample Paper II", "60 MCQs – University Style (Paper Code B)"],
      ["Answer Key I", "Complete answers + explanations for Sample Paper I"],
      ["Answer Key II", "Complete answers + explanations for Sample Paper II"],
    ], 3120, 6240),
    space(), space(),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: "Exam Pattern: 60 MCQs × 1 mark each | Negative Marking: –0.25 per wrong answer | Time: 3 Hours", size: 22, italics: true })] }),
    pageBreak(),
  ];
}

// ─── ASSEMBLE DOCUMENT ────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 36, bold: true, font: "Arial", color: "2E4057" },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 },
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "Arial", color: "1A5276" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 },
      },
      {
        id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "Arial", color: "117A65" },
        paragraph: { spacing: { before: 180, after: 100 }, outlineLevel: 2 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1080, bottom: 1440, left: 1080 },
        },
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: "PEL121 Communication Skills-I  |  Page ", size: 18 }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18 }),
              new TextRun({ text: " of ", size: 18 }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18 }),
            ],
            alignment: AlignmentType.CENTER,
          })],
        }),
      },
      children: [
        ...titlePage(),
        ...unit1(),
        ...unit2(),
        ...unit3(),
        ...unit4(),
        ...unit5(),
        ...unit6(),
        ...samplePaper1(),
        ...samplePaper2(),
        ...answerKey1(),
        ...answerKey2(),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/PEL121_Complete_Notes_and_Sample_Papers.docx', buf);
  console.log('Done! File written.');
}).catch(e => { console.error(e); process.exit(1); });
