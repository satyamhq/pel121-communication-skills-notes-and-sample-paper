const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, TabStopType, TabStopPosition, LevelFormat,
  UnderlineType
} = require('docx');
const fs = require('fs');

// ─── HELPER FUNCTIONS ───────────────────────────────────────────────────────

function bold(text, size = 20) {
  return new TextRun({ text, bold: true, size });
}
function reg(text, size = 20) {
  return new TextRun({ text, size });
}
function italic(text, size = 20) {
  return new TextRun({ text, italics: true, size });
}
function underlineBold(text, size = 22) {
  return new TextRun({ text, bold: true, underline: { type: UnderlineType.SINGLE }, size });
}

function centerPara(children, spacingBefore = 0, spacingAfter = 0) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: spacingBefore, after: spacingAfter },
    children
  });
}

function leftPara(children, spacingBefore = 0, spacingAfter = 40) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: spacingBefore, after: spacingAfter },
    children
  });
}

function pageBreakPara() {
  return new Paragraph({ children: [new PageBreak()] });
}

function borderBox(children) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
              left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
              right: { style: BorderStyle.SINGLE, size: 4, color: "000000" }
            },
            children
          })
        ]
      })
    ]
  });
}

function singleLineBox(children) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
              bottom: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
              left: { style: BorderStyle.SINGLE, size: 2, color: "000000" },
              right: { style: BorderStyle.SINGLE, size: 2, color: "000000" }
            },
            children
          })
        ]
      })
    ]
  });
}

// Two-column table for Time/Marks row
function twoColRow(leftText, rightText) {
  const border = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4680, 4680],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 4680, type: WidthType.DXA },
            borders,
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [bold(leftText, 20)] })]
          }),
          new TableCell({
            width: { size: 4680, type: WidthType.DXA },
            borders,
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [bold(rightText, 20)] })]
          })
        ]
      })
    ]
  });
}

// ─── ANSWER KEY TABLE ───────────────────────────────────────────────────────

function answerKeyTable(answers) {
  // answers: array of { q, ans, explanation }
  // Build 3-column table: Q No | Answer | Explanation
  const headerBorder = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
  const cellBorder = { style: BorderStyle.SINGLE, size: 2, color: "000000" };
  const hBorders = { top: headerBorder, bottom: headerBorder, left: headerBorder, right: headerBorder };
  const cBorders = { top: cellBorder, bottom: cellBorder, left: cellBorder, right: cellBorder };

  const headerRow = new TableRow({
    children: [
      new TableCell({
        width: { size: 700, type: WidthType.DXA },
        borders: hBorders,
        shading: { fill: "D9E1F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [bold("Q.No.", 18)] })]
      }),
      new TableCell({
        width: { size: 1100, type: WidthType.DXA },
        borders: hBorders,
        shading: { fill: "D9E1F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [bold("Answer", 18)] })]
      }),
      new TableCell({
        width: { size: 7560, type: WidthType.DXA },
        borders: hBorders,
        shading: { fill: "D9E1F2", type: ShadingType.CLEAR },
        margins: { top: 60, bottom: 60, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [bold("Explanation", 18)] })]
      })
    ]
  });

  const rows = [headerRow];
  answers.forEach(({ q, ans, explanation }) => {
    rows.push(new TableRow({
      children: [
        new TableCell({
          width: { size: 700, type: WidthType.DXA },
          borders: cBorders,
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [reg(`Q${q}`, 18)] })]
        }),
        new TableCell({
          width: { size: 1100, type: WidthType.DXA },
          borders: cBorders,
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [bold(ans, 18)] })]
        }),
        new TableCell({
          width: { size: 7560, type: WidthType.DXA },
          borders: cBorders,
          margins: { top: 40, bottom: 40, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [reg(explanation, 18)] })]
        })
      ]
    }));
  });

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [700, 1100, 7560],
    rows
  });
}

// ─── QUESTION BUILDER ───────────────────────────────────────────────────────

function qPara(qNum, qText, co = "") {
  return leftPara([
    bold(`Q(${qNum})  `, 20),
    reg(qText, 20)
  ], 80, 20);
}

function optionsPara(a, b, c, d) {
  return leftPara([
    reg(`(a) ${a}`, 20),
    reg("          ", 20),
    reg(`(b) ${b}`, 20),
    reg("          ", 20),
    reg(`(c) ${c}`, 20),
    reg("          ", 20),
    reg(`(d) ${d}`, 20),
  ], 0, 0);
}

function optionsParaVertical(a, b, c, d) {
  return new Paragraph({
    spacing: { before: 0, after: 60 },
    children: [
      reg(`(a) ${a}    `, 20),
      reg(`(b) ${b}    `, 20),
      reg(`(c) ${c}    `, 20),
      reg(`(d) ${d}`, 20),
    ]
  });
}

function optionsBlock(a, b, c, d) {
  return [
    leftPara([reg(`(a) ${a}`, 20), reg("          ", 20), reg(`(b) ${b}`, 20)], 0, 0),
    leftPara([reg(`(c) ${c}`, 20), reg("          ", 20), reg(`(d) ${d}`, 20)], 0, 60),
  ];
}

function coTag(co) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: 60 },
    children: [italic(co, 18)]
  });
}

// ─── PAPER HEADER ───────────────────────────────────────────────────────────

function buildHeader(paperNum, paperCode, codeStr) {
  return [
    centerPara([bold("Registration No.:_______________________", 22)], 0, 80),

    borderBox([
      centerPara([underlineBold("IMPORTANT NOTE", 20)], 40, 40),
      leftPara([bold("1) ", 20), reg("This sample paper serves solely as a visual representation of the question paper.", 20)], 0, 20),
      leftPara([bold("2) ", 20), reg("Each question in the question paper is mapped with corresponding Course Outcome(s) (CO) and Revised Bloom's Taxonomy (RBT) Level. However the question-wise CO-RBT mapping given below is indicative only and will vary in the final question paper.", 20)], 0, 20),
      leftPara([bold("3) ", 20), reg("Students should refer to the latest course syllabus and question paper pattern for exam preparation.", 20)], 0, 20),
      leftPara([bold("4) ", 20), reg("The final version of the End Term Examination question paper may differ in terms of layout, design, and content. This sample paper should be viewed for reference purpose only and not as final/official question paper.", 20)], 0, 40),
    ]),

    leftPara([], 20, 0),
    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 60, after: 0 }, children: [bold(codeStr, 20)] }),
    new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 20, after: 80 }, children: [bold(`Paper Code: ${paperCode}`, 20)] }),

    centerPara([bold("Course Code: PEL121", 24)], 40, 20),
    centerPara([bold("Course Title: COMMUNICATION SKILLS-I", 24)], 0, 60),

    twoColRow("Time Allowed: 03:00 hrs.", "Max Marks: 60"),
    leftPara([], 20, 0),

    singleLineBox([
      leftPara([italic("Read the following instructions carefully before attempting the question paper.", 19)], 20, 20),
      leftPara([italic("1. Match the Paper Code shaded on the OMR Sheet with the Paper code mentioned on the question paper and ensure that both are the same.", 19)], 0, 20),
      leftPara([italic("2. This question paper contains 60 questions of 1 mark each. 0.25 marks will be deducted for each wrong answer.", 19)], 0, 20),
      leftPara([italic("3. Attempt all the questions in serial order.", 19)], 0, 20),
      leftPara([italic("4. Do not write or mark anything on the question paper and/or on rough sheet(s) which could be helpful to any student in copying, except your registration number on the designated space.", 19)], 0, 20),
      leftPara([italic("5. Submit the question paper and the rough sheet(s) along with the OMR sheet to the invigilator before leaving the examination hall.", 19)], 0, 20),
    ]),
    leftPara([], 40, 0),
  ];
}

// ─── PAPER 2 QUESTIONS ──────────────────────────────────────────────────────

const paper2Questions = [
  // PARTS OF SPEECH (Q1–Q14)
  { q: 1, text: "Choose the correct part of speech for the underlined word: She speaks English very _fluently_.", opts: ["Noun", "Adjective", "Adverb", "Verb"], co: "CO1,L2" },
  { q: 2, text: "Identify the noun in the following sentence: 'Honesty is the best policy.'", opts: ["is", "best", "Honesty", "policy — both (c) and (d)"], co: "CO1,L1" },
  { q: 3, text: "Which of the following is a transitive verb?\n(a) He sleeps early.\n(b) She reads a novel.\n(c) The baby cried.\n(d) They arrived late.", opts: ["He sleeps early.", "She reads a novel.", "The baby cried.", "They arrived late."], co: "CO1,L2", vertical: true },
  { q: 4, text: "Identify the preposition in: 'The book is on the shelf.'", opts: ["book", "is", "on", "shelf"], co: "CO1,L1" },
  { q: 5, text: "Which word is a conjunction in the sentence: 'I wanted to go, but it was raining.'?", opts: ["wanted", "but", "raining", "go"], co: "CO1,L1" },
  { q: 6, text: "Choose the correct pronoun: Neither the manager nor the employees ____ happy.", opts: ["is", "was", "were", "has been"], co: "CO1,L3" },
  { q: 7, text: "Identify the adjective: 'The golden retriever is a friendly dog.'", opts: ["retriever", "friendly, golden", "dog", "is"], co: "CO1,L1" },
  { q: 8, text: "Which sentence uses a possessive pronoun correctly?", opts: ["That book is their.", "That book is theirs.", "That book is they.", "That book is them."], co: "CO1,L2", vertical: true },
  { q: 9, text: "What is the verb form in: 'They have been studying for five hours.'?", opts: ["Present Simple", "Past Perfect", "Present Perfect Continuous", "Past Continuous"], co: "CO1,L2" },
  { q: 10, text: "Identify the correct use of an interjection:", opts: ["Oh! What a surprise it was.", "He oh was surprised.", "She was oh very surprised.", "They gave an oh reaction."], co: "CO1,L3", vertical: true },
  { q: 11, text: "Choose the correct reflexive pronoun: He hurt ____ while cooking.", opts: ["him", "himself", "his", "he"], co: "CO1,L1" },
  { q: 12, text: "Which sentence uses an abstract noun?", opts: ["The table is heavy.", "Courage is a great virtue.", "She bought a red car.", "The dog barked loudly."], co: "CO1,L2", vertical: true },
  { q: 13, text: "Which of the following is a linking verb?", opts: ["run", "eat", "seem", "throw"], co: "CO1,L2" },
  { q: 14, text: "Identify the demonstrative pronoun: ____ is the house I grew up in.", opts: ["That", "Which", "What", "Who"], co: "CO1,L1" },

  // TENSES (Q15–Q27)
  { q: 15, text: "Choose the correct tense form: By the time she arrived, they ____ dinner.", opts: ["had finished", "have finished", "finish", "finished"], co: "CO1,L3" },
  { q: 16, text: "Which sentence is in the Simple Future tense?", opts: ["I was going to school.", "I have been studying.", "I shall complete the task tomorrow.", "I completed the task."], co: "CO1,L2", vertical: true },
  { q: 17, text: "Fill in: She ____ in this company since 2018.", opts: ["is working", "worked", "has been working", "was working"], co: "CO1,L3" },
  { q: 18, text: "Choose the correct tense: When I reached the station, the train ____.", opts: ["already left", "has already left", "had already left", "was already leaving off"], co: "CO1,L3" },
  { q: 19, text: "Identify the tense: 'The children are playing in the yard.'", opts: ["Present Simple", "Past Continuous", "Present Perfect", "Present Continuous"], co: "CO1,L1" },
  { q: 20, text: "Choose the correct form: She ____ her homework before her father came home.", opts: ["completes", "completed", "has completed", "had completed"], co: "CO2,L3" },
  { q: 21, text: "Which sentence is in Present Perfect tense?", opts: ["She sings every day.", "She sang yesterday.", "She has sung this song before.", "She will sing tomorrow."], co: "CO1,L2", vertical: true },
  { q: 22, text: "Fill in: They ____ football every Sunday. (routine)", opts: ["play", "are playing", "have played", "had played"], co: "CO1,L2" },
  { q: 23, text: "Identify the tense: 'By next year, I will have completed my degree.'", opts: ["Simple Future", "Future Perfect", "Future Continuous", "Present Perfect"], co: "CO1,L2" },
  { q: 24, text: "Choose the correct option: He told me that he ____ tired.", opts: ["is", "will be", "was", "has been"], co: "CO2,L3" },
  { q: 25, text: "Which sentence correctly uses Past Perfect Continuous?", opts: ["She had been waiting for two hours.", "She has been waiting for two hours.", "She was waiting for two hours.", "She waited for two hours."], co: "CO1,L3", vertical: true },
  { q: 26, text: "Fill in the blank: I ____ (not/eat) anything since morning.", opts: ["did not eat", "have not eaten", "had not eaten", "was not eating"], co: "CO1,L3" },
  { q: 27, text: "Choose the correct sentence in Simple Past tense:", opts: ["He has written a letter.", "He writes a letter.", "He wrote a letter.", "He was writing a letter."], co: "CO1,L2", vertical: true },

  // ARTICLES (Q28–Q34)
  { q: 28, text: "Fill in the blank: ____ Himalayas are the highest mountains in the world.", opts: ["A", "An", "The", "No article"], co: "CO1,L3" },
  { q: 29, text: "Choose the correct article: She is ____ honest woman.", opts: ["a", "an", "the", "no article"], co: "CO1,L3" },
  { q: 30, text: "Which sentence uses articles correctly?", opts: ["He is a European.", "He is an European.", "She plays an piano.", "He is the European."], co: "CO1,L3", vertical: true },
  { q: 31, text: "Fill in: I saw ____ one-eyed man near ____ market.", opts: ["a / a", "an / the", "a / the", "the / a"], co: "CO1,L3" },
  { q: 32, text: "Choose the correct sentence with articles:", opts: ["Life is a wonderful gift.", "The life is a wonderful gift.", "A life is wonderful gift.", "Life is the wonderful gift."], co: "CO2,L5", vertical: true },
  { q: 33, text: "Fill in: We need ____ umbrella when it rains.", opts: ["a", "an", "the", "no article"], co: "CO1,L3" },
  { q: 34, text: "Choose the correct article: ____ sun rises in the east.", opts: ["A", "An", "The", "No article"], co: "CO1,L1" },

  // ACTIVE-PASSIVE VOICE (Q35–Q44)
  { q: 35, text: "Change to passive voice: 'Someone has stolen my wallet.'", opts: ["My wallet is stolen.", "My wallet has been stolen.", "My wallet was stolen.", "My wallet had been stolen."], co: "CO1,L2", vertical: true },
  { q: 36, text: "Identify the passive voice sentence:", opts: ["The teacher praises the student.", "The student was praised by the teacher.", "The teacher is praising the student.", "The teacher praised the student."], co: "CO1,L2", vertical: true },
  { q: 37, text: "Change to active voice: 'The letter was written by her.'", opts: ["She write the letter.", "She had written the letter.", "She wrote the letter.", "She has written the letter."], co: "CO1,L2", vertical: true },
  { q: 38, text: "Choose the correct passive transformation: 'They will announce the results tomorrow.'", opts: ["The results will be announced tomorrow.", "The results are announced tomorrow.", "The results were announced tomorrow.", "The results have been announced tomorrow."], co: "CO1,L2", vertical: true },
  { q: 39, text: "Which sentence is in active voice?", opts: ["A book was being read by him.", "He was reading a book.", "The book had been read.", "The book is read by him."], co: "CO1,L2", vertical: true },
  { q: 40, text: "Change to passive: 'People speak English all over the world.'", opts: ["English is spoken all over the world.", "English was spoken all over the world.", "English has been spoken all over the world.", "English had been spoken all over the world."], co: "CO3,L3", vertical: true },
  { q: 41, text: "Choose the correct passive form: 'He is writing a report.'", opts: ["A report was written by him.", "A report is being written by him.", "A report has been written by him.", "A report will be written by him."], co: "CO1,L2", vertical: true },
  { q: 42, text: "Change to active: 'A new bridge is being built by the government.'", opts: ["The government builds a new bridge.", "The government built a new bridge.", "The government is building a new bridge.", "The government has built a new bridge."], co: "CO1,L2", vertical: true },
  { q: 43, text: "Which is correctly changed to passive? 'Did they invite you to the party?'", opts: ["Were you invited to the party by them?", "Are you invited to the party?", "Was you invited to the party?", "Had you been invited to the party?"], co: "CO3,L3", vertical: true },
  { q: 44, text: "Identify the correct passive sentence:", opts: ["She was given the prize by jury.", "She was given the prize by the jury.", "She has given the prize by the jury.", "She is being given prize by the jury."], co: "CO1,L3", vertical: true },

  // PHRASAL VERBS (Q45–Q54)
  { q: 45, text: "Choose the correct meaning of 'break down' in: 'Her car broke down on the highway.'", opts: ["to improve suddenly", "to stop working", "to slow down", "to speed up"], co: "CO2,L3" },
  { q: 46, text: "Fill in the blank: Please ____ the form before submitting it. (complete carefully)", opts: ["fill up", "fill in", "fill on", "fill out"], co: "CO3,L3" },
  { q: 47, text: "Choose the correct phrasal verb: She could not ____ the sadness after losing her job.", opts: ["get over", "get in", "get up", "get by"], co: "CO3,L3" },
  { q: 48, text: "Choose the meaning of 'bring up' in: 'He was brought up by his grandparents.'", opts: ["to carry upstairs", "to raise or rear a child", "to mention", "to throw up"], co: "CO2,L3" },
  { q: 49, text: "Which sentence uses 'run out' correctly?", opts: ["We ran out of sugar.", "We ran out sugar of.", "We ran sugar out.", "We out ran of sugar."], co: "CO3,L3", vertical: true },
  { q: 50, text: "Choose the correct phrasal verb: The meeting was ____ because of the storm.", opts: ["called off", "called in", "called by", "called up"], co: "CO3,L3" },
  { q: 51, text: "What does 'turn down' mean in: 'She turned down the job offer.'?", opts: ["to accept", "to refuse", "to delay", "to forward"], co: "CO2,L3" },
  { q: 52, text: "Choose the correct meaning: 'He looks after his ailing parents.'", opts: ["to look behind", "to take care of", "to search for", "to look for"], co: "CO2,L3" },
  { q: 53, text: "Fill in: I need to ____ a plan before the meeting. (think of)", opts: ["come up with", "come down with", "come across", "come out of"], co: "CO3,L3" },
  { q: 54, text: "Choose the correct meaning of 'give in':", opts: ["to distribute", "to surrender or yield", "to donate", "to put away"], co: "CO2,L3" },

  // SENTENCE BUILDING / GRAMMAR (Q55–Q60)
  { q: 55, text: "Choose the correct sentence:", opts: ["Neither he nor she are coming.", "Neither he nor she is coming.", "Neither he nor she were coming.", "Neither he nor she come."], co: "CO1,L3", vertical: true },
  { q: 56, text: "Which sentence uses the comparative degree correctly?", opts: ["She is more smarter than her sister.", "She is smartest than her sister.", "She is smarter than her sister.", "She is smart than her sister."], co: "CO1,L3", vertical: true },
  { q: 57, text: "Identify the correctly punctuated sentence:", opts: ["She said, I am fine.", "She said, \"I am fine.\"", "She said I am fine.", "She said 'I am fine."], co: "CO1,L5", vertical: true },
  { q: 58, text: "Choose the correct sentence using 'too … to':", opts: ["She is too tired for work.", "She is too tired that she cannot work.", "She is too tired to work.", "She is too tired as she cannot work."], co: "CO2,L3", vertical: true },
  { q: 59, text: "Select the sentence with correct subject-verb agreement:", opts: ["The news are shocking.", "The news is shocking.", "The news were shocking.", "The news have shocked."], co: "CO1,L3", vertical: true },
  { q: 60, text: "Choose the sentence that is grammatically correct:", opts: ["He has gave his best in the exam.", "He have given his best in the exam.", "He has given his best in the exam.", "He had gave his best in the exam."], co: "CO1,L3", vertical: true },
];

// Paper 2 answers
const paper2Answers = [
  { q: 1, ans: "(c)", explanation: "'Fluently' modifies the verb 'speaks' — it tells how she speaks, making it an adverb." },
  { q: 2, ans: "(d)", explanation: "'Honesty' and 'policy' are both nouns. (d) correctly identifies both." },
  { q: 3, ans: "(b)", explanation: "'Reads' takes a direct object 'novel', making it a transitive verb." },
  { q: 4, ans: "(c)", explanation: "'On' shows the relationship between 'book' and 'shelf' — it is a preposition." },
  { q: 5, ans: "(b)", explanation: "'But' joins two contrasting clauses — it is a coordinating conjunction." },
  { q: 6, ans: "(c)", explanation: "When 'neither…nor' is used, the verb agrees with the nearer noun. 'Employees' (plural) takes 'were'." },
  { q: 7, ans: "(b)", explanation: "'Golden' and 'friendly' both modify nouns — 'golden' modifies retriever, 'friendly' modifies dog." },
  { q: 8, ans: "(b)", explanation: "'Theirs' is the correct possessive pronoun form used independently (without a following noun)." },
  { q: 9, ans: "(c)", explanation: "'Have been studying' = has/have + been + V-ing = Present Perfect Continuous." },
  { q: 10, ans: "(a)", explanation: "Interjections express sudden emotion. 'Oh!' at the start of a sentence is the correct usage." },
  { q: 11, ans: "(b)", explanation: "'Himself' is the reflexive pronoun for 'he' — used when the subject and object are the same person." },
  { q: 12, ans: "(b)", explanation: "'Courage' is an abstract noun — it names a quality or concept that cannot be physically touched." },
  { q: 13, ans: "(c)", explanation: "'Seem' links the subject to a subject complement without expressing action — it is a linking verb." },
  { q: 14, ans: "(a)", explanation: "'That' pointing to a specific house is a demonstrative pronoun used independently." },
  { q: 15, ans: "(a)", explanation: "When one past action was completed before another, use Past Perfect ('had finished') for the earlier action." },
  { q: 16, ans: "(c)", explanation: "Simple Future uses 'will/shall + base verb'. 'Shall complete…tomorrow' fits this pattern." },
  { q: 17, ans: "(c)", explanation: "An ongoing action from a point in the past until now uses Present Perfect Continuous." },
  { q: 18, ans: "(c)", explanation: "The departure (leaving) occurred before the arrival at the station — Past Perfect is required." },
  { q: 19, ans: "(d)", explanation: "'Are playing' = am/is/are + V-ing = Present Continuous tense." },
  { q: 20, ans: "(d)", explanation: "Homework was completed before father came home — the earlier completed action needs Past Perfect." },
  { q: 21, ans: "(c)", explanation: "'Has sung' = has/have + past participle = Present Perfect tense." },
  { q: 22, ans: "(a)", explanation: "A habitual or routine action is expressed using Simple Present tense ('play')." },
  { q: 23, ans: "(b)", explanation: "'Will have completed' indicates an action that will be finished by a future point — Future Perfect." },
  { q: 24, ans: "(c)", explanation: "In indirect speech, the present 'is' shifts to past 'was' after a past reporting verb ('told')." },
  { q: 25, ans: "(a)", explanation: "'Had been waiting' = had been + V-ing = Past Perfect Continuous, showing duration before a past point." },
  { q: 26, ans: "(b)", explanation: "'Since morning' with a present-time perspective uses Present Perfect — 'have not eaten'." },
  { q: 27, ans: "(c)", explanation: "'Wrote' is the Simple Past form of 'write', used for a completed action in the past." },
  { q: 28, ans: "(c)", explanation: "Definite article 'The' is used with unique geographical features like mountain ranges." },
  { q: 29, ans: "(b)", explanation: "'Honest' starts with a vowel sound /ɒ/, so 'an' is used before it." },
  { q: 30, ans: "(a)", explanation: "'European' starts with a consonant sound /j/ (yoo-), so 'a European' is correct." },
  { q: 31, ans: "(c)", explanation: "'One-eyed' starts with consonant sound /w/, so 'a'. 'Market' is a specific known place, so 'the'." },
  { q: 32, ans: "(a)", explanation: "Abstract nouns used in a general sense take no article. 'Life' here is general, 'a wonderful gift' is correct." },
  { q: 33, ans: "(b)", explanation: "'Umbrella' begins with a vowel sound /ʌ/, requiring the article 'an'." },
  { q: 34, ans: "(c)", explanation: "Unique objects like celestial bodies (sun, moon, earth) take the definite article 'the'." },
  { q: 35, ans: "(b)", explanation: "Present Perfect active 'has stolen' becomes passive 'has been stolen'. Agent 'someone' is dropped." },
  { q: 36, ans: "(b)", explanation: "Passive structure: subject + was/were + past participle + by + agent." },
  { q: 37, ans: "(c)", explanation: "Past Simple passive 'was written by her' reverts to active: subject 'she' + simple past 'wrote'." },
  { q: 38, ans: "(a)", explanation: "Future Simple passive: will + be + past participle. 'Will be announced' is correct." },
  { q: 39, ans: "(b)", explanation: "'He was reading a book' has a subject performing the action — this is active voice." },
  { q: 40, ans: "(a)", explanation: "Simple Present passive: subject + is/are + past participle. 'English is spoken' is correct." },
  { q: 41, ans: "(b)", explanation: "Present Continuous passive: is/are + being + past participle. 'Is being written' is correct." },
  { q: 42, ans: "(c)", explanation: "Present Continuous passive 'is being built' reverts to active: 'The government is building'." },
  { q: 43, ans: "(a)", explanation: "Passive of Past Simple question: Were + subject + past participle + by + agent?" },
  { q: 44, ans: "(b)", explanation: "'By the jury' requires the definite article 'the'. Option (b) is complete and grammatically correct." },
  { q: 45, ans: "(b)", explanation: "'Break down' means to stop functioning — commonly used for machines, cars, or systems." },
  { q: 46, ans: "(b)", explanation: "'Fill in' means to complete a form or document by writing required information in spaces." },
  { q: 47, ans: "(a)", explanation: "'Get over' means to recover from something difficult like illness, grief, or disappointment." },
  { q: 48, ans: "(b)", explanation: "'Bring up' means to raise or rear a child from childhood." },
  { q: 49, ans: "(a)", explanation: "'Run out of' is the correct form meaning to exhaust a supply of something." },
  { q: 50, ans: "(a)", explanation: "'Call off' means to cancel a planned event — the meeting was cancelled due to the storm." },
  { q: 51, ans: "(b)", explanation: "'Turn down' means to reject or refuse an offer, invitation, or request." },
  { q: 52, ans: "(b)", explanation: "'Look after' means to take care of or tend to someone — caring for ailing parents." },
  { q: 53, ans: "(a)", explanation: "'Come up with' means to think of or devise something — a plan, idea, or solution." },
  { q: 54, ans: "(b)", explanation: "'Give in' means to surrender, yield, or stop resisting pressure." },
  { q: 55, ans: "(b)", explanation: "With 'neither…nor', the verb agrees with the subject closest to it. 'She' (singular) takes 'is'." },
  { q: 56, ans: "(c)", explanation: "Comparative of 'smart' is 'smarter' — double comparison ('more smarter') is incorrect." },
  { q: 57, ans: "(b)", explanation: "Quoted speech requires double quotation marks around the spoken words." },
  { q: 58, ans: "(c)", explanation: "'Too + adjective + to + infinitive' is the correct structure to show inability due to excess." },
  { q: 59, ans: "(b)", explanation: "'News' is an uncountable singular noun — it always takes a singular verb ('is')." },
  { q: 60, ans: "(c)", explanation: "'Has given' is the correct Present Perfect form; 'gave' (not 'gave') follows 'has'." },
];

// ─── PAPER 3 QUESTIONS ──────────────────────────────────────────────────────

const paper3Questions = [
  // PARTS OF SPEECH (Q1–Q14)
  { q: 1, text: "Identify the verb in the sentence: 'The scientist discovered a new planet.'", opts: ["scientist", "discovered", "new", "planet"], co: "CO1,L1" },
  { q: 2, text: "Which of the following words is an abstract noun?", opts: ["Mountain", "Wisdom", "Table", "River"], co: "CO1,L1" },
  { q: 3, text: "Choose the correct relative pronoun: The girl ____ won the prize was my classmate.", opts: ["whom", "whose", "which", "who"], co: "CO1,L3" },
  { q: 4, text: "Identify the adverb in: 'She smiled warmly at the guests.'", opts: ["smiled", "warmly", "guests", "at"], co: "CO1,L1" },
  { q: 5, text: "Which sentence uses a collective noun correctly?", opts: ["A flock of birds are flying.", "A flock of birds is flying.", "A flock of birds were flying.", "A flock of bird is flying."], co: "CO1,L3", vertical: true },
  { q: 6, text: "Choose the correct option: He is the person ____ I told you about.", opts: ["who", "whom", "whose", "which"], co: "CO1,L3" },
  { q: 7, text: "Which of the following is a proper noun?", opts: ["city", "river", "Amazon", "book"], co: "CO1,L1" },
  { q: 8, text: "Identify the conjunction: 'She worked hard, yet she failed.'", opts: ["worked", "hard", "yet", "failed"], co: "CO1,L1" },
  { q: 9, text: "What part of speech is 'beautiful' in: 'She has a beautiful smile.'?", opts: ["Adverb", "Noun", "Adjective", "Verb"], co: "CO1,L1" },
  { q: 10, text: "Which sentence uses an interrogative pronoun?", opts: ["This is my pen.", "He is clever.", "Who told you this?", "They went home."], co: "CO1,L2", vertical: true },
  { q: 11, text: "Choose the correct indefinite pronoun: ____ of the students passed the test.", opts: ["None", "Neither", "Nor", "Either"], co: "CO1,L3" },
  { q: 12, text: "Identify the preposition: 'She hid the letter beneath the mattress.'", opts: ["hid", "letter", "beneath", "mattress"], co: "CO1,L1" },
  { q: 13, text: "Which word is a modal auxiliary?", opts: ["go", "run", "should", "write"], co: "CO1,L1" },
  { q: 14, text: "Choose the correct option for subject-verb agreement: The team ____ its best game last night.", opts: ["played", "play", "playing", "have played"], co: "CO1,L3" },

  // TENSES (Q15–Q27)
  { q: 15, text: "Choose the correct form: Right now, the chef ____ a special dish.", opts: ["cooks", "cooked", "is cooking", "has cooked"], co: "CO1,L2" },
  { q: 16, text: "Fill in: I ____ this movie twice before.", opts: ["see", "saw", "have seen", "had seen"], co: "CO1,L3" },
  { q: 17, text: "Which sentence uses Future Continuous correctly?", opts: ["She will study at 8 PM.", "She will be studying at 8 PM.", "She studies at 8 PM.", "She would study at 8 PM."], co: "CO1,L2", vertical: true },
  { q: 18, text: "Choose the correct tense: He ____ to Paris many times.", opts: ["went", "goes", "has gone", "is going"], co: "CO1,L3" },
  { q: 19, text: "Fill in: By the time you read this, she ____ the city.", opts: ["will leave", "leaves", "will have left", "is leaving"], co: "CO2,L3" },
  { q: 20, text: "Identify the error: 'They are knowing the answer.'", opts: ["are — should be 'were'", "'knowing' should be 'know' — stative verb", "answer — should be 'answers'", "No error"], co: "CO1,L6", vertical: true },
  { q: 21, text: "Choose the correct option: She ____ for the interview since morning.", opts: ["prepares", "is preparing", "was preparing", "has been preparing"], co: "CO1,L3" },
  { q: 22, text: "Which tense is used for scientific facts and universal truths?", opts: ["Past Simple", "Present Continuous", "Present Simple", "Future Simple"], co: "CO2,L2" },
  { q: 23, text: "Fill in: He said that he ____ complete the work the next day.", opts: ["will", "would", "shall", "had"], co: "CO2,L3" },
  { q: 24, text: "Choose the correct sentence:", opts: ["She already has left.", "She has left already.", "She had left already.", "Both (b) and (c)"], co: "CO1,L3" },
  { q: 25, text: "Identify the tense: 'We were watching TV when the power went out.'", opts: ["Past Simple", "Past Continuous", "Past Perfect", "Present Continuous"], co: "CO1,L1" },
  { q: 26, text: "Fill in: They ____ (work) on this project for three months when the funding stopped.", opts: ["worked", "have been working", "had been working", "were working"], co: "CO2,L3" },
  { q: 27, text: "Choose the correct form for conditional: If it ____, we will cancel the trip.", opts: ["rained", "rains", "has rained", "had rained"], co: "CO3,L3" },

  // ARTICLES (Q28–Q34)
  { q: 28, text: "Fill in: He is ____ MA in English Literature.", opts: ["a", "an", "the", "no article"], co: "CO1,L3" },
  { q: 29, text: "Choose the correct sentence with articles:", opts: ["He is an university student.", "He is a university student.", "He is the university student.", "He is university student."], co: "CO1,L3", vertical: true },
  { q: 30, text: "Fill in: ____ poor need our help. (as a general class)", opts: ["A", "An", "The", "No article"], co: "CO1,L5" },
  { q: 31, text: "Which sentence uses articles correctly?", opts: ["She plays the piano every morning.", "She plays a piano every morning.", "She plays piano every morning.", "She plays an piano every morning."], co: "CO2,L5", vertical: true },
  { q: 32, text: "Fill in: We had ____ dinner at ____ nice French restaurant.", opts: ["a/a", "the/a", "no article/a", "a/the"], co: "CO1,L3" },
  { q: 33, text: "Choose the correct article use: It was ____ best performance I have ever seen.", opts: ["a", "an", "the", "no article"], co: "CO1,L3" },
  { q: 34, text: "Fill in: ____ gold is a precious metal. (general statement)", opts: ["A", "An", "The", "No article"], co: "CO1,L5" },

  // ACTIVE-PASSIVE VOICE (Q35–Q44)
  { q: 35, text: "Change to passive: 'The children broke the window.'", opts: ["The window is broken by the children.", "The window was broken by the children.", "The window has been broken by the children.", "The window had been broken by the children."], co: "CO1,L2", vertical: true },
  { q: 36, text: "Identify the sentence in passive voice:", opts: ["He mailed the letter.", "The letter is being mailed by him.", "She posted the parcel.", "They announced the results."], co: "CO1,L2", vertical: true },
  { q: 37, text: "Change to active: 'Mistakes are often made by beginners.'", opts: ["Beginners often make mistakes.", "Beginners often made mistakes.", "Beginners are often making mistakes.", "Beginners will often make mistakes."], co: "CO1,L2", vertical: true },
  { q: 38, text: "Change to passive: 'He is teaching the students.'", opts: ["The students are taught by him.", "The students are being taught by him.", "The students were being taught by him.", "The students have been taught by him."], co: "CO1,L2", vertical: true },
  { q: 39, text: "Choose the correct passive form of: 'They had built the bridge before the floods.'", opts: ["The bridge was built before the floods.", "The bridge had been built before the floods.", "The bridge has been built before the floods.", "The bridge is built before the floods."], co: "CO1,L2", vertical: true },
  { q: 40, text: "Which transformation is correct? 'Do they serve food here?'", opts: ["Is food served here?", "Was food served here?", "Is food being served here?", "Food is served here?"], co: "CO3,L3", vertical: true },
  { q: 41, text: "Change to passive: 'She was making tea when I arrived.'", opts: ["Tea was being made by her when I arrived.", "Tea is made by her when I arrived.", "Tea had been made by her when I arrived.", "Tea was made by her when I arrived."], co: "CO1,L2", vertical: true },
  { q: 42, text: "Choose the correct option: 'The news ____ announced on television.' (passive)", opts: ["is being", "was being", "was", "were"], co: "CO1,L3" },
  { q: 43, text: "Change to active: 'The project will have been completed by July.'", opts: ["They complete the project by July.", "They will complete the project by July.", "They will have completed the project by July.", "They had completed the project by July."], co: "CO1,L2", vertical: true },
  { q: 44, text: "Identify the correctly formed passive sentence:", opts: ["The cake was ate by him.", "The cake was eaten by him.", "The cake has ate by him.", "The cake is eaten him."], co: "CO1,L3", vertical: true },

  // PHRASAL VERBS (Q45–Q54)
  { q: 45, text: "Choose the correct meaning of 'set up' in: 'They set up a new business.'", opts: ["to destroy", "to establish or start", "to move away", "to fail"], co: "CO2,L3" },
  { q: 46, text: "Fill in: He had to ____ his plans because of bad weather.", opts: ["put off", "put up", "put on", "put by"], co: "CO3,L3" },
  { q: 47, text: "What does 'come across' mean?", opts: ["to move forward", "to discover or find by chance", "to come from a direction", "to approach angrily"], co: "CO2,L3" },
  { q: 48, text: "Choose the correct sentence using 'take off':", opts: ["The plane took off at 6 AM.", "The plane took off the runway.", "The plane took the off at 6 AM.", "The plane was took off at 6 AM."], co: "CO3,L3", vertical: true },
  { q: 49, text: "Choose the meaning of 'hold on':", opts: ["to grasp tightly only", "to wait briefly", "to release something", "to push forward"], co: "CO2,L3" },
  { q: 50, text: "Fill in: She found it hard to ____ the loss of her pet.", opts: ["get over", "get in", "get off", "get around"], co: "CO3,L3" },
  { q: 51, text: "What does 'show up' mean?", opts: ["to perform on stage", "to arrive or appear", "to display something", "to pretend"], co: "CO2,L3" },
  { q: 52, text: "Choose the correct phrasal verb: I ____ an old friend at the supermarket yesterday.", opts: ["came across", "came over", "came up", "came in"], co: "CO3,L3" },
  { q: 53, text: "What does 'back out' mean in: 'He backed out of the deal at the last minute.'?", opts: ["to reverse a vehicle", "to withdraw from an agreement", "to support strongly", "to enter from behind"], co: "CO2,L3" },
  { q: 54, text: "Fill in: The company decided to ____ a new product next month. (introduce/launch)", opts: ["bring out", "bring up", "bring about", "bring in"], co: "CO3,L3" },

  // SENTENCE BUILDING / GRAMMAR (Q55–Q60)
  { q: 55, text: "Choose the correct sentence using reported speech: He said, 'I will help you.' →", opts: ["He said that he will help me.", "He said that he would help me.", "He said that he helps me.", "He said that he helped me."], co: "CO2,L3", vertical: true },
  { q: 56, text: "Identify the correct conditional sentence:", opts: ["If I knew the answer, I would told you.", "If I know the answer, I would tell you.", "If I knew the answer, I would tell you.", "If I had known the answer, I would told you."], co: "CO3,L3", vertical: true },
  { q: 57, text: "Choose the sentence with correct use of 'although':", opts: ["Although it was raining, but we went out.", "Although it was raining, we went out.", "Although it was raining, yet we went out.", "Although it was raining, so we stayed in."], co: "CO1,L5", vertical: true },
  { q: 58, text: "Which sentence uses the superlative degree correctly?", opts: ["Mount Everest is most high mountain.", "Mount Everest is the most highest mountain.", "Mount Everest is the highest mountain.", "Mount Everest is highest mountain."], co: "CO1,L3", vertical: true },
  { q: 59, text: "Identify the sentence with a dangling modifier:", opts: ["Running fast, he caught the bus.", "Having finished lunch, the plates were cleared.", "After studying hard, she passed the test.", "Feeling tired, he went to bed early."], co: "CO1,L5", vertical: true },
  { q: 60, text: "Choose the sentence where 'used to' is used correctly:", opts: ["He used to plays football when he was young.", "He used to played football when he was young.", "He used to play football when he was young.", "He use to play football when he was young."], co: "CO2,L4", vertical: true },
];

// Paper 3 answers
const paper3Answers = [
  { q: 1, ans: "(b)", explanation: "'Discovered' is the action word — the verb — in this sentence, expressing what the scientist did." },
  { q: 2, ans: "(b)", explanation: "'Wisdom' names a quality that cannot be seen or touched — it is an abstract noun." },
  { q: 3, ans: "(d)", explanation: "'Who' is used for people as the subject of a relative clause — 'the girl who won'." },
  { q: 4, ans: "(b)", explanation: "'Warmly' modifies the verb 'smiled', describing how she smiled — it is an adverb." },
  { q: 5, ans: "(b)", explanation: "Collective nouns like 'flock' are treated as singular in British/academic English — 'is flying'." },
  { q: 6, ans: "(b)", explanation: "'Whom' is used as the object of the preposition 'about' — 'the person about whom I told you'." },
  { q: 7, ans: "(c)", explanation: "'Amazon' is the name of a specific river — it is a proper noun, capitalised." },
  { q: 8, ans: "(c)", explanation: "'Yet' joins two contrasting ideas — it is an adversative coordinating conjunction." },
  { q: 9, ans: "(c)", explanation: "'Beautiful' describes the noun 'smile' — it is an adjective." },
  { q: 10, ans: "(c)", explanation: "'Who' in a question asking for identification is an interrogative pronoun." },
  { q: 11, ans: "(a)", explanation: "'None' is an indefinite pronoun meaning 'not one' and can refer to multiple people." },
  { q: 12, ans: "(c)", explanation: "'Beneath' shows the position of the letter relative to the mattress — it is a preposition." },
  { q: 13, ans: "(c)", explanation: "'Should' is a modal auxiliary used to express obligation or advice." },
  { q: 14, ans: "(a)", explanation: "'Team' is a collective noun, treated as singular in formal writing — 'played' (past simple) is correct." },
  { q: 15, ans: "(c)", explanation: "'Right now' signals an action in progress — Present Continuous 'is cooking' is correct." },
  { q: 16, ans: "(c)", explanation: "'Before' with a present-time reference indicates an experience — Present Perfect 'have seen' is correct." },
  { q: 17, ans: "(b)", explanation: "Future Continuous: will + be + V-ing. 'Will be studying at 8 PM' correctly describes an ongoing future action." },
  { q: 18, ans: "(c)", explanation: "Repeated experience without specifying time uses Present Perfect — 'has gone'." },
  { q: 19, ans: "(c)", explanation: "'By the time you read' = future reference; the action will already be complete then — Future Perfect." },
  { q: 20, ans: "(b)", explanation: "'Know' is a stative verb and cannot be used in continuous forms — 'They know the answer' is correct." },
  { q: 21, ans: "(d)", explanation: "'Since morning' with an ongoing activity uses Present Perfect Continuous — 'has been preparing'." },
  { q: 22, ans: "(c)", explanation: "Scientific facts and universal truths always use the Simple Present tense." },
  { q: 23, ans: "(b)", explanation: "In reported speech, 'will' changes to 'would' after a past reporting verb." },
  { q: 24, ans: "(d)", explanation: "Both 'She has left already' (Present Perfect) and 'She had left already' (Past Perfect) are grammatically correct." },
  { q: 25, ans: "(b)", explanation: "Two simultaneous past actions where one was ongoing uses Past Continuous — 'were watching'." },
  { q: 26, ans: "(c)", explanation: "An action ongoing over a period before another past event uses Past Perfect Continuous." },
  { q: 27, ans: "(b)", explanation: "Type 1 (real) conditional: 'If + Simple Present, will + base verb' — 'If it rains, we will cancel'." },
  { q: 28, ans: "(b)", explanation: "'MA' is pronounced 'em-ay', starting with a vowel sound — 'an MA' is correct." },
  { q: 29, ans: "(b)", explanation: "'University' begins with a consonant sound /j/ (yoo-), so 'a university' is correct." },
  { q: 30, ans: "(c)", explanation: "'The + adjective' refers to a class of people — 'the poor' means poor people in general." },
  { q: 31, ans: "(a)", explanation: "Musical instruments use the definite article — 'plays the piano' is the standard expression." },
  { q: 32, ans: "(c)", explanation: "'Dinner' (uncountable, mealtime) takes no article; 'restaurant' (first mention, countable) takes 'a'." },
  { q: 33, ans: "(c)", explanation: "Superlatives require 'the' — 'the best performance' is correct." },
  { q: 34, ans: "(d)", explanation: "Metals used in general statements do not take any article — 'Gold is a precious metal'." },
  { q: 35, ans: "(b)", explanation: "Past Simple active becomes Past Simple passive: 'was/were + past participle'." },
  { q: 36, ans: "(b)", explanation: "'Is being mailed by him' has the passive structure: is/are + being + past participle." },
  { q: 37, ans: "(a)", explanation: "Simple Present passive 'are made by beginners' → active: 'Beginners often make mistakes'." },
  { q: 38, ans: "(b)", explanation: "Present Continuous passive: is/are + being + past participle — 'are being taught'." },
  { q: 39, ans: "(b)", explanation: "Past Perfect active becomes Past Perfect passive: had + been + past participle." },
  { q: 40, ans: "(a)", explanation: "Yes/no question passive (Present Simple): 'Is + subject + past participle + here?'" },
  { q: 41, ans: "(a)", explanation: "Past Continuous passive: was/were + being + past participle — 'was being made by her'." },
  { q: 42, ans: "(c)", explanation: "'News' is singular; this is a Simple Past passive — 'was announced' is correct." },
  { q: 43, ans: "(c)", explanation: "Future Perfect passive 'will have been completed' → active: 'They will have completed the project by July'." },
  { q: 44, ans: "(b)", explanation: "The correct past participle of 'eat' is 'eaten'. Option (b) correctly uses 'was eaten by him'." },
  { q: 45, ans: "(b)", explanation: "'Set up' means to establish, start, or create something new — a business, system, or organisation." },
  { q: 46, ans: "(a)", explanation: "'Put off' means to postpone or delay — he had to delay his plans due to bad weather." },
  { q: 47, ans: "(b)", explanation: "'Come across' means to find or encounter something/someone unexpectedly or by chance." },
  { q: 48, ans: "(a)", explanation: "'Take off' (intransitive) means to leave the ground — 'The plane took off at 6 AM' is correct." },
  { q: 49, ans: "(b)", explanation: "'Hold on' primarily means to wait for a short time — often used on phone or in conversations." },
  { q: 50, ans: "(a)", explanation: "'Get over' means to recover emotionally from a difficult experience like loss." },
  { q: 51, ans: "(b)", explanation: "'Show up' means to arrive or appear, especially when expected — 'She finally showed up'." },
  { q: 52, ans: "(a)", explanation: "'Came across' means to meet or find unexpectedly — correct for a chance meeting at a supermarket." },
  { q: 53, ans: "(b)", explanation: "'Back out' means to withdraw from an agreement, deal, or commitment." },
  { q: 54, ans: "(a)", explanation: "'Bring out' means to launch, publish, or introduce a new product to the market." },
  { q: 55, ans: "(b)", explanation: "Reported speech shifts 'will' to 'would' after a past reporting verb. Pronoun 'you' changes to 'me'." },
  { q: 56, ans: "(c)", explanation: "Type 2 conditional: 'If + Past Simple, would + base verb' — 'If I knew, I would tell you'." },
  { q: 57, ans: "(b)", explanation: "'Although' is already a concessive conjunction — it must NOT be followed by 'but', 'yet', or 'so'." },
  { q: 58, ans: "(c)", explanation: "Superlatives require 'the' before them. The correct form is 'the highest' — not 'most highest'." },
  { q: 59, ans: "(b)", explanation: "'Having finished lunch' modifies no logical subject — 'plates' didn't finish lunch. This is a dangling modifier." },
  { q: 60, ans: "(c)", explanation: "'Used to + base verb' is the correct structure for past habits — 'used to play'." },
];

// ─── BUILD QUESTION SECTION ─────────────────────────────────────────────────

function buildQuestions(questions) {
  const paras = [];
  questions.forEach(({ q, text, opts, co, vertical }) => {
    paras.push(leftPara([bold(`Q(${q})  `, 20), reg(text, 20)], 80, 20));
    if (vertical) {
      paras.push(leftPara([reg(`(a) ${opts[0]}`, 20)], 0, 10));
      paras.push(leftPara([reg(`(b) ${opts[1]}`, 20)], 0, 10));
      paras.push(leftPara([reg(`(c) ${opts[2]}`, 20)], 0, 10));
      paras.push(leftPara([reg(`(d) ${opts[3]}`, 20)], 0, 0));
    } else {
      paras.push(leftPara([
        reg(`(a) ${opts[0]}    `, 20),
        reg(`(b) ${opts[1]}    `, 20),
        reg(`(c) ${opts[2]}    `, 20),
        reg(`(d) ${opts[3]}`, 20),
      ], 0, 0));
    }
    paras.push(new Paragraph({
      alignment: AlignmentType.RIGHT,
      spacing: { before: 0, after: 60 },
      children: [italic(co, 18)]
    }));
  });
  return paras;
}

// ─── ASSEMBLE DOCUMENT ──────────────────────────────────────────────────────

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 20 } }
    }
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 }
        }
      },
      children: [
        // ── PAPER 2 ──────────────────────────────────────────────────────────
        ...buildHeader(2, "B", "24252PEL65190"),

        centerPara([bold("Q(1)", 20)], 40, 20),
        ...buildQuestions(paper2Questions),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 60 },
          children: [italic("--End of Question Paper--", 20)]
        }),

        // Page number footer marker
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 40 },
          children: [italic("Page 6 of 6", 18)]
        }),

        pageBreakPara(),

        // ── ANSWER KEY – PAPER 2 ─────────────────────────────────────────────
        centerPara([bold("ANSWER KEY — SAMPLE PAPER 2 (Code B)", 26)], 40, 20),
        centerPara([bold("Course: PEL121 — Communication Skills-I", 22)], 0, 60),
        answerKeyTable(paper2Answers),

        pageBreakPara(),

        // ── PAPER 3 ──────────────────────────────────────────────────────────
        ...buildHeader(3, "C", "24252PEL65191"),

        centerPara([bold("Q(1)", 20)], 40, 20),
        ...buildQuestions(paper3Questions),

        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 60, after: 60 },
          children: [italic("--End of Question Paper--", 20)]
        }),

        new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { before: 0, after: 40 },
          children: [italic("Page 6 of 6", 18)]
        }),

        pageBreakPara(),

        // ── ANSWER KEY – PAPER 3 ─────────────────────────────────────────────
        centerPara([bold("ANSWER KEY — SAMPLE PAPER 3 (Code C)", 26)], 40, 20),
        centerPara([bold("Course: PEL121 — Communication Skills-I", 22)], 0, 60),
        answerKeyTable(paper3Answers),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("/mnt/user-data/outputs/PEL121_Sample_Papers_2_and_3.docx", buffer);
  console.log("Done!");
});
