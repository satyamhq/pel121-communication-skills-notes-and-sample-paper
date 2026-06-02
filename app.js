// DOM Elements
const contentContainer = document.getElementById('content-container');
const pageTitle = document.getElementById('page-title');
const navLinks = document.querySelectorAll('.nav-link');
const themeBtn = document.getElementById('theme-btn');
const mobileMenuBtn = document.getElementById('mobile-menu-btn');
const sidebar = document.querySelector('.sidebar');

// Theme Toggle
let isDark = localStorage.getItem('theme') === 'dark';
if (isDark) document.body.setAttribute('data-theme', 'dark');

themeBtn.addEventListener('click', () => {
    isDark = !isDark;
    if (isDark) {
        document.body.setAttribute('data-theme', 'dark');
        themeBtn.innerHTML = '☀️ Light Mode';
        localStorage.setItem('theme', 'dark');
    } else {
        document.body.removeAttribute('data-theme');
        themeBtn.innerHTML = '🌙 Dark Mode';
        localStorage.setItem('theme', 'light');
    }
});
if(isDark) themeBtn.innerHTML = '☀️ Light Mode';

// Mobile Menu
mobileMenuBtn.addEventListener('click', () => {
    sidebar.classList.toggle('open');
});

// Routing
navLinks.forEach(link => {
    link.addEventListener('click', (e) => {
        e.preventDefault();
        // Update active class
        navLinks.forEach(l => l.classList.remove('active'));
        e.target.classList.add('active');
        
        // Hide sidebar on mobile
        if(window.innerWidth <= 768) sidebar.classList.remove('open');

        const target = e.target.getAttribute('data-target');
        pageTitle.innerText = e.target.innerText;
        
        renderContent(target);
    });
});

function renderContent(target) {
    contentContainer.innerHTML = '';
    // Small fade animation trick
    contentContainer.style.animation = 'none';
    contentContainer.offsetHeight; /* trigger reflow */
    contentContainer.style.animation = null;

    if (target.startsWith('unit')) {
        renderNote(target);
    } else if (target.startsWith('paper')) {
        renderPaper(target);
    }
}

function renderNote(unitId) {
    // We embedded notesHTML in notes_data.js. We need to extract the specific section.
    const parser = new DOMParser();
    const doc = parser.parseFromString(notesHTML, 'text/html');
    const section = doc.getElementById(unitId);
    
    if (section) {
        contentContainer.innerHTML = section.innerHTML;
    } else {
        contentContainer.innerHTML = '<p>Content not found.</p>';
    }
}

function renderPaper(paperId) {
    const data = courseData[paperId];
    let questions = [];
    let answers = [];
    
    if (paperId === 'paper1') {
        questions = courseData.paper1;
        // Paper 1 answers were not explicitly exported as an array, we'll try to extract from text or just not show explanations
        // Actually, looking at the data, paper 1 didn't have explicit answers array in extract.js.
        // For simplicity, we'll just render it as a reading material if answers missing.
    } else {
        questions = data.questions;
        answers = data.answers;
    }
    
    if(!questions || questions.length === 0) {
        contentContainer.innerHTML = '<p>No questions available for this paper.</p>';
        return;
    }

    const form = document.createElement('div');
    form.id = 'quiz-form';

    questions.forEach((q, index) => {
        const qBlock = document.createElement('div');
        qBlock.className = 'question-block';
        qBlock.id = `qblock-${q.q}`;
        
        const qText = document.createElement('div');
        qText.className = 'question-text';
        qText.innerText = `Q${q.q}. ${q.text}`;
        qBlock.appendChild(qText);

        const optionsList = document.createElement('div');
        optionsList.className = 'options-list';

        q.opts.forEach((optText, optIndex) => {
            const label = document.createElement('label');
            label.className = 'option-label';
            label.innerHTML = `
                <input type="radio" name="q${q.q}" value="${optIndex}">
                <span>${optText}</span>
            `;
            // Selection effect
            label.querySelector('input').addEventListener('change', (e) => {
                const siblings = label.parentElement.querySelectorAll('.option-label');
                siblings.forEach(s => s.style.borderColor = 'var(--border-color)');
                label.style.borderColor = 'var(--accent-color)';
            });
            optionsList.appendChild(label);
        });
        qBlock.appendChild(optionsList);

        const expBox = document.createElement('div');
        expBox.className = 'explanation-box';
        expBox.id = `exp-${q.q}`;
        qBlock.appendChild(expBox);

        form.appendChild(qBlock);
    });

    const submitBtn = document.createElement('button');
    submitBtn.className = 'submit-btn';
    submitBtn.innerText = 'Submit Test';
    
    const scoreDisplay = document.createElement('div');
    scoreDisplay.className = 'score-display';
    scoreDisplay.id = 'score-display';
    
    submitBtn.addEventListener('click', () => {
        let score = 0;
        let attempted = 0;

        questions.forEach((q) => {
            const selected = document.querySelector(`input[name="q${q.q}"]:checked`);
            const qBlock = document.getElementById(`qblock-${q.q}`);
            const expBox = document.getElementById(`exp-${q.q}`);
            
            // Find correct answer info
            let correctAnsObj = answers ? answers.find(a => a.q === q.q) : null;
            
            if (correctAnsObj) {
                // Correct ans is like "(c)"
                const correctChar = correctAnsObj.ans.replace(/[\(\)]/g, '').toLowerCase();
                const charToIndex = {'a':0, 'b':1, 'c':2, 'd':3};
                const correctIndex = charToIndex[correctChar];

                // Disable all inputs
                const inputs = qBlock.querySelectorAll('input');
                inputs.forEach(i => i.disabled = true);

                // Highlight correct option
                const labels = qBlock.querySelectorAll('.option-label');
                if (labels[correctIndex]) {
                    labels[correctIndex].classList.add('correct');
                }

                if (selected) {
                    attempted++;
                    const selectedIndex = parseInt(selected.value);
                    if (selectedIndex === correctIndex) {
                        score += 1;
                    } else {
                        score -= 0.25; // Negative marking as per notes
                        labels[selectedIndex].classList.add('wrong');
                    }
                }
                
                expBox.innerHTML = `<strong>Correct Answer: ${correctAnsObj.ans}</strong><br>${correctAnsObj.explanation}`;
                expBox.classList.add('visible');
            } else {
                // If no answers available (like Paper 1)
                expBox.innerHTML = "<em>Answer key not available for this question.</em>";
                expBox.classList.add('visible');
            }
        });

        submitBtn.style.display = 'none';
        scoreDisplay.innerText = `Your Score: ${score.toFixed(2)} / ${questions.length} (Attempted: ${attempted})`;
        scoreDisplay.style.display = 'block';
        window.scrollTo({top: 0, behavior: 'smooth'});
    });

    contentContainer.appendChild(scoreDisplay);
    contentContainer.appendChild(form);
    contentContainer.appendChild(submitBtn);
}

// Initial render
renderContent('unit1');
