// On Ready function must be called in order for the add in to be registered.
Office.onReady();

async function run() {
    findRhyme();
}

function getRandomWhole(max) {
    return Math.floor(Math.random() * Math.floor(max));
}

function getRandomIndex(array) {
    const max = array.length;
    const index = getRandomWhole(max);
    return array[index];
}

function hasWhiteSpace(s) {
    return /\s/g.test(s);
}

async function getWordRhymes(word) {
    const query = "https://api.datamuse.com/words?rel_rhy=" + word;

    const response = await fetch(query);
    const o = await response.json();

    // get the list of words
    const list = o.map((item) => item.word).filter((word) => !hasWhiteSpace(word));

    return list;
}

async function findRhyme() {
    const word = await getSelectedText();

    const rhymes = await getWordRhymes(word);

    if (rhymes.length === 0) {
        writeRhyme("No Rhymes uncovered, select another word to discover.");
    } else {
        const rhyme = getRandomIndex(rhymes);
        writeRhyme(rhyme);
    }
}

async function getSelectedText() {
    // Gets the current selection and changes the font color to red.

    let selectedText = "";
    await Word.run(async (context) => {
        // would be neat to be able to highlight what was selected but this turns our to be pretty difficult
        const range = context.document.getSelection();

        range.load("text");

        await context.sync();
        selectedText = range.text;
    });

    return selectedText.trim();
}

function writeRhyme(word) {
    document.getElementById("rhyme-id").innerText = word;
}
