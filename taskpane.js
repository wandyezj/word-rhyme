// On Ready function must be called in order for the add in to be registered.
Office.onReady();
async function run() {
    const word = await getSelectedText();
    // word error cases
    if (hasWhiteSpace(word)) {
        writeMessage("A space was selected, multiple words, thus rejected.");
        return;
    }
    if (word.length > 28) {
        writeMessage("The longest non-contrived and nontechnical word is antidisestablishmentarianism.");
        return;
    }
    // show the word selected
    writeWord(word);
    const rhymes = await getWordRhymes(word);
    if (rhymes.length === 0) {
        writeMessage("No Rhymes to uncover, select another word to discover.");
        return;
    }
    // get a random rhyme
    const rhyme = getRandomIndex(rhymes);
    writeRhyme(rhyme);
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
async function getWordRhymesFromDatamuse(word) {
    const query = "https://api.datamuse.com/words?rel_rhy=" + word;
    const response = await fetch(query);
    const o = await response.json();
    // get the list of words
    const list = o.map((item) => item.word).filter((word) => !hasWhiteSpace(word));
    return list;
}
// dictionary to keep track of previous queries to reduce queries to datamuse
const dictionary = new Map();
async function getWordRhymes(word) {
    let words = dictionary.get(word);
    if (words === undefined) {
        words = await getWordRhymesFromDatamuse(word);
        // store words so that datamuse is not requeried.
        dictionary.set(word, words);
    }
    return words;
}
/**
 * get teh currenty selected text in the word document
 */
async function getSelectedText() {
    // Gets the current selection
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
/**
 * write out the rhyming word
 * @param word
 */
function writeRhyme(rhyme) {
    document.getElementById("rhyme").innerText = rhyme;
}
function writeWord(word) {
    document.getElementById("word").innerText = word;
}
function writeMessage(message) {
    writeWord("");
    writeRhyme("");
    document.getElementById("message").innerText = message;
}
//# sourceMappingURL=taskpane.js.map