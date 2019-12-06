// On Ready function must be called in order for the add in to be registered.
let host: Office.HostType = undefined;
Office.onReady((info)=> {
    host = info.host;
    
    // Office.HostType.Word
    // Office.HostType.Outlook
});

// /**
//  * log for testing
//  * @param message 
//  */
// function log(message: string) {

//     const logDiv = document.getElementById("log") as HTMLDivElement;

//     logDiv.innerHTML = logDiv.innerHTML + `<p>${message}</p>`;

// }

// function runF(f: () => void) {
//     try {
//         f();
//     } catch(e) {
//         log(`ERROR: ${e}`);
//     }
// }

/**
 * get the currenty selected text in the word document
 */
async function getSelectedTextWord() {
    // Gets the current selection
 

    let selectedText = "";
    await Word.run(async (context) => {
        // would be neat to be able to highlight what was selected but this turns our to be pretty difficult
        const range = context.document.getSelection();

        range.load("text");

        await context.sync();
        selectedText = range.text;
    });

    return selectedText;
}

async function getSelectedTextOutlook() {
    return new Promise<string>((resolve) => {
        Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, (result: Office.AsyncResult<any>)=> {

            const value = result.value;

            const text = value.data;

            const selectedText = text? text : "";

            resolve(selectedText);
        });
    });
}


async function getSelectedText() {

    // Get text using the method specific to the host
    if (host === Office.HostType.Word) {
        return await getSelectedTextWord();
    }
    
    if (host === Office.HostType.Outlook) {
        return await getSelectedTextOutlook();
    }

    console.log("Unsupported Host");
    return "Unsupported";
}




async function run() {
    writeClear();
    const selectedText = await getSelectedText();
    
    // trim any text to remove starting and ending spaces
    const word = selectedText.trim();

    // word error cases
    if (word === "") {
        writeMessage("Highlight a word to rhyme!");
        return;
    }

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
const dictionary: Map<string, string[]> = new Map();

async function getWordRhymes(word): Promise<string[]> {

    let words = dictionary.get(word);

    if (words === undefined) {
        words = await getWordRhymesFromDatamuse(word);
        // store words so that datamuse is not requeried.
        dictionary.set(word, words);
    }

    return words;
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
    document.getElementById("message").innerText = message;
}

function writeClear() {
    writeWord("");
    writeRhyme("");
    writeMessage("");
}