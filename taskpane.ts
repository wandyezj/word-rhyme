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

    // need to wait for Office to be ready before the APIs can be called.
    await Office.onReady();

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

function buttonRhymeSelection(): HTMLButtonElement {
    return document.getElementById("button-rhyme-selection") as HTMLButtonElement;
}

/**
 * Is the rhyme function already running?
 */
let running = false;

let previousWord: string | undefined = undefined;

async function run() {
    if (running) {
        // Only allow one instance to run at a time.
        return;
    }
    
    // only allow a single run
    // since JavaScript is single threaded and functions run to completion there should never be a race condition
    running = true;
    
    clearRhyme();
    clearMessage();

    const selectedText = await getSelectedText();
    
    // trim any text to remove starting and ending spaces
    const word = selectedText.trim();
    let message: undefined | string = undefined;

    // Only modify what changes in the UI so there is no annoying refersh of the word from clearing then rewriting it.
    if (word !== previousWord) {
        clearWord();
    }

    // word error cases
    if (word === "") {
        message = "Highlight a word to rhyme!";

    } else if (hasWhiteSpace(word)) {
        message = "A space was selected, multiple words, thus rejected.";

    } else if (word.length > 28) {
        message = "The longest non-contrived and nontechnical word is antidisestablishmentarianism.";

    } else {

        // show the word selected
        writeWord(word);
        previousWord = word;

        let rhymes = [];
        const hasRhymes = wordRhymes(word);

        if (hasRhymes === undefined) {

            // only disable the buttonw when loading in case it takes a while
            // the disable changes the color of the button to indicate loading
            buttonRhymeSelection().disabled = true;
            
            try {
                rhymes = await getWordRhymes(word);
            } catch(e) {
                console.error(e);
                // set rhymes to jump past the message
                rhymes = ["An error! Oh my! Give another word a try."];
            }

            buttonRhymeSelection().disabled = false;
        } else {
            rhymes = hasRhymes;
        }

        
        if (rhymes.length === 0) {
            message = "No Rhymes to uncover, select another word to discover.";

        } else {
            // get a random rhyme
            const rhyme = getRandomIndex(rhymes);
            writeRhyme(rhyme);
        }
    }

    if (message !== undefined) {
        writeMessage(message);
    }

    // enable another run
    // since JavaScript is single threaded and functions run to completion there should never be a race condition
    running = false;
    
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

    const response = await fetch(query, {
        mode: 'cors'
    });

    const o = await response.json();

    // get the list of words
    const list = o.map((item) => item.word).filter((word) => !hasWhiteSpace(word));

    return list;
}

// dictionary to keep track of previous queries to reduce queries to datamuse
const dictionary: Map<string, string[]> = new Map();

function wordRhymes(word): string[] | undefined {
    let words = dictionary.get(word);
    return words;
}

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

    
    
}

function clearRhyme() {
    writeRhyme("");
}

function clearMessage() {
    writeMessage("");
}

function clearWord() {
    writeWord("");
}