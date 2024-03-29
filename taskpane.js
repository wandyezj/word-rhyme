var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
var host = undefined;
Office.onReady(function (info) {
    host = info.host;
});
function getSelectedTextWord() {
    return __awaiter(this, void 0, void 0, function () {
        var selectedText;
        var _this = this;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    selectedText = "";
                    return [4, Word.run(function (context) { return __awaiter(_this, void 0, void 0, function () {
                            var range;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        range = context.document.getSelection();
                                        range.load("text");
                                        return [4, context.sync()];
                                    case 1:
                                        _a.sent();
                                        selectedText = range.text;
                                        return [2];
                                }
                            });
                        }); })];
                case 1:
                    _a.sent();
                    return [2, selectedText];
            }
        });
    });
}
function getSelectedTextOutlook() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            return [2, new Promise(function (resolve) {
                    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, function (result) {
                        var value = result.value;
                        var text = value.data;
                        var selectedText = text ? text : "";
                        resolve(selectedText);
                    });
                })];
        });
    });
}
function getSelectedText() {
    return __awaiter(this, void 0, void 0, function () {
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0: return [4, Office.onReady()];
                case 1:
                    _a.sent();
                    if (!(host === Office.HostType.Word)) return [3, 3];
                    return [4, getSelectedTextWord()];
                case 2: return [2, _a.sent()];
                case 3:
                    if (!(host === Office.HostType.Outlook)) return [3, 5];
                    return [4, getSelectedTextOutlook()];
                case 4: return [2, _a.sent()];
                case 5:
                    console.log("Unsupported Host");
                    return [2, "Unsupported"];
            }
        });
    });
}
function buttonRhymeSelection() {
    return document.getElementById("button-rhyme-selection");
}
var running = false;
var previousWord = undefined;
function run() {
    return __awaiter(this, void 0, void 0, function () {
        var selectedText, word, message, rhymes, hasRhymes, e_1, rhyme;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    if (running) {
                        return [2];
                    }
                    running = true;
                    clearRhyme();
                    clearMessage();
                    return [4, getSelectedText()];
                case 1:
                    selectedText = _a.sent();
                    word = selectedText.trim();
                    message = undefined;
                    if (word !== previousWord) {
                        clearWord();
                    }
                    if (!(word === "")) return [3, 2];
                    message = "Highlight a word to rhyme!";
                    return [3, 11];
                case 2:
                    if (!hasWhiteSpace(word)) return [3, 3];
                    message = "A space was selected, multiple words, thus rejected.";
                    return [3, 11];
                case 3:
                    if (!(word.length > 28)) return [3, 4];
                    message = "The longest non-contrived and nontechnical word is antidisestablishmentarianism.";
                    return [3, 11];
                case 4:
                    writeWord(word);
                    previousWord = word;
                    rhymes = [];
                    hasRhymes = wordRhymes(word);
                    if (!(hasRhymes === undefined)) return [3, 9];
                    buttonRhymeSelection().disabled = true;
                    _a.label = 5;
                case 5:
                    _a.trys.push([5, 7, , 8]);
                    return [4, getWordRhymes(word)];
                case 6:
                    rhymes = _a.sent();
                    return [3, 8];
                case 7:
                    e_1 = _a.sent();
                    console.error(e_1);
                    rhymes = ["An error! Oh my! Give another word a try."];
                    return [3, 8];
                case 8:
                    buttonRhymeSelection().disabled = false;
                    return [3, 10];
                case 9:
                    rhymes = hasRhymes;
                    _a.label = 10;
                case 10:
                    if (rhymes.length === 0) {
                        message = "No Rhymes to uncover, select another word to discover.";
                    }
                    else {
                        rhyme = getRandomIndex(rhymes);
                        writeRhyme(rhyme);
                    }
                    _a.label = 11;
                case 11:
                    if (message !== undefined) {
                        writeMessage(message);
                    }
                    running = false;
                    return [2];
            }
        });
    });
}
function getRandomWhole(max) {
    return Math.floor(Math.random() * Math.floor(max));
}
function getRandomIndex(array) {
    var max = array.length;
    var index = getRandomWhole(max);
    return array[index];
}
function hasWhiteSpace(s) {
    return /\s/g.test(s);
}
function getWordRhymesFromDatamuse(word) {
    return __awaiter(this, void 0, void 0, function () {
        var query, response, o, list;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    query = "https://api.datamuse.com/words?rel_rhy=" + word;
                    return [4, fetch(query, {
                            mode: 'cors'
                        })];
                case 1:
                    response = _a.sent();
                    return [4, response.json()];
                case 2:
                    o = _a.sent();
                    list = o.map(function (item) { return item.word; }).filter(function (word) { return !hasWhiteSpace(word); });
                    return [2, list];
            }
        });
    });
}
var dictionary = new Map();
function wordRhymes(word) {
    var words = dictionary.get(word);
    return words;
}
function getWordRhymes(word) {
    return __awaiter(this, void 0, void 0, function () {
        var words;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    words = dictionary.get(word);
                    if (!(words === undefined)) return [3, 2];
                    return [4, getWordRhymesFromDatamuse(word)];
                case 1:
                    words = _a.sent();
                    dictionary.set(word, words);
                    _a.label = 2;
                case 2: return [2, words];
            }
        });
    });
}
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
//# sourceMappingURL=taskpane.js.map