/*
	Speech Recognition Editor

	A single page web-app text editor using speech recognition.

	Author: LJ Bubb
	License: GPL 3.0
	
	Libraries
	---------
	Incorporates "compromise" library under the MIT license:
	https://github.com/spencermountain/compromise
*/

const editorStorageKey = "speech-editor-a44ddce83936:";

const realTimeSpeechDivName = "real-time-speech";
const aggregatedSpeechDivName = "aggregated-speech";
const sentenceEditorDivNameOuter = "sentence-editor-outer";
const sentenceEditorDivName = "sentence-editor-content";
const paragraphEditorDivNameOuter = "paragraph-editor-outer";
const paragraphEditorDivName = "paragraph-editor-content";
const wordsEditorDivNameOuter = "words-editor-outer";
const wordsEditorDivName = "words-editor-content";
const fullDocumentDivName = "full-document-editable";
const alternativesDivNameOuter = "real-time-alternatives-outer";
const alternativesDivName = "real-time-alternatives-content";
const configAreaDivName = "config-area";

const realTimeSpeechElement = document.getElementById(realTimeSpeechDivName);
const aggregateElement = document.getElementById(aggregatedSpeechDivName);
const fullDocTopElement = document.getElementById(fullDocumentDivName);
const sentencesEditorDivOuter = document.getElementById(
	sentenceEditorDivNameOuter,
);
const sentencesEditorDiv = document.getElementById(sentenceEditorDivName);
const paragraphEditorDivOuter = document.getElementById(paragraphEditorDivNameOuter);
const paragraphEditorDiv = document.getElementById(paragraphEditorDivName);
const wordsEditorDivOuter = document.getElementById(wordsEditorDivNameOuter);
const wordsEditorDiv = document.getElementById(wordsEditorDivName);
const alternativesDiv = document.getElementById(alternativesDivName);
const alternativesDivOuter = document.getElementById(alternativesDivNameOuter);
const configAreaDiv = document.getElementById(configAreaDivName);

const microphoneButton = document.getElementById("microphone-icon");
const spellButtonElement = document.getElementById("spell-button");

let micTimeout = 60 * 5; //Five minute timeout for the mic
let timeoutAt = 0;

let bufferStateBeforeCurrentUtternace = ""; //What was the buffer text before anything was appended (used for alternatives feature)?
let bufferUndoStateBeforeUserEdit = { buffer: "", alternatives: [] }; //What would the undo state have before the user manually edited the buffer?
let documentUndoStateBeforeUserEdit = { document: "" }; //The full document state prior to any edits made by the user
let undoBufferPointer = -1;
let undoBuffer = [];

// Word to prefix commands with
var commandTriggerWord = "command";
var escapeTriggerWord = "literal";

// In paragraph, sentence or word editing mode or neither.
const standardEditorMode = "standard";
const wordsEditorMode = "words";
const sentenceEditorMode = "sentence";
const paragraphEditorMode = "paragraph";
let editorFocus = standardEditorMode;
let spellModeOn = false;
let withSpace = true;

let currentlyOpenEditors = {
	words: false,
	sentence: false,
	paragraph: false,
};

//Any currently stored text from cut/copy commands.
let pasteBuffer = "";

const appendCommandName = "return";
const emptyCommandName = "empty";
const chooseAlternative = "alternative";
const spellModeCommandName = "spell";
const undoCommandName = "undo";

const wordsEditorNormalisedCommandName = "words editor";
const displaySentencesCommandName = "sentences editor";
const displayParagraphsCommandName = "paragraphs editor";
const closeEditorCommandName = "close editor";
const cutNumberedItem = "cut";
const cutBufferCommandName = "cut buffer";
const cutAfterCommandName = "cut after";
const cutBeforeCommandName = "cut before";
const pasteIntoBufferCommandName = "paste into buffer";
const pasteAfterNumberedItem = "paste after";
const pasteBeforeNumberedItem = "paste before";
const insertAfterNumberedItem = "insert after";
const insertBeforeNumberedItem = "insert before";
const replaceCommandName = "replace";
const restoreNormalisedCommandName = "restore";
const upperCaseNormalisedCommandName = "upper case";
const lowerCaseNormalisedCommandName = "lower case";
const titleCaseNormalisedCommandName = "title case";
const shortenNormalisedCommandName = "shorten";
const moveUpCommandName = "move up";
const moveDownCommandName = "move down";
const cutLastSentenceCommand = "cut last sentence";
const scrollToNumberCommand = "scroll to item";
const scrollUpCommand = "scroll up";
const scrollDownCommand = "scroll down";

let fullCommandList = [
	{
		normalisedCommand: appendCommandName,
		userPhrase: "return",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: emptyCommandName,
		userPhrase: "empty",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: chooseAlternative,
		userPhrase: "alternative",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: spellModeCommandName,
		userPhrase: "spell",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: undoCommandName,
		userPhrase: "undo|back",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: cutLastSentenceCommand,
		userPhrase: "cut last sentence",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: scrollUpCommand,
		userPhrase: "scroll up",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: scrollDownCommand,
		userPhrase: "scroll down",
		param: false,
		paramType: null,
	},

	{
		normalisedCommand: wordsEditorNormalisedCommandName,
		userPhrase: "words",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: displaySentencesCommandName,
		userPhrase: "sentences",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: displayParagraphsCommandName,
		userPhrase: "paragraphs",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: closeEditorCommandName,
		userPhrase: "close",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: cutBufferCommandName,
		userPhrase: "cut buffer",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: pasteIntoBufferCommandName,
		userPhrase: "paste into buffer",
		param: false,
		paramType: null,
	},
	{
		normalisedCommand: cutNumberedItem,
		userPhrase: "cut",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: cutAfterCommandName,
		userPhrase: "cut after",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: cutBeforeCommandName,
		userPhrase: "cut before",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: pasteAfterNumberedItem,
		userPhrase: "paste after",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: pasteBeforeNumberedItem,
		userPhrase: "paste before",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: insertAfterNumberedItem,
		userPhrase: "insert after",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: insertBeforeNumberedItem,
		userPhrase: "insert before",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: restoreNormalisedCommandName,
		userPhrase: "restore",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: replaceCommandName,
		userPhrase: "replace",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: upperCaseNormalisedCommandName,
		userPhrase: "upper case|uppercase",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: lowerCaseNormalisedCommandName,
		userPhrase: "lower case|lowercase",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: titleCaseNormalisedCommandName,
		userPhrase: "title case",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: shortenNormalisedCommandName,
		userPhrase: "shorten",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: moveUpCommandName,
		userPhrase: "move up",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: moveDownCommandName,
		userPhrase: "move down",
		param: true,
		paramType: "single_int_first_last",
	},
	{
		normalisedCommand: scrollToNumberCommand,
		userPhrase: "scroll",
		param: true,
		paramType: "single_int_first_last",
	},
];

const punctRuleLikeComma = "Like Comma"; //Usually insert a space afterwards but no space prior to the character
const punctRuleLikeQuote = "Like Quote"; //Space before open quote but not before closing quote
const punctRuleLikeOpenBracket = "Like Open Bracket"; //Space before character but not after character
const punctRuleLikeWord = "Like Word"; //Usually space before and after
const punctRuleOnlySpaceAfter = "Space After"; //Always insert a space after character but not before
const punctRuleOnlySpaceBefore = "Space Before"; //Always insert a space before the character
const punctRuleNone = "No Spaces"; //Insert no spaces

export const configPunctuationMenuChoices = [
	punctRuleLikeComma,
	punctRuleLikeQuote,
	punctRuleLikeOpenBracket,
];

let autoReplacementPhrases = [
	{
		userPhrase: "period|full stop",
		replacement: ".",
		spaceRules: punctRuleLikeComma,
	},
	{
		userPhrase: "comma",
		replacement: ",",
		spaceRules: punctRuleLikeComma,
	},
	{
		userPhrase: "exclamation mark",
		replacement: "!",
		spaceRules: punctRuleLikeComma,
	},
	{
		userPhrase: "semi colon|semicolon",
		replacement: ";",
		spaceRules: punctRuleLikeComma,
	},
	{
		userPhrase: "question mark",
		replacement: "?",
		spaceRules: punctRuleLikeComma,
	},
	{
		userPhrase: "open bracket",
		replacement: "(",
		spaceRules: punctRuleLikeOpenBracket,
	},
	{
		userPhrase: "close bracket",
		replacement: ")",
		spaceRules: punctRuleLikeComma,
	},
	{
		userPhrase: "double quote",
		replacement: '"',
		spaceRules: punctRuleLikeQuote,
	},
	{ userPhrase: "quote", replacement: "'", spaceRules: punctRuleLikeQuote },
];

export function getAutoReplacements() {
	return autoReplacementPhrases;
}

export function updateAutoReplacements(newPhrases) {
	autoReplacementPhrases = newPhrases;
}

export function getCommandList() {
	return fullCommandList;
}

export function setCommandList(newCommandList) {
	fullCommandList = newCommandList;
}

export function getStorageKey() {
	return editorStorageKey;
}

export function getTriggerWord() {
	return commandTriggerWord;
}

export function setTriggerWord(newWord) {
	if(newWord)
		commandTriggerWord = newWord.trim();
}

export function getEscapeWord() {
	return escapeTriggerWord;
}

export function setEscapeWord(newWord) {
	if(newWord)
		escapeTriggerWord = newWord.trim();
}

export function setRecognitionLanguage(newLangauge){
	if(recognition){
		recognition.lang = newLangauge;
	}
}

const numRecognitionAlternatives = 4;

let recognitionRunning = false;
let recognition = configureRecognition();

document.addEventListener("DOMContentLoaded", () => {
	setEvents();
	restoreDocumentFromLocalStorage();
});

function displayMessage(messageHTML) {
	document.getElementById("message-content").innerHTML = messageHTML;
	showDiv("floating-message");
}

function configureRecognition() {
	let recognition = window.SpeechRecognition || window.webkitSpeechRecognition;

	if (recognition) {
		recognition = new recognition();

		let JSGFGrammar = buildJSGFGrammar();
		var speechRecognitionList = new webkitSpeechGrammarList();
		speechRecognitionList.addFromString(JSGFGrammar, 1);
		recognition.grammars = speechRecognitionList; //Doesn't have much effect in Chrome

		recognition.continuous = true;
		recognition.interimResults = true;
		recognition.lang = 'en-US';
		recognition.maxAlternatives = numRecognitionAlternatives;

		recognition.onstart = () => {
			console.log("Recording started");
		};

		recognition.onresult = function (event) {
			let result = "";
			let partialResult = "";

			resetMicInactivityTimer();

			for (let i = event.resultIndex; i < event.results.length; i++) {
				if (event.results[i].isFinal) {
					result += event.results[i][0].transcript;
					let alternativeTranscripts = generateAlternatives(event.results[i]);

					let commandCheckResponse = filterUtterance(
						result.trim(),
						spellModeOn,
					);

					if (!commandCheckResponse.ignoreUtternance) {
						if (commandCheckResponse.filteredUtterance.length > 0) {
							//Push the current state into the undo buffer just before it is modified
							setUndoBufferCurrent(["buffer", "alternatives"]);
							refreshAlternatives(alternativeTranscripts);
							appendToBuffer(commandCheckResponse.filteredUtterance, withSpace);
							if (spellModeOn) withSpace = false;
							
							//Chrome has an annoying bug where it seems to forget the recognition language
							//part way through speaking and default to US English for part of the phrase
							//This forces a restart which seems to improve it
							recognition.stop();
						}
					}
					processCommands(commandCheckResponse);
				} else {
					partialResult += event.results[i][0].transcript;
				}
			}

			realTimeSpeechElement.innerText = partialResult;
		};

		recognition.onerror = function (event) {
			console.log("Speech recognition error:", event.error);
			switch (event.error) {
				case "network":
					console.log("Network error occurred. Can't contact service");
					displayMessage(
						"Network error when trying to contact speech recognition service. Is the network available?",
					);
					recognitionRunning = false;
					microPhoneIconState(false);
					break;
				case "not-allowed":
					console.log("no-allowed erorr.");
					displayMessage(
						"The browser has denied permission for speech recognition.",
					);
					recognitionRunning = false;
					microPhoneIconState(false);
					break;
			}
		};

		recognition.onend = function () {
			console.log("Speech recognition ended");

			//Restart recognition if it should be running
			let timeNow = Math.floor(Date.now() / 1000);
			if (recognitionRunning && timeNow < timeoutAt){
				recognition.start();
			}else
				microPhoneIconState(false);
		};
	} else {
		console.log("Speech recognition not supported.");
		displayMessage("This browser does not support speech recognition.");
		recognitionRunning = false;
		microPhoneIconState(false);
	}

	return recognition;
}

function setUndoBufferCurrent(elements) {
	setUndoBuffer(makeUndoBufferState(elements));
}

function makeUndoBufferState(whichElements) {
	let undoChanges = {};

	for (let element of whichElements) {
		switch (element) {
			case "alternatives":
				undoChanges["alternatives"] = getCurrentAlternatives();
				break;
			case "buffer":
				undoChanges["buffer"] = aggregateElement.innerText;
				break;
			case "document":
				undoChanges["document"] = fullDocTopElement.innerText;
				break;
			case "words":
				undoChanges["words"] = getAllWordsFromWordsEditor();
				break;
			case "sentences":
				undoChanges["sentences"] = getAllSentencesFromSentenceEditor();
				break;
			case "paragraphs":
				undoChanges["paragraphs"] = getAllParagraphsFromParagraphEditor();
				break;
			case "windowState":
				undoChanges["windowState"] = cloneWindowStateDict();
				break;
		}
	}

	return undoChanges;
}

function setUndoBuffer(whatWillChange) {
	if (undoBufferPointer == -1) undoBuffer = [];

	if (undoBufferPointer >= 0 && undoBufferPointer <= undoBuffer.length - 2)
		undoBuffer.splice(undoBufferPointer + 1);

	whatWillChange["bufferStateBeforeCurrentUtternace"] =
		bufferStateBeforeCurrentUtternace;
	whatWillChange["spellModeOn"] = spellModeOn;
	whatWillChange["withSpace"] = withSpace;

	undoBuffer.push(whatWillChange);
	undoBufferPointer = undoBuffer.length - 1;
}

function handleUndo() {
	if (undoBufferPointer >= 0 && undoBuffer.length > 0) {
		let undoState = undoBuffer[undoBufferPointer];

		let redoState = {};

		if ("alternatives" in undoState) {
			redoState["alternatives"] = getCurrentAlternatives();
			refreshAlternatives(undoState["alternatives"]);
		}
		if ("buffer" in undoState) {
			redoState["buffer"] = aggregateElement.innerText;
			aggregateElement.innerText = undoState["buffer"];
			scrollDivEnd(aggregateElement, false);
			moveCursorToEnd( aggregateElement );
		}
		if ("document" in undoState) {
			redoState["document"] = fullDocTopElement.innerText;
			fullDocTopElement.innerText = undoState["document"];
			preserveStateLocalStorage();
		}
		if ("words" in undoState) {
			redoState["words"] = getAllWordsFromWordsEditor();
			refreshWordsEditor(undoState["words"]);
		}
		if ("sentences" in undoState) {
			redoState["sentences"] = getAllSentencesFromSentenceEditor();
			refreshSentenceEditor(undoState["sentences"]);
		}
		if ("paragraphs" in undoState) {
			redoState["paragraphs"] = getAllParagraphsFromParagraphEditor();
			refreshParagraphEditor(undoState["paragraphs"]);
		}
		if ("windowState" in undoState) {
			redoState["windowState"] = cloneWindowStateDict();

			if (undoState["windowState"].words) showDiv(wordsEditorDivNameOuter);
			else hideDiv(wordsEditorDivNameOuter);

			if (undoState["windowState"].sentence)
				showDiv(sentenceEditorDivNameOuter);
			else hideDiv(sentenceEditorDivNameOuter);

			if (undoState["windowState"].paragraph) {
				showDiv(paragraphEditorDivNameOuter);
				hideDiv(fullDocumentDivName);
			} else {
				hideDiv(paragraphEditorDivNameOuter);
				showDiv(fullDocumentDivName);
			}

			currentlyOpenEditors = undoState["windowState"];
			resolveNewEditorFocus();
		}

		redoState["bufferStateBeforeCurrentUtternace"] =
			bufferStateBeforeCurrentUtternace;
		redoState["spellModeOn"] = spellModeOn;
		redoState["withSpace"] = withSpace;

		if (undoBufferPointer === undoBuffer.length - 1) undoBuffer.push(redoState);

		bufferUndoStateBeforeUserEdit = redoState;

		if ("bufferStateBeforeCurrentUtternace" in undoState)
			bufferStateBeforeCurrentUtternace =
				undoState["bufferStateBeforeCurrentUtternace"];

		if ("spellModeOn" in undoState) {
			spellModeOn = undoState["spellModeOn"];
			setButtonHighlight(spellButtonElement, spellModeOn);
		}
		if ("withSpace" in undoState) withSpace = undoState["withSpace"];

		undoBufferPointer--;
	}
}

function cloneWindowStateDict() {
	return JSON.parse(JSON.stringify(currentlyOpenEditors));
}

function generateAlternatives(recogniserAlternatives) {
	let alternativeTranscripts  = []
	for(let alt of recogniserAlternatives){
		let filteredAlt = filterUtterance(alt.transcript.trim(), spellModeOn).filteredUtterance;
		if(filteredAlt.length > 0)
			alternativeTranscripts.push( filteredAlt );
	}
	
	if (alternativeTranscripts.length > 0) {
		//Assume altnerative 1 is the most probable
		//Also add the lower case version of the most probable transcript to avoid random upper cases
		alternativeTranscripts.push(alternativeTranscripts[0].toLowerCase());
	}
	return deduplicateArray(alternativeTranscripts);
}

function filterUtterance(utterance, spellMode = false) {
	utterance = doAutoReplacements(
		utterance.trim(),
		autoReplacementPhrases,
		escapeTriggerWord,
	);
	let commandCheckResponse = interpretCommands(
		commandTriggerWord,
		fullCommandList,
		utterance,
		escapeTriggerWord,
	);
	commandCheckResponse.filteredUtterance = removeInstancesOfEscapeKeyword(
		commandCheckResponse.filteredUtterance,
		escapeTriggerWord,
	).trim();
	commandCheckResponse.filteredUtterance = fixPunctuationSpacing(
		commandCheckResponse.filteredUtterance,
	).trim();
	if (spellMode)
		commandCheckResponse.filteredUtterance = spellFromWords(
			commandCheckResponse.filteredUtterance,
		).trim();
	return commandCheckResponse;
}

function appendToBuffer(text, withSpace = true) {
	//This is the buffer state just prior to the text being appended
	//This is needed for the "alternatives" feature such that it can undo the current utterance and then re-append the alternative
	bufferStateBeforeCurrentUtternace = aggregateElement.innerText.trim();

	if (withSpace) aggregateElement.innerText += " ";
	aggregateElement.innerText += text.trim();
	aggregateElement.innerText = autoFixCapitalLetters(aggregateElement.innerText,);
	aggregateElement.innerText = fixPunctuationSpacing(aggregateElement.innerText).trim();

	//The user can freely edit the buffer area. We need to record what the state of the buffer area was before they edited it
	//so it can be pushed into the undo buffer retrospectively.
	bufferUndoStateBeforeUserEdit = {
		buffer: aggregateElement.innerText.trim(),
		alternatives: getCurrentAlternatives(),
	};
	
	scrollDivEnd(aggregateElement, false);
	moveCursorToEnd( aggregateElement );
}

function buildJSGFGrammar() {
	let phrases = [];

	for (let command of fullCommandList) {
		for (let version of command.userPhrase.split("|"))
			phrases.push(`${commandTriggerWord} ${version}`);
	}

	for (let ar of autoReplacementPhrases) {
		for (let version of ar.userPhrase.split("|")) phrases.push(version);
	}

	let JSGFgrammar =
		"#JSGF V1.0; grammar words; public <commands> = " +
		phrases.join(" | ") +
		" ;";
	return JSGFgrammar;
}

//TODO: Bug where it splits possessive contractions
//e.g. splits Chris' into Chris and ' but should be one word
function splitStringIntoTokens(inputString) {
	const pattern =
		/\b(?:\d+[\p{Letter}_-]|\p{Letter})[\p{Letter}\d_-]*(?:[''][\p{Letter}\d_-]+)?(?:[-_][\p{Letter}\d_-]+)*\b|\d+(?:\.\d+)?%?|'|[^\p{Letter}\s]|[\p{Letter}]+/giu;

	// Find all matches in the input string
	const result = inputString.match(pattern);

	return result || [];
}

function spellFromWords(inputString) {
	let outputString = "";
	if (inputString.trim().length > 0) {
		let tokens = splitStringIntoTokens(inputString);
		for (let token of tokens) {
			if (token.length > 0) {
				outputString += token.charAt(0).toLowerCase();
			}
		}
	}

	return outputString;
}

function parseSentencesFromParagraph(paragraph) {
	const nlp_parser = nlp(paragraph);
	const sentences = nlp_parser.sentences().out("array");
	return sentences;
}

function cutLastSentenceFromParagraph(paragraph) {
	let sentences = parseSentencesFromParagraph(paragraph);
	if (sentences.length < 2) return "";

	let lastSentence = sentences.pop();
	pasteBuffer = lastSentence;

	return sentences.join(" ");
}

function refreshSentenceEditor(parsed_sentences_array) {
	refreshEditor(sentencesEditorDiv, parsed_sentences_array);
}

function refreshParagraphEditor(parsed_paragraphs_array, resplit=false) {
	if(resplit)
		parsed_paragraphs_array  = (parsed_paragraphs_array.join("\n\n")).split("\n\n");
		
	refreshEditor(paragraphEditorDiv, parsed_paragraphs_array);
	preserveStateLocalStorage();
}

function refreshWordsEditor(parsed_tokens, resplitWords = false) {
	if (resplitWords)
		parsed_tokens = splitStringIntoTokens(joinWords(parsed_tokens));
	refreshEditor(wordsEditorDiv, parsed_tokens, false);
}

function deduplicateArray(arr) {
	return Object.keys(
		arr.reduce((acc, curr) => {
			acc[curr] = true;
			return acc;
		}, {}),
	);
}

function refreshAlternatives(altSet) {
	let alts = deduplicateArray(altSet);
	//If there's only one alternative then it's already in the editor and the choice is no use
	if (alts.length < 2) alts = [];
		refreshEditor(alternativesDiv, alts);
	if (alts.length < 1) 
		hideDiv(alternativesDivNameOuter);
	else
		showDiv(alternativesDivNameOuter);
}

function processChooseAlternativeCommand(itemNum) {
	let alts = getCurrentAlternatives();
	if (alts.length > 0) {
		if (itemNum >= 1 && itemNum <= alts.length) {
			setUndoBufferCurrent(["buffer", "alternatives"]);
			let altPhrase = alts[itemNum - 1];
			aggregateElement.innerText = bufferStateBeforeCurrentUtternace;
			appendToBuffer(altPhrase);
			refreshAlternatives(alts);
			bufferUndoStateBeforeUserEdit = {
				buffer: aggregateElement.innerText.trim(),
				alternatives: getCurrentAlternatives(),
			};
			scrollDivEnd(aggregateElement, false);
			moveCursorToEnd( aggregateElement );
		}
	}
}

function refreshEditor(
	editorDivElement,
	parsed_text_array,
	div_per_token = true,
) {
	editorDivElement.innerHTML = "";
	for (let i = 0; i < parsed_text_array.length; i++) {
		let div = document.createElement("div");
		let idNumSpan = document.createElement("span");
		idNumSpan.textContent = i + 1;
		idNumSpan.className = "editor-id-num-active";
		let sentenceTextSpan = document.createElement("span");
		sentenceTextSpan.textContent = parsed_text_array[i].trim();
		sentenceTextSpan.className = "editor-parsed-token";
		if (!div_per_token) {
			editorDivElement.appendChild(idNumSpan);
			editorDivElement.appendChild(sentenceTextSpan);
		} else {
			div.appendChild(idNumSpan);
			div.appendChild(sentenceTextSpan);
			editorDivElement.appendChild(div);
		}
	}
}

function getCurrentAlternatives() {
	return gatherSpanText(alternativesDivName, "editor-parsed-token");
}

function getAllSentencesFromSentenceEditor() {
	return gatherSpanText(sentenceEditorDivName, "editor-parsed-token");
}

function getAllParagraphsFromParagraphEditor() {
	return gatherSpanText(paragraphEditorDivName, "editor-parsed-token");
}

function getAllWordsFromWordsEditor() {
	return gatherSpanText(wordsEditorDivName, "editor-parsed-token");
}

function joinWords(wordList) {
	let joined_words = wordList.join(" ");
	return autoFixCapitalLetters(
		fixPunctuationSpacing(joined_words.trim()),
	).trim();
}

function closeWordsEditor( overWriteBuffer=true ) {
	if (currentlyOpenEditors.words) {
		let words = getAllWordsFromWordsEditor();
		if (words.length > 0 && overWriteBuffer) {
			setUndoBufferCurrent(["windowState", "buffer", "alternatives", "words"]);
			emptyBuffer();
			aggregateElement.innerText = joinWords(words);
		}

		wordsEditorDiv.innerHTML = "";
		currentlyOpenEditors.words = false;
		hideDiv(wordsEditorDivNameOuter);
		resolveNewEditorFocus();
		scrollDivEnd(aggregateElement, false);
		moveCursorToEnd( aggregateElement );
	}
}

function closeSentenceEditor( overWriteBuffer=true ) {
	if (currentlyOpenEditors.sentence) {
		setUndoBufferCurrent([
			"windowState",
			"buffer",
			"alternatives",
			"sentences",
		]);
		let sentences = getAllSentencesFromSentenceEditor();
		if( overWriteBuffer && sentences.length > 0){
			emptyBuffer();
			let joined_sentences = sentences.join(" ");
			aggregateElement.innerText = fixPunctuationSpacing(joined_sentences.trim());
			scrollDivEnd(aggregateElement, false);
			moveCursorToEnd( aggregateElement );
		}
		
		currentlyOpenEditors.sentence = false;		
		resolveNewEditorFocus();
		hideDiv(sentenceEditorDivNameOuter);
		sentencesEditorDiv.innerHTML = "";
	}
}

//Cut a sentence from the sentence editor (1 indexed number) and store it
function cutSentenceFromEditor(sentence_number) {
	let sentences = getAllSentencesFromSentenceEditor();
	if (sentence_number >= 0) {
		if(sentence_number == 0)
			sentence_number = 1;
		if(sentence_number >= sentences.length)
			sentence_number = sentences.length;
		setUndoBufferCurrent(["sentences"]);
		let cutSentence = sentences.splice(sentence_number - 1, 1)[0];
		pasteBuffer = cutSentence.trim();
		refreshSentenceEditor(sentences);
	}
}

function cutParagraphFromEditor(paragraph_number) {
	let paragraphs = getAllParagraphsFromParagraphEditor();
	if (paragraph_number >= 0) {
		if(paragraph_number == 0)
			paragraph_number = 1
		if(paragraph_number > paragraphs.length)
			paragraphs = paragraphs.length
	
		setUndoBufferCurrent(["paragraphs"]);
		let cutParagraph = paragraphs.splice(paragraph_number - 1, 1)[0];
		pasteBuffer = cutParagraph.trim();
		refreshParagraphEditor(paragraphs);
	}
}

function cutWordFromEditor(word_number) {
	let allWords = getAllWordsFromWordsEditor();
	if (word_number >= 0) {
		if(word_number == 0)
			word_number = 1;
		if(word_number > allWords.length)
			word_number = allWords.length;
		setUndoBufferCurrent(["words"]);
		let cutWord = allWords.splice(word_number - 1, 1)[0];
		pasteBuffer = cutWord.trim();
		refreshWordsEditor(allWords);
	}
}


function cutSentenceAfter(sentence_number, cut_before=false, copy=false) {
	let sentences = getAllSentencesFromSentenceEditor();
	
	if(sentences.length >= 1){	
		if( cut_before ){
			if( sentence_number >= 1 ){
				if(!copy)
					setUndoBufferCurrent(["sentences"]);
				if( sentence_number > sentences.length ){
					pasteBuffer = sentences.join(' ').trim();
					if(!copy)
						closeSentenceEditor(false);
				}else{		
					let beforeSentences = sentences.slice(0, sentence_number-1);
					let afterSentences = sentences.splice(sentence_number-1);
					pasteBuffer = beforeSentences.join(' ').trim();
					if(!copy)
						refreshSentenceEditor(afterSentences);
				}
			}	
		}else{	
			if (sentence_number >= 0 && sentence_number <= (sentences.length -1)) {
				if(!copy)
					setUndoBufferCurrent(["sentences"]);
				
				if( sentence_number == 0 ){
					pasteBuffer= sentences.join(' ').trim();
					if(!copy)
						closeSentenceEditor(false);
				}else{
					let initialSentences = sentences.slice(0, sentence_number);
					let cutSentences = sentences.splice(sentence_number);
					pasteBuffer = cutSentences.join(' ').trim();
					if(!copy)
						refreshSentenceEditor(initialSentences);
				}			
			}
		}
	}
}


function cutWordAfter(word_number, cut_before=false, copy=false) {
	let words = getAllWordsFromWordsEditor();
	
	if(words.length >= 1){	
		if( cut_before ){
			if( word_number >= 1 ){
				if(!copy)
					setUndoBufferCurrent(["words"]);
				if( word_number > words.length ){
					pasteBuffer = joinWords(words).trim();
					if(!copy)
						closeWordsEditor(false);
				}else{		
					let beforewords = words.slice(0, word_number-1);
					let afterwords = words.splice(word_number-1);
					pasteBuffer = joinWords(beforewords).trim();
					if(!copy)
						refreshWordsEditor(afterwords);
				}
			}	
		}else{	
			if (word_number >= 0 && word_number <= (words.length -1)) {
				if(!copy)
					setUndoBufferCurrent(["words"]);
				
				if( word_number == 0 ){
					pasteBuffer= joinWords(words).trim();
					if(!copy)
						closeWordsEditor(false);
				}else{
					let initialwords = words.slice(0, word_number);
					let cutwords = words.splice(word_number);
					pasteBuffer = joinWords(cutwords).trim();
					if(!copy)
						refreshWordsEditor(initialwords);
				}			
			}
		}
	}
}


function cutParagraphAfter(paragraph_num, cut_before=false, copy=false) {
	let paragraphs = getAllParagraphsFromParagraphEditor();
	
	if(paragraphs.length >= 1){	
		if( cut_before ){
			if( paragraph_num >= 1 ){
				if(!copy)
					setUndoBufferCurrent(["paragraphs"]);
				if( paragraph_num > paragraphs.length ){
					pasteBuffer = paragraphs.join("\n\n").trim();
					if(!copy)
						paragraphEditorDiv.innerHTML='';
				}else{		
					let beforeparagraphs = paragraphs.slice(0, paragraph_num-1);
					let afterparagraphs = paragraphs.splice(paragraph_num-1);
					pasteBuffer = join("\n\n").trim();
					if(!copy)
						refreshParagraphEditor(afterparagraphs);
				}
			}	
		}else{	
			if (paragraph_num >= 0 && paragraph_num <= (paragraphs.length -1)) {
				if(!copy)
					setUndoBufferCurrent(["paragraphs"]);
				
				if( paragraph_num == 0 ){
					pasteBuffer= paragraphs.join('\n\n').trim();
					if(!copy)
						paragraphEditorDiv.innerHTML='';
				}else{
					let initialparagraphs = paragraphs.slice(0, paragraph_num);
					let cutparagraphs = paragraphs.splice(paragraph_num);
					pasteBuffer = cutparagraphs.join("\n\n").trim();
					if(!copy)
						refreshParagraphEditor(initialparagraphs);
				}			
			}
		}
	}
}



function replaceSentenceAtEditor(sentenceNumber, replacement) {
	if (replacement.length > 0) {
		let normalisedWord = replacement.trim();

		let sentences = getAllSentencesFromSentenceEditor();

		if (sentenceNumber >= 0) {
			if(sentenceNumber == 0)
				sentenceNumber = 1;
			if(sentenceNumber > sentences.length)
				sentenceNumber = sentences.length;
			setUndoBufferCurrent(["sentences"]);
			sentences[sentenceNumber - 1] = replacement.trim();
			refreshSentenceEditor(parseSentencesFromParagraph(sentences.join(" ")));
			return true;
		}
	}
	return false;
}

//Insert text after (or before) the specified sentence from the sentence editor (1 indexed number) and store it
function insertIntoSentenceEditor(
	sentence_number,
	additionalSentence,
	before = false,
) {
	if (additionalSentence.trim().length > 0) {
		let sentences = getAllSentencesFromSentenceEditor();
		if (sentences.length == 0) {
			setUndoBufferCurrent(["buffer", "alternatives", "sentences"]);
			sentences = [additionalSentence.trim()];
		} else {
			if (sentence_number == -1) sentence_number = sentences.length;

			if (sentence_number >= 0 && sentence_number <= sentences.length) {
				setUndoBufferCurrent(["buffer", "alternatives", "sentences"]);
				let position = sentence_number;
				if (before && sentence_number >= 1) position--;
				sentences.splice(position, 0, additionalSentence.trim());
			}
		}
		refreshSentenceEditor(parseSentencesFromParagraph(sentences.join(" ")));
		return true;
	}
	return false;
}

function moveSentenceInEditor(itemNum, up = true) {
	let sentences = getAllSentencesFromSentenceEditor();
	if (itemNum >= 0 && itemNum <= sentences.length) {
		setUndoBufferCurrent(["sentences"]);
		sentences = moveItemInArray(sentences, itemNum - 1, up);
		refreshSentenceEditor(parseSentencesFromParagraph(sentences.join(" ")));
	}
}

function moveParagraphInEditor(itemNum, up = true) {
	let allParagraphs = getAllParagraphsFromParagraphEditor();
	if (itemNum >= 0 && itemNum <= allParagraphs.length) {
		setUndoBufferCurrent(["paragraphs"]);
		allParagraphs = moveItemInArray(allParagraphs, itemNum - 1, up);
		refreshParagraphEditor(allParagraphs);
	}
}

function replaceParagraphAtEditor(paragraphNumber, replacement) {
	if (replacement.length > 0) {
		let allParagraphs = getAllParagraphsFromParagraphEditor();

		if (paragraphNumber >= 0) {
			if(paragraphNumber == 0)
				paragraphNumber = 1;
			if(paragraphNumber > allParagraphs.length)
				paragraphNumber = allParagraphs.length;
			setUndoBufferCurrent(["paragraphs"]);
			allParagraphs[paragraphNumber - 1] = replacement.trim();
			refreshParagraphEditor(allParagraphs, true);
			return true;
		}
	}
	return false;
}

function insertIntoParagraphEditor(paraNumber, additionalPara, before = false) {
	if (additionalPara.trim().length > 0) {
		let allParagraphs = getAllParagraphsFromParagraphEditor();

		if (allParagraphs.length == 0) {
			setUndoBufferCurrent(["buffer", "alternatives", "paragraphs"]);
			allParagraphs = [additionalPara.trim()];
		} else {
			if (paraNumber == -1) paraNumber = allParagraphs.length;

			if (paraNumber >= 0 && paraNumber <= allParagraphs.length) {
				setUndoBufferCurrent(["buffer", "alternatives", "paragraphs"]);
				let position = paraNumber;
				if (before && paraNumber >= 1) position--;
				allParagraphs.splice(position, 0, additionalPara.trim() );
			}
		}
		refreshParagraphEditor(allParagraphs, true);
		return true;
	}
	return false;
}

function processCaseCommand(wordNumber, commandType) {
	if (editorFocus === wordsEditorMode) {
		let theWord = getWordAtEditor(wordNumber);
		let replacement = theWord;
		if (theWord != null) {
			switch (commandType) {
				case upperCaseNormalisedCommandName:
					//Check if uppercase and ends with an "S". This is for conversion of BOBS into BOBs for plurals of initialisms.
					replacement = replacement.replace(
						/^[^\p{Lowercase_Letter}]+S$/,
						(match) => match.slice(0, -1) + "s",
					);
					if (replacement === theWord) {
						replacement = replacement.toUpperCase();
					}
					break;
				case lowerCaseNormalisedCommandName:
					replacement = replacement.toLowerCase();
					break;
				case titleCaseNormalisedCommandName:
					replacement = capitaliseFirstLetter(replacement.toLowerCase());
					break;
			}

			replaceWordAtEditor(wordNumber, replacement);
		}
	}
}

function processShortenCommand(wordNumber) {
	if (editorFocus === wordsEditorMode) {
		let theWord = getWordAtEditor(wordNumber);
		if (theWord != null && theWord.length > 1) {
			theWord = theWord.slice(0, -1);
			replaceWordAtEditor(wordNumber, theWord);
		}
	}
}

function getWordAtEditor(wordNumber) {
	if (editorFocus === wordsEditorMode) {
		let allWords = getAllWordsFromWordsEditor();
		if (wordNumber >= 0 && allWords.length > 0) {
			if(wordNumber == 0)
				wordNumber = 1;
			if(wordNumber > allWords.length)
				wordNumber = allWords.length;
			return allWords[wordNumber - 1];
		}
	}
	return null;
}

function replaceWordAtEditor(wordNumber, replacement, caseNormalise = false) {
	if (replacement.length > 0) {
		setUndoBufferCurrent(["words"]);
		let normalisedWord = replacement.trim();
		if (caseNormalise) {
			//Unless starts with 2 upper case characters then set to lower case
			if (!/^\p{Uppercase_Letter}{2}/u.test(normalisedWord))
				normalisedWord = capitaliseFirstLetter(normalisedWord, true);
		}
		let allWords = getAllWordsFromWordsEditor();
		if (wordNumber >= 0) {
			if(wordNumber == 0)
				wordNumber = 1;
			if(wordNumber > allWords.length)
				wordNumber = allWords.length;
			allWords[wordNumber - 1] = normalisedWord;
			refreshWordsEditor(allWords, true);
			return true;
		}
	}
	return false;
}

function insertIntoWordsEditor(wordNumber, additionalWord, before = false) {
	if (additionalWord.trim().length > 0) {
		let normalisedWord = additionalWord.trim();

		//Unless starts with 2 upper case characters then set to lower case
		if (!/^\p{Uppercase_Letter}{2}/u.test(normalisedWord))
			normalisedWord = capitaliseFirstLetter(normalisedWord, true);

		let allWords = getAllWordsFromWordsEditor();
		if (allWords.length == 0) {
			setUndoBufferCurrent(["buffer", "alternatives", "words"]);
			allWords = [normalisedWord];
		} else {
			if (wordNumber == -1 || wordNumber > allWords.length)
				wordNumber = allWords.length;

			if (wordNumber >= 0 && wordNumber <= allWords.length) {
				setUndoBufferCurrent(["buffer", "alternatives", "words"]);
				let position = wordNumber;
				if (before && wordNumber >= 1) position--;
				allWords.splice(position, 0, normalisedWord);
			}
		}
		refreshWordsEditor(allWords, true);
		return true;
	}
	return false;
}

//Retrieve a sentence from the sentence editor back into the buffer
function getSentenceFromEditor(item_num) {
	return getNumberedItemFromEditor(getAllSentencesFromSentenceEditor(),item_num);
}

function getParagraphFromEditor(item_num) {
	return getNumberedItemFromEditor(
		getAllParagraphsFromParagraphEditor(),
		item_num,
	);
}

function getNumberedItemFromEditor(text_list, item_num) {
	if (text_list.length == 0) return "";

	if(item_num >= 0){	
		if(item_num == 0)
			item_num = 1;
		if(item_num > text_list.length)
			item_num = text_list.length;

		return text_list[item_num - 1];
	}
	
	return "";
}

//Get all text from a specific span class in a specific div into an array
function gatherSpanText(divName, spanClassName) {
	return gatherTextFromElements(divName, `span.${spanClassName}`);
}

function gatherParagraphElementText(divName, paragraphClassName = null) {
	let element = "p";
	if (paragraphClassName != null) element = `p.${paragraphClassName}`;

	return gatherTextFromElements(divName, element);
}

function gatherTextFromElements(divName, elementSpec) {
	// Select the div with the specific class
	const targetDiv = document.getElementById(divName);

	// If the div doesn't exist, return an empty array
	if (!targetDiv) return [];

	// Select all elements with the specific class within the div
	const spans = targetDiv.querySelectorAll(elementSpec);

	// Use Array.from() to convert NodeList to array and map to get innerText
	const spanTexts = Array.from(spans).map((span) => span.innerText.trim());

	return spanTexts;
}

function scrollToNumberedPositionInParagraphEditor(itemNum) {
	scrollToNumberedPositionInEditor(
		paragraphEditorDiv,
		"span.editor-parsed-token",
		itemNum,
	);
}

function scrollToNumberedPositionInSentencesEditor(itemNum) {
	scrollToNumberedPositionInEditor(
		sentencesEditorDiv,
		"span.editor-parsed-token",
		itemNum,
	);
}

function scrollToNumberedPositionInWordsEditor(itemNum) {
	scrollToNumberedPositionInEditor(
		wordsEditorDiv,
		"span.editor-parsed-token",
		itemNum,
	);
}

//elementSpec = a specification that defines a numbered item
function scrollToNumberedPositionInEditor(editorDiv, elementSpec, itemNum) {
	if (itemNum <= 1) {
		scrollDivEnd(editorDiv);
	} else {
		const spans = editorDiv.querySelectorAll(elementSpec);
		if (spans.length > 0) {
			if (itemNum < 1)
				itemNum = 1;
			if (itemNum > spans.length)
				itemNum = spans.length;

			spans[itemNum - 1].scrollIntoView({ behavior: "smooth" });
		}
	}
}
function hasVerticalScrollbar(element) {
	return element.scrollHeight > element.clientHeight;
}

function scrollDivEnd(divElement, top = true) {
	if( hasVerticalScrollbar(divElement) ){
		if (top)
			divElement.scrollTo({ top: 0, behavior: "smooth" });
		else
			divElement.scrollTo({ top: divElement.scrollHeight, behavior: "smooth" });
	}
}

function getScrollPercentageOfDiv(divElement) {
	if( hasVerticalScrollbar(divElement) ){
		let percent = divElement.scrollTop / (divElement.scrollHeight - divElement.clientHeight);
		return percent;
	}
	return 0;
}

function setScrollPercentageOfDiv(divElement, percent) {
	if( hasVerticalScrollbar(divElement) ){
		divElement.scrollTop = percent * (divElement.scrollHeight - divElement.clientHeight);
	}
}

function scrollDivDirection(divElement, up = true) {

	if( hasVerticalScrollbar(divElement) ){
		const halfHeight = divElement.clientHeight / 2;
		if (up) 
			divElement.scrollTop -= halfHeight;
		else
			divElement.scrollTop += halfHeight;
	}
}

function setRecordingState() {
	if (!recognitionRunning) {
		startRecording();
	} else {
		stopRecording();
	}
}

function startRecording() {
	if (recognition) {
		realTimeSpeechElement.innerText = "";
		recognition.start();
		resetMicInactivityTimer();
		microPhoneIconState(true);
	}
}

function resetMicInactivityTimer() {
	timeoutAt = Math.floor(Date.now() / 1000) + micTimeout;
}

function stopRecording() {
	if (recognition) {
		recognition.stop();
		microPhoneIconState(false);
	}
}

function microPhoneIconState(state) {
	if (state) {
		if (microphoneButton.classList.contains("microphone-button-black"))
			microphoneButton.classList.remove("microphone-button-black");
		if (!microphoneButton.classList.contains("microphone-button-red"))
			microphoneButton.classList.add("microphone-button-red");
		recognitionRunning = true;
	} else {
		if (microphoneButton.classList.contains("microphone-button-red"))
			microphoneButton.classList.remove("microphone-button-red");
		if (!microphoneButton.classList.contains("microphone-button-black"))
			microphoneButton.classList.add("microphone-button-black");
		recognitionRunning = false;
	}
}

function emptyBuffer() {
	aggregateElement.innerText = "";
	realTimeSpeechElement.innerText = "";
	refreshAlternatives([]);
	bufferStateBeforeCurrentUtternace = "";

	bufferUndoStateBeforeUserEdit = makeUndoBufferState([
		"buffer",
		"alternatives",
	]);
}

function processInsertNumberedItemCommand(
	item_number,
	text,
	before = false,
	empty_buffer = false,
) {
	switch (editorFocus) {
		case sentenceEditorMode:
			{
				let success = insertIntoSentenceEditor(item_number, text.trim(), before);
				if (empty_buffer && success)
					emptyBuffer();
			}
			break;
			
		case paragraphEditorMode:
			insertIntoParagraphEditor(item_number, text.trim(), before);
			break;
		case wordsEditorMode:
			{
				let successInsert = insertIntoWordsEditor(item_number, text.trim(), before);
				if (empty_buffer && successInsert)
					emptyBuffer();
			}
			break;
	}
}

function appendParagraphToFullDocument(fullDocumentDivName, text) {
	if(text.trim().length > 0){
		const divElement = document.getElementById(fullDocumentDivName);
		
		let allParagraphs = parseParagraphs(fullDocTopElement.innerText);
		allParagraphs.push( text );
		fullDocTopElement.innerText = allParagraphs.join("\n\n");
		
		preserveStateLocalStorage();
		scrollDivEnd(fullDocTopElement, false);
		moveCursorToEnd( fullDocTopElement );
	}
}

function processAppendCommand(text) {
	switch (editorFocus) {
		case standardEditorMode:
			setUndoBufferCurrent(["buffer", "alternatives", "document"]);
			appendParagraphToFullDocument(fullDocumentDivName, text.trim());
			emptyBuffer();
			documentUndoStateBeforeUserEdit = {
				document: fullDocTopElement.innerText,
			};
			scrollDivEnd(fullDocumentDivName, false);
			break;
		case wordsEditorMode:
			processInsertNumberedItemCommand(-1, text.trim(), false, true);
			emptyBuffer();
			scrollDivEnd(wordsEditorDiv, false);
			break;
		case sentenceEditorMode:
			processInsertNumberedItemCommand(-1, text.trim(), false, true);
			emptyBuffer();
			scrollDivEnd(sentencesEditorDiv, false);
			break;
		case paragraphEditorMode:
			processInsertNumberedItemCommand(-1, text.trim(), false, true);
			emptyBuffer();
			scrollDivEnd(paragraphEditorDiv, false);
			break;
	}
}

function hideDiv(divName) {
	const targetDiv = document.getElementById(divName);
	targetDiv.style.display = "none";
}

function showDiv(divName) {
	const targetDiv = document.getElementById(divName);
	targetDiv.style.display = "block";
}

function toggleDivVisible(divName) {
	const targetDiv = document.getElementById(divName);

	if (window.getComputedStyle(targetDiv).display === "none") {
		targetDiv.style.display = "block";
	} else {
		targetDiv.style.display = "none";
	}
}

function emptyDivByName(divName) {
	const divElement = document.getElementById(divName);
	divElement.innerHTML = "";
}

//Handle seting the current focus if something was closed
function resolveNewEditorFocus() {
	if (currentlyOpenEditors.sentence) setEditorFocus(sentenceEditorMode);
	else if (currentlyOpenEditors.paragraph) setEditorFocus(paragraphEditorMode);
	else if (currentlyOpenEditors.words) {
		setEditorFocus(wordsEditorMode);
	} else {
		setEditorFocus(standardEditorMode);
	}
}

function setFocusVisualState(divName, isFocussed) {
	let focusDiv = document.getElementById(divName);
	if (isFocussed) {
		focusDiv.style.borderWidth = "2px";
		focusDiv.style.borderColor = "red";
	} else {
		focusDiv.style.borderWidth = "1px";
		focusDiv.style.borderColor = "grey";
	}
}

function defocusAll() {
	setFocusVisualState(sentenceEditorDivNameOuter, false);
	setFocusVisualState(paragraphEditorDivNameOuter, false);
	setFocusVisualState(wordsEditorDivNameOuter, false);
}

function setEditorFocus(newFocus) {
	defocusAll();
	switch (newFocus) {
		case paragraphEditorMode:
			if (currentlyOpenEditors.paragraph) {
				editorFocus = newFocus;
				setFocusVisualState(paragraphEditorDivNameOuter, true);
			}
			break;
		case sentenceEditorMode:
			if (currentlyOpenEditors.sentence) {
				editorFocus = newFocus;
				setFocusVisualState(sentenceEditorDivNameOuter, true);
			}
			break;
		case wordsEditorMode:
			if (currentlyOpenEditors.words) {
				editorFocus = newFocus;
				setFocusVisualState(wordsEditorDivNameOuter, true);
			}
			break;
		case standardEditorMode:
			if (!currentlyOpenEditors.paragraph && !currentlyOpenEditors.sentence) {
				editorFocus = newFocus;
			}
			break;
	}
}

function parseParagraphs(text) {
	// Split the text into paragraphs based on two or more newline characters
	const paragraphs = text.split(/\n{2,}/);
	let output = []
	for(let paragraph of paragraphs){
		paragraph = paragraph.trim();
		if(paragraph.length > 0)
			output.push( paragraph )	
	}
	
	return output;
}

function processDisplayParagraphsCommand() {
	if (
		!currentlyOpenEditors.paragraph &&
		fullDocTopElement.innerText.trim().length > 0
	) {
		setUndoBufferCurrent(["document", "windowState"]);
		let allParagraphs = parseParagraphs(fullDocTopElement.innerText);
		showDiv(paragraphEditorDivNameOuter);
		refreshParagraphEditor(allParagraphs);
		setScrollPercentageOfDiv(
			paragraphEditorDiv,
			getScrollPercentageOfDiv(fullDocTopElement),
		);
		hideDiv(fullDocumentDivName);
		currentlyOpenEditors.paragraph = true;
	}
	setEditorFocus(paragraphEditorMode);
}

function processDisplaySentencesCommand() {
	if (!currentlyOpenEditors.sentence) {
		let sentences_array = parseSentencesFromParagraph(
			aggregateElement.innerText.trim(),
		);
		if (sentences_array.length > 0) {
			setUndoBufferCurrent(["buffer", "alternatives", "windowState"]);
			refreshSentenceEditor(sentences_array);
			emptyBuffer();
			showDiv(sentenceEditorDivNameOuter);
			currentlyOpenEditors.sentence = true;
		}
	}
	setEditorFocus(sentenceEditorMode);
}

function processDisplayWordsEditorCommand() {
	if (!currentlyOpenEditors.words) {
		let tokens = splitStringIntoTokens(aggregateElement.innerText.trim());
		if (tokens.length > 0) {
			setUndoBufferCurrent(["buffer", "alternatives", "windowState"]);
			refreshWordsEditor(tokens);
			emptyBuffer();
			showDiv(wordsEditorDivNameOuter);
			currentlyOpenEditors.words = true;
		}
	}
	setEditorFocus(wordsEditorMode);
}

function refreshFullDocument(fullDocumentDivName, paragraphs) {
	emptyDivByName(fullDocumentDivName);
	if( paragraphs.length > 0 ){
		let output = []
		for (let paragraph of paragraphs) {
			paragraph = paragraph.trim();
			if(paragraph.length > 0)
				output.push(paragraph);
		}
		fullDocTopElement.innerText = output.join("\n\n");
	}
}

function closeParagraphEditor() {
	if (currentlyOpenEditors.paragraph) {
		setUndoBufferCurrent(["windowState", "paragraphs", "document"]);
		let allParagraphs = getAllParagraphsFromParagraphEditor();
		refreshFullDocument(fullDocumentDivName, allParagraphs);
		showDiv(fullDocumentDivName);
		//Set the document to the same position
		setScrollPercentageOfDiv(
			fullDocTopElement,
			getScrollPercentageOfDiv(paragraphEditorDiv),
		);

		hideDiv(paragraphEditorDivNameOuter);
		emptyDivByName(paragraphEditorDivName);
		currentlyOpenEditors.paragraph = false;
		resolveNewEditorFocus();
		preserveStateLocalStorage();
	}
}

function processCloseEditorCommand() {
	switch (editorFocus) {
		case sentenceEditorMode:
			closeSentenceEditor();
			break;
		case paragraphEditorMode:
			closeParagraphEditor();
			break;
		case wordsEditorMode:
			closeWordsEditor();
			break;
	}
}

function processScrollToNumberedItem(item_num) {
	switch (editorFocus) {
		case wordsEditorMode:
			scrollToNumberedPositionInWordsEditor(item_num);
			break;
		case sentenceEditorMode:
			scrollToNumberedPositionInSentencesEditor(item_num);
			break;
		case paragraphEditorMode:
			scrollToNumberedPositionInParagraphEditor(item_num);
			break;
	}
}

function processScrollDirectionCommand(up = true) {
	switch (editorFocus) {
		case wordsEditorMode:
			scrollDivDirection(wordsEditorDiv, up);
			break;
		case sentenceEditorMode:
			scrollDivDirection(sentencesEditorDiv, up);
			break;
		case paragraphEditorMode:
			scrollDivDirection(paragraphEditorDiv, up);
			break;
		default:
			scrollDivDirection(fullDocTopElement, up);
			break;
	}
}

function processCutBufferCommand(){
	const buffer = aggregateElement.innerText.trim();
	if(buffer.length > 0){
		setUndoBufferCurrent(["buffer"]);
		pasteBuffer = buffer;
		emptyBuffer();
	}
}

function processCutNumberedItemCommand(item_num) {
	switch (editorFocus) {
		case sentenceEditorMode:
			cutSentenceFromEditor(item_num);
			break;
		case paragraphEditorMode:
			cutParagraphFromEditor(item_num);
			break;
		case wordsEditorMode:
			cutWordFromEditor(item_num);
			break;
	}
}

function processCutAfterCommand(item_num, before=false) {
	switch (editorFocus) {
		case sentenceEditorMode:
			cutSentenceAfter(item_num, before);
			break;
		case paragraphEditorMode:
			cutParagraphAfter(item_num, before);
			break;
		case wordsEditorMode:
			cutWordAfter(item_num, before);
			break;
	}
}

function processRestoreNumberedItemCommand(item_num) {
	switch (editorFocus) {
		case sentenceEditorMode:
			let sentence = getSentenceFromEditor(item_num).trim();
			if (sentence != "") {
				emptyBuffer();
				aggregateElement.innerText = sentence;
				scrollDivEnd(aggregateElement, false);
				moveCursorToEnd( aggregateElement );
			}
			break;
		case paragraphEditorMode:
			let paragraph = getParagraphFromEditor(item_num).trim();
			if (paragraph != "") {
				emptyBuffer();
				aggregateElement.innerText = paragraph;
				scrollDivEnd(aggregateElement, false);
				scrollDivEnd(aggregateElement, false);
				moveCursorToEnd( aggregateElement );
			}
			break;
	}
}

function processReplaceNumberedItemCommand(item_num) {
	switch (editorFocus) {
		case wordsEditorMode:
			if (replaceWordAtEditor(item_num, aggregateElement.innerText, true));
			emptyBuffer();
			break;
		case sentenceEditorMode:
			if (replaceSentenceAtEditor(item_num, aggregateElement.innerText));
			emptyBuffer();
			break;
		case paragraphEditorMode:
			if (replaceParagraphAtEditor(item_num, aggregateElement.innerText))
				emptyBuffer();
			break;
	}
}

function processMoveInEditor(itemNum, up = true) {
	switch (editorFocus) {
		case sentenceEditorMode:
			moveSentenceInEditor(itemNum, up);
			break;
		case paragraphEditorMode:
			moveParagraphInEditor(itemNum, up);
			break;
	}
}

function processCutLastSentence() {
	setUndoBuffer({
		buffer: aggregateElement.innerText,
		alternatives: getCurrentAlternatives(),
	});
	aggregateElement.innerText = cutLastSentenceFromParagraph(
		aggregateElement.innerText,
	);
	scrollDivEnd(aggregateElement, false);
	moveCursorToEnd( aggregateElement );
}

function processPasteIntoBuffer() {
	if(pasteBuffer.length > 0){
		setUndoBufferCurrent(["buffer", "alternatives"]);
		appendToBuffer( pasteBuffer, true );
	}
}

function processCommands(commandResponse) {
	for (let command of commandResponse.commandList) {
		switch (command.normalisedCommandName) {
			case appendCommandName:
				processAppendCommand(aggregateElement.innerText);
				break;
			case emptyCommandName:
				setUndoBuffer({
					buffer: aggregateElement.innerText,
					alternatives: getCurrentAlternatives(),
				});
				emptyBuffer();
				break;
			case chooseAlternative:
				processChooseAlternativeCommand(command.parameter);
				break;
			case spellModeCommandName:
				toggleSpellMode();
				break;
			case wordsEditorNormalisedCommandName:
				processDisplayWordsEditorCommand();
				break;
			case displaySentencesCommandName:
				processDisplaySentencesCommand();
				break;
			case displayParagraphsCommandName:
				processDisplayParagraphsCommand();
				break;
			case closeEditorCommandName:
				processCloseEditorCommand();
				break;
			case cutBufferCommandName:
				processCutBufferCommand();
				break;
			case cutNumberedItem:
				processCutNumberedItemCommand(command.parameter);
				break;
			case cutAfterCommandName:
				processCutAfterCommand(command.parameter);
				break;
			case cutBeforeCommandName:
				processCutAfterCommand(command.parameter, true);
				break;
			case pasteIntoBufferCommandName:
				processPasteIntoBuffer();
				break;
			case pasteAfterNumberedItem:
				processInsertNumberedItemCommand(
					command.parameter,
					pasteBuffer,
					false,
					false,
				);
				break;
			case pasteBeforeNumberedItem:
				processInsertNumberedItemCommand(
					command.parameter,
					pasteBuffer,
					true,
					false,
				);
				break;
			case insertAfterNumberedItem:
				processInsertNumberedItemCommand(
					command.parameter,
					aggregateElement.innerText,
					false,
					true,
				);
				break;
			case insertBeforeNumberedItem:
				processInsertNumberedItemCommand(
					command.parameter,
					aggregateElement.innerText,
					true,
					true,
				);
				break;
			case restoreNormalisedCommandName:
				processRestoreNumberedItemCommand(command.parameter);
				break;
			case replaceCommandName:
				processReplaceNumberedItemCommand(command.parameter);
				break;
			case upperCaseNormalisedCommandName:
			case lowerCaseNormalisedCommandName:
			case titleCaseNormalisedCommandName:
				processCaseCommand(command.parameter, command.normalisedCommandName);
				break;
			case shortenNormalisedCommandName:
				processShortenCommand(command.parameter);
				break;
			case moveUpCommandName:
				processMoveInEditor(command.parameter, true);
				break;
			case moveDownCommandName:
				processMoveInEditor(command.parameter, false);
				break;
			case undoCommandName:
				handleUndo();
				break;
			case cutLastSentenceCommand:
				processCutLastSentence();
				break;
			case scrollToNumberCommand:
				processScrollToNumberedItem(command.parameter);
				break;
			case scrollUpCommand:
				processScrollDirectionCommand(true);
				break;
			case scrollDownCommand:
				processScrollDirectionCommand(false);
				break;
		}
	}
}

//Take a string of characters and escape special characters appropriately for use in as a regular expression
function escapeRegExp(string) {
	return string.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

//Make sure each sentence starts with a capital letter
function autoFixCapitalLetters(buffer) {
	let sentences = parseSentencesFromParagraph(buffer);
	let final_result_sentences = [];
	for (let sentence of sentences) {
		final_result_sentences.push(capitaliseFirstLetter(sentence.trim()));
	}
	return final_result_sentences.join(" ");
}

//Handle special case exceptions where spaces may have been inserted incorrectly around punctation marks
//by the conventional set of replacements
function handlePunctuationExceptions(input) {
	input = removeSpacesFromInitialisms(input);
	input = removeSpacesFromDecimals(input);
	return input;
}

//Fix initialisms to remove spaces between dots
//Converts e.g. I. B. M. into I.B.M.
function removeSpacesFromInitialisms(input) {
	return input.replace(
		/\b(([\p{Uppercase_Letter}]\.\s){2,}[\p{Uppercase_Letter}]\.)/gu,
		(match) => {
			return match.replace(/\.\s/g, ".");
		},
	);
}

//Fix for incorrect spacing in decimal numbers. Convert e.g.  "12. 5" into "12.5"
function removeSpacesFromDecimals(input) {
	return input.replace(/\b(\d+\. \d+)/g, (match) => {
		return match.replace(/\.\s/g, ".");
	});
}

function insertStringAtIndex(originalString, stringToInsert, index) {
	return (
		originalString.slice(0, index) +
		stringToInsert +
		originalString.slice(index)
	);
}

//Remove the character at a specific index in a string
function removeCharacterAtIndex(str, index) {
	return str.substring(0, index) + str.substring(index + 1);
}

//Removes multiple spaces
function replaceMultipleSpacesWithSingle(text) {
	return text.replace(/\s+/g, " ");
}

function testForWordContractionAtIndex(utterance, charPos) {
	//If the quote is at the end or start of a string then accept
	if (charPos == utterance.length || charPos == 0) accepted = false;
	else {
		let beforeChar = utterance.charAt(charPos - 1);
		let afterChar = utterance.charAt(charPos + 1);
		if (/[\p{Letter}]/u.test(beforeChar) && /[\p{Letter}]/u.test(afterChar))
			return true;

		if (beforeChar == "s" && !/[\p{Letter}]/u.test(afterChar)) {
			return true;
		}

		return false;
	}
}

//Auto fix the spacing before and and after punctuation marks.
//e.g. a comma should not have a space before it but should have a space afterwards
function fixPunctuationSpacing(utterance) {
	utterance = replaceMultipleSpacesWithSingle(utterance);

	for (let substitution of autoReplacementPhrases) {
		let openQuote = true;

		//Look for instances where a space prior needs to be removed
		let searchTextEscaped = escapeRegExp(substitution.replacement);
		switch (substitution.spaceRules) {
			case punctRuleLikeComma:
			case punctRuleOnlySpaceAfter:
				{
					const regex = new RegExp(`(\\s${searchTextEscaped})`, "ig");
					utterance = utterance.replace(regex, substitution.replacement);
				}
				break;
		}

		//Look for instances where a space following needs to be removed
		switch (substitution.spaceRules) {
			case punctRuleLikeOpenBracket:
				{
					const regex = new RegExp(`(${searchTextEscaped}\\s)`, "ig");
					utterance = utterance.replace(regex, substitution.replacement);
				}
				break;
		}

		//Look for where a space after the substitution needs to be inserted
		switch (substitution.spaceRules) {
			case punctRuleLikeComma:
				{
					//Insert a space only if the character is immediately followed by a word
					const regex = new RegExp(`${searchTextEscaped}(\\p{Letter})`, "igu");
					utterance = utterance.replace(
						regex,
						substitution.replacement + " " + "$1",
					);
				}
				break;
			case punctRuleOnlySpaceAfter:
				{
					const regex = new RegExp(`${searchTextEscaped}([^\\s])`, "ig");
					utterance = utterance.replace(
						regex,
						substitution.replacement + " " + "$1",
					);
				}
				break;
		}

		//Look for where a space before the substitution needs to be inserted
		switch (substitution.spaceRules) {
			case punctRuleLikeOpenBracket:
				{
					//Insert a space only if the character is immediately preceeded by a word
					const regex = new RegExp(`(\\p{Letter})${searchTextEscaped}`, "igu");
					utterance = utterance.replace(
						regex,
						"$1 " + substitution.replacement,
					);
				}
				break;
		}

		//Handle quote-like characters where the rule on spaces is different depending on whether it is an open or close quote
		//Tries to handle contractions and possessives.
		if (substitution.spaceRules === punctRuleLikeQuote) {
			let startIndex = 0;
			let matchIndex = 0;

			while (matchIndex != -1 && startIndex < utterance.length) {
				let bytesAppended = 0;

				let accepted = false;

				matchIndex = utterance.indexOf(substitution.replacement, startIndex);

				while (!accepted && matchIndex >= 0 && startIndex < utterance.length) {
					matchIndex = utterance.indexOf(substitution.replacement, startIndex);
					accepted = true;

					//If we are dealing with a single quote, then check for contractions in words and ignore.
					if (substitution.replacement == "'" && matchIndex > 0) {
						if (testForWordContractionAtIndex(utterance, matchIndex)) {
							accepted = false;
							startIndex = matchIndex + 1;
							if (startIndex >= utterance.length) matchIndex = -1;
						}
					}
				}

				//if the match was not found
				if (matchIndex < 0) break;

				//if we're current on an open quote, we want to insert a space before the quote (if not already present) and remove any space after the quote
				if (openQuote) {
					if (matchIndex != 0 && utterance.charAt(matchIndex - 1) != " ") {
						utterance = insertStringAtIndex(utterance, " ", matchIndex);
						bytesAppended++;
					}
					let charIndexAfterMatch =
						matchIndex + substitution.replacement.length + bytesAppended;

					if (charIndexAfterMatch < utterance.length) {
						if (utterance.charAt(charIndexAfterMatch) == " ") {
							utterance = removeCharacterAtIndex(
								utterance,
								charIndexAfterMatch,
							);
							bytesAppended--;
						}
					}
					//Just had an open quote so must be close quote next
					openQuote = false;
				} else {
					//We're current on a close quote. Remove any space prior to the quote and insert a space after the quote (if not present).

					let charIndexAfterMatch =
						matchIndex + substitution.replacement.length;

					if (charIndexAfterMatch < utterance.length) {
						let charValAfter = utterance.charAt(charIndexAfterMatch);
						if (/\p{Letter}/u.test(charValAfter)) {
							utterance = insertStringAtIndex(
								utterance,
								" ",
								charIndexAfterMatch,
							);
							bytesAppended++;
						}
					}

					if (matchIndex != 0 && utterance.charAt(matchIndex - 1) == " ") {
						utterance = removeCharacterAtIndex(utterance, matchIndex - 1);
						bytesAppended--;
					}

					openQuote = true;
				}

				startIndex =
					matchIndex + substitution.replacement.length + bytesAppended + 1;
			}
		}
	}

	utterance = handlePunctuationExceptions(utterance);
	return utterance.trim();
}

//Removes instances of "literal" (or the current user escape command phrase) except for "literal literal" which is converted to "literal".
function removeInstancesOfEscapeKeyword(input, escapeTriggerWord) {
	// Construct the regex patterns using the escapeTriggerWord
	const doubleWordPattern = new RegExp(
		`\\b${escapeTriggerWord} ${escapeTriggerWord}\\b`,
		"g",
	);
	const singleWordPattern = new RegExp(
		`(?<!\\b${escapeTriggerWord})(?:\\s|^)${escapeTriggerWord}\\b`,
		"g",
	);

	// Remove any remaining standalone "escape"
	let result = input.replace(singleWordPattern, "");

	// Replace "escape escape" with "escape"
	result = result.replace(doubleWordPattern, escapeTriggerWord);

	// Trim any leading or trailing whitespace
	return result.trim();
}

function doAutoReplacements(
	utterance,
	autoReplacementPhrases,
	escapeTriggerWord,
) {
	//Sort phrases by length descending
	//Prevents subsets being considered first e.g. the phrase "semi colon" is substituted prior to "colon" which is a subset of the phrase
	let lengthOrderedReplacementList = autoReplacementPhrases.sort(
		(item1, item2) => item2.userPhrase.length - item1.userPhrase.length,
	);

	//First pass: just replace all of the punctuation phrases with the relevant character and don't worry about the correct spacing
	//There will be spaces in the wrong place with respect punctation symbols. This will be corrected later.
	for (let phrase of lengthOrderedReplacementList) {
		let userPhrase = phrase.userPhrase;
		let replacement = phrase.replacement;
		const regex = new RegExp(
			`(?<!\\${escapeTriggerWord}\\s)\\b((?:${phrase.userPhrase}))\\b`,
			"ig",
		);
		if (regex.test(utterance)) {
			utterance = utterance.replace(regex, replacement);
		}
	}

	return utterance;
}

function stringToInt(number) {
	switch (number) {
		case "zero":
		case "first":
		case "start":
		case "beginning":
			return 0;
		case "one":
			return 1;
		case "two":
			return 2;
		case "to":
			return 2;
		case "too":
			return 2;
		case "three":
			return 3;
		case "four":
			return 4;
		case "for":
			return 4;
		case "five":
			return 5;
		case "six":
			return 6;
		case "seven":
			return 7;
		case "eight":
			return 8;
		case "nine":
			return 9;
		case "last":
		case "end":
			return 999999;
		default:
			let value = parseInt(number, 10);
			if (isNaN(value)) {
				return -1;
			}
			return value;
	}
}

//Determine what commands were spoken and remove the from the utterance
function interpretCommands(
	triggerWord,
	commandList,
	utterance,
	escapeTriggerWord,
) {
	let response = {
		ignoreUtternance: false, //If "command ignore" was said or
		filteredUtterance: utterance.trim(), //The utterance with any commands removed
		commandList: [],
	};

	if (!wasCommandWordSaid(triggerWord, utterance)) return response;

	let filteredUtterance = utterance;

	//Prevent subsets of longer phrases being recognised first by sorting the command phrases in descending order of length so e.g. command deleted comes before command delete
	//Sort items with a parameter as slightly longer so e.g. command delete 5 sorts before command delete
	let lengthOrderedCommandList = commandList.sort(
		(item1, item2) =>
			item2.userPhrase.length +
			(item2.param ? 1 : 0) -
			(item1.userPhrase.length + (item1.param ? 1 : 0)),
	);

	for (const command of lengthOrderedCommandList) {
		//Check for commands which do not require any parameters
		const userPhrase = command.userPhrase;
		if (command.param === false) {
			const regex = new RegExp(
				`(?<!\\${escapeTriggerWord}\\s)(\\b(?:${triggerWord})\\s(?:${userPhrase}))\\b`,
				"ig",
			);
			const matches = filteredUtterance.matchAll(regex);
			for (const match of matches) {
				response.commandList.push({
					normalisedCommandName: command.normalisedCommand,
				});

				if (command.normalisedCommand === undoCommandName)
					response.ignoreUtternance = true;
				
			}
			filteredUtterance = filteredUtterance.replace(regex, "");
		} else {
			//Check for commands that have a single integer parameter
			if (command.paramType === "single_int_first_last") {
				const numbers =
					"(?:\\d+|first|start|beginning|zero|one|two|to|too|three|four|for|five|six|seven|eight|nine|last|end)";
				const regex = new RegExp(
					`(?<!\\${escapeTriggerWord}\\s)(\\b(?:${triggerWord})\\s(?:${userPhrase})\\s${numbers})\\b`,
					"ig",
				);
				const matches = filteredUtterance.matchAll(regex);
				for (const match of matches) {
					const regexExtractNumber = new RegExp(
						`\\b(?:${triggerWord})\\s(?:${userPhrase})\\s(${numbers})\\b`,
						"i",
					);
					const match = utterance.match(regexExtractNumber);
					if (match) {
						let parameter = stringToInt(match[1]);
						if (parameter >= 0)
							response.commandList.push({
								normalisedCommandName: command.normalisedCommand,
								parameter: parameter,
							});
					}
				}
				filteredUtterance = filteredUtterance.replace(regex, "");
			}
		}
	}

	response.filteredUtterance = filteredUtterance.trim();

	return response;
}

function aggregateElementContentChange() {
	//Set the undo state to what is would have been before the user changed it
	setUndoBuffer(bufferUndoStateBeforeUserEdit);
	refreshAlternatives([]);
	bufferUndoStateBeforeUserEdit = {
		buffer: aggregateElement.innerText.trim(),
		alternatives: getCurrentAlternatives(),
	};
	resetMicInactivityTimer();
}

function documentElementContentChange() {
	//Set the undo state to what is would have been before the user changed it
	setUndoBuffer(documentUndoStateBeforeUserEdit);
	bufferUndoStateBeforeUserEdit = {
		document: fullDocTopElement.innerText.trim(),
	};
	preserveStateLocalStorage();
	resetMicInactivityTimer();
}

function wasCommandWordSaid(triggerWord, utterance) {
	const regex = new RegExp(`(?:^|\\s)${triggerWord}(?:\\s+|\$)`, "i");
	if (regex.test(utterance)) {
		return true;
	}

	return false;
}

function capitaliseFirstLetter(str, lowerCase = false) {
	// Use a regular expression to find the first letter
	const regex = /[\p{Letter}]/u;
	const match = str.match(regex);

	// If no letter is found, return the original string
	if (!match) {
		return str;
	}

	// Get the index of the first letter
	const index = match.index;

	let capitalisedLetter = str[index].toUpperCase();
	if (lowerCase) capitalisedLetter = str[index].toLowerCase();

	// Reconstruct the string
	return str.slice(0, index) + capitalisedLetter + str.slice(index + 1);
}

function debounce(func, delay) {
	let timeout;
	return function (...args) {
		clearTimeout(timeout);
		timeout = setTimeout(() => func.apply(this, args), delay);
	};
}

function moveItemInArray(array, index, moveUp = true) {
	// Create a copy of the array to avoid modifying the original
	let newArray = [...array];

	if (moveUp && index > 0) {
		[newArray[index - 1], newArray[index]] = [
			newArray[index],
			newArray[index - 1],
		];
	} else if (!moveUp && index < newArray.length - 1) {
		[newArray[index], newArray[index + 1]] = [
			newArray[index + 1],
			newArray[index],
		];
	}

	return newArray;
}

function preserveStateLocalStorage() {
	try {
		if( currentlyOpenEditors['paragraph'] == true ){
		
			localStorage.setItem(
				editorStorageKey + "lastDocument",
				getAllParagraphsFromParagraphEditor().join("\n\n").trim()
			);
			
		}else{
			localStorage.setItem(
				editorStorageKey + "lastDocument",
				fullDocTopElement.innerHTML.trim()
			);
		}
		
	} catch (e) {
		if (e instanceof DOMException && e.name === "QuotaExceededError") {
			displayMessage(
				"Warning: can't auto-save document because local storage is full.",
			);
		}
	}
}

function restoreDocumentFromLocalStorage() {
	const documentKey = editorStorageKey + "lastDocument";

	if (localStorage.hasOwnProperty(documentKey)) {
		let store_j = localStorage.getItem(documentKey);
		fullDocTopElement.innerHTML = store_j;
	}
}

function setEvents() {
	microphoneButton.addEventListener("click", setRecordingState);

	document
		.getElementById("close-button")
		.addEventListener("click", function () {
			hideDiv("floating-message");
		});

	document.getElementById("undo-button").addEventListener("click", function () {
		handleUndo();
	});

	document
		.getElementById("spell-button")
		.addEventListener("click", function () {
			toggleSpellMode();
		});

	document
		.getElementById("config-button")
		.addEventListener("click", function () {
			toggleDivVisible(configAreaDivName);
		});

	aggregateElement.addEventListener(
		"input",
		debounce(aggregateElementContentChange, 1000),
	);
	fullDocTopElement.addEventListener(
		"input",
		debounce(documentElementContentChange, 1000),
	);
}

function toggleSpellMode() {
	if (spellModeOn) {
		spellModeOn = false;
		withSpace = true;
		setButtonHighlight(spellButtonElement, false);
	} else {
		spellModeOn = true;
		setButtonHighlight(spellButtonElement, true);
	}
}

function setButtonHighlight(buttonElement, state) {
	if (state) buttonElement.style.backgroundColor = "#A7C7E7";
	else buttonElement.style.backgroundColor = "";
}


function moveCursorToEnd(element) {
	const range = document.createRange();
	range.selectNodeContents(element);
	range.collapse(false);
	const selection = window.getSelection();
	selection.removeAllRanges();
	selection.addRange(range);
	element.focus();
}
