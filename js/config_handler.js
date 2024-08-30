import {
	getStorageKey,
	getAutoReplacements,
	configPunctuationMenuChoices,
	updateAutoReplacements,
	getCommandList,
	setCommandList,
	getTriggerWord,
	setTriggerWord,
	getEscapeWord,
	setEscapeWord,
	setRecognitionLanguage,
} from "./main.js";


let supportedLanguages = {
	'English (US)' : 'en-US',
	'English (UK)' : 'en-GB'	
};

function isLanguageSupported( language ){
	for (const language in supportedLanguages) {
		if( supportedLanguages[language] ===  navigator.language)
			return true;
	}
	return false;
}

function addLanguageListBox() {
  // Create the select element
  
  const langSelectDiv = document.getElementById("language-selection-area");
  
  const selectElement = document.createElement('select');
   
  // Iterate over the keys of the supportedLanguages object
  
  
  for (const language in supportedLanguages) {
    if (supportedLanguages.hasOwnProperty(language)) {
      // Create an option element for each language
      const optionElement = document.createElement('option');
      optionElement.value = supportedLanguages[language];
      if( supportedLanguages[language] ===  navigator.language){
      	optionElement.selected = true;
      	setRecognitionLanguage(navigator.language);
      }
      optionElement.textContent = language;
      // Append the option to the select element
      selectElement.appendChild(optionElement);
    }
  }
  
  // Add an event listener for the change event
  selectElement.addEventListener('change', function(event) {
    // Call the provided callback function with the selected value
    setRecognitionLanguage(event.target.value);
  });
  
  langSelectDiv.appendChild(selectElement);

}

const autoreplacementsConfig = document.getElementById(
	"autoreplacements-config",
);
const commandsConfig = document.getElementById("commands-config");
const commandWordInput = document.getElementById("command-word-config");
const escapeWordInput = document.getElementById("escape-word-config");

function populateAutoReplacements() {
	autoreplacementsConfig.innerHTML = "";

	const addReplcementButton = document.getElementById("add-replacement");

	addReplcementButton.addEventListener("click", () => {
		addAutoReplacement();
	});

	getAutoReplacements().forEach((phrase) => {
		addAutoReplacement(
			phrase.userPhrase,
			phrase.replacement,
			phrase.spaceRules,
		);
	});
}

function populateCommandListing() {
	commandsConfig.innerHTML = "";
	let commandList = getCommandList();
	commandList.sort((a, b) =>
		a.normalisedCommand.localeCompare(b.normalisedCommand),
	);

	for (let i = 0; i < commandList.length; i++) {
		let command = commandList[i];

		const comContainer = document.createElement("div");
		comContainer.className = "comContainer";

		const userPhraseInput = document.createElement("input");
		userPhraseInput.type = "text";
		userPhraseInput.value = command.userPhrase;
		userPhraseInput.classList.add("user-phrase-list-box");
		userPhraseInput.placeholder = "Spoken Phrase";
		userPhraseInput.id = "command-" + i;

		const userPhraseLabel = document.createElement("label");
		userPhraseLabel.htmlFor = userPhraseInput.id; // Associate the label with the input
		userPhraseLabel.textContent = command.normalisedCommand;
		userPhraseLabel.classList.add("user-phrase-label"); // Optional: add a class for styling

		userPhraseInput.addEventListener("change", function (event) {
			updateCommandPhrase(command.normalisedCommand, event.target.value);
		});

		comContainer.appendChild(userPhraseInput);
		comContainer.appendChild(userPhraseLabel);
		commandsConfig.appendChild(comContainer);
	}
}

function populateCommandLiteral(){
	commandWordInput.value = getTriggerWord();
	escapeWordInput.value = getEscapeWord();
}

function updateCommandPhrase(normalisedName, newPhrase) {
	let commandList = getCommandList();
	for (let command of commandList) {
		if (command.normalisedCommand === normalisedName) {
			command.userPhrase = newPhrase.trim();
			setCommandList(commandList);
			break;
		}
	}
	updateLocalStorageConfig();
}

function createPunctuationOptionsSelect(spaceRules) {
	const selectElement = document.createElement("select");
	selectElement.id = "punctuation_rules";
	selectElement.name = "punctuation_rules";

	configPunctuationMenuChoices.forEach((choice) => {
		const optionElement = document.createElement("option");
		optionElement.value = choice;
		optionElement.textContent = choice;
		if (choice === spaceRules) optionElement.selected = true;
		else optionElement.selected = false;
		selectElement.appendChild(optionElement);
	});

	return selectElement;
}

function addAutoReplacement(userPhrase, replacement, spaceRules) {
	const arContainer = document.createElement("div");
	arContainer.className = "arContainer";
	
	if( userPhrase )
		userPhrase = userPhrase.trim()

	const userPhraseInput = document.createElement("input");
	userPhraseInput.type = "text";
	userPhraseInput.value = userPhrase || "";
	userPhraseInput.classList.add("user-phrase-list-box");
	userPhraseInput.placeholder = "Spoken Phrase";

	let punctSelect = createPunctuationOptionsSelect(spaceRules);

	const replacementInput = document.createElement("input");
	replacementInput.type = "text";
	replacementInput.value = replacement || "";
	replacementInput.classList.add("replacement-list-box");
	replacementInput.placeholder = "Replacement";

	const deleteButton = document.createElement("button");
	deleteButton.classList.add("delete-icon");
	deleteButton.setAttribute("aria-label", "Delete phrase: " + userPhrase);
	deleteButton.addEventListener("click", () => {
		autoreplacementsConfig.removeChild(arContainer);
	});

	userPhraseInput.addEventListener("change", function (event) {
		saveReplacements();
	});
	replacementInput.addEventListener("change", function (event) {
		saveReplacements();
	});
	punctSelect.addEventListener("change", function (event) {
		saveReplacements();
	});

	arContainer.appendChild(userPhraseInput);
	arContainer.appendChild(replacementInput);
	arContainer.appendChild(punctSelect);
	arContainer.appendChild(deleteButton);

	autoreplacementsConfig.appendChild(arContainer);
}

function saveReplacements() {
	let replacements = Array.from(
		autoreplacementsConfig.getElementsByTagName("input"),
	).map((input) => input.value);
	let punctRules = Array.from(
		autoreplacementsConfig.getElementsByTagName("select"),
	).map((input) => input.value);

	let newPhrases = [];

	// You can add code here to send the tasks array to a server or save it locally
	for (let i = 0; i < replacements.length / 2; i++) {
		let phrase = replacements[i * 2];
		let replacement = replacements[i * 2 + 1];
		let punctChoice = punctRules[i];
		if (phrase.length > 0 && replacement.length > 0)
			newPhrases.push({
				userPhrase: phrase.trim(),
				replacement: replacement,
				spaceRules: punctChoice,
			});
	}

	updateAutoReplacements(newPhrases);
	updateLocalStorageConfig();
}

function updateLocalStorageConfig() {
	const configKey = getStorageKey() + "config";
	let config = {
		autoreplacements: getAutoReplacements(),
		commands: getCommandList(),
		triggerWord: getTriggerWord(),
		escapeWord: getEscapeWord()
	};
	const serializedData = JSON.stringify(config);
	localStorage.setItem(configKey, serializedData);
	console.log("wrote storage");
}

function retrieveConfigLocalStorage() {
	const configKey = getStorageKey() + "config";

	if (localStorage.hasOwnProperty(configKey)) {
		let store_j = localStorage.getItem(configKey);
		const retrievedData = JSON.parse(store_j);
		
		if( retrievedData.autoreplacements )
			updateAutoReplacements(retrievedData.autoreplacements);
		if( retrievedData.commands )
			setCommandList(updateUserCommandPhrases(retrievedData.commands));
		if( retrievedData.triggerWord )
			setTriggerWord(retrievedData.triggerWord);
		if( retrievedData.escapeWord )
			setEscapeWord(retrievedData.escapeWord);
	}
}

function updateUserCommandPhrases(userSetFromStorage) {
	let originalCommands = getCommandList();
	let userReplacements = userSetFromStorage.reduce((acc, command) => {
		acc[command.normalisedCommand] = command.userPhrase;
		return acc;
	}, {});

	for (let command of originalCommands) {
		if (command.normalisedCommand in userReplacements) {
			command["userPhrase"] = userReplacements[command.normalisedCommand];
		}
	}
	return originalCommands;
}

function deleteAllLocalStorage(){
	for (let i = localStorage.length - 1; i >= 0; i--) {
		const key = localStorage.key(i);
		const prefix = getStorageKey()
		if ( key.startsWith( prefix ) ) {
			localStorage.removeItem(key);
		}
	}
	location.reload();
}

function updateCommandLiteral(){
	const commandWordVal = commandWordInput.value;
	const escapeWordVal = escapeWordInput.value;
	if( commandWordVal )
		setTriggerWord(commandWordVal.trim());
	if( escapeWordVal )
		setEscapeWord(escapeWordVal.trim());
	updateLocalStorageConfig();
}

document.addEventListener("DOMContentLoaded", () => {
	retrieveConfigLocalStorage();
	populateAutoReplacements();
	populateCommandListing();
	populateCommandLiteral();
	addLanguageListBox();
			
	const deleteAllDataButton = document.getElementById("delete-all-button");
	deleteAllDataButton.addEventListener("click", () => { deleteAllLocalStorage(); });
	
	commandWordInput.addEventListener("change", function (event) {
		updateCommandLiteral();
	});
	
	escapeWordInput.addEventListener("change", function (event) {
		updateCommandLiteral();
	});
});
