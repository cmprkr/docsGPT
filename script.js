function onOpen() {
  var ui = DocumentApp.getUi();
  var menu = ui.createMenu('docsGPT');
  
  menu.addItem('Tone', 'toggleOptions')
  menu.addSubMenu(ui.createMenu('Writing')
    .addItem('Write Paragraph', 'writeEssay')
    .addItem('Expand', 'expandText')
    .addItem('Expand Thesis', 'thesisExpand')
    .addItem('Summarize Text', 'summarizeText'));

  menu.addSubMenu(ui.createMenu('Research')
    .addItem('Generate Ideas', 'generateIdeas')
    .addItem('Find Sources', 'findSources')
    .addItem('Research Topic', 'researchTopic'));

  menu.addItem('Answer Question', 'answerQuestion');
  
  menu.addToUi();
}

// FIXED VARIABLES. Your API and Model Type
var apiKey = "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";
var model = "gpt-3.5-turbo";
// ****END VARIABLES****

function toggleOptions() {
  var ui = DocumentApp.getUi();
  var response = ui.prompt('Tone Selector', 'Formatting tone:', ui.ButtonSet.OK_CANCEL);
  console.log(response)

  if (response.getSelectedButton() == ui.Button.OK) {
    var isChecked = response.getResponseText() == 'Yes';
    // Do something with the isChecked value
    ui.alert('Tone has been updated successfully.');
  }
}

function generateIdeas() {
 var doc = DocumentApp.getActiveDocument();
 var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();


 // Check if there is selected text in the document
 if (!selectedTextElement.asText) {
   return;
 }


 var selectedText = selectedTextElement.asText().getText();
 var body = doc.getBody();
 var prompt = "Generate ideas for: " + selectedText;
 var messages = [{
   "role": "user",
   "content": prompt
 }];


 var requestOptions = {
   "method": "POST",
   "headers": {
     "Content-Type": "application/json",
     "Authorization": "Bearer " + apiKey
   },
   "payload": JSON.stringify({
     "model": model,
     "messages": messages
   })
 };


 var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
 var responseText = response.getContentText();
 var jsonResponse = JSON.parse(responseText);


 var generatedText = jsonResponse["choices"][0]["message"]["content"];


 // Check if generatedText is defined before calling .trim()
 if (generatedText) {
   generatedText = generatedText.trim();
 }


 // Set a default value for generatedText if it is undefined or empty
 generatedText = generatedText || "";


 // Convert the generatedText to plain text
 var plainText = Utilities.formatString("%s", generatedText);


 // Add the plain text to the document
 var para = body.appendParagraph(plainText);
}

function writeEssay() {
  var doc = DocumentApp.getActiveDocument();
  var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();

  // Check if there is selected text in the document
  if (!selectedTextElement.asText) {
    return;
  }

  var selectedText = selectedTextElement.asText();
  var body = doc.getBody();
  var prompt = {
    "prompt": "Write a formal paragraph on: " + selectedText.getText() + '. Make sure that the paragraph is formatted in a scholarly manner with 5-8 sentances. The lamguage used should be above high school level.',
    "temperature": 0.5,
    "max_tokens": 1000,
    "frequency_penalty": 0.4,
    "presence_penalty": 0.4,
    "model": "text-davinci-002"
  };

  var requestOptions = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(prompt)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", requestOptions);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);
  var generatedText = jsonResponse["choices"][0]["text"];

  // Check if generatedText is defined before calling .trim()
  if (generatedText) {
    generatedText = generatedText.trim();
  }

  // Replace the selected text with the generated essay
  selectedText.replaceText(selectedText.getText(), generatedText);
}

function summarizeText() {
  var doc = DocumentApp.getActiveDocument();
  var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();

  // Check if there is selected text in the document
  if (!selectedTextElement.asText) {
    return;
  }

  var selectedText = selectedTextElement.asText();
  var body = doc.getBody();
  var prompt = {
    "prompt": "Summarize the following into concise bullet points: " + selectedText.getText(),
    "temperature": 0.5,
    "max_tokens": 1000,
    "frequency_penalty": 0.4,
    "presence_penalty": 0.4,
    "model": "text-davinci-002"
  };

  var requestOptions = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(prompt)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", requestOptions);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);
  var generatedText = jsonResponse["choices"][0]["text"];

  // Check if generatedText is defined before calling .trim()
  if (generatedText) {
    generatedText = generatedText.trim();
  }

  // Replace the selected text with the generated essay
  selectedText.replaceText(selectedText.getText(), generatedText);
}

function answerQuestion() {
  var doc = DocumentApp.getActiveDocument();
  var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();

  // Check if there is selected text in the document
  if (!selectedTextElement.asText) {
    return;
  }

  var selectedText = selectedTextElement.asText();
  var body = doc.getBody();
  var prompt = {
    "prompt": "Answer the following question in a clear manner, explaining it thoroughly: " + selectedText.getText() + ".",
    "temperature": 0.2,
    "max_tokens": 1000,
    "frequency_penalty": 0.4,
    "presence_penalty": 0.4,
    "model": "text-davinci-002"
  };

  var requestOptions = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(prompt)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", requestOptions);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);
  var generatedText = jsonResponse["choices"][0]["text"];

  // Check if generatedText is defined before calling .trim()
  if (generatedText) {
    generatedText = generatedText.trim();
  }

  // Replace the selected text with the generated essay
  selectedText.replaceText(selectedText.getText(), generatedText);
}

function researchTopic() {
  var doc = DocumentApp.getActiveDocument();
  var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();

  // Check if there is selected text in the document
  if (!selectedTextElement.asText) {
    return;
  }

  var selectedText = selectedTextElement.asText();
  var body = doc.getBody();
  var prompt = {
    "prompt": "Research the following topic as thouroughly as possible. Make sure to display the information clearly:  " + selectedText.getText() + ".",
    "temperature": 0.4,
    "max_tokens": 1000,
    "frequency_penalty": 0.4,
    "presence_penalty": 0.4,
    "model": "text-davinci-002"
  };

  var requestOptions = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(prompt)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", requestOptions);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);
  var generatedText = jsonResponse["choices"][0]["text"];

  // Check if generatedText is defined before calling .trim()
  if (generatedText) {
    generatedText = generatedText.trim();
  }

  // Replace the selected text with the generated essay
  selectedText.replaceText(selectedText.getText(), generatedText);
}

function expandText() {
  var doc = DocumentApp.getActiveDocument();
  var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();

  // Check if there is selected text in the document
  if (!selectedTextElement.asText) {
    return;
  }

  var selectedText = selectedTextElement.asText();
  var body = doc.getBody();
  var prompt = {
    "prompt": "Take the following text and expand it, delving into the topic more clearly. Use a formal writing style:  " + selectedText.getText(),
    "temperature": 0.4,
    "max_tokens": 1000,
    "frequency_penalty": 0.4,
    "presence_penalty": 0.4,
    "model": "text-davinci-002"
  };

  var requestOptions = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(prompt)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", requestOptions);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);
  var generatedText = jsonResponse["choices"][0]["text"];

  // Check if generatedText is defined before calling .trim()
  if (generatedText) {
    generatedText = generatedText.trim();
  }

  // Replace the selected text with the generated essay
  selectedText.replaceText(selectedText.getText(), generatedText);
}

function thesisExpand() {
  var doc = DocumentApp.getActiveDocument();
  var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();

  // Check if there is selected text in the document
  if (!selectedTextElement.asText) {
    return;
  }

  var selectedText = selectedTextElement.asText();
  var body = doc.getBody();
  var prompt = {
    "prompt": "Using the following thesis, write a formal, in-depth essay:  " + selectedText.getText(),
    "temperature": 0.4,
    "max_tokens": 1000,
    "frequency_penalty": 0.4,
    "presence_penalty": 0.4,
    "model": "text-davinci-002"
  };

  var requestOptions = {
    "method": "POST",
    "headers": {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    "payload": JSON.stringify(prompt)
  };

  var response = UrlFetchApp.fetch("https://api.openai.com/v1/completions", requestOptions);
  var responseText = response.getContentText();
  var jsonResponse = JSON.parse(responseText);
  var generatedText = jsonResponse["choices"][0]["text"];

  // Check if generatedText is defined before calling .trim()
  if (generatedText) {
    generatedText = generatedText.trim();
  }

  // Replace the selected text with the generated essay
  selectedText.replaceText(selectedText.getText(), generatedText);
}

function findSources() {
 var doc = DocumentApp.getActiveDocument();
 var selectedTextElement = doc.getSelection().getRangeElements()[0].getElement();


 // Check if there is selected text in the document
 if (!selectedTextElement.asText) {
   return;
 }


 var selectedText = selectedTextElement.asText().getText();
 var body = doc.getBody();
 var prompt = "Find scholarly sources (websites and books) for : " + selectedText + ". Just give me a few potential sources that might be relevant, do not write about how you cannot perform the task as an AI Language Model.";
 var messages = [{
   "role": "user",
   "content": prompt
 }];


 var requestOptions = {
   "method": "POST",
   "headers": {
     "Content-Type": "application/json",
     "Authorization": "Bearer " + apiKey
   },
   "payload": JSON.stringify({
     "model": "gpt-3.5-turbo",
     "messages": messages
   })
 };


 var response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", requestOptions);
 var responseText = response.getContentText();
 var jsonResponse = JSON.parse(responseText);


 var generatedText = jsonResponse["choices"][0]["message"]["content"];


 // Check if generatedText is defined before calling .trim()
 if (generatedText) {
   generatedText = generatedText.trim();
 }


 // Set a default value for generatedText if it is undefined or empty
 generatedText = generatedText || "";


 // Convert the generatedText to plain text
 var plainText = Utilities.formatString("%s", generatedText);


 // Add the plain text to the document
 var para = body.appendParagraph("\n");
 var para = body.appendParagraph(plainText);
}
// ****END PROMPT****