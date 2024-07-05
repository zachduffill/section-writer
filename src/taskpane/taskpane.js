/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

var targetCountDict = new Map();
var wordCountDict = new Map();
var totalWordCountExcludingSelected = 0;

var ccIdToEdit;

var lastCcIdSelected = 0;

var defaultTargetCount;
var defaultColor;

var totalWordCountElement;
var colorPickerElement;
var defaultTargetCountElement;

Office.onReady(() => {
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  totalWordCountElement = document.getElementById("word-count-num");
  colorPickerElement = document.getElementById("sectionColor");
  defaultTargetCountElement = document.getElementById("wordTargetCount");

  // Set default settings from local storage
  defaultTargetCount = getFromLocalStorage("defaultTarget");
  defaultColor = getFromLocalStorage("defaultColor");
  if (defaultTargetCount == null){
    defaultTargetCount = 500;
    setInLocalStorage("defaultTarget",defaultTargetCount);
  }
  if (defaultColor == null){
    defaultColor = "#F0F000";
    setInLocalStorage("defaultColor",defaultColor);
  }
  defaultTargetCountElement.value = defaultTargetCount;
  colorPickerElement.value = defaultColor;

  loadAllCcTargetCounts();
  eventSubscribe();
});

// Edit section dialog popup
async function editSection(event){
  await Word.run(async (context) => {
    try{
      let selection = context.document.getSelection();
      let cc = selection.parentContentControl; //catches if selection is not inside cc
      cc.load(["id","color"]);
      await context.sync();

      console.log(ccIdToEdit);
      ccIdToEdit = cc.id;

      // set input current values in popup
      let sColor = cc.color;
      let target = targetCountDict.get(cc.id);

      Office.context.ui.displayDialogAsync(`https://zachduffill.github.io/section-writer/dist/editDialog.html?color=${sColor.slice(1)}&target=${target}`, {height: 14, width: 15},  // Had to get the cc again from id within the callback event, as cannot retain cc object in both contexts, even with trackedobjects
        function (asyncResult) {
            let dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, 
              async function(msg){ 
                // set new values if ok btn pressed
                if (msg.message != ""){
                  let ccs = context.document.contentControls;
                  ccs.load("items");
                  await context.sync();

                  let cc = ccs.getById(ccIdToEdit);
                  cc.load(["color,title"]);
                  await context.sync();

                  let newValues = JSON.parse(msg.message);
                  // cc color
                  cc.color = newValues[0];
                  // target count
                  targetCountDict.set(cc.id,newValues[1]);
                  cc.title = wordCountDict.get(cc.id) + " / " + newValues[1];
                  await context.sync();
                }
                dialog.close();
              }
            );
        }
      );
    }
    catch(err){
      if (err.message != "ItemNotFound" && err.message != "Wait until the previous call completes.") console.log(err);
    }
    finally{
      event.completed();
    }
  });
}

// load CC target and wordcount info from all CC titles
async function loadAllCcTargetCounts(){
  await Word.run(async (context) => {
    var totalWordCount = 0;

    contentControls = context.document.contentControls;
    contentControls.load("items");
    await context.sync();

    let ccs = contentControls.items;
    for (i=0; i<ccs.length; i++){
      try{
        let countAndTargetCount = ccs[i].title.split(" / ");
        wordCountDict.set(ccs[i].id,Number(countAndTargetCount[0]));
        targetCountDict.set(ccs[i].id,Number(countAndTargetCount[1]));
        totalWordCount += Number(countAndTargetCount[0]);
      }
      catch(err){console.error(err);}
    }
    totalWordCountElement.innerHTML = totalWordCount;
  }); 
}

async function eventSubscribe(){
  await Word.run(async (context) => {
    var events = setInterval(
      checkForCcChanges,
      250,
    );
  });  
}

// allow subsections?
// checks for wordcount changes periodically in all CCs
async function checkForCcChanges(){
  await Word.run(async (context) => {
    try{
      let selection = context.document.getSelection();
      let cc = selection.parentContentControl; //catches if selection is not inside cc
      
      cc.load(["id","text","title"]);
      await context.sync();

      // recalculate totalwordcountexcludingselected if selected cc changes
      if (cc.id != lastCcIdSelected) await selectedCcChanged(cc.id);

      let wordTargetCount = String(targetCountDict.get(cc.id));
      let words = splitWords(cc.text);
      let wordCount = words.length;

      if (wordCount!=wordCountDict.get(cc.id)) { // if there has been a change in wordcount
        wordCountDict.set(cc.id,wordCount);
        totalWordCountElement.innerHTML = totalWordCountExcludingSelected+wordCount;

        lastCcIdSelected = cc.id;

        // Update Title Information
        if (wordCount>=0 && wordTargetCount>0) {
          let newtitle = wordCount + " / " + wordTargetCount;
          if (newtitle != cc.title) cc.title = newtitle;
          await context.sync();
        }
      }

      // Over-limit red highlight
      if (wordCount>wordTargetCount){
        let range = cc.getRange();

        let wordRanges = range.split([" "],false,true,true);
        context.load(wordRanges,"items");
        await context.sync();

        let wordsOverLimitRanges = wordRanges.items.slice(wordTargetCount);
        for (r=0;r<wordsOverLimitRanges.length;r++){
          wordsOverLimitRanges[r].font.color = "red";
        }

        await context.sync();
      }
    }
    catch(err){
      lastCcIdSelected=null;
      if (err.message != "ItemNotFound" && err.message != "Wait until the previous call completes.") console.log(err);
    }
  });
}

async function addSection(event) {
  try{
      await Word.run(async (context) => {
        var selectedRange = context.document.getSelection();
        let parent = selectedRange.parentContentControlOrNullObject;
        await context.sync();
        
        let cc;

        // If CC parent, insert new CC after parent CC
        var isParentCc = !parent.isNullObject;
        if (!isParentCc){ 
          cc = selectedRange.insertContentControl();
        }
        else{
          let ccMarker = parent.parentBody.getRange("Whole").insertText("\n","After");
          cc = ccMarker.insertContentControl();
        }

        cc.appearance = Word.ContentControlAppearance.boundingBox;
        cc.color = defaultColor;
        
        // Load into context and sync
        context.load(selectedRange, "text");
        context.load(cc,"id");
        await context.sync();

        targetCountDict.set(cc.id,String(defaultTargetCount));

        if (selectedRange.text.length < 2 && !isParentCc){
          selectedRange.insertText("\n","Before");
        }

        // Register event handlers
        cc.onDeleted.add(contentControlDeleted);
        cc.track();
        await context.sync();
      });
    event.completed();
  }
  catch(err){
    console.error(err);
  }
  finally{
    event.completed();
  }
}

async function contentControlDeleted(event) {
  await Word.run(async (context) => {
    let wc = wordCountDict.get(event.ids[0])
    totalWordCountElement.innerHTML = Number(totalWordCountElement.innerHTML)-wc;
    totalWordCountExcludingSelected -= wc;
    wordCountDict.delete(event.ids[0]);
    event.completed();
  });
}

async function selectedCcChanged(id) { // this saves processing time by summing all wordCounts when selected cc is changed
  await Word.run(async (context) => {  // so that when changes are made it only has to sum the changed wordCount to the precalced value
    let sum = 0; 
    wordCountDict.forEach((value,key) => {
      if (key != id) sum+=value;
    })                    
    totalWordCountExcludingSelected = sum;
  });
}

async function changeDefaultColor(){
  await Word.run(async (context) => {
    defaultColor = colorPickerElement.value;
    setInLocalStorage("defaultColor",defaultColor);
  });
}

async function changeDefaultTargetCount(){
  let newVal = defaultTargetCountElement.value;
  if (!(newVal>=5)) defaultTargetCountElement.value=5;
  else defaultTargetCount = newVal;
  setInLocalStorage("defaultTarget",defaultTargetCount);
}

function splitWords(text){
  text = text.trim();
  if (text==="") return [];
  return text.split(/\s+/);
}

function setInLocalStorage(key, value) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned. 
  // If so, use the partition to ensure the data is only accessible by your add-in.
  if (myPartitionKey) {
    localStorage.setItem(myPartitionKey + key, value);
  } else {
    localStorage.setItem(key, value);
  }
}

function getFromLocalStorage(key) {
  const myPartitionKey = Office.context.partitionKey;

  // Check if local storage is partitioned.
  if (myPartitionKey) {
    return localStorage.getItem(myPartitionKey + key);
  } else {
    return localStorage.getItem(key);
  }
}

