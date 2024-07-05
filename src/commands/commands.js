// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */

// /* global Office */

// var contentControls;
// var targetCounts;
// var defaultTargetCount;

// Office.onReady(() => {
//   Office.addin.setStartupBehavior(Office.StartupBehavior.load);

//   targetCounts = new Map();
//   defaultTargetCount = 500;
//   loadAllCcLimits();
//   eventSubscribe();
// });

// async function loadAllCcLimits(){
//   await Word.run(async (context) => {
//     contentControls = context.document.contentControls;
//     contentControls.load("items");
//     await context.sync();

//     let ccs = contentControls.items;
//     for (i=0; i<ccs.length; i++){
//       try{
//         let limit = ccs[i].title.split(" / ")[1];
//         targetCounts.set(ccs[i].id,limit);
//       }
//       catch(err){console.error(err);}
//     }
//   });  
// }

// async function eventSubscribe(){
//   await Word.run(async (context) => {
//     var events = setInterval(
//       checkForCcChanges,
//       250,
//     );
//   });  
// }

// async function checkForCcChanges(){
//   await Word.run(async (context) => {
//     try{
//       let selection = context.document.getSelection();
//       let cc = selection.parentContentControl; 
//       //throws an error if a cc is not currently selected

//       context.load(cc,'id');
//       context.load(cc,'text');
//       context.load(cc,'title');
//       await context.sync();

//       let wordCount = String(countWords(cc.text));
//       let wordLimit = String(targetCounts.get(cc.id));

//       if (wordCount>=0 && wordLimit>0) {
//         let newtitle = wordCount + " / " + wordLimit;
//         if (newtitle != cc.title) cc.title = newtitle;
//         await context.sync();
//       }
//     }
//     catch(err){console.error(err);}
//   });
// }

// async function addSection(event) {
//   try{
//     await Word.run(async (context) => {
//       var selectedRange = context.document.getSelection();
//       context.load(selectedRange, "text");

//       let cc = selectedRange.insertContentControl();
//       cc.appearance = Word.ContentControlAppearance.boundingBox;
//       cc.color = "#ffd970";
      
//       cc.title = "WordCount";
//       context.load(cc,"id");

//       await context.sync();

//       contentControls = context.document.contentControls;
//       contentControls.load("items");

//       targetCounts.set(cc.id,String(defaultTargetCount));

//       if (selectedRange.text.length < 2){
//         selectedRange.insertText("\n","Before");
//       }
//     });
//     event.completed();
//   }
//   catch(err){
//     console.error(err);
//   }
//   finally{
//     event.completed();
//   }
// }


// function countWords(text){
//   text = text.trim();
//   if (text === "") return 0;
//   return text.split(/\s+/).length;
// }

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message.
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete.
//   event.completed();
// }

// // Register the function with Office.
// Office.actions.associate("action", action);
