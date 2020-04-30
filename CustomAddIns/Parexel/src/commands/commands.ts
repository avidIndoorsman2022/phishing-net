/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event 
 */
function action(event: Office.AddinCommands.Event) {

/*  
  let errMessage: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, 
    message: "Performed action, but with errors."
  }

  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, 
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true
  }
  
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", errMessage);
*/
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() { 
  return (typeof self !== "undefined") ? self :
    (typeof window !== "undefined") ? window : 
    (typeof global !== "undefined") ? global :
    undefined;
}

const g = getGlobal() as any;

// the add-in command functions need to be available in global scope
g.action = action;
