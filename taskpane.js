/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  const item = Office.context.mailbox.item;
  const subject = item.subject;
  const body = await getBody(item);

  // Send subject and body to backend service
  const response = await fetch('https://1a3c-46-18-28-243.ngrok-free.app/score-email/', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ subject, body })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
  }


  const data = await response.json();
  const projectTitle = data.project_title;

  // Update UI with project title and provide option to move email
  let insertAt = document.getElementById("item-subject");
  insertAt.innerHTML = ''; // Clear previous content
  let label = document.createElement("b").appendChild(document.createTextNode("Project Title: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(projectTitle));
  insertAt.appendChild(document.createElement("br"));

  // Create button to move email
  let moveButton = document.createElement("button");
  moveButton.textContent = "Move to Folder";
  moveButton.onclick = () => moveToFolder(projectTitle);
  insertAt.appendChild(moveButton);
}

async function getBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error);
      }
    });
  });
}

async function moveToFolder(folderName) {
  const mailbox = Office.context.mailbox;
  const item = mailbox.item;

  // Check if folder exists, if not create it
  const folder = await getOrCreateFolder(folderName);

  // Move the email to the folder
  item.moveAsync(folder.id, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      console.log('Email moved successfully');
    } else {
      console.error('Error moving email:', result.error);
    }
  });
}

async function getOrCreateFolder(folderName) {
  const mailbox = Office.context.mailbox;
  return new Promise((resolve, reject) => {
    mailbox.folders.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const folder = result.value.find(f => f.displayName === folderName);
        if (folder) {
          resolve(folder);
        } else {
          // Create folder if it doesn't exist
          mailbox.folders.addAsync(folderName, (addResult) => {
            if (addResult.status === Office.AsyncResultStatus.Succeeded) {
              resolve(addResult.value);
            } else {
              reject(addResult.error);
            }
          });
        }
      } else {
        reject(result.error);
      }
    });
  });
}
