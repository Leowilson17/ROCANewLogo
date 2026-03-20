/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { debug } from "webpack";

/* global document, Office, msal */

// Office.onReady((info) => {
//   if (info.host === Office.HostType.Outlook) {
//     document.getElementById("sideload-msg").style.display = "none";
//     document.getElementById("app-body").style.display = "flex";
//     document.getElementById("run").onclick = run;
//   }
// });

// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */

//   const item = Office.context.mailbox.item;
//   let insertAt = document.getElementById("item-subject");
//   let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
//   insertAt.appendChild(label);
//   insertAt.appendChild(document.createElement("br"));
//   insertAt.appendChild(document.createTextNode(item.subject));
//   insertAt.appendChild(document.createElement("br"));
// }
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("userForm").addEventListener("submit", submitForm);
  }
});

const msalConfig = {
  auth: {
    clientId: "8a4c9c58-0108-453a-9309-6b69d3656138",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

// const msalInstance = new msal.PublicClientApplication(msalConfig);

const graphScopes = ["User.Read", "Sites.ReadWrite.All"];

async function getAccessToken() {
  try {
    const apiUrl =
      "https://default3e8e53bea48f4147adf87e90a6e46b.57.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/2ff412299f724c619c608c3a1d430290/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=s8gBVzJ-ner3YkcvjZEoTU3MIYNd6iaWq-GpCxgIGHk";

    const body = {
      secret: "xJ-8Q~J~zayk-gxZ5X~nE4ftP0vFv0J7b615~bxX",
      tenantId: "3e8e53be-a48f-4147-adf8-7e90a6e46b57",
      clientId: "8a4c9c58-0108-453a-9309-6b69d3656138",
    };

    const response = await fetch(apiUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(body),
    });
    const token = await response.json();
    return token.token;
  } catch (err) {
    console.error(err, "getAccessToken");
  }
}
// async function getAccessToken() {
//   try {
//     // let account = msalInstance.getAllAccounts()[0];

//     // if (!account) {
//     //   const loginResponse = await msalInstance.loginPopup({
//     //     scopes: graphScopes,
//     //   });
//     //   account = loginResponse.account;
//     // }

//     // const tokenResponse = await msalInstance.acquireTokenSilent({
//     //   scopes: graphScopes,
//     //   account: account,
//     // });

//     // return tokenResponse.accessToken;
//     return "";
//   } catch (error) {
//     alert("Authentication failed");
//   }
// }

/**
 * Main form submission han   dler using Microsoft Graph
 */
export async function submitForm(event) {
  event.preventDefault();

  const statusElement = document.getElementById("status-message");
  if (statusElement) {
    statusElement.innerText = "Submitting data...";
    statusElement.style.color = "blue";
  }

  try {
    const accessToken = await getAccessToken();

    const userData = {
      name: (document.getElementById("name") as HTMLInputElement).value,
      email: (document.getElementById("email") as HTMLInputElement).value,
      phone: (document.getElementById("phone") as HTMLInputElement).value,
      gender: (document.getElementById("gender") as HTMLSelectElement).value,
      dob: (document.getElementById("dob") as HTMLInputElement).value,
    };

    // 1. Get Site ID from Graph
    const siteUrl = "https://chandrudemo.sharepoint.com/sites/GaneshDevsite";
    const siteId = await getSiteId(siteUrl, accessToken);

    const attachmentBlobs: any = await getAllAttachmentsAsBlobs();

    await createGraphListItem(siteId, "UserDataList", userData, accessToken);

    if (attachmentBlobs.length > 0) {
      if (statusElement)
        statusElement.innerText = `Uploading ${attachmentBlobs.length} attachments...`;
      for (const item of attachmentBlobs) {
        await uploadToGraphLibrary(siteId, "Documents", item.name, item.blob, accessToken);
      }
    }

    if (statusElement) {
      statusElement.innerText = "Submission successful!";
      statusElement.style.color = "green";
    }
    (document.getElementById("userForm") as HTMLFormElement).reset();
  } catch (error) {
    console.error("Submission failed:", error);
    if (statusElement) {
      statusElement.innerText = "Submission failed. See console.";
      statusElement.style.color = "red";
    }
  }
}

/**
 * Fetches the Microsoft Graph Site ID for a given SharePoint Site URL
 */
async function getSiteId(siteUrl, accessToken) {
  const urlObj = new URL(siteUrl);
  const hostname = urlObj.hostname;
  const path = urlObj.pathname.startsWith("/") ? urlObj.pathname.substring(1) : urlObj.pathname;

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${hostname}:/${path}?$select=id`,
    {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(`Failed to get Site ID: ${error.error.message}`);
  }

  const data = await response.json();
  return data.id;
}

/**
 * Microsoft Graph: Create List Item
 */
async function createGraphListItem(siteId, listName, data, accessToken) {
  const _data = {
    fields: {
      Title: data.name,
      Email: data.email,
      PhoneNumber: data.phone,
      Gender: data.gender,
      DateOfBirth: data.dob,
    },
  };

  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/lists/${listName}/items`,
    {
      method: "POST",
      body: JSON.stringify(_data),
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(`Failed to create list item: ${error.error.message}`);
  }
  return response.json();
}

/**
 * Microsoft Graph: Upload File to Library
 */
async function uploadToGraphLibrary(siteId, libraryName, fileName, blob, accessToken) {
  console.log(`Uploading ${fileName} to Graph Library '${libraryName}'`);

  // Using the PUT method for small files (< 4MB)
  // Endpoint: sites/{site-id}/drive/root:/{folder-path}/{filename}:/content
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${fileName}:/content`,
    {
      method: "PUT",
      body: blob,
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(`Failed to upload file ${fileName}: ${error.error.message}`);
  }
  return response.json();
}

/**
 * Retrieves all attachments from the current email and converts them to Blobs
 */
async function getAllAttachmentsAsBlobs() {
  const item = Office.context.mailbox.item;
  if (!item.attachments || item.attachments.length === 0) {
    return [];
  }

  const attachmentPromises = item.attachments.map((att) => getAttachmentBlob(att));
  return Promise.all(attachmentPromises);
}

/**
 * Fetches content for a single attachment and creates a Blob
 */
function getAttachmentBlob(attachment) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const base64Content = result.value.content;
        const contentType = attachment.contentType || "application/octet-stream";
        const blob = base64ToBlob(base64Content, contentType);
        resolve({ name: attachment.name, blob: blob });
      } else {
        reject(new Error(`Failed to get attachment ${attachment.name}: ${result.error.message}`));
      }
    });
  });
}

function base64ToBlob(base64, contentType) {
  const byteCharacters = atob(base64);
  const byteNumbers = new Array(byteCharacters.length);
  for (let i = 0; i < byteCharacters.length; i++) {
    byteNumbers[i] = byteCharacters.charCodeAt(i);
  }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], { type: contentType });
}
