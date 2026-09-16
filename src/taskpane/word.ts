/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, fetch, HTMLElement, HTMLTextAreaElement, HTMLButtonElement, HTMLSelectElement */

const systemPrompt =
  "Tu es un assistant pour Microsoft Office. Réponds clairement en français, en conservant le contexte du document quand il est fourni.";
let latestAnswer = "";

function element<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

async function getSelectionText(): Promise<string> {
  return Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();
    return selection.text;
  });
}

async function askOmniroute(): Promise<void> {
  const question = element<HTMLTextAreaElement>("question").value.trim();
  const status = element("status");
  const answer = element("answer");
  const send = element<HTMLButtonElement>("send");
  if (!question) {
    status.textContent = "Saisissez une question";
    return;
  }
  send.disabled = true;
  element<HTMLButtonElement>("insert").disabled = true;
  status.textContent = "Réflexion...";
  answer.textContent = "";
  try {
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: element<HTMLSelectElement>("model").value,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: question },
        ],
      }),
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Le service IA a renvoyé une erreur.");
    }
    latestAnswer = data.choices?.[0]?.message?.content || "Aucune réponse reçue.";
    answer.textContent = latestAnswer;
    element<HTMLButtonElement>("insert").disabled = !latestAnswer;
    status.textContent = "Réponse reçue";
  } catch (error) {
    status.textContent = "Erreur";
    answer.textContent =
      error instanceof Error ? error.message : "Impossible de contacter le service IA.";
  } finally {
    send.disabled = false;
  }
}

async function useSelection(): Promise<void> {
  const status = element("status");
  try {
    const selectedText = (await getSelectionText()).trim();
    if (!selectedText) {
      status.textContent = "Aucune sélection";
      return;
    }
    element<HTMLTextAreaElement>("question").value =
      `Voici le texte sélectionné :\n\n${selectedText}\n\n`;
    status.textContent = "Sélection ajoutée";
  } catch (error) {
    status.textContent =
      error instanceof Error ? error.message : "Impossible de lire la sélection.";
  }
}

async function insertAnswer(): Promise<void> {
  if (!latestAnswer) return;
  try {
    await Word.run(async (context) => {
      context.document.getSelection().insertText(latestAnswer, Word.InsertLocation.replace);
      await context.sync();
    });
    element("status").textContent = "Insérée dans le document";
  } catch (error) {
    element("status").textContent =
      error instanceof Error ? error.message : "Insertion impossible.";
  }
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    element("sideload-msg").hidden = true;
    element("app-body").hidden = false;
    element<HTMLButtonElement>("send").onclick = () => void askOmniroute();
    element<HTMLButtonElement>("use-selection").onclick = () => void useSelection();
    element<HTMLButtonElement>("insert").onclick = () => void insertAnswer();
  }
});
