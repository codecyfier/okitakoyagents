/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, fetch, HTMLElement, HTMLTextAreaElement, HTMLButtonElement, HTMLSelectElement */

const systemPrompt =
  "Tu es un assistant expert de Microsoft Word. Réponds en français avec une rédaction professionnelle, claire et directement exploitable dans un document Word. N'utilise jamais de Markdown visible : pas de #, ##, ###, *, **, _, backticks, puces avec tirets ou symboles de séparation. Structure le contenu avec des titres courts, des paragraphes et des listes numérotées ou à puces en texte simple. N'ajoute pas de préambule inutile ni de mention de ton formatage. Quand une réponse doit être insérée dans Word, privilégie une formulation naturelle et soignée.";
let latestAnswer = "";

function element<T extends HTMLElement>(id: string): T {
  return document.getElementById(id) as T;
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function inlineMarkdownToHtml(value: string): string {
  return escapeHtml(value)
    .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
    .replace(/__(.+?)__/g, "<strong>$1</strong>")
    .replace(/\*(.+?)\*/g, "<em>$1</em>")
    .replace(/_(.+?)_/g, "<em>$1</em>")
    .replace(/`(.+?)`/g, "$1");
}

function answerToWordHtml(value: string): string {
  const lines = value.replace(/\r\n/g, "\n").split("\n");
  const blocks: string[] = [];
  let paragraph: string[] = [];
  let listItems: string[] = [];
  let listType = "";

  const flushParagraph = (): void => {
    if (paragraph.length) {
      blocks.push(`<p>${inlineMarkdownToHtml(paragraph.join(" ").trim())}</p>`);
      paragraph = [];
    }
  };

  const flushList = (): void => {
    if (listItems.length) {
      blocks.push(`<${listType}>${listItems.join("")}</${listType}>`);
      listItems = [];
      listType = "";
    }
  };

  lines.forEach((line) => {
    const trimmedLine = line.trim();
    const headingMatch = trimmedLine.match(/^#{1,3}\s+(.+)$/);
    const unorderedMatch = trimmedLine.match(/^(?:[-*+]\s+|•\s+)(.+)$/);
    const orderedMatch = trimmedLine.match(/^\d+[.)]\s+(.+)$/);

    if (!trimmedLine) {
      flushParagraph();
      flushList();
    } else if (headingMatch) {
      flushParagraph();
      flushList();
      const headingLevel = trimmedLine.match(/^#+/)?.[0].length || 2;
      blocks.push(`<h${headingLevel}>${inlineMarkdownToHtml(headingMatch[1])}</h${headingLevel}>`);
    } else if (unorderedMatch || orderedMatch) {
      flushParagraph();
      const nextListType = orderedMatch ? "ol" : "ul";
      if (listType && listType !== nextListType) flushList();
      listType = nextListType;
      const listText = orderedMatch ? orderedMatch[1] : unorderedMatch?.[1] || "";
      listItems.push(`<li>${inlineMarkdownToHtml(listText)}</li>`);
    } else {
      flushList();
      paragraph.push(trimmedLine);
    }
  });

  flushParagraph();
  flushList();
  return blocks.join("") || "<p></p>";
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
      context.document
        .getSelection()
        .insertHtml(answerToWordHtml(latestAnswer), Word.InsertLocation.replace);
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
