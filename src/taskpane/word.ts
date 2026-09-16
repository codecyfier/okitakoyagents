/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word, fetch, HTMLElement, HTMLTextAreaElement, HTMLButtonElement, HTMLSelectElement */

const systemPrompt =
  "Tu es un assistant expert de Microsoft Word. Le document courant et la sélection éventuelle sont fournis dans chaque demande : utilise-les directement, sans demander à l'utilisateur de les recoller ou de les joindre. Pour une correction, retourne le texte corrigé prêt à remplacer le passage concerné. Pour une analyse, retourne une analyse structurée et directement exploitable. Pour une demande de mise en forme, retourne le contenu final correctement structuré. Réponds en français avec une rédaction professionnelle. N'utilise jamais de Markdown visible : pas de #, *, _, backticks, puces avec tirets ou séparateurs. Structure naturellement le contenu avec un titre court au début, des sous-titres sur des lignes séparées terminés par deux-points, des paragraphes aérés, des listes numérotées avec 1. ou des listes à puces avec •, des citations entre guillemets si nécessaire, et des tableaux simples avec le caractère | uniquement quand c'est utile. N'ajoute pas de préambule ni de mention de ton formatage.";
let latestAnswer = "";
let replaceWholeDocument = false;
const maximumDocumentCharacters = 50000;

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
  let hasContent = false;

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

  const isTableRow = (line: string): boolean => line.includes("|") && line.split("|").length >= 3;

  const tableHtml = (tableLines: string[]): string => {
    const rows = tableLines
      .filter((line) => !/^\s*\|?\s*:?-{2,}:?\s*(\|\s*:?-{2,}:?\s*)+\|?\s*$/.test(line))
      .map((line) => line.split("|").map((cell) => cell.trim()).filter(Boolean));
    if (!rows.length) return "";
    const header = rows[0].map((cell) => `<th style="background-color:#eee8df;color:#315e48">${inlineMarkdownToHtml(cell)}</th>`).join("");
    const body = rows.slice(1).map((row) => `<tr>${row.map((cell) => `<td>${inlineMarkdownToHtml(cell)}</td>`).join("")}</tr>`).join("");
    return `<table><thead><tr>${header}</tr></thead><tbody>${body}</tbody></table>`;
  };

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const trimmedLine = line.trim();
    const headingMatch = trimmedLine.match(/^#{1,3}\s+(.+)$/);
    const unorderedMatch = trimmedLine.match(/^(?:[-*+]\s+|•\s+)(.+)$/);
    const orderedMatch = trimmedLine.match(/^\d+[.)]\s+(.+)$/);
    const plainHeading = trimmedLine.match(/^([^.!?]{2,90}):$/);
    const tableLines: string[] = [];

    if (isTableRow(trimmedLine) && isTableRow(lines[index + 1] || "")) {
      flushParagraph();
      flushList();
      while (index < lines.length && isTableRow(lines[index].trim())) {
        tableLines.push(lines[index].trim());
        index += 1;
      }
      index -= 1;
      blocks.push(tableHtml(tableLines));
      hasContent = true;
      continue;
    }

    if (!trimmedLine) {
      flushParagraph();
      flushList();
    } else if (/^[—-]{3,}$/.test(trimmedLine)) {
      flushParagraph();
      flushList();
    } else if (headingMatch) {
      flushParagraph();
      flushList();
      const headingLevel = trimmedLine.match(/^#+/)?.[0].length || 2;
      blocks.push(`<h${headingLevel} style="color:#315e48">${inlineMarkdownToHtml(headingMatch[1])}</h${headingLevel}>`);
      hasContent = true;
    } else if (plainHeading || (!hasContent && trimmedLine.length <= 90)) {
      flushParagraph();
      flushList();
      const headingText = plainHeading ? plainHeading[1] : trimmedLine;
      blocks.push(`<h${hasContent ? 2 : 1} style="color:#315e48">${inlineMarkdownToHtml(headingText.replace(/:$/, ""))}</h${hasContent ? 2 : 1}>`);
      hasContent = true;
    } else if (unorderedMatch || orderedMatch) {
      flushParagraph();
      const nextListType = orderedMatch ? "ol" : "ul";
      if (listType && listType !== nextListType) flushList();
      listType = nextListType;
      const listText = orderedMatch ? orderedMatch[1] : unorderedMatch?.[1] || "";
      listItems.push(`<li>${inlineMarkdownToHtml(listText)}</li>`);
      hasContent = true;
    } else if (/^["«].+["»]$/.test(trimmedLine)) {
      flushParagraph();
      flushList();
      blocks.push(`<blockquote style="border-left:3px solid #9b4d2f;color:#665f56">${inlineMarkdownToHtml(trimmedLine)}</blockquote>`);
      hasContent = true;
    } else {
      flushList();
      paragraph.push(trimmedLine);
      hasContent = true;
    }
  }

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

async function getDocumentContext(): Promise<{ documentText: string; selectionText: string }> {
  return Word.run(async (context) => {
    const body = context.document.body;
    const selection = context.document.getSelection();
    body.load("text");
    selection.load("text");
    await context.sync();
    return {
      documentText: body.text.slice(0, maximumDocumentCharacters),
      selectionText: selection.text.trim(),
    };
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
  status.textContent = "Réflexion...";
  answer.textContent = "";
  try {
    status.textContent = "Lecture du document...";
    const documentContext = await getDocumentContext();
    replaceWholeDocument =
      !documentContext.selectionText &&
      /\b(corrig|orthograph|grammaire|faute|réécri|reécri)/i.test(question);
    const contextMessage = [
      "DOCUMENT WORD COURANT :",
      documentContext.documentText || "(Le document est vide.)",
      "",
      "TEXTE ACTUELLEMENT SÉLECTIONNÉ :",
      documentContext.selectionText || "(Aucune sélection.)",
      "",
      "DEMANDE DE L'UTILISATEUR :",
      question,
    ].join("\n");
    status.textContent = "Analyse du document...";
    const response = await fetch("/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: element<HTMLSelectElement>("model").value,
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: contextMessage },
        ],
      }),
    });
    const data = await response.json();
    if (!response.ok) {
      throw new Error(data.error || "Le service IA a renvoyé une erreur.");
    }
    latestAnswer = data.choices?.[0]?.message?.content || "Aucune réponse reçue.";
    answer.textContent = latestAnswer;
    status.textContent = "Insertion dans Word...";
    await insertAnswer();
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
      const target = replaceWholeDocument
        ? context.document.body
        : context.document.getSelection();
      target.insertHtml(answerToWordHtml(latestAnswer), Word.InsertLocation.replace);
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
  }
});
