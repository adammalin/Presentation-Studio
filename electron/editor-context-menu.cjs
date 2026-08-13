function uniqueSuggestions(values) {
  return [...new Set((values ?? []).filter((value) => typeof value === "string" && value.trim()).map((value) => value.trim()))].slice(0, 8);
}

function dictionaryLabel(word) {
  const compact = word.length > 36 ? `${word.slice(0, 33)}…` : word;
  return `Add “${compact}” to dictionary`;
}

function buildEditorContextMenuTemplate(params, actions) {
  if (!params?.isEditable) return [];
  const template = [];
  const misspelledWord = typeof params.misspelledWord === "string" ? params.misspelledWord.trim() : "";
  const suggestions = uniqueSuggestions(params.dictionarySuggestions);
  if (misspelledWord) {
    if (suggestions.length > 0) {
      for (const suggestion of suggestions) template.push({ label: suggestion, click: () => actions.replaceMisspelling(suggestion) });
    } else {
      template.push({ label: "No spelling suggestions", enabled: false });
    }
    template.push({ type: "separator" }, { label: dictionaryLabel(misspelledWord), click: () => actions.addToDictionary(misspelledWord) }, { type: "separator" });
  }
  const flags = params.editFlags ?? {};
  template.push(
    { role: "undo", enabled: flags.canUndo !== false },
    { role: "redo", enabled: flags.canRedo !== false },
    { type: "separator" },
    { role: "cut", enabled: flags.canCut !== false },
    { role: "copy", enabled: flags.canCopy !== false },
    { role: "paste", enabled: flags.canPaste !== false },
    { role: "pasteAndMatchStyle", enabled: flags.canPaste !== false },
    { type: "separator" },
    { role: "selectAll", enabled: flags.canSelectAll !== false },
  );
  return template;
}

function installEditorContextMenu(window, Menu) {
  window.webContents.on("context-menu", (_event, params) => {
    const template = buildEditorContextMenuTemplate(params, {
      replaceMisspelling: (suggestion) => {
        if (!window.isDestroyed() && !window.webContents.isDestroyed()) window.webContents.replaceMisspelling(suggestion);
      },
      addToDictionary: (word) => {
        if (!window.isDestroyed() && !window.webContents.isDestroyed()) window.webContents.session.addWordToSpellCheckerDictionary(word);
      },
    });
    if (template.length === 0 || window.isDestroyed()) return;
    Menu.buildFromTemplate(template).popup({ window });
  });
}

module.exports = { buildEditorContextMenuTemplate, installEditorContextMenu, uniqueSuggestions };
