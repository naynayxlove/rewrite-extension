import {
    eventSource,
    event_types,
    messageFormatting,
    addCopyToCodeBlocks,
} from "../../../../script.js";
import { extension_settings, getContext } from "../../../extensions.js";

const extensionName = "rewrite-extension-del-only";
const extensionFolderPath = `scripts/extensions/third-party/${extensionName}`;

const undo_steps = 15;

// Default settings
const defaultSettings = {
    showDelete: true,
    showRewrite: true,
};

let deleteMenu = null;
let lastSelection = null;

let changeHistory = [];

function ensureSettings() {
    const existing = extension_settings[extensionName] || {};
    extension_settings[extensionName] = { ...defaultSettings, ...existing };
    return extension_settings[extensionName];
}

// Load settings
function loadSettings() {
    const settings = ensureSettings();

    // Helper function to get a setting with a default value
    const getSetting = (key, defaultValue) => {
        return settings[key] !== undefined ? settings[key] : defaultValue;
    };

    // Load settings, using defaults if not set
    $("#show_delete").prop('checked', getSetting('showDelete', defaultSettings.showDelete));
    $("#show_rewrite").prop('checked', getSetting('showRewrite', defaultSettings.showRewrite));
}

function saveSettings() {
    extension_settings[extensionName] = {
        ...defaultSettings,
        showDelete: $("#show_delete").is(':checked'),
        showRewrite: $("#show_rewrite").is(':checked'),
    };

    ensureSettings();
}

// Initialize
jQuery(async () => {
    try {
        const settingsHtml = await $.get(`${extensionFolderPath}/delete_settings.html`);
        $("#extensions_settings2").append(settingsHtml);
    } catch (error) {
        console.error("[Delete Extension] Failed to load settings UI:", error);
    }

    ensureSettings();

    // Add event listeners
    $("#show_delete, #show_rewrite").on("change", saveSettings);

    // Load settings
    loadSettings();

    eventSource.on(event_types.CHAT_CHANGED, () => {
        changeHistory = [];
        updateUndoButtons();
    });

    eventSource.on(event_types.MESSAGE_EDITED, (editedMesId) => {
        removeUndoButton(editedMesId);
    });

    // Initialize the delete menu functionality after settings are loaded
    initDeleteMenu();
});

function initDeleteMenu() {
    document.addEventListener('selectionchange', handleSelectionChange);
    document.addEventListener('mousedown', hideMenuOnOutsideClick);
    document.addEventListener('touchstart', hideMenuOnOutsideClick);

    let chatContainer = document.getElementById('chat');
    if (chatContainer) {
        chatContainer.addEventListener('scroll', positionMenu);
    }
}

function handleSelectionChange() {
    // Use a small timeout to ensure the selection has been updated
    setTimeout(processSelection, 50);
}

function processSelection() {
    // First, check if getContext().chatId is defined
    if (getContext().chatId === undefined) {
        return; // Exit the function if chatId is undefined
    }

    let selection = window.getSelection();
    let selectedText = selection.toString().trim();

    // Always remove the existing menu first
    removeDeleteMenu();

    if (selectedText.length > 0) {
        let range = selection.getRangeAt(0);

        // Find the mes_text elements for both start and end of the selection
        let startMesText = range.startContainer.nodeType === Node.ELEMENT_NODE
            ? range.startContainer.closest('.mes_text')
            : range.startContainer.parentElement.closest('.mes_text');

        let endMesText = range.endContainer.nodeType === Node.ELEMENT_NODE
            ? range.endContainer.closest('.mes_text')
            : range.endContainer.parentElement.closest('.mes_text');

        // Check if both start and end are within the same mes_text element
        if (startMesText && endMesText && startMesText === endMesText) {
            createDeleteMenu();
        }
    }

    lastSelection = selectedText.length > 0 ? selectedText : null;
}

async function handleMenuItemClick(e) {
    e.preventDefault();
    e.stopPropagation();

    const option = e.target.dataset.option;
    const selection = window.getSelection();

    // Ensure there's a selection and a range
    if (!selection || selection.rangeCount === 0) {
        removeDeleteMenu();
        return;
    }

    // Capture the range *before* any awaits or potential selection changes
    const initialRange = selection.getRangeAt(0).cloneRange();
    const selectedText = initialRange.toString().trim();

    if (selectedText) {
        const mesTextElement = findClosestMesText(selection.anchorNode);
        if (mesTextElement) {
            const messageDiv = findMessageDiv(mesTextElement);
            if (messageDiv) {
                const mesId = messageDiv.getAttribute('mesid');
                const swipeId = messageDiv.getAttribute('swipeid');

                if (option === 'Delete') {
                    // Pass the initially captured range to handleDeleteSelection
                    await handleDeleteSelection(mesId, swipeId, initialRange);
                } else if (option === 'Rewrite') {
                    // Get the new text from user input
                    const newText = await getRewriteTextFromPopup();
                    // Treat explicit false as cancel to avoid inserting "false"
                    if (newText !== null && newText !== undefined && newText !== false) {
                        await handleRewriteSelection(mesId, swipeId, initialRange, newText);
                    }
                } else if (option === 'Custom') {
                    const customInstructions = await getCustomInstructionsFromPopup();
                    if (typeof customInstructions === 'string' && customInstructions.trim() !== '') { // Proceed only if user entered text and didn't cancel
                        // Get selectionInfo *after* await and *before* handleRewrite
                        // Pass the initially captured range
                        const selectionInfo = getSelectedTextInfo(mesId, mesTextElement, initialRange);
                        if (!selectionInfo) {
                             console.error("[Rewrite Extension] Failed to get selectionInfo for Custom rewrite!");
                             return; // Prevent calling with undefined
                        }
                        await handleRewrite(mesId, swipeId, option, customInstructions.trim(), selectionInfo); // Use trimmed instructions
                    }
                }
            }
        }
    }

    removeDeleteMenu();
    window.getSelection().removeAllRanges();
}

async function getRewriteTextFromPopup() {
    const { callPopup } = getContext();
    try {
        const newText = await callPopup('Enter new text to replace selection:', 'input');
        // Introduce a zero-delay setTimeout to yield to the event loop
        await new Promise(resolve => setTimeout(resolve, 0));
        return newText;
    } catch (error) {
        console.error("[Delete Extension] Error during rewrite text popup:", error);
        return null;
    }
}

async function handleDeleteSelection(mesId, swipeId, range) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    // Use the passed-in range to get selection info
    let { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, selectionText } = getSelectedTextInfo(mesId, mesDiv, range);

    const resolved = resolveOffsets(fullMessage, rawStartOffset, rawEndOffset, selectedRawText, selectionText);
    if (!resolved) {
        return;
    }
    rawStartOffset = resolved.start;
    rawEndOffset = resolved.end;

    // Create the new message with the deleted section removed
    let newMessage = fullMessage.slice(0, rawStartOffset) + fullMessage.slice(rawEndOffset);

    // Fallback: if nothing changed, try a loose regex-based match against the raw message
    if (newMessage === fullMessage) {
        const fallbackText = selectedRawText || selectionText || '';
        const looseMatch = findLooseMatchInRaw(fullMessage, fallbackText);
        if (looseMatch) {
            rawStartOffset = looseMatch.start;
            rawEndOffset = looseMatch.end;
            newMessage = fullMessage.slice(0, rawStartOffset) + fullMessage.slice(rawEndOffset);
        }
    }

    // Save the change to the history (this also calls updateUndoButtons)
    saveLastChange(mesId, swipeId, fullMessage, newMessage);

    // Update the message in the chat context
    getContext().chat[mesId].mes = newMessage;
    if (swipeId !== undefined && getContext().chat[mesId].swipes) {
        getContext().chat[mesId].swipes[swipeId] = newMessage;
    }

    // Update the UI
    mesDiv.innerHTML = messageFormatting(newMessage, getContext().name2, getContext().chat[mesId].isSystem, getContext().chat[mesId].isUser, mesId);
    addCopyToCodeBlocks(mesDiv);

    // Save the chat
    await getContext().saveChat();
}

async function handleRewriteSelection(mesId, swipeId, range, newText) {
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`);
    // Use the passed-in range to get selection info
    let { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, selectionText } = getSelectedTextInfo(mesId, mesDiv, range);

    const resolved = resolveOffsets(fullMessage, rawStartOffset, rawEndOffset, selectedRawText, selectionText);
    if (!resolved) {
        return;
    }
    rawStartOffset = resolved.start;
    rawEndOffset = resolved.end;

    // Create the new message with the rewritten section
    let newMessage = fullMessage.slice(0, rawStartOffset) + newText + fullMessage.slice(rawEndOffset);

    // Fallback: if nothing changed, try a loose regex-based match against the raw message
    if (newMessage === fullMessage) {
        const fallbackText = selectedRawText || selectionText || '';
        const looseMatch = findLooseMatchInRaw(fullMessage, fallbackText);
        if (looseMatch) {
            rawStartOffset = looseMatch.start;
            rawEndOffset = looseMatch.end;
            newMessage = fullMessage.slice(0, rawStartOffset) + newText + fullMessage.slice(rawEndOffset);
        }
    }

    // Save the change to the history (this also calls updateUndoButtons)
    saveLastChange(mesId, swipeId, fullMessage, newMessage);

    // Update the message in the chat context
    getContext().chat[mesId].mes = newMessage;
    if (swipeId !== undefined && getContext().chat[mesId].swipes) {
        getContext().chat[mesId].swipes[swipeId] = newMessage;
    }

    // Update the UI
    mesDiv.innerHTML = messageFormatting(newMessage, getContext().name2, getContext().chat[mesId].isSystem, getContext().chat[mesId].isUser, mesId);
    addCopyToCodeBlocks(mesDiv);

    // Save the chat
    await getContext().saveChat();
}

function hideMenuOnOutsideClick(e) {
    if (deleteMenu && !deleteMenu.contains(e.target)) {
        removeDeleteMenu();
    }
}

function createDeleteMenu() {
    removeDeleteMenu();

    deleteMenu = document.createElement('ul');
    deleteMenu.className = 'list-group ctx-menu';
    deleteMenu.style.position = 'absolute';
    deleteMenu.style.zIndex = '1000';
    deleteMenu.style.position = 'fixed';

    const settings = ensureSettings();

    const options = [
        { name: 'Rewrite', show: settings.showRewrite },
        { name: 'Delete', show: settings.showDelete }
    ];
    options.forEach(option => {
        if (option.show) {
            let li = document.createElement('li');
            li.className = 'list-group-item ctx-item';
            li.textContent = option.name;
            li.addEventListener('mousedown', handleMenuItemClick);
            li.addEventListener('touchstart', handleMenuItemClick);
            li.dataset.option = option.name;
            deleteMenu.appendChild(li);
        }
    });

    document.body.appendChild(deleteMenu);
    positionMenu();
}

function positionMenu() {
    if (!deleteMenu) return;

    let selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
        removeDeleteMenu();
        return;
    }
    let range = selection.getRangeAt(0);
    let rect = range.getBoundingClientRect();

    // Calculate the menu's position
    let left = rect.left + window.pageXOffset;
    let top = rect.bottom + window.pageYOffset + 5;

    // Get the viewport dimensions
    let viewportWidth = window.innerWidth;
    let viewportHeight = window.innerHeight;

    // Get the menu's dimensions
    let menuWidth = deleteMenu.offsetWidth;
    let menuHeight = deleteMenu.offsetHeight;

    // Adjust the position if the menu overflows the viewport
    if (left + menuWidth > viewportWidth) {
        left = viewportWidth - menuWidth;
    }
    if (top + menuHeight > viewportHeight) {
        top = rect.top + window.pageYOffset - menuHeight - 5;
    }

    deleteMenu.style.left = `${left}px`;
    deleteMenu.style.top = `${top}px`;
}

function removeDeleteMenu() {
    if (deleteMenu) {
        deleteMenu.remove();
        deleteMenu = null;
    }
}

function addUndoButton(mesId) {
    const messageDiv = document.querySelector(`[mesid="${mesId}"]`);
    if (messageDiv) {
        const mesButtons = messageDiv.querySelector('.mes_buttons');
        if (mesButtons) {
            const undoButton = document.createElement('div');
            undoButton.className = 'mes_button mes_undo_delete fa-solid fa-undo interactable';
            undoButton.title = 'Undo delete';
            undoButton.dataset.mesId = mesId;
            undoButton.addEventListener('click', handleUndo);

            if (mesButtons.children.length >= 1) {
                mesButtons.insertBefore(undoButton, mesButtons.children[1]);
            } else {
                mesButtons.appendChild(undoButton);
            }
        }
    }
}

function removeUndoButton(editedMesId) {
    // Remove all changes for this message from the changeHistory
    changeHistory = changeHistory.filter(change => change.mesId !== editedMesId);

    // Update undo buttons for other messages
    updateUndoButtons();
}

function findClosestMesText(element) {
    while (element && element.nodeType !== 1) {
        element = element.parentElement;
    }
    while (element) {
        if (element.classList && element.classList.contains('mes_text')) {
            return element;
        }
        element = element.parentElement;
    }
    return null;
}

function findMessageDiv(element) {
    while (element) {
        if (element.hasAttribute('mesid')) {
            return element;
        }
        element = element.parentElement;
    }
    return null;
}

function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function findLooseMatchInRaw(rawText, selectionText) {
    if (!selectionText) return null;
    const pattern = escapeRegex(selectionText.trim()).replace(/\s+/g, '\\s+');
    if (!pattern) return null;
    const re = new RegExp(pattern, 'm');
    const match = rawText.match(re);
    if (match) {
        return { start: match.index, end: match.index + match[0].length };
    }
    return null;
}

function createTextMapping(rawText, formattedHtml) {
    const formattedText = stripHtml(formattedHtml);
    const mapping = [];
    let rawIndex = 0;
    let formattedIndex = 0;

    const isWs = (ch) => ch === ' ' || ch === '\n' || ch === '\r' || ch === '\t' || ch === '\u00A0';

    while (rawIndex < rawText.length && formattedIndex < formattedText.length) {
        const r = rawText[rawIndex];
        const f = formattedText[formattedIndex];

        if (r === f) {
            mapping.push([rawIndex, formattedIndex]);
            rawIndex++;
            formattedIndex++;
            continue;
        }

        if (isWs(r) && isWs(f)) {
            mapping.push([rawIndex, formattedIndex]);
            rawIndex++;
            formattedIndex++;
            continue;
        }

        if (isWs(r)) {
            mapping.push([rawIndex, formattedIndex]);
            rawIndex++;
            continue;
        }

        if (isWs(f)) {
            mapping.push([rawIndex, formattedIndex]);
            formattedIndex++;
            continue;
        }

        // Fallback: advance raw to try to realign
        rawIndex++;
    }

    return {
        formattedToRaw: (formattedOffset) => {
            if (mapping.length === 0) {
                return Math.min(formattedOffset, rawText.length);
            }

            let low = 0;
            let high = mapping.length - 1;

            while (low <= high) {
                let mid = Math.floor((low + high) / 2);
                if (mapping[mid][1] === formattedOffset) {
                    return mapping[mid][0];
                } else if (mapping[mid][1] < formattedOffset) {
                    low = mid + 1;
                } else {
                    high = mid - 1;
                }
            }

            let refIndex;
            if (low >= mapping.length) {
                refIndex = mapping.length - 1;
            } else if (low > 0) {
                refIndex = low - 1;
            } else {
                refIndex = 0;
            }

            const delta = formattedOffset - mapping[refIndex][1];
            return Math.min(rawText.length, mapping[refIndex][0] + delta);
        },
    };
}

function stripHtml(html) {
    const tmp = document.createElement('DIV');
    tmp.innerHTML = html;
    return tmp.textContent || tmp.innerText || '';
}

function getTextOffset(parent, node) {
    const treeWalker = document.createTreeWalker(
        parent,
        NodeFilter.SHOW_TEXT,
        null,
        false
    );

    let offset = 0;
    while (treeWalker.nextNode() !== node) {
        offset += treeWalker.currentNode.length;
    }

    return offset;
}

function getSelectedTextInfo(mesId, mesDiv, range) {
    // Get the full message content
    const fullMessage = getContext().chat[mesId].mes;
    const selectionText = range.toString();

    // Get the formatted message text currently shown in the DOM (fallback to formatter)
    const formattedMessage = mesDiv ? mesDiv.textContent : messageFormatting(fullMessage, undefined, getContext().chat[mesId].isSystem, getContext().chat[mesId].isUser, mesId);

    // Create a mapping between raw and formatted text
    const mapping = createTextMapping(fullMessage, formattedMessage);

    // Calculate the start and end offsets relative to the formatted text content
    const startOffset = getTextOffset(mesDiv, range.startContainer) + range.startOffset;
    const endOffset = getTextOffset(mesDiv, range.endContainer) + range.endOffset;

    // Map these offsets back to the raw message
    let rawStartOffset = mapping.formattedToRaw(startOffset);
    let rawEndOffset = mapping.formattedToRaw(endOffset);

    // Heuristic: Adjust offsets to include surrounding markdown if selection seems to abut it
    // Check for italics (*)
    if (rawStartOffset > 0 && rawEndOffset < fullMessage.length &&
        fullMessage[rawStartOffset - 1] === '*' && fullMessage[rawEndOffset] === '*') {
        // Avoid expanding if it looks like bold/bold-italics boundary
        const prevChar = rawStartOffset > 1 ? fullMessage[rawStartOffset - 2] : null;
        const nextChar = rawEndOffset + 1 < fullMessage.length ? fullMessage[rawEndOffset + 1] : null;
        if (prevChar !== '*' && nextChar !== '*') {
            rawStartOffset--;
            rawEndOffset++;
        }
    }
    // Check for bold (**) - ensure we don't double-adjust if italics check already expanded
    else if (rawStartOffset > 1 && rawEndOffset < fullMessage.length - 1 &&
             fullMessage.substring(rawStartOffset - 2, rawStartOffset) === '**' &&
             fullMessage.substring(rawEndOffset, rawEndOffset + 2) === '**') {
        // Avoid expanding if it looks like bold-italics boundary
        const prevChar = rawStartOffset > 2 ? fullMessage[rawStartOffset - 3] : null;
        const nextChar = rawEndOffset + 2 < fullMessage.length ? fullMessage[rawEndOffset + 2] : null;
        if (prevChar !== '*' && nextChar !== '*') {
            rawStartOffset -= 2;
            rawEndOffset += 2;
        }
    }

    // Get the selected raw text using potentially adjusted offsets
    const selectedRawText = fullMessage.substring(rawStartOffset, rawEndOffset);

    return {
        fullMessage,
        selectedRawText,
        rawStartOffset,
        rawEndOffset,
        range,
        selectionText,
    };
}

function resolveOffsets(fullMessage, rawStartOffset, rawEndOffset, selectedRawText, selectionText) {
    const mappedLen = Math.max(0, rawEndOffset - rawStartOffset);
    const trimmedSel = (selectionText || '').trim();
    const trimmedRaw = (selectedRawText || '').trim();

    const isMappedSane = mappedLen > 0 && rawStartOffset >= 0 && rawEndOffset <= fullMessage.length;
    if (isMappedSane) {
        return { start: rawStartOffset, end: rawEndOffset };
    }

    const loose = findLooseMatchInRaw(fullMessage, trimmedSel || trimmedRaw);
    if (loose) {
        return loose;
    }

    // Give up if we cannot resolve safely
    return null;
}

function saveLastChange(mesId, swipeId, originalContent, newContent) {
    changeHistory.push({
        mesId,
        swipeId,
        originalContent,
        newContent,
        timestamp: Date.now()
    });

    // Limit history to last n changes
    if (changeHistory.length > undo_steps) {
        changeHistory.shift();
    }

    updateUndoButtons();
}

function updateUndoButtons() {
    // Remove all existing undo buttons
    document.querySelectorAll('.mes_undo_delete').forEach(button => button.remove());

    // Add undo buttons for all messages with changes
    const changedMessageIds = [...new Set(changeHistory.map(change => change.mesId))];
    changedMessageIds.forEach(mesId => addUndoButton(mesId));
}

// Updated handleRewrite signature to accept selectionInfo
async function handleRewrite(mesId, swipeId, option, customInstructions = null, selectionInfo) {
    if (!selectionInfo) {
        console.error("[Rewrite Extension] handleRewrite called without selectionInfo!");
        return; // Cannot proceed without selection info
    }

    if (main_api === 'openai') {
        const selectedModel = extension_settings[extensionName].selectedModel;
        if (selectedModel === 'chat_completion') {
            return handleChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo); // Pass selectionInfo
        } else {
            return handleSimplifiedChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo); // Pass selectionInfo
        }
    } else {
        return handleTextBasedRewrite(mesId, swipeId, option, customInstructions, selectionInfo); // Pass selectionInfo
    }
}

// Updated signature to accept selectionInfo
async function handleChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo) {
    // Use pre-captured selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = selectionInfo;
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`); // Keep getting mesDiv for highlight/DOM ops
    if (!mesDiv) { // Add check for mesDiv existence
        console.error("[Rewrite Extension] Could not find mesDiv in handleChatCompletionRewrite.");
        return;
    }

    // Get the selected preset based on the option
    let selectedPreset;
    switch (option) {
        case 'Rewrite':
            selectedPreset = extension_settings[extensionName].rewritePreset;
            break;
        case 'Shorten':
            selectedPreset = extension_settings[extensionName].shortenPreset;
            break;
        case 'Expand':
            selectedPreset = extension_settings[extensionName].expandPreset;
            break;
        case 'Custom': // New case
            selectedPreset = extension_settings[extensionName].customPreset;
            break;
        default:
            console.error("Unknown rewrite option:", option);
            return; // Exit if the option is not recognized
    }

    // Fetch the settings
    const result = await fetch('/api/settings/get', {
        method: 'POST',
        headers: getContext().getRequestHeaders(),
        body: JSON.stringify({}),
    });

    if (!result.ok) {
        console.error('Failed to fetch settings');
        return;
    }

    const data = await result.json();
    const presetIndex = data.openai_setting_names.indexOf(selectedPreset);
    if (presetIndex === -1) {
        console.error('Selected preset not found');
        return;
    }

    // Save the current settings
    const prev_oai_settings = Object.assign({}, oai_settings);

    // Parse the selected preset settings
    let selectedPresetSettings;
    try {
        selectedPresetSettings = JSON.parse(data.openai_settings[presetIndex]);
    } catch (error) {
        console.error('Error parsing preset settings:', error);
        return;
    }

    // Extension streaming overrides preset streaming
    selectedPresetSettings.stream_openai = extension_settings[extensionName].useStreaming;

    if (extension_settings[extensionName].overrideMaxTokens) {
        selectedPresetSettings.openai_max_tokens = calculateTargetTokenCount(selectedRawText, option);
    }

    // Override oai_settings with the selected preset
    Object.assign(oai_settings, selectedPresetSettings);

    // Always generate the base prompt using the selected preset
    const promptReadyPromise = new Promise(resolve => {
        eventSource.once(event_types.CHAT_COMPLETION_PROMPT_READY, resolve);
    });
    getContext().generate('normal', {}, true); // Trigger prompt generation
    const promptData = await promptReadyPromise; // Wait for the generated prompt
    let chatToSend = promptData.chat; // Start with the generated chat array

    // Inject custom instructions if applicable
    if (option === 'Custom' && customInstructions) {
        // Find the last user message to append to
        let targetMessageIndex = -1;
        for (let i = chatToSend.length - 1; i >= 0; i--) {
            if (chatToSend[i].role === 'user') {
                targetMessageIndex = i;
                break;
            }
        }

        if (targetMessageIndex !== -1) {
            const targetMessage = chatToSend[targetMessageIndex];
            const instructionText = `\n\nAdditional Instructions:\n${customInstructions}`;

            if (Array.isArray(targetMessage.content)) {
                // Find the last text part or add a new one
                let lastTextPartIndex = -1;
                for (let j = targetMessage.content.length - 1; j >= 0; j--) {
                    if (targetMessage.content[j].type === 'text') {
                        lastTextPartIndex = j;
                        break;
                    }
                }
                if (lastTextPartIndex !== -1) {
                    targetMessage.content[lastTextPartIndex].text += instructionText;
                } else {
                    // Should not happen with standard prompts, but handle just in case
                    targetMessage.content.push({ type: 'text', text: instructionText });
                }
            } else if (typeof targetMessage.content === 'string') {
                targetMessage.content += instructionText;
            }
        } else {
            console.warn('[Rewrite Extension] Could not find a user message in the generated prompt to inject custom instructions into.');
            // Optionally, could append a new user message, but might break formatting
            // chatToSend.push({ role: "user", content: `Additional Instructions:\n${customInstructions}` });
        }
    }

    // Substitute standard macros AFTER potential custom instruction injection
    const wordCount = extractAllWords(selectedRawText).length;
    chatToSend = chatToSend.map(message => {
        if (Array.isArray(message.content)) {
            message.content = message.content.map(item => {
                if (item.type === 'text') {
                    item.text = item.text.replace(/{{rewrite}}/gi, selectedRawText);
                    item.text = item.text.replace(/{{targetmessage}}/gi, fullMessage);
                    item.text = item.text.replace(/{{rewritecount}}/gi, wordCount);
                }
                return item;
            });
        } else if (typeof message.content === 'string') {
            message.content = message.content.replace(/{{rewrite}}/gi, selectedRawText);
            message.content = message.content.replace(/{{targetmessage}}/gi, fullMessage);
            message.content = message.content.replace(/{{rewritecount}}/gi, wordCount);
        }
        return message;
    });

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.prev_oai_settings = prev_oai_settings;
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;

    // Show the stop button
    getContext().deactivateSendButtons();

    let res;
    try {

        // Send the request with the prepared chat
        res = await sendOpenAIRequest('normal', chatToSend, abortController.signal);
    } catch (error) {
        console.error('[Rewrite Extension] Error during sendOpenAIRequest:', error);
        toastr.error("Rewrite failed. Check browser console (F12) for details.", "Rewrite Error");
        // Ensure cleanup happens even on error
    } finally {
        window.getSelection().removeAllRanges();
        // Restore the original settings (moved to finally)
        Object.assign(oai_settings, prev_oai_settings);
        getContext().activateSendButtons();
    }

    // If the request failed, res will be undefined, stop further processing
    if (res === undefined) {
        // Remove highlight immediately if the request failed before starting streaming/display
        removeHighlight(mesDiv, mesId, swipeId);
        return;
    }

    let newText = '';
    try {
        if (typeof res === 'function') {
            // Streaming case
            const streamingSpan = document.createElement('span');
            streamingSpan.className = 'animated-highlight';

            // Replace the selected text with the streaming span
            range.deleteContents();
            range.insertNode(streamingSpan);

            for await (const chunk of res()) {
                newText = chunk.text;
                streamingSpan.textContent = newText;
            }
        } else {
            // Non-streaming case
            newText = res?.choices?.[0]?.message?.content ?? res?.choices?.[0]?.text ?? res?.text ?? '';
            const highlightedNewText = document.createElement('span');
            highlightedNewText.className = 'animated-highlight';
            highlightedNewText.textContent = newText;

            range.deleteContents();
            range.insertNode(highlightedNewText);
        }

        // Remove highlight after x seconds when processing is complete
        const highlightDuration = extension_settings[extensionName].highlightDuration;
        setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

        await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);

    } catch (error) {
        console.error('[Rewrite Extension] Error processing API response:', error);
        toastr.error("Failed to process rewrite response. Check console.", "Processing Error");
        // Ensure highlight is removed if processing fails
        removeHighlight(mesDiv, mesId, swipeId);
    }
    // activateSendButtons is now handled in the finally block above
}

// Updated signature to accept selectionInfo
async function handleSimplifiedChatCompletionRewrite(mesId, swipeId, option, customInstructions, selectionInfo) {
    // Use pre-captured selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = selectionInfo;
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`); // Keep getting mesDiv for highlight/DOM ops
    if (!mesDiv) { // Add check for mesDiv existence
        console.error("[Rewrite Extension] Could not find mesDiv in handleSimplifiedChatCompletionRewrite.");
        return;
    }
    // Get the text completion prompt based on the option
    let promptTemplate;
    switch (option) {
        case 'Rewrite':
            promptTemplate = extension_settings[extensionName].textRewritePrompt;
            break;
        case 'Shorten':
            promptTemplate = extension_settings[extensionName].textShortenPrompt;
            break;
        case 'Expand':
            promptTemplate = extension_settings[extensionName].textExpandPrompt;
            break;
        case 'Custom': // New case
            promptTemplate = extension_settings[extensionName].textCustomPrompt;
            break;
        default:
            console.error("Unknown rewrite option:", option);
            return; // Exit if the option is not recognized
    }

    // Get amount of words
    const wordCount = extractAllWords(selectedRawText).length;

    // Replace macros in the prompt template
    let prompt = getContext().substituteParams(promptTemplate);

    prompt = prompt
        .replace(/{{rewrite}}/gi, selectedRawText)
        .replace(/{{targetmessage}}/gi, fullMessage)
        .replace(/{{rewritecount}}/gi, wordCount);

    // Inject custom instructions if applicable
    if (option === 'Custom') {
        if (prompt.includes('{{custom_instructions}}')) {
            prompt = prompt.replace(/{{custom_instructions}}/gi, customInstructions);
        } else {
            // Append if macro is missing (basic fallback)
            prompt += `\n\nInstructions: ${customInstructions}`;
        }
    }

    // Create a simplified chat format
    const simplifiedChat = [
        {
            role: "system",
            content: prompt
        }
    ];

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;

    // Show the stop button
    getContext().deactivateSendButtons();

    const res = await sendOpenAIRequest('normal', simplifiedChat, abortController.signal);
    window.getSelection().removeAllRanges();

    let newText = '';

    if (typeof res === 'function') {
        // Streaming case
        const streamingSpan = document.createElement('span');
        streamingSpan.className = 'animated-highlight';

        // Replace the selected text with the streaming span
        range.deleteContents();
        range.insertNode(streamingSpan);

        for await (const chunk of res()) {
            newText = chunk.text;
            streamingSpan.textContent = newText;
        }
    } else {
        // Non-streaming case
        newText = res?.choices?.[0]?.message?.content ?? '';
        const highlightedNewText = document.createElement('span');
        highlightedNewText.className = 'animated-highlight';
        highlightedNewText.textContent = newText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

    // Remove highlight after x seconds when streaming is complete
    const highlightDuration = extension_settings[extensionName].highlightDuration;
    setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
    getContext().activateSendButtons();
}

// Updated signature to accept selectionInfo
async function handleTextBasedRewrite(mesId, swipeId, option, customInstructions, selectionInfo) {
    // Use pre-captured selection info
    const { fullMessage, selectedRawText, rawStartOffset, rawEndOffset, range } = selectionInfo;
    const mesDiv = document.querySelector(`[mesid="${mesId}"] .mes_text`); // Keep getting mesDiv for highlight/DOM ops
    if (!mesDiv) { // Add check for mesDiv existence
        console.error("[Rewrite Extension] Could not find mesDiv in handleTextBasedRewrite.");
        return;
    }
    // Get the selected model and option-specific prompt
    const selectedModel = extension_settings[extensionName].selectedModel;
    let promptTemplate;
    switch (option) {
        case 'Rewrite':
            promptTemplate = extension_settings[extensionName].textRewritePrompt;
            break;
        case 'Shorten':
            promptTemplate = extension_settings[extensionName].textShortenPrompt;
            break;
        case 'Expand':
            promptTemplate = extension_settings[extensionName].textExpandPrompt;
            break;
        case 'Custom': // New case
            promptTemplate = extension_settings[extensionName].textCustomPrompt;
            break;
        default:
            console.error('Unknown rewrite option:', option);
            return;
    }

    // Get amount of words
    const wordCount = extractAllWords(selectedRawText).length;

    // Replace macros in the prompt template
    let prompt = getContext().substituteParams(promptTemplate);

    prompt = prompt
        .replace(/{{rewrite}}/gi, selectedRawText)
        .replace(/{{targetmessage}}/gi, fullMessage)
        .replace(/{{rewritecount}}/gi, wordCount);

    // Inject custom instructions if applicable
    if (option === 'Custom') {
        if (prompt.includes('{{custom_instructions}}')) {
            prompt = prompt.replace(/{{custom_instructions}}/gi, customInstructions);
        } else {
            // Append if macro is missing (basic fallback)
            prompt += `\n\nInstructions: ${customInstructions}`;
        }
    }

    let generateData;
    let amount_gen;

    if (extension_settings[extensionName].useDynamicTokens) {
        amount_gen = calculateTargetTokenCount(selectedRawText, option);
    } else {
        switch (option) {
            case 'Rewrite':
                amount_gen = extension_settings[extensionName].rewriteTokens;
                break;
            case 'Shorten':
                amount_gen = extension_settings[extensionName].shortenTokens;
                break;
            case 'Expand':
                amount_gen = extension_settings[extensionName].expandTokens;
                break;
            case 'Custom': // New case
                amount_gen = extension_settings[extensionName].customTokens;
                break;
        }
    }

    // Prepare generation data based on the selected model
    switch (main_api) {
        case 'novel':
            const novelSettings = novelai_settings[novelai_setting_names[nai_settings.preset_settings_novel]];
            generateData = getNovelGenerationData(prompt, novelSettings, amount_gen, false, false, null, 'quiet');
            break;
        case 'textgenerationwebui':
            generateData = getTextGenGenerationData(prompt, amount_gen, false, false, null, 'quiet');
            break;
        case 'koboldhorde':
            if (option === 'Custom') {
                // For Custom Horde, use the manually constructed prompt directly
                // We need a basic structure for generateHorde, mimicking what getContext().generate would provide
                generateData = {
                    prompt: prompt, // Use the manually constructed prompt
                    max_length: Math.max(amount_gen, MIN_LENGTH),
                    // Include other necessary default parameters if generateHorde requires them
                    // Based on generateHorde usage, 'quiet' and potentially others might be needed.
                    quiet: true, // Often used in background generation
                };
            } else {
                // Existing logic for non-custom Horde rewrites
                const promptReadyPromise = new Promise(resolve => {
                    eventSource.once(event_types.GENERATE_AFTER_DATA, resolve);
                });
                getContext().generate('normal', {}, true); // Trigger standard prompt generation
                generateData = await promptReadyPromise; // Wait for the generated data
                generateData.max_length = Math.max(amount_gen, MIN_LENGTH);
            }
            break;
        // Add more cases for other text-based models as needed
        default:
            toastr.error('Unsupported model:', main_api);
            return;
    }

    // Create a new AbortController
    abortController = new AbortController();

    // Store the necessary data in the signal
    abortController.signal.mesDiv = mesDiv;
    abortController.signal.mesId = mesId;
    abortController.signal.swipeId = swipeId;
    abortController.signal.highlightDuration = extension_settings[extensionName].highlightDuration;

    // Show the stop button
    getContext().deactivateSendButtons();
    let res;
    if (extension_settings[extensionName].useStreaming) {
        switch (main_api) {
            case 'textgenerationwebui':
                res = await generateTextGenWithStreaming(generateData, abortController.signal);
                break;
            case 'novel':
                res = await generateNovelWithStreaming(generateData, abortController.signal);
                break;
            case 'koboldhorde':
                toastr.warning('Rewrite streaming not supported for Kobold. Turn off in rewrite settings.');
            default:
                throw new Error('Streaming is enabled, but the current API does not support streaming.');
        }
    } else {
        if (main_api === 'koboldhorde') {
            res = await generateHorde(prompt, generateData, abortController.signal, true);
        } else {
            const response = await generateRaw(prompt, null, false, false, null, generateData.max_length);
            res = {text: response};
            // Shamelessly copied from script.js
            /*function getGenerateUrl(api) {
                switch (api) {
                    case 'textgenerationwebui':
                        return '/api/backends/text-completions/generate';
                    case 'novel':
                        return '/api/novelai/generate';
                    default:
                        throw new Error(`Unknown API: ${api}`);
                }
            }

            const response = await fetch(getGenerateUrl(main_api), {
                method: 'POST',
                headers: getRequestHeaders(),
                cache: 'no-cache',
                body: JSON.stringify(generateData),
                signal: abortController.signal,
            });

            if (!response.ok) {
                const error = await response.json();
                throw error;
            }

            res = await response.json();*/
        }
    }

    window.getSelection().removeAllRanges();

    let newText = '';

    if (typeof res === 'function') {
        // Streaming case

        const streamingSpan = document.createElement('span');
        streamingSpan.className = 'animated-highlight';

        // Replace the selected text with the streaming span
        range.deleteContents();
        range.insertNode(streamingSpan);

        for await (const chunk of res()) {
            newText = chunk.text;
            streamingSpan.textContent = newText;
        }
    } else {
        // Non-streaming case
        newText = res?.choices?.[0]?.message?.content ?? res?.choices?.[0]?.text ?? res?.text ?? '';
        if (main_api === 'novel') newText = res.output;
        const highlightedNewText = document.createElement('span');
        highlightedNewText.className = 'animated-highlight';
        highlightedNewText.textContent = newText;

        range.deleteContents();
        range.insertNode(highlightedNewText);
    }

    // Remove highlight after x seconds when streaming is complete
    const highlightDuration = extension_settings[extensionName].highlightDuration;
    setTimeout(() => removeHighlight(mesDiv, mesId, swipeId), highlightDuration);

    await saveRewrittenText(mesId, swipeId, fullMessage, rawStartOffset, rawEndOffset, newText);
    getContext().activateSendButtons();
}

function calculateTargetTokenCount(selectedText, option) {
    const baseTokenCount = getTokenCount(selectedText);
    const useDynamicTokens = extension_settings[extensionName].useDynamicTokens;
    const dynamicTokenMode = extension_settings[extensionName].dynamicTokenMode;
    let result;

    if (useDynamicTokens) {
        if (dynamicTokenMode === 'additive') {
            let modifier;
            switch (option) {
                case 'Rewrite':
                    modifier = extension_settings[extensionName].rewriteTokensAdd;
                    break;
                case 'Shorten':
                    modifier = extension_settings[extensionName].shortenTokensAdd;
                    break;
                case 'Expand':
                    modifier = extension_settings[extensionName].expandTokensAdd;
                    break;
                case 'Custom':
                    modifier = extension_settings[extensionName].customTokensAdd;
                    break;
            }
            result = baseTokenCount + modifier;
        } else { // multiplicative
            let multiplier;
            switch (option) {
                case 'Rewrite':
                    multiplier = extension_settings[extensionName].rewriteTokensMult;
                    break;
                case 'Shorten':
                    multiplier = extension_settings[extensionName].shortenTokensMult;
                    break;
                case 'Expand':
                    multiplier = extension_settings[extensionName].expandTokensMult;
                    break;
                case 'Custom':
                    multiplier = extension_settings[extensionName].customTokensMult;
                    break;
            }
            result = baseTokenCount * multiplier;
        }
    } else {
        switch (option) {
            case 'Rewrite':
                result = extension_settings[extensionName].rewriteTokens;
                break;
            case 'Shorten':
                result = extension_settings[extensionName].shortenTokens;
                break;
            case 'Expand':
                result = extension_settings[extensionName].expandTokens;
                break;
            case 'Custom':
                result = extension_settings[extensionName].customTokens;
                break;
        }
    }

    return Math.max(1, Math.round(result)); // Ensure at least 1 token and round to nearest integer
}

async function handleUndo(event) {
    const mesId = event.target.dataset.mesId;
    const change = changeHistory.findLast(change => change.mesId === mesId);

    if (change) {
        const context = getContext();
        const messageDiv = document.querySelector(`[mesid="${mesId}"]`);

        if (!messageDiv || !context.chat[mesId]) {
            console.error('Message not found for undo operation');
            return;
        }

        // Update the chat context
        context.chat[mesId].mes = change.originalContent;

        // Only update swipes if they exist
        if (change.swipeId !== undefined && context.chat[mesId].swipes) {
            context.chat[mesId].swipes[change.swipeId] = change.originalContent;
        }

        // Update the UI
        const mesTextElement = messageDiv.querySelector('.mes_text');
        if (mesTextElement) {
            mesTextElement.innerHTML = messageFormatting(
                change.originalContent,
                context.name2,
                context.chat[mesId].isSystem,
                context.chat[mesId].isUser,
                mesId
            );
            addCopyToCodeBlocks(mesTextElement);
        }

        // Save the chat
        await context.saveChat();

        // Remove this change from history
        changeHistory = changeHistory.filter(c => c !== change);

        // Update undo buttons
        updateUndoButtons();
    }
}
