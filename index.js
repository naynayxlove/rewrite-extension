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

    while (rawIndex < rawText.length && formattedIndex < formattedText.length) {
        if (rawText[rawIndex] === formattedText[formattedIndex]) {
            mapping.push([rawIndex, formattedIndex]);
            rawIndex++;
            formattedIndex++;
            continue;
        }

        if (formattedText[formattedIndex] === ' ' || formattedText[formattedIndex] === '\n') {
            formattedIndex++;
            continue;
        }

        rawIndex++;
    }

    return {
        formattedToRaw: (formattedOffset) => {
            if (mapping.length === 0) {
                return formattedOffset;
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

            if (low >= mapping.length) {
                low = mapping.length - 1;
            } else if (low > 0) {
                low -= 1;
            }

            return mapping[low][0] + (formattedOffset - mapping[low][1]);
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
