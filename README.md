# Delete & Rewrite Extension

A powerful utility for SillyTavern that allows you to surgically edit chat messages by selecting text. This extension is a fork of Splitclover's Rewrite Extension, enhanced with deletion capabilities and improved selection handling.

## Features

### 1. Selection-Based Actions
When you select text within a chat message, a context menu appears with the following options:
- **Delete**: Instantly removes the selected text from the message.
- **Rewrite**: Opens a popup allowing you to manually enter new text to replace the selection.
- **Custom (LLM Powered)**: (If configured) Allows you to provide instructions to an LLM to rewrite the selected portion of the message.

### 2. Intelligent Mapping
The extension uses a sophisticated text mapping system to ensure that when you select formatted text (Markdown/HTML), the changes are applied correctly to the underlying raw message content. It even includes heuristics to automatically include surrounding Markdown syntax (like bold `**` or italics `*`) if your selection abuts them.

### 3. Undo System
- Every action is tracked in a local history.
- An **Undo** button (counter-clockwise arrow icon) appears in the message button area after an edit.
- Supports up to **15 undo steps** across the current chat session.
- *Note: History is cleared when switching chats.*

### 4. Customizable UI
Toggle buttons on or off via the Extension Settings panel:
- **Show Rewrite Button**: Enable/Disable the manual rewrite option.
- **Show Delete Button**: Enable/Disable the quick delete option.

## Installation

1. Clone or download this repository into your SillyTavern `extensions` folder.
2. The extension will automatically load on the next start or when you refresh the UI.

## Requirements
- Works best with the latest version of SillyTavern.
- For "Custom" rewrite features, an active LLM API connection (OpenAI, NovelAI, KoboldCpp, etc.) is required.

---
*Original Extension by [splitclover](https://github.com/splitclover/rewrite-extension)*  
*Enhanced and Forked by Revivalist-Dev*
