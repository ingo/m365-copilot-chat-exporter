#!/usr/bin/env node

/**
 * convert.js
 * 
 * Converts Microsoft 365 Copilot chat exports into clean, formatted Markdown (.md)
 * optimized for Google NotebookLM and Microsoft OneNote.
 * 
 * Supports both formats:
 *  1. Raw Substrate API Captures (`copilot_raw_capture_*.json`)
 *     - Extracts full bot responses from `adaptiveCards` (TextBlock, RichTextBlock)
 *     - Preserves web citations / source references
 *  2. ChatGPT-compatible Conversations (`copilot_conversations_*.json` or `conversations.json`)
 *     - Traverses conversation mapping tree
 * 
 * Usage:
 *   node convert.js [inputFile] [outputDir]
 * 
 * Examples:
 *   node convert.js
 *   node convert.js "copilot_raw_capture_2026-10-11 (1).json"
 *   node convert.js "copilot_conversations_2026-10-11.json" ./my_markdown_chats
 */

const fs = require('fs');
const path = require('path');

// ── HTML & Text Sanitization ──────────────────────────────────────────────

/**
 * Clean up HTML tags commonly injected by Copilot / Substrate in user and bot text.
 */
function cleanHtml(text) {
  if (!text || typeof text !== 'string') return '';

  return text
    // Replace <br> and <br/> with linebreaks
    .replace(/<br\s*\/?>/gi, '\n')
    // Convert <a> links into markdown: <a href="URL"...>TEXT</a> -> [TEXT](URL)
    .replace(/<a\s+[^>]*href=["']([^"']+)["'][^>]*>(.*?)<\/a>/gi, (match, url, label) => {
      const cleanLabel = label.replace(/<[^>]+>/g, '').trim();
      return `[${cleanLabel || url}](${url})`;
    })
    // Strip empty spans / font wrappers while preserving inner text
    .replace(/<span[^>]*>(.*?)<\/span>/gi, '$1')
    .replace(/<div[^>]*>(.*?)<\/div>/gi, '\n$1\n')
    .replace(/<p[^>]*>(.*?)<\/p>/gi, '\n$1\n')
    // Strip any remaining unwanted HTML tags
    .replace(/<[^>]+>/g, '')
    // Normalize excessive consecutive blank lines
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

/**
 * Sanitize strings for safe cross-platform file names (macOS, Windows, Linux).
 */
function sanitizeFilename(name) {
  if (!name) return 'untitled_chat';
  const clean = name
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '') // remove diacritics
    .replace(/[^a-z0-9\s_-]/g, '')
    .trim()
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .slice(0, 70);
  return clean || 'untitled_chat';
}

function formatDate(isoOrTs) {
  if (!isoOrTs) return { dateStr: 'unknown_date', fullStr: 'Unknown Date' };
  let d;
  if (typeof isoOrTs === 'number') {
    d = new Date(isoOrTs > 1e12 ? isoOrTs : isoOrTs * 1000);
  } else {
    d = new Date(isoOrTs);
  }

  if (isNaN(d.getTime())) {
    return { dateStr: 'unknown_date', fullStr: 'Unknown Date' };
  }

  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  const hh = String(d.getUTCHours()).padStart(2, '0');
  const min = String(d.getUTCMinutes()).padStart(2, '0');
  const ss = String(d.getUTCSeconds()).padStart(2, '0');

  return {
    dateStr: `${yyyy}-${mm}-${dd}`,
    fullStr: `${yyyy}-${mm}-${dd} ${hh}:${min}:${ss} UTC`
  };
}

// ── AdaptiveCard & Reference Extractors (Raw Substrate Format) ────────────

/**
 * Recursively extract text from AdaptiveCard elements.
 */
function extractAdaptiveCardText(card) {
  if (!card) return '';
  const parts = [];

  function walk(node) {
    if (!node) return;

    if (node.type === 'TextBlock' && node.text) {
      parts.push(node.text);
    } else if (node.type === 'RichTextBlock' && Array.isArray(node.inlines)) {
      const inlineText = node.inlines.map((i) => i.text || '').join('');
      if (inlineText) parts.push(inlineText);
    }

    if (Array.isArray(node.body)) {
      node.body.forEach(walk);
    }
    if (Array.isArray(node.items)) {
      node.items.forEach(walk);
    }
    if (Array.isArray(node.columns)) {
      node.columns.forEach((col) => {
        if (Array.isArray(col.items)) col.items.forEach(walk);
      });
    }
  }

  walk(card);
  return parts.join('\n\n').trim();
}

/**
 * Format source citations/references attached to Copilot bot messages.
 */
function formatReferences(referencesObj) {
  if (!referencesObj || typeof referencesObj !== 'object') return '';
  const entries = Object.entries(referencesObj);
  if (entries.length === 0) return '';

  const refLines = [];

  for (const [key, ref] of entries) {
    const link = ref.targetLink;
    if (!link) continue;

    let title = link;
    let label = key;

    if (ref.displayData?.content) {
      try {
        const parsed = JSON.parse(ref.displayData.content);
        if (parsed.Title) title = parsed.Title;
        else if (parsed.providerDisplayName) title = parsed.providerDisplayName;
        if (parsed.label) label = parsed.label;
      } catch {}
    }

    refLines.push(`- [${label}] [${title}](${link})`);
  }

  if (refLines.length === 0) return '';
  return `\n\n#### 📚 References & Sources\n${refLines.join('\n')}`;
}

// ── Parsers ───────────────────────────────────────────────────────────────

/**
 * Format 1: Parse Raw Substrate Captures (copilot_raw_capture_*.json)
 */
function parseRawCapture(jsonArray) {
  const conversations = [];

  const convRecords = jsonArray.filter(
    (item) => item?.url && item.url.includes('GetConversation') && item.data?.messages
  );

  for (const rec of convRecords) {
    const data = rec.data;
    const convId = data.conversationId || 'unknown_id';
    const rawMessages = data.messages || [];

    const messages = [];

    for (const msg of rawMessages) {
      if (msg.author === 'system' || msg.author === 'tool') continue;
      if (msg.messageType && msg.messageType !== 'GeneratedCode' && !msg.text && !msg.adaptiveCards?.length) {
        continue;
      }

      let role = msg.author === 'user' ? 'user' : 'assistant';
      let content = '';

      // User prompt
      if (role === 'user') {
        content = cleanHtml(msg.text || '');
      } else {
        // Assistant response: check msg.text first, then adaptiveCards
        if (msg.text && msg.text.trim()) {
          content = cleanHtml(msg.text);
        } else if (Array.isArray(msg.adaptiveCards) && msg.adaptiveCards.length > 0) {
          const cardTexts = msg.adaptiveCards
            .map(extractAdaptiveCardText)
            .filter(Boolean);
          content = cardTexts.join('\n\n').trim();
        }

        // Include generated code if present
        if (msg.messageType === 'GeneratedCode' && msg.text && !content.includes(msg.text.slice(0, 30))) {
          content = content ? `${content}\n\n${msg.text}` : msg.text;
        }

        // Attach citations/references if available
        if (msg.references) {
          const refsStr = formatReferences(msg.references);
          if (refsStr) content += refsStr;
        }
      }

      if (!content.trim()) continue;

      messages.push({
        role,
        content,
        timestamp: msg.createdAt || msg.timestamp || null
      });
    }

    if (messages.length === 0) continue;

    // Determine title
    let title = data.chatName ? cleanHtml(data.chatName) : '';
    if (!title) {
      const firstUserMsg = messages.find((m) => m.role === 'user');
      title = firstUserMsg ? firstUserMsg.content.slice(0, 60).replace(/\n/g, ' ') : 'Untitled Chat';
    }

    conversations.push({
      id: convId,
      title,
      createTime: data.createTimeUtc || messages[0]?.timestamp || null,
      updateTime: data.updateTimeUtc || messages[messages.length - 1]?.timestamp || null,
      messages
    });
  }

  return conversations;
}

/**
 * Format 2: Parse ChatGPT-compatible Conversations (conversations.json)
 */
function parseChatGPTFormat(jsonData) {
  const convList = Array.isArray(jsonData)
    ? jsonData
    : Array.isArray(jsonData.conversations)
    ? jsonData.conversations
    : [];

  const conversations = [];

  for (const conv of convList) {
    if (!conv.mapping || typeof conv.mapping !== 'object') continue;

    const messages = [];

    // Traverse the mapping tree by finding nodes in chronological order
    const nodes = Object.values(conv.mapping).filter(
      (n) => n.message && n.message.author && n.message.author.role !== 'system'
    );

    // Sort by message create_time
    nodes.sort((a, b) => (a.message.create_time || 0) - (b.message.create_time || 0));

    for (const node of nodes) {
      const msg = node.message;
      const role = msg.author.role === 'user' ? 'user' : 'assistant';
      const parts = msg.content?.parts || [];
      const text = cleanHtml(parts.join('\n'));

      if (!text.trim()) continue;

      messages.push({
        role,
        content: text,
        timestamp: msg.create_time ? new Date(msg.create_time * 1000).toISOString() : null
      });
    }

    if (messages.length === 0) continue;

    conversations.push({
      id: conv.id || conv.conversation_id || 'unknown_id',
      title: conv.title ? cleanHtml(conv.title) : 'Untitled Chat',
      createTime: conv.create_time ? conv.create_time * 1000 : null,
      updateTime: conv.update_time ? conv.update_time * 1000 : null,
      messages
    });
  }

  return conversations;
}

// ── Markdown Formatting ───────────────────────────────────────────────────

function renderConversationMarkdown(conv) {
  const { fullStr: createdFull } = formatDate(conv.createTime);
  const { fullStr: updatedFull } = formatDate(conv.updateTime);

  let md = `# ${conv.title}\n\n`;
  md += `> **Date:** ${createdFull}  \n`;
  if (conv.updateTime && conv.updateTime !== conv.createTime) {
    md += `> **Last Updated:** ${updatedFull}  \n`;
  }
  md += `> **Source:** Microsoft 365 Copilot  \n`;
  md += `> **Conversation ID:** \`${conv.id}\`\n\n`;
  md += `---\n\n`;

  for (let i = 0; i < conv.messages.length; i++) {
    const msg = conv.messages[i];
    const isUser = msg.role === 'user';
    const icon = isUser ? '👤' : '🤖';
    const roleName = isUser ? 'User' : 'Copilot';
    const timeInfo = msg.timestamp ? ` *(${formatDate(msg.timestamp).fullStr})*` : '';

    md += `## ${icon} ${roleName}${timeInfo}\n\n`;
    md += `${msg.content}\n\n`;

    if (i < conv.messages.length - 1) {
      md += `---\n\n`;
    }
  }

  return md;
}

// ── CLI & Orchestration ───────────────────────────────────────────────────

function detectFormatAndParse(filePath) {
  const rawData = fs.readFileSync(filePath, 'utf8');
  const json = JSON.parse(rawData);

  // Check Format 1: Raw capture array
  if (Array.isArray(json) && json.some((item) => item?.url && item?.data)) {
    console.log(`[Format Detected] Raw Substrate API Capture (${json.length} records)`);
    return parseRawCapture(json);
  }

  // Check Format 2: ChatGPT format
  if (Array.isArray(json) || (json && Array.isArray(json.conversations))) {
    console.log(`[Format Detected] ChatGPT-compatible format`);
    return parseChatGPTFormat(json);
  }

  throw new Error(`Unrecognized JSON format in: ${filePath}`);
}

function findDefaultInputFile() {
  const files = fs.readdirSync(process.cwd());

  // Prioritize raw capture with "(1)" or largest size
  const rawFiles = files.filter((f) => f.startsWith('copilot_raw_capture') && f.endsWith('.json'));
  if (rawFiles.length > 0) {
    // pick largest raw file
    rawFiles.sort((a, b) => fs.statSync(b).size - fs.statSync(a).size);
    return rawFiles[0];
  }

  const convFiles = files.filter((f) => f.startsWith('copilot_conversations') && f.endsWith('.json'));
  if (convFiles.length > 0) {
    return convFiles[0];
  }

  const standardFile = files.find((f) => f === 'conversations.json');
  if (standardFile) return standardFile;

  return null;
}

function main() {
  const args = process.argv.slice(2);
  let inputFile = args[0];
  let outputDir = args[1] || path.join(process.cwd(), 'output');

  if (!inputFile) {
    inputFile = findDefaultInputFile();
    if (!inputFile) {
      console.error('Error: No input JSON file found.');
      console.log('Usage: node convert.js [inputFile] [outputDir]');
      process.exit(1);
    }
    console.log(`[Auto-selected input file] ${inputFile}`);
  }

  if (!fs.existsSync(inputFile)) {
    console.error(`Error: File not found: ${inputFile}`);
    process.exit(1);
  }

  console.log(`Reading: ${inputFile}`);
  const conversations = detectFormatAndParse(inputFile);
  console.log(`Successfully parsed ${conversations.length} conversations.\n`);

  // Create output directories
  const individualDir = path.join(outputDir, 'individual_chats');
  fs.mkdirSync(individualDir, { recursive: true });

  const usedFilenames = new Map();
  const summaryEntries = [];

  let totalMessages = 0;

  // 1. Export Individual Markdown Files (Ideal for OneNote)
  conversations.forEach((conv, idx) => {
    totalMessages += conv.messages.length;
    const { dateStr } = formatDate(conv.createTime);
    const baseSlug = `${dateStr}_${sanitizeFilename(conv.title)}`;

    let filename = `${baseSlug}.md`;
    let count = 1;
    while (usedFilenames.has(filename)) {
      count++;
      filename = `${baseSlug}-${count}.md`;
    }
    usedFilenames.set(filename, true);

    const filePath = path.join(individualDir, filename);
    const mdContent = renderConversationMarkdown(conv);
    fs.writeFileSync(filePath, mdContent, 'utf8');

    summaryEntries.push({
      title: conv.title,
      date: dateStr,
      filename,
      messagesCount: conv.messages.length
    });
  });

  console.log(`✅ Exported ${conversations.length} individual .md files into:\n   ${individualDir}\n`);

  // 2. Export Single Consolidated File (Ideal for Google NotebookLM 50-source limit)
  const notebookLmFile = path.join(outputDir, 'notebooklm_all_conversations.md');
  let consolidatedMd = `# Microsoft 365 Copilot — Complete Chat Archive\n\n`;
  consolidatedMd += `> Generated on: ${new Date().toISOString().replace('T', ' ').slice(0, 19)} UTC  \n`;
  consolidatedMd += `> Total Conversations: ${conversations.length}  \n`;
  consolidatedMd += `> Total Messages: ${totalMessages}\n\n`;
  consolidatedMd += `## Table of Contents\n\n`;

  summaryEntries.forEach((s, i) => {
    const anchor = s.title
      .toLowerCase()
      .replace(/[^a-z0-9\s-]/g, '')
      .replace(/\s+/g, '-');
    consolidatedMd += `${i + 1}. [${s.title} (${s.date})](#${anchor}) — *${s.messagesCount} messages*\n`;
  });

  consolidatedMd += `\n---\n\n`;

  conversations.forEach((conv, idx) => {
    consolidatedMd += renderConversationMarkdown(conv);
    if (idx < conversations.length - 1) {
      consolidatedMd += `\n\n---\n\n`;
    }
  });

  fs.writeFileSync(notebookLmFile, consolidatedMd, 'utf8');
  console.log(`✅ Exported consolidated NotebookLM document into:\n   ${notebookLmFile}\n`);

  console.log(`Summary:`);
  console.log(`- Total Conversations: ${conversations.length}`);
  console.log(`- Total Messages: ${totalMessages}`);
  console.log(`- Ready for Google NotebookLM (upload: output/notebooklm_all_conversations.md)`);
  console.log(`- Ready for OneNote (import/drag: output/individual_chats/*.md)`);
}

main();
