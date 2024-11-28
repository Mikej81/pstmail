const fs = require('fs');
const path = require('path');
const { PSTFile, PSTFolder, PSTMessage, PSTAttachment } = require('pst-extractor');

const saveToFS = true;
const verbose = true;
const displaySender = true;
const displayBody = true;
let depth = -1;
let col = 0;

// ANSI escape codes for console highlighting
const ANSI_RED = 31;
const ANSI_YELLOW = 93;
const ANSI_GREEN = 32;
const ANSI_BLUE = 34;

const highlight = (str, code = ANSI_RED) => `\u001b[${code}m${str}\u001b[0m`;

// Get PST file path from arguments
const pstFilePath = process.argv[2];

if (!pstFilePath || !fs.existsSync(pstFilePath)) {
    console.error(highlight('Usage: node processPST.js <path-to-pst-file>', ANSI_RED));
    process.exit(1);
}

// Top-level output folder
const topOutputFolder = './output/';
fs.mkdirSync(topOutputFolder, { recursive: true });

/**
 * Returns a string with visual indication of depth in the tree.
 * @param {number} depth
 * @returns {string}
 */
function getDepth(depth) {
    return ' '.repeat(depth * 2) + '|- ';
}

/**
 * Save email and its attachments to the file system.
 * @param {PSTMessage} msg
 * @param {string} emailFolder
 * @param {string} sender
 * @param {string} recipients
 */
function doSaveToFS(msg, emailFolder, sender, recipients) {
    try {
        const filename = path.join(emailFolder, `${msg.descriptorNodeId}.txt`);
        if (verbose) console.log(highlight(`Saving email to ${filename}`, ANSI_BLUE));
        const fd = fs.openSync(filename, 'w');
        fs.writeSync(fd, `${msg.clientSubmitTime}\r\n`);
        fs.writeSync(fd, `Type: ${msg.messageClass}\r\n`);
        fs.writeSync(fd, `From: ${sender}\r\n`);
        fs.writeSync(fd, `To: ${recipients}\r\n`);
        fs.writeSync(fd, `Subject: ${msg.subject}\r\n`);
        fs.writeSync(fd, msg.body || '');
        fs.closeSync(fd);

        // Save attachments
        for (let i = 0; i < msg.numberOfAttachments; i++) {
            const attachment = msg.getAttachment(i);
            if (attachment && attachment.filename) {
                const attachmentFilename = path.join(
                    emailFolder,
                    `${msg.descriptorNodeId}-${attachment.longFilename}`
                );
                if (verbose) console.log(highlight(`Saving attachment to ${attachmentFilename}`, ANSI_BLUE));
                const fd = fs.openSync(attachmentFilename, 'w');
                const attachmentStream = attachment.fileInputStream;
                if (attachmentStream) {
                    const buffer = Buffer.alloc(8176);
                    let bytesRead;
                    do {
                        bytesRead = attachmentStream.read(buffer);
                        if (bytesRead > 0) fs.writeSync(fd, buffer, 0, bytesRead);
                    } while (bytesRead === 8176);
                    fs.closeSync(fd);
                }
            }
        }
    } catch (err) {
        console.error(err);
    }
}

/**
 * Process each folder recursively, maintaining the PST folder structure.
 * @param {PSTFolder} folder
 * @param {string} currentPath - Current folder path for saving content.
 */
function processFolder(folder, currentPath) {
    depth++;
    const folderPath = path.join(currentPath, sanitizeFolderName(folder.displayName));

    // Create the folder in the output directory
    if (!fs.existsSync(folderPath)) fs.mkdirSync(folderPath, { recursive: true });

    if (depth > 0) console.log(getDepth(depth) + folder.displayName);

    // Process subfolders
    if (folder.hasSubfolders) {
        const subFolders = folder.getSubFolders();
        subFolders.forEach((subFolder) => processFolder(subFolder, folderPath));
    }

    // Process emails
    if (folder.contentCount > 0) {
        depth++;
        let email = folder.getNextChild();
        while (email) {
            if (email instanceof PSTMessage) {
                const sender = getSender(email);
                const recipients = getRecipients(email);

                if (verbose && displayBody) {
                    console.log(highlight('Email body:', ANSI_YELLOW), email.body);
                }

                if (saveToFS) {
                    doSaveToFS(email, folderPath, sender, recipients);
                }
            }
            email = folder.getNextChild();
        }
        depth--;
    }
    depth--;
}

/**
 * Sanitize folder names to remove invalid characters.
 * @param {string} name
 * @returns {string}
 */
function sanitizeFolderName(name) {
    if (!name) return 'unknown_folder';
    return name.replace(/[^a-zA-Z0-9-_ ]/g, '_');
}

/**
 * Get the sender and display it.
 * @param {PSTMessage} email
 * @returns {string}
 */
function getSender(email) {
    let sender = email.senderName || 'Unknown';
    if (sender !== email.senderEmailAddress) sender += ` (${email.senderEmailAddress})`;
    if (verbose && displaySender) console.log(getDepth(depth) + `Sender: ${sender}`);
    return sender;
}

/**
 * Get the recipients and display them.
 * @param {PSTMessage} email
 * @returns {string}
 */
function getRecipients(email) {
    return email.displayTo || 'Unknown';
}

// Process the provided PST file
console.log(highlight(`Processing file: ${pstFilePath}`, ANSI_GREEN));
const pstFile = new PSTFile(fs.readFileSync(pstFilePath));

const outputRootFolder = path.join(
    topOutputFolder,
    sanitizeFolderName(path.basename(pstFilePath, path.extname(pstFilePath)))
);
fs.mkdirSync(outputRootFolder, { recursive: true });

processFolder(pstFile.getRootFolder(), outputRootFolder);

console.log(highlight(`Finished processing ${pstFilePath}`, ANSI_GREEN));
