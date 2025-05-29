# Redactinator: Securely Remove Sensitive Information from Documents

**Protect your privacy and comply with data regulations by automatically redacting Personally Identifiable Information (PII) and Potentially Sensitive Information (PSI) from your Word (.docx) and PowerPoint (.pptx) documents.**

Redactinator provides a simple web interface to upload your documents. It then intelligently scans them for common types of sensitive data (like names, phone numbers, email addresses, locations, etc.) and "blacks out" or replaces this information before allowing you to download the secured version.

## Key Features

*   **Easy to Use:** Simple drag-and-drop or click-to-select file upload.
*   **Supports Common Formats:** Works with Microsoft Word (.docx) and PowerPoint (.pptx) files.
*   **Automatic PII/PSI Detection:** Leverages advanced Natural Language Processing (NLP) via [Microsoft Presidio](https://microsoft.github.io/presidio/) to identify a wide range of sensitive information.
*   **Secure Redaction:**
    *   For Word documents, identified sensitive text is highlighted in black, effectively obscuring it.
    *   For PowerPoint documents, text containing sensitive information is replaced with black block characters (e.g., '██████').
*   **Privacy-Focused:** Your original uploaded file is processed in memory and not stored permanently on the server after redaction. The redacted file is provided back to you for download.

## How It Works (Simplified)

1.  **Upload:** You upload a `.docx` or `.pptx` file through the web page.
2.  **Scan:** Redactinator reads the text content of your document (including text in paragraphs, tables, and text boxes).
3.  **Identify:** Using a powerful PII detection engine (Presidio), the tool identifies potential sensitive information based on patterns and context.
4.  **Redact:**
    *   In Word files, the identified text segments are highlighted with black.
    *   In PowerPoint files, the text within sections containing PII is replaced with '█' characters.
5.  **Download:** You receive a new version of your document with the sensitive information redacted, ready for safe sharing or storage.

## Who Is This For?

*   Individuals needing to share documents while protecting personal details.
*   Businesses handling customer or employee data that need to remove PII before internal or external sharing for specific purposes (e.g., creating anonymized datasets for analysis, sharing examples without exposing real data).
*   Anyone needing a quick and automated way to reduce the risk of exposing sensitive information in Word and PowerPoint files.

## How to Use Redactinator

1.  **Access the Application:** Open the Redactinator web page in your browser.
2.  **Upload Your Document:**
    *   Drag and drop your `.docx` or `.pptx` file onto the designated "Drop file here" area.
    *   Alternatively, click within the drop zone to open a file selection dialog and choose your file.
3.  **Redact:** Click the "Upload and Redact" button.
4.  **Download:** Once processing is complete, your browser will automatically download the redacted version of your file (usually named `redacted_yourfilename.docx` or `redacted_yourfilename.pptx`).

**Important Notes:**

*   **Accuracy:** While Redactinator uses advanced tools, no automated PII detection is 100% perfect. It's always a good idea to review the redacted document to ensure all desired information has been addressed, especially for highly sensitive scenarios.
*   **Complex Layouts:** Redaction in documents with very complex formatting or embedded objects might have limitations. The tool primarily focuses on standard text content.
*   **Supported PII:** The tool is configured to detect a broad range of common PII types recognized by the English language model of Presidio. For specific or custom PII types (e.g., unique internal identifiers), the underlying recognizers would need to be extended (a task for a technical user).

## Supported File Types

*   Microsoft Word Documents (`.docx`)
*   Microsoft PowerPoint Presentations (`.pptx`)

---
