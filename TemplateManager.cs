using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ArabicDocumentGenerator
{
    public static class TemplateManager
    {
        public static void GenerateDocument(string templatePath, string outputPath, Dictionary<string, string> placeholders)
        {
            // Validate inputs
            if (string.IsNullOrWhiteSpace(templatePath))
                throw new ArgumentException("Template path cannot be null or empty", nameof(templatePath));

            if (string.IsNullOrWhiteSpace(outputPath))
                throw new ArgumentException("Output path cannot be null or empty", nameof(outputPath));

            if (placeholders == null)
                throw new ArgumentNullException(nameof(placeholders));

            // Convert relative paths to absolute paths
            string absoluteTemplatePath = Path.GetFullPath(templatePath);
            string absoluteOutputPath = Path.GetFullPath(outputPath);

            if (!File.Exists(absoluteTemplatePath))
                throw new FileNotFoundException($"Template file not found at: {absoluteTemplatePath}", absoluteTemplatePath);

            try
            {
                // Validate template file integrity first
                ValidateWordDocument(absoluteTemplatePath);

                // Ensure output directory exists
                string? outputDirectory = Path.GetDirectoryName(absoluteOutputPath);
                if (!string.IsNullOrWhiteSpace(outputDirectory) && !Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                }

                // Copy template to output location
                File.Copy(absoluteTemplatePath, absoluteOutputPath, overwrite: true);

                // Open the document for editing
                using (WordprocessingDocument doc = WordprocessingDocument.Open(absoluteOutputPath, true))
                {
                    // Get the main document part
                    MainDocumentPart? mainPart = doc.MainDocumentPart;
                    if (mainPart?.Document == null)
                        throw new InvalidOperationException("The document doesn't contain a main document part or body.");

                    // Replace all placeholders in the document
                    foreach (var placeholder in placeholders)
                    {
                        if (string.IsNullOrWhiteSpace(placeholder.Key))
                            continue;

                        string searchValue = "{" + placeholder.Key + "}";
                        string newValue = placeholder.Value ?? string.Empty;

                        // Search and replace in the document
                        ReplacePlaceholder(mainPart, searchValue, newValue);
                    }

                    // Set Right-to-Left direction for the whole document
                    SetRightToLeftDirection(mainPart);

                    // Save changes
                    mainPart.Document.Save();
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException($"An error occurred while generating the document: {ex.Message}", ex);
            }
        }

        private static void ReplacePlaceholder(MainDocumentPart mainPart, string searchValue, string newValue)
        {
            // Handle text that might be split across multiple Text elements
            var body = mainPart.Document.Body;
            if (body == null) return;

            // First, try to find and replace direct matches
            foreach (var text in body.Descendants<Text>())
            {
                if (text.Text != null && text.Text.Contains(searchValue))
                {
                    text.Text = text.Text.Replace(searchValue, newValue);
                }
            }

            // Handle cases where placeholder text might be split across runs
            MergeConsecutiveRuns(body);
            
            // Try replacement again after merging
            foreach (var text in body.Descendants<Text>())
            {
                if (text.Text != null && text.Text.Contains(searchValue))
                {
                    text.Text = text.Text.Replace(searchValue, newValue);
                }
            }
        }

        private static void MergeConsecutiveRuns(OpenXmlElement element)
        {
            // Process paragraphs directly without creating intermediate lists
            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                // Use a reverse iteration approach to avoid index shifting issues
                var runs = paragraph.Elements<Run>().ToArray(); // Single conversion to array
                
                for (int i = runs.Length - 1; i > 0; i--)
                {
                    var currentRun = runs[i - 1];
                    var nextRun = runs[i];
                    
                    // Check if runs have similar formatting and can be merged
                    if (CanMergeRuns(currentRun, nextRun))
                    {
                        var currentText = currentRun.GetFirstChild<Text>();
                        var nextText = nextRun.GetFirstChild<Text>();
                        
                        if (currentText != null && nextText != null)
                        {
                            // Merge text content
                            currentText.Text += nextText.Text;
                            // Remove the next run from the document
                            nextRun.Remove();
                        }
                    }
                }
            }
        }

        private static bool CanMergeRuns(Run run1, Run run2)
        {
            var props1 = run1.RunProperties;
            var props2 = run2.RunProperties;
            
            // If both have no properties, they can be merged
            if (props1 == null && props2 == null) return true;
            
            // If one has properties and the other doesn't, they cannot be merged
            if ((props1 == null) != (props2 == null)) return false;
            
            // Both have properties - compare key formatting attributes
            if (props1 != null && props2 != null)
            {
                // Compare bold formatting
                var bold1 = props1.GetFirstChild<Bold>() != null;
                var bold2 = props2.GetFirstChild<Bold>() != null;
                if (bold1 != bold2) return false;
                
                // Compare italic formatting
                var italic1 = props1.GetFirstChild<Italic>() != null;
                var italic2 = props2.GetFirstChild<Italic>() != null;
                if (italic1 != italic2) return false;
                
                // Compare font size
                var fontSize1 = props1.GetFirstChild<FontSize>()?.Val?.Value;
                var fontSize2 = props2.GetFirstChild<FontSize>()?.Val?.Value;
                if (fontSize1 != fontSize2) return false;
                
                // Compare RTL text property
                var rtl1 = props1.GetFirstChild<RightToLeftText>() != null;
                var rtl2 = props2.GetFirstChild<RightToLeftText>() != null;
                if (rtl1 != rtl2) return false;
            }
            
            return true;
        }

        private static void SetRightToLeftDirection(MainDocumentPart mainPart)
        {
            try
            {
                // Set BiDi (bidirectional) for all paragraphs and runs
                var body = mainPart.Document.Body;
                if (body == null) return;

                foreach (var paragraph in body.Descendants<Paragraph>())
                {
                    // Get or create paragraph properties
                    ParagraphProperties? pPr = paragraph.GetFirstChild<ParagraphProperties>();
                    if (pPr == null)
                    {
                        pPr = new ParagraphProperties();
                        paragraph.InsertAt(pPr, 0);
                    }

                    // Set BiDi (bidirectional text) - this enables RTL
                    if (pPr.GetFirstChild<BiDi>() == null)
                    {
                        pPr.AppendChild(new BiDi());
                    }

                    // Only set right alignment for paragraphs that don't already have center alignment
                    var existingJustification = pPr.GetFirstChild<Justification>();
                    if (existingJustification == null || existingJustification.Val == null || 
                        existingJustification.Val.Value == JustificationValues.Left)
                    {
                        if (existingJustification != null)
                            existingJustification.Remove();
                        pPr.AppendChild(new Justification() { Val = JustificationValues.Right });
                    }

                    // Set RTL properties for all runs in the paragraph
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        RunProperties? rPr = run.GetFirstChild<RunProperties>();
                        if (rPr == null)
                        {
                            rPr = new RunProperties();
                            run.InsertAt(rPr, 0);
                        }

                        // Set RTL for the run
                        if (rPr.GetFirstChild<RightToLeftText>() == null)
                        {
                            rPr.AppendChild(new RightToLeftText());
                        }
                    }
                }

                // Try to set document-level RTL settings
                SetDocumentRTLSettings(mainPart);
            }
            catch (Exception)
            {
                // RTL setting failed, but continue - not critical for basic functionality
            }
        }

        private static void SetDocumentRTLSettings(MainDocumentPart mainPart)
        {
            try
            {
                // Get or create document settings part
                DocumentSettingsPart? settingsPart = mainPart.DocumentSettingsPart;
                if (settingsPart == null)
                {
                    settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                    settingsPart.Settings = new Settings();
                }

                var settings = settingsPart.Settings;
                if (settings == null)
                {
                    settings = new Settings();
                    settingsPart.Settings = settings;
                }

                // Add default tab stop for RTL-friendly settings
                if (settings.GetFirstChild<DefaultTabStop>() == null)
                {
                    settings.AppendChild(new DefaultTabStop() { Val = 708 }); // 0.5 inch
                }
            }
            catch (Exception)
            {
                // Document settings failed, but continue
            }
        }

        private static void ValidateWordDocument(string filePath)
        {
            try
            {
                // Try to open the document to validate its structure
                using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
                {
                    if (doc.MainDocumentPart?.Document == null)
                    {
                        throw new InvalidDataException("File contains corrupted data.");
                    }

                    // Additional validation: check if the document has a valid body
                    var body = doc.MainDocumentPart.Document.Body;
                    if (body == null)
                    {
                        throw new InvalidDataException("File contains corrupted data.");
                    }
                }
            }
            catch (OpenXmlPackageException ex)
            {
                throw new InvalidDataException("File contains corrupted data.", ex);
            }
            catch (InvalidDataException)
            {
                throw; // Re-throw our custom message
            }
            catch (Exception ex)
            {
                throw new InvalidDataException("File contains corrupted data.", ex);
            }
        }
    }
}