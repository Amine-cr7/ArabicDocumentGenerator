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

                        string searchValue = $"{{{{{placeholder.Key}}}}}";
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
            var paragraphs = element.Descendants<Paragraph>().ToList();
            
            foreach (var paragraph in paragraphs)
            {
                var runs = paragraph.Elements<Run>().ToList();
                
                for (int i = 0; i < runs.Count - 1; i++)
                {
                    var currentRun = runs[i];
                    var nextRun = runs[i + 1];
                    
                    // Check if runs have similar formatting and can be merged
                    if (CanMergeRuns(currentRun, nextRun))
                    {
                        var currentText = currentRun.GetFirstChild<Text>();
                        var nextText = nextRun.GetFirstChild<Text>();
                        
                        if (currentText != null && nextText != null)
                        {
                            currentText.Text += nextText.Text;
                            nextRun.Remove();
                            runs.RemoveAt(i + 1);
                            i--; // Adjust index after removal
                        }
                    }
                }
            }
        }

        private static bool CanMergeRuns(Run run1, Run run2)
        {
            // Simple check - both runs should have no special formatting or same formatting
            var props1 = run1.RunProperties;
            var props2 = run2.RunProperties;
            
            // If both have no properties, they can be merged
            if (props1 == null && props2 == null) return true;
            
            // For now, only merge if both have no properties
            // This could be enhanced to compare actual properties
            return false;
        }

        private static void SetRightToLeftDirection(MainDocumentPart mainPart)
        {
            try
            {
                // Set BiDi (bidirectional) for all paragraphs
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

                    // Set text alignment to right
                    if (pPr.GetFirstChild<Justification>() == null)
                    {
                        pPr.AppendChild(new Justification() { Val = JustificationValues.Right });
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