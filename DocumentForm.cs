using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace ArabicDocumentGenerator
{
    public partial class DocumentForm : Form
    {
        private readonly string _documentType;
        private readonly Dictionary<string, string> _placeholders = new Dictionary<string, string>();

        public DocumentForm(string documentType)
        {
            _documentType = documentType;
            InitializeComponent();
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;
            SetupForm();
        }

        private void SetupForm()
        {
            this.Text = _documentType;

            switch (_documentType)
            {
                case "شهادة عدم العمل":
                    CreateUnemploymentCertificateForm();
                    break;
                case "شهادة السكنى":
                    CreateResidenceCertificateForm();
                    break;
                case "طلب خطي":
                    CreateWrittenRequestForm();
                    break;
                case "شهادة الحياة":
                    CreateLifeCertificateForm();
                    break;
            }

            btnGenerate.Text = "إنشاء المستند";
            btnGenerate.Click += BtnGenerate_Click;
        }

        private void CreateUnemploymentCertificateForm()
        {
            AddField("الاسم الكامل", "fullName");
            AddField("رقم البطاقة الوطنية", "idNumber");
            AddField("العنوان", "address");
            AddField("التاريخ", "date");
        }

        private void CreateResidenceCertificateForm()
        {
            AddField("الاسم الكامل", "fullName");
            AddField("رقم البطاقة الوطنية", "idNumber");
            AddField("عنوان السكن", "residenceAddress");
            AddField("مدة السكن", "duration");
            AddField("التاريخ", "date");
        }

        private void CreateWrittenRequestForm()
        {
            AddField("الاسم الكامل", "fullName");
            AddField("رقم البطاقة الوطنية", "idNumber");
            AddField("الجهة المقدمة إليها الطلب", "recipient");
            AddField("محتوى الطلب", "requestContent");
            AddField("التاريخ", "date");
        }

        private void CreateLifeCertificateForm()
        {
            AddField("الاسم الكامل", "fullName");
            AddField("رقم البطاقة الوطنية", "idNumber");
            AddField("مكان الإقامة", "residence");
            AddField("التاريخ", "date");
        }

        private void AddField(string labelText, string fieldName)
        {
            Label label = new Label
            {
                Text = labelText,
                RightToLeft = RightToLeft.Yes,
                Width = 150,
                TextAlign = System.Drawing.ContentAlignment.MiddleRight
            };

            TextBox textBox = new TextBox
            {
                Name = fieldName,
                RightToLeft = RightToLeft.Yes,
                Width = 200
            };

            flowLayoutPanel1.Controls.Add(label);
            flowLayoutPanel1.Controls.Add(textBox);
        }

        private void BtnGenerate_Click(object? sender, EventArgs e)
        {
            try
            {
                // Clear previous placeholders
                _placeholders.Clear();

                // Collect all field values
                foreach (Control control in flowLayoutPanel1.Controls)
                {
                    if (control is TextBox textBox && !string.IsNullOrWhiteSpace(textBox.Name))
                    {
                        _placeholders[textBox.Name] = textBox.Text ?? string.Empty;
                    }
                }

                // Validate that we have some data
                if (_placeholders.Count == 0)
                {
                    MessageBox.Show("لم يتم العثور على بيانات لإدخالها في المستند", "تحذير", 
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Map document type names to template file names
                string templateFileName = GetTemplateFileName(_documentType);
                
                // Generate paths
                string templatePath = Path.Combine("Templates", templateFileName);
                string outputDirectory = "Generated";
                string outputFileName = $"{_documentType}_{DateTime.Now:yyyyMMddHHmmss}.docx";
                string outputPath = Path.Combine(outputDirectory, outputFileName);

                // Debug information
                string debugInfo = $"Template Path: {Path.GetFullPath(templatePath)}\n" +
                                 $"Output Path: {Path.GetFullPath(outputPath)}\n" +
                                 $"Template Exists: {File.Exists(templatePath)}\n" +
                                 $"Placeholders Count: {_placeholders.Count}";

                Console.WriteLine(debugInfo);

                // Check if template exists before proceeding
                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"لم يتم العثور على ملف القالب:\n{Path.GetFullPath(templatePath)}\n\n" +
                                  "تأكد من وجود ملف القالب في مجلد Templates", 
                                  "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Generate the document
                TemplateManager.GenerateDocument(templatePath, outputPath, _placeholders);

                // Show success message with option to open the file
                DialogResult result = MessageBox.Show(
                    $"تم إنشاء المستند بنجاح في:\n{Path.GetFullPath(outputPath)}\n\nهل تريد فتح المجلد؟", 
                    "نجاح", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    // Open the output directory
                    System.Diagnostics.Process.Start("explorer.exe", Path.GetFullPath(outputDirectory));
                }
            }
            catch (FileNotFoundException ex)
            {
                MessageBox.Show($"لم يتم العثور على الملف:\n{ex.Message}", 
                    "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (DirectoryNotFoundException ex)
            {
                MessageBox.Show($"لم يتم العثور على المجلد:\n{ex.Message}", 
                    "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"ليس لديك صلاحية للوصول إلى الملف:\n{ex.Message}", 
                    "خطأ في الصلاحيات", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"حدث خطأ أثناء إنشاء المستند:\n{ex.Message}\n\nتفاصيل إضافية:\n{ex.InnerException?.Message}", 
                    "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetTemplateFileName(string documentType)
        {
            // Map the display names to actual template file names
            return documentType switch
            {
                "شهادة عدم العمل" => "شهادة عدم العمل.docx",
                "شهادة السكنى" => "شهادة السكنى.docx", // Fixed: matches actual template file name
                "طلب خطي" => "طلب خطي.docx",
                "شهادة الحياة" => "شهادة الحياة.docx",
                _ => $"{documentType}.docx"
            };
        }
    }
}