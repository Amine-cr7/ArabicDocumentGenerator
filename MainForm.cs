using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace ArabicDocumentGenerator
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            LoadDocumentTypes();
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;
        }

        private void LoadDocumentTypes()
        {
            List<string> documentTypes = new List<string>
            {
                "شهادة عدم العمل",
                "شهادة السكنى",
                "طلب خطي",
                "شهادة الحياة"
            };

            cmbDocumentType.DataSource = documentTypes;
        }

        private void btnOpenForm_Click(object sender, EventArgs e)
        {
            // Add null check for SelectedItem
            if (cmbDocumentType.SelectedItem == null)
            {
                MessageBox.Show("الرجاء اختيار نوع المستند", "تحذير", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string selectedDocument = cmbDocumentType.SelectedItem.ToString()!; // Add null-forgiving operator
            DocumentForm documentForm = new DocumentForm(selectedDocument);
            documentForm.ShowDialog();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Center the form on the screen
            this.StartPosition = FormStartPosition.CenterScreen;
        }
    }
}