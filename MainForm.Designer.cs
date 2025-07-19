namespace ArabicDocumentGenerator
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.cmbDocumentType = new System.Windows.Forms.ComboBox();
            this.btnOpenForm = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cmbDocumentType
            // 
            this.cmbDocumentType.FormattingEnabled = true;
            this.cmbDocumentType.Location = new System.Drawing.Point(150, 50);
            this.cmbDocumentType.Name = "cmbDocumentType";
            this.cmbDocumentType.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.cmbDocumentType.Size = new System.Drawing.Size(200, 23);
            this.cmbDocumentType.TabIndex = 0;
            // 
            // btnOpenForm
            // 
            this.btnOpenForm.Location = new System.Drawing.Point(150, 100);
            this.btnOpenForm.Name = "btnOpenForm";
            this.btnOpenForm.Size = new System.Drawing.Size(200, 30);
            this.btnOpenForm.TabIndex = 1;
            this.btnOpenForm.Text = "فتح نموذج";
            this.btnOpenForm.UseVisualStyleBackColor = true;
            this.btnOpenForm.Click += new System.EventHandler(this.btnOpenForm_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(356, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "نوع المستند:";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 211);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnOpenForm);
            this.Controls.Add(this.cmbDocumentType);
            this.Name = "MainForm";
            this.Text = "منشئ المستندات العربية";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ComboBox cmbDocumentType;
        private Button btnOpenForm;
        private Label label1;
    }
}