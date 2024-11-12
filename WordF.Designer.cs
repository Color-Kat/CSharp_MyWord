namespace MyWord
{
    partial class WordF
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.textBox = new System.Windows.Forms.RichTextBox();
            this.createWordDocumentButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBox
            // 
            this.textBox.Location = new System.Drawing.Point(22, 66);
            this.textBox.Name = "textBox";
            this.textBox.Size = new System.Drawing.Size(742, 655);
            this.textBox.TabIndex = 0;
            this.textBox.Text = "";
            // 
            // createWordDocumentButton
            // 
            this.createWordDocumentButton.Location = new System.Drawing.Point(22, 741);
            this.createWordDocumentButton.Name = "createWordDocumentButton";
            this.createWordDocumentButton.Size = new System.Drawing.Size(742, 41);
            this.createWordDocumentButton.TabIndex = 1;
            this.createWordDocumentButton.Text = "Create Word Documtn";
            this.createWordDocumentButton.UseVisualStyleBackColor = true;
            this.createWordDocumentButton.Click += new System.EventHandler(this.createWordDocumentButton_Click);
            // 
            // WordF
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(792, 794);
            this.Controls.Add(this.createWordDocumentButton);
            this.Controls.Add(this.textBox);
            this.Name = "WordF";
            this.Text = "My WOrd";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox textBox;
        private System.Windows.Forms.Button createWordDocumentButton;
    }
}

