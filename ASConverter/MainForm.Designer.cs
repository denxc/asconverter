namespace ASConverter {
    partial class MainForm {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent() {
            this.sourceBox = new System.Windows.Forms.ListBox();
            this.destBox = new System.Windows.Forms.ListBox();
            this.sourceFilesLabel = new System.Windows.Forms.Label();
            this.destFilesLabel = new System.Windows.Forms.Label();
            this.addSourceFileButton = new System.Windows.Forms.Button();
            this.addDestFileButton = new System.Windows.Forms.Button();
            this.removeSourceFileButton = new System.Windows.Forms.Button();
            this.removeDestFileButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.exportAndOpenButton = new System.Windows.Forms.Button();
            this.sourceFileInfoBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.destFileInfoBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // sourceBox
            // 
            this.sourceBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.sourceBox.FormattingEnabled = true;
            this.sourceBox.Location = new System.Drawing.Point(12, 26);
            this.sourceBox.Name = "sourceBox";
            this.sourceBox.Size = new System.Drawing.Size(211, 275);
            this.sourceBox.TabIndex = 1;
            this.sourceBox.SelectedIndexChanged += new System.EventHandler(this.sourceBox_SelectedIndexChanged);
            // 
            // destBox
            // 
            this.destBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.destBox.FormattingEnabled = true;
            this.destBox.Location = new System.Drawing.Point(385, 26);
            this.destBox.Name = "destBox";
            this.destBox.Size = new System.Drawing.Size(211, 275);
            this.destBox.TabIndex = 2;
            this.destBox.SelectedIndexChanged += new System.EventHandler(this.destBox_SelectedIndexChanged);
            // 
            // sourceFilesLabel
            // 
            this.sourceFilesLabel.AutoSize = true;
            this.sourceFilesLabel.Location = new System.Drawing.Point(9, 10);
            this.sourceFilesLabel.Name = "sourceFilesLabel";
            this.sourceFilesLabel.Size = new System.Drawing.Size(98, 13);
            this.sourceFilesLabel.TabIndex = 3;
            this.sourceFilesLabel.Text = "Исходные файлы:";
            // 
            // destFilesLabel
            // 
            this.destFilesLabel.AutoSize = true;
            this.destFilesLabel.Location = new System.Drawing.Point(382, 10);
            this.destFilesLabel.Name = "destFilesLabel";
            this.destFilesLabel.Size = new System.Drawing.Size(109, 13);
            this.destFilesLabel.TabIndex = 4;
            this.destFilesLabel.Text = "Файлы назначения:";
            // 
            // addSourceFileButton
            // 
            this.addSourceFileButton.Location = new System.Drawing.Point(229, 26);
            this.addSourceFileButton.Name = "addSourceFileButton";
            this.addSourceFileButton.Size = new System.Drawing.Size(33, 27);
            this.addSourceFileButton.TabIndex = 5;
            this.addSourceFileButton.Tag = "";
            this.addSourceFileButton.Text = "+";
            this.addSourceFileButton.UseVisualStyleBackColor = true;
            this.addSourceFileButton.Click += new System.EventHandler(this.addSourceFileButton_Click);
            // 
            // addDestFileButton
            // 
            this.addDestFileButton.Location = new System.Drawing.Point(602, 26);
            this.addDestFileButton.Name = "addDestFileButton";
            this.addDestFileButton.Size = new System.Drawing.Size(33, 27);
            this.addDestFileButton.TabIndex = 6;
            this.addDestFileButton.Tag = "";
            this.addDestFileButton.Text = "+";
            this.addDestFileButton.UseVisualStyleBackColor = true;
            this.addDestFileButton.Click += new System.EventHandler(this.addDestFileButton_Click);
            // 
            // removeSourceFileButton
            // 
            this.removeSourceFileButton.Location = new System.Drawing.Point(229, 59);
            this.removeSourceFileButton.Name = "removeSourceFileButton";
            this.removeSourceFileButton.Size = new System.Drawing.Size(33, 27);
            this.removeSourceFileButton.TabIndex = 7;
            this.removeSourceFileButton.Tag = "";
            this.removeSourceFileButton.Text = "-";
            this.removeSourceFileButton.UseVisualStyleBackColor = true;
            this.removeSourceFileButton.Click += new System.EventHandler(this.removeSourceFileButton_Click);
            // 
            // removeDestFileButton
            // 
            this.removeDestFileButton.Location = new System.Drawing.Point(602, 59);
            this.removeDestFileButton.Name = "removeDestFileButton";
            this.removeDestFileButton.Size = new System.Drawing.Size(33, 27);
            this.removeDestFileButton.TabIndex = 8;
            this.removeDestFileButton.Tag = "";
            this.removeDestFileButton.Text = "-";
            this.removeDestFileButton.UseVisualStyleBackColor = true;
            this.removeDestFileButton.Click += new System.EventHandler(this.removeDestFileButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.Location = new System.Drawing.Point(244, 125);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(126, 27);
            this.exportButton.TabIndex = 9;
            this.exportButton.Tag = "";
            this.exportButton.Text = "Экспорт >";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // exportAndOpenButton
            // 
            this.exportAndOpenButton.Location = new System.Drawing.Point(244, 187);
            this.exportAndOpenButton.Name = "exportAndOpenButton";
            this.exportAndOpenButton.Size = new System.Drawing.Size(126, 27);
            this.exportAndOpenButton.TabIndex = 10;
            this.exportAndOpenButton.Tag = "";
            this.exportAndOpenButton.Text = "Экспорт и открыть >>";
            this.exportAndOpenButton.UseVisualStyleBackColor = true;
            this.exportAndOpenButton.Click += new System.EventHandler(this.exportAndOpenButton_Click);
            // 
            // sourceFileInfoBox
            // 
            this.sourceFileInfoBox.BackColor = System.Drawing.SystemColors.Control;
            this.sourceFileInfoBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.sourceFileInfoBox.Enabled = false;
            this.sourceFileInfoBox.Location = new System.Drawing.Point(12, 320);
            this.sourceFileInfoBox.Multiline = true;
            this.sourceFileInfoBox.Name = "sourceFileInfoBox";
            this.sourceFileInfoBox.Size = new System.Drawing.Size(211, 36);
            this.sourceFileInfoBox.TabIndex = 11;
            this.sourceFileInfoBox.Text = "Файл не выбран.";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 304);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Информация о файле:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(382, 304);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 13);
            this.label2.TabIndex = 14;
            this.label2.Text = "Информация о файле:";
            // 
            // destFileInfoBox
            // 
            this.destFileInfoBox.BackColor = System.Drawing.SystemColors.Control;
            this.destFileInfoBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.destFileInfoBox.Enabled = false;
            this.destFileInfoBox.Location = new System.Drawing.Point(385, 320);
            this.destFileInfoBox.Multiline = true;
            this.destFileInfoBox.Name = "destFileInfoBox";
            this.destFileInfoBox.Size = new System.Drawing.Size(211, 36);
            this.destFileInfoBox.TabIndex = 13;
            this.destFileInfoBox.Text = "Файл не выбран.";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(647, 371);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.destFileInfoBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sourceFileInfoBox);
            this.Controls.Add(this.exportAndOpenButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.removeDestFileButton);
            this.Controls.Add(this.removeSourceFileButton);
            this.Controls.Add(this.addDestFileButton);
            this.Controls.Add(this.addSourceFileButton);
            this.Controls.Add(this.destFilesLabel);
            this.Controls.Add(this.sourceFilesLabel);
            this.Controls.Add(this.destBox);
            this.Controls.Add(this.sourceBox);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертер выписок счетов (с) AlvaSoft";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ListBox sourceBox;
        private System.Windows.Forms.ListBox destBox;
        private System.Windows.Forms.Label sourceFilesLabel;
        private System.Windows.Forms.Label destFilesLabel;
        private System.Windows.Forms.Button addSourceFileButton;
        private System.Windows.Forms.Button addDestFileButton;
        private System.Windows.Forms.Button removeSourceFileButton;
        private System.Windows.Forms.Button removeDestFileButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.Button exportAndOpenButton;
        private System.Windows.Forms.TextBox sourceFileInfoBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox destFileInfoBox;
    }
}

