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
            this.sourceFilesLabel = new System.Windows.Forms.Label();
            this.destFilesLabel = new System.Windows.Forms.Label();
            this.openSourceFileButton = new System.Windows.Forms.Button();
            this.openDestFileButton = new System.Windows.Forms.Button();
            this.exportButton = new System.Windows.Forms.Button();
            this.exportAndOpenButton = new System.Windows.Forms.Button();
            this.sourceFileInfoBox = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.destFileInfoBox = new System.Windows.Forms.TextBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.новыйОтчетToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оПрограммеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.оПрограммеToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.sourceFileNameBox = new System.Windows.Forms.TextBox();
            this.destFileNameBox = new System.Windows.Forms.TextBox();
            this.selectedShieldBox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.exitButton = new System.Windows.Forms.Button();
            this.openButton = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // sourceFilesLabel
            // 
            this.sourceFilesLabel.AutoSize = true;
            this.sourceFilesLabel.Location = new System.Drawing.Point(9, 44);
            this.sourceFilesLabel.Name = "sourceFilesLabel";
            this.sourceFilesLabel.Size = new System.Drawing.Size(87, 13);
            this.sourceFilesLabel.TabIndex = 3;
            this.sourceFilesLabel.Text = "Исходный файл";
            // 
            // destFilesLabel
            // 
            this.destFilesLabel.AutoSize = true;
            this.destFilesLabel.Location = new System.Drawing.Point(9, 143);
            this.destFilesLabel.Name = "destFilesLabel";
            this.destFilesLabel.Size = new System.Drawing.Size(98, 13);
            this.destFilesLabel.TabIndex = 4;
            this.destFilesLabel.Text = "Файл назначения";
            // 
            // openSourceFileButton
            // 
            this.openSourceFileButton.Location = new System.Drawing.Point(325, 60);
            this.openSourceFileButton.Name = "openSourceFileButton";
            this.openSourceFileButton.Size = new System.Drawing.Size(76, 20);
            this.openSourceFileButton.TabIndex = 5;
            this.openSourceFileButton.Tag = "";
            this.openSourceFileButton.Text = "Выбрать";
            this.openSourceFileButton.UseVisualStyleBackColor = true;
            this.openSourceFileButton.Click += new System.EventHandler(this.openSourceFileButton_Click);
            // 
            // openDestFileButton
            // 
            this.openDestFileButton.Location = new System.Drawing.Point(325, 159);
            this.openDestFileButton.Name = "openDestFileButton";
            this.openDestFileButton.Size = new System.Drawing.Size(76, 20);
            this.openDestFileButton.TabIndex = 6;
            this.openDestFileButton.Tag = "";
            this.openDestFileButton.Text = "Выбрать";
            this.openDestFileButton.UseVisualStyleBackColor = true;
            this.openDestFileButton.Click += new System.EventHandler(this.openDestFileButton_Click);
            // 
            // exportButton
            // 
            this.exportButton.Location = new System.Drawing.Point(12, 301);
            this.exportButton.Name = "exportButton";
            this.exportButton.Size = new System.Drawing.Size(148, 27);
            this.exportButton.TabIndex = 9;
            this.exportButton.Tag = "";
            this.exportButton.Text = "Экспорт >";
            this.exportButton.UseVisualStyleBackColor = true;
            this.exportButton.Click += new System.EventHandler(this.exportButton_Click);
            // 
            // exportAndOpenButton
            // 
            this.exportAndOpenButton.Location = new System.Drawing.Point(253, 301);
            this.exportAndOpenButton.Name = "exportAndOpenButton";
            this.exportAndOpenButton.Size = new System.Drawing.Size(148, 27);
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
            this.sourceFileInfoBox.Location = new System.Drawing.Point(12, 98);
            this.sourceFileInfoBox.Multiline = true;
            this.sourceFileInfoBox.Name = "sourceFileInfoBox";
            this.sourceFileInfoBox.Size = new System.Drawing.Size(389, 36);
            this.sourceFileInfoBox.TabIndex = 11;
            this.sourceFileInfoBox.Text = "Файл не выбран.";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 82);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "Информация о файле:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 182);
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
            this.destFileInfoBox.Location = new System.Drawing.Point(12, 198);
            this.destFileInfoBox.Multiline = true;
            this.destFileInfoBox.Name = "destFileInfoBox";
            this.destFileInfoBox.Size = new System.Drawing.Size(389, 36);
            this.destFileInfoBox.TabIndex = 13;
            this.destFileInfoBox.Text = "Файл не выбран.";
            this.destFileInfoBox.TextChanged += new System.EventHandler(this.destFileInfoBox_TextChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.оПрограммеToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(414, 24);
            this.menuStrip1.TabIndex = 15;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.новыйОтчетToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.файлToolStripMenuItem.Text = "Файл";
            // 
            // новыйОтчетToolStripMenuItem
            // 
            this.новыйОтчетToolStripMenuItem.Name = "новыйОтчетToolStripMenuItem";
            this.новыйОтчетToolStripMenuItem.Size = new System.Drawing.Size(140, 22);
            this.новыйОтчетToolStripMenuItem.Text = "Новый отчет";
            this.новыйОтчетToolStripMenuItem.Click += new System.EventHandler(this.новыйОтчетToolStripMenuItem_Click_1);
            // 
            // оПрограммеToolStripMenuItem
            // 
            this.оПрограммеToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.оПрограммеToolStripMenuItem1});
            this.оПрограммеToolStripMenuItem.Name = "оПрограммеToolStripMenuItem";
            this.оПрограммеToolStripMenuItem.Size = new System.Drawing.Size(82, 20);
            this.оПрограммеToolStripMenuItem.Text = "Информация";
            // 
            // оПрограммеToolStripMenuItem1
            // 
            this.оПрограммеToolStripMenuItem1.Name = "оПрограммеToolStripMenuItem1";
            this.оПрограммеToolStripMenuItem1.Size = new System.Drawing.Size(138, 22);
            this.оПрограммеToolStripMenuItem1.Text = "О программе";
            this.оПрограммеToolStripMenuItem1.Click += new System.EventHandler(this.оПрограммеToolStripMenuItem1_Click);
            // 
            // sourceFileNameBox
            // 
            this.sourceFileNameBox.Enabled = false;
            this.sourceFileNameBox.Location = new System.Drawing.Point(12, 60);
            this.sourceFileNameBox.Name = "sourceFileNameBox";
            this.sourceFileNameBox.Size = new System.Drawing.Size(307, 20);
            this.sourceFileNameBox.TabIndex = 16;
            // 
            // destFileNameBox
            // 
            this.destFileNameBox.Enabled = false;
            this.destFileNameBox.Location = new System.Drawing.Point(12, 159);
            this.destFileNameBox.Name = "destFileNameBox";
            this.destFileNameBox.Size = new System.Drawing.Size(307, 20);
            this.destFileNameBox.TabIndex = 17;
            // 
            // selectedShieldBox
            // 
            this.selectedShieldBox.Enabled = false;
            this.selectedShieldBox.FormattingEnabled = true;
            this.selectedShieldBox.Location = new System.Drawing.Point(12, 264);
            this.selectedShieldBox.Name = "selectedShieldBox";
            this.selectedShieldBox.Size = new System.Drawing.Size(389, 21);
            this.selectedShieldBox.TabIndex = 18;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 248);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(392, 13);
            this.label3.TabIndex = 19;
            this.label3.Text = "Лист назначения (выберите из списка или введите название нового листа)";
            // 
            // exitButton
            // 
            this.exitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.exitButton.Location = new System.Drawing.Point(254, 362);
            this.exitButton.Name = "exitButton";
            this.exitButton.Size = new System.Drawing.Size(147, 39);
            this.exitButton.TabIndex = 20;
            this.exitButton.Text = "Выход";
            this.exitButton.UseVisualStyleBackColor = true;
            this.exitButton.Click += new System.EventHandler(this.button1_Click);
            // 
            // openButton
            // 
            this.openButton.Enabled = false;
            this.openButton.Location = new System.Drawing.Point(325, 185);
            this.openButton.Name = "openButton";
            this.openButton.Size = new System.Drawing.Size(76, 21);
            this.openButton.TabIndex = 21;
            this.openButton.Tag = "";
            this.openButton.Text = "Открыть";
            this.openButton.UseVisualStyleBackColor = true;
            this.openButton.Click += new System.EventHandler(this.openButton_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.exitButton;
            this.ClientSize = new System.Drawing.Size(414, 413);
            this.Controls.Add(this.openButton);
            this.Controls.Add(this.exitButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.selectedShieldBox);
            this.Controls.Add(this.destFileNameBox);
            this.Controls.Add(this.sourceFileNameBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.destFileInfoBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.sourceFileInfoBox);
            this.Controls.Add(this.exportAndOpenButton);
            this.Controls.Add(this.exportButton);
            this.Controls.Add(this.openDestFileButton);
            this.Controls.Add(this.openSourceFileButton);
            this.Controls.Add(this.destFilesLabel);
            this.Controls.Add(this.sourceFilesLabel);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Конвертер банковских выписок";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label sourceFilesLabel;
        private System.Windows.Forms.Label destFilesLabel;
        private System.Windows.Forms.Button openSourceFileButton;
        private System.Windows.Forms.Button openDestFileButton;
        private System.Windows.Forms.Button exportButton;
        private System.Windows.Forms.Button exportAndOpenButton;
        private System.Windows.Forms.TextBox sourceFileInfoBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox destFileInfoBox;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem новыйОтчетToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem оПрограммеToolStripMenuItem1;
        private System.Windows.Forms.TextBox sourceFileNameBox;
        private System.Windows.Forms.TextBox destFileNameBox;
        private System.Windows.Forms.ComboBox selectedShieldBox;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button exitButton;
        private System.Windows.Forms.Button openButton;
    }
}

