using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Windows.Forms;

namespace ASConverter {
    public partial class MainForm : Form {
        public MainForm() {
            InitializeComponent();
        }

        private void addSourceFileButton_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;    
            openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";

            if (openFileDialog.ShowDialog() != DialogResult.OK) {
                return;
            }

            var sourceFiles = openFileDialog.FileNames;

            if (sourceFiles == null || sourceFiles.Length == 0) {
                return;
            }

            var configuration = ASConverter.Default;
            try {                
                var existedSourceFiles = configuration.SourceFiles;
                if (existedSourceFiles == null) {
                    existedSourceFiles = new StringCollection();
                }

                foreach (var newSourceFile in sourceFiles) {
                    var isFind = false;
                    foreach (var existedSourceFile in existedSourceFiles) {
                        if (existedSourceFile.Equals(newSourceFile)) {
                            isFind = true;
                            break;
                        }
                    }

                    if (!isFind) {
                        existedSourceFiles.Add(newSourceFile);
                    }
                }

                if (configuration.SourceFiles == null) {
                    configuration.SourceFiles = existedSourceFiles;
                }
                configuration.Save();
            } catch (Exception ex) {
                MessageBox.Show("При работе с конфигурацией входных файлов возникла проблема: " + ex.Message);
            } finally {
                UpdateSourceFiles();
            }                        
        }

        private void UpdateSourceFiles() {
            var configuration = ASConverter.Default;
            var sourceFiles = configuration.SourceFiles;
            if (sourceFiles == null || sourceFiles.Count == 0) {
                return;
            }

            var removed = new List<string>();
            foreach (var sourceFile in sourceFiles) {
                if (!File.Exists(sourceFile)) {
                    removed.Add(sourceFile);
                }
            }
            foreach (var remFile in removed) {
                sourceFiles.Remove(remFile);
            }

            sourceBox.Items.Clear();
            foreach (var sourceFile in sourceFiles) {
                var fileName = sourceFile.Substring(sourceFile.LastIndexOf('\\') + 1);
                sourceBox.Items.Add(fileName);
            }
        }

        private void UpdateDestFiles() {
            var configuration = ASConverter.Default;
            var destFiles = configuration.DestFiles;
            if (destFiles == null || destFiles.Count == 0) {
                return;
            }

            var removed = new List<string>();
            foreach (var destFile in destFiles)
            {
                if (!File.Exists(destFile))
                {
                    removed.Add(destFile);
                }
            }
            foreach (var remFile in removed)
            {
                destFiles.Remove(remFile);
            }

            destBox.Items.Clear();
            foreach (var destFile in destFiles) {
                var fileName = destFile.Substring(destFile.LastIndexOf('\\') + 1);
                destBox.Items.Add(fileName);
            }
        }

        private void MainForm_Load(object sender, EventArgs e) {
            UpdateSourceFiles();
            UpdateDestFiles();
        }

        private void removeSourceFileButton_Click(object sender, EventArgs e) {
            if (sourceBox.SelectedIndex == -1) {
                MessageBox.Show("Выберите файл для удаления из списка.");
                return;
            }

            try {
                var configuration = ASConverter.Default;
                var sourceFiles = configuration.SourceFiles;
                sourceFiles.RemoveAt(sourceBox.SelectedIndex);
                configuration.Save();
            } catch (Exception ex) {
                MessageBox.Show("Ошибка при работе с конфигурацией: " + ex.Message);
            }
            sourceBox.Items.RemoveAt(sourceBox.SelectedIndex);            
        }

        private void addDestFileButton_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";

            if (openFileDialog.ShowDialog() != DialogResult.OK) {
                return;
            }

            var destFiles = openFileDialog.FileNames;

            if (destFiles == null || destFiles.Length == 0) {
                return;
            }

            var configuration = ASConverter.Default;
            try {
                var existedDestFiles = configuration.DestFiles;
                if (existedDestFiles == null) {
                    existedDestFiles = new StringCollection();
                }

                foreach (var newDestFile in destFiles) {
                    var isFind = false;
                    foreach (var existedDestFile in existedDestFiles) {
                        if (existedDestFile.Equals(newDestFile)) {
                            isFind = true;
                            break;
                        }
                    }

                    if (!isFind) {
                        existedDestFiles.Add(newDestFile);
                    }
                }

                if (configuration.DestFiles == null) {
                    configuration.DestFiles = existedDestFiles;
                }
                configuration.Save();
            } catch (Exception ex) {
                MessageBox.Show("При работе с конфигурацией входных файлов возникла проблема: " + ex.Message);
            } finally {
                UpdateDestFiles();
            }
        }

        private void removeDestFileButton_Click(object sender, EventArgs e) {
            if (destBox.SelectedIndex == -1) {
                MessageBox.Show("Выберите файл для удаления из списка.");
                return;
            }

            try {
                var configuration = ASConverter.Default;
                var destFiles = configuration.DestFiles;
                destFiles.RemoveAt(destBox.SelectedIndex);
                configuration.Save();
            } catch (Exception ex) {
                MessageBox.Show("Ошибка при работе с конфигурацией: " + ex.Message);
            }
            destBox.Items.RemoveAt(destBox.SelectedIndex);
        }

        private void sourceBox_SelectedIndexChanged(object sender, EventArgs e) {
            if (sourceBox.SelectedIndex == -1) {
                sourceFileInfoBox.Text = "Файл не выбран.";
                return;
            }

            var configuration = ASConverter.Default;
            var sourceFiles = configuration.SourceFiles;
            sourceFileInfoBox.Text = sourceFiles[sourceBox.SelectedIndex];
        }

        private void destBox_SelectedIndexChanged(object sender, EventArgs e) {
            if (destBox.SelectedIndex == -1) {
                destFileInfoBox.Text = "Файл не выбран.";
                return;
            }

            var configuration = ASConverter.Default;
            var destFiles = configuration.DestFiles;
            destFileInfoBox.Text = destFiles[destBox.SelectedIndex];
        }

        private void новыйОтчетToolStripMenuItem_Click(object sender, EventArgs e) {

        }

        private void exportButton_Click(object sender, EventArgs e) {
            exportButton.Enabled = false;
            exportAndOpenButton.Enabled = false;
            try {
                if (sourceBox.SelectedIndex == -1 || destBox.SelectedIndex == -1) {
                    MessageBox.Show("Выберите файл экспорта и файл импорта.");
                    return;
                }

                var configuration = ASConverter.Default;

                var sourceFile = configuration.SourceFiles[sourceBox.SelectedIndex];
                var destFile = configuration.DestFiles[destBox.SelectedIndex];

                if (!File.Exists(sourceFile)) {
                    MessageBox.Show("Отсутствует файл для импорта.");
                    return;
                }

                if (!File.Exists(destFile)) {
                    MessageBox.Show("Отсутствует файл для экспорта.");
                    return;
                }

                try {                    
                    var exporter = new ASExporter();
                    string message;
                    var count = exporter.Export(sourceFile, destFile, out message);
                    MessageBox.Show("Успешно добавлено операций: " + count + "\n" + message);

                } catch (Exception ex) {
                    MessageBox.Show("При экспорте данных произошла ошибка: " + ex.Message);
                    return;
                }
            } catch (Exception ex) {
                MessageBox.Show("Непредвиенная ошибка: " + ex.Message);
                return;
            } finally {
                exportButton.Enabled = true;
                exportAndOpenButton.Enabled = true;
            }
        }

        private void exportAndOpenButton_Click(object sender, EventArgs e) {
            exportButton.Enabled = false;
            exportAndOpenButton.Enabled = false;
            try {
                if (sourceBox.SelectedIndex == -1 || destBox.SelectedIndex == -1) {
                    MessageBox.Show("Выберите файл экспорта и файл импорта.");
                    return;
                }

                var configuration = ASConverter.Default;

                var sourceFile = configuration.SourceFiles[sourceBox.SelectedIndex];
                var destFile = configuration.DestFiles[destBox.SelectedIndex];

                if (!File.Exists(sourceFile)) {
                    MessageBox.Show("Отсутствует файл для импорта.");
                    return;
                }

                if (!File.Exists(destFile)) {
                    MessageBox.Show("Отсутствует файл для экспорта.");
                    return;
                }

                try {
                    var exporter = new ASExporter();
                    string message;
                    var count = exporter.Export(sourceFile, destFile, out message);
                    MessageBox.Show("Успешно добавлено операций: " + count + "\n" + message);
                } catch (Exception ex) {
                    MessageBox.Show("При экспорте данных произошла ошибка: " + ex.Message);
                    return;
                }

                System.Diagnostics.Process.Start(destFile);
            } catch (Exception ex) {
                MessageBox.Show("Непредвиенная ошибка: " + ex.Message);
                return;
            } finally {
                exportButton.Enabled = true;
                exportAndOpenButton.Enabled = true;
            }
            
        }

        private void новыйОтчетToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файл Excel (*.xlsx)|*.xlsx";
            if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                var fileName = saveFileDialog.FileName;
                if (File.Exists(fileName)) {
                    MessageBox.Show("Такой файл уже существует.");
                    return;
                }
                try
                {
                    new ASExporter().Generate(fileName);
                    var configuration = ASConverter.Default;
                    if (configuration.DestFiles == null) {
                        configuration.DestFiles = new StringCollection();
                    }
                    if (!configuration.DestFiles.Contains(fileName)) {
                        configuration.DestFiles.Add(fileName);
                        UpdateDestFiles();
                    }
                }
                catch (Exception ex) {
                    MessageBox.Show("При создании отчета произошла ошибка: " + ex.Message);
                }
            }
        }

        private void оПрограммеToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Версия 1.0");
        }
    }
}
