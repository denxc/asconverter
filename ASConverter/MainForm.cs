using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ASConverter {
    public partial class MainForm : Form {
        private string sourceFilePath = string.Empty;
        private string destFilePath = string.Empty;

        public MainForm() {
            InitializeComponent();            
        }

        private void openSourceFileButton_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Текстовые файлы (*.txt)|*.txt";

            if (openFileDialog.ShowDialog() != DialogResult.OK) {
                return;
            }

            var sourceFile = openFileDialog.FileName;

            SelectSourceFile(sourceFile);
        }

        private void SelectSourceFile(string sourceFile) {
            if (sourceFile == null || sourceFile.Length == 0) {
                return;
            }

            var sourceFileName = sourceFile.Substring(sourceFile.LastIndexOf('\\') + 1);
            sourceFileNameBox.Text = sourceFileName;
            sourceFileInfoBox.Text = sourceFile;

            sourceFilePath = sourceFile;
            sourceFileNameBox.BackColor = Color.White;            
        }

        private void MainForm_Load(object sender, EventArgs e) {
        }

        private void openDestFileButton_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Файлы Excel 1997-2003 (*.xls)|*.xls";

            if (openFileDialog.ShowDialog() != DialogResult.OK) {
                return;
            }

            var destFile = openFileDialog.FileName;

            SelectDestFile(destFile);
        }

        private void SelectDestFile(string destFile) {
            if (destFile == null || destFile.Length == 0) {
                return;
            }

            var destFileName = destFile.Substring(destFile.LastIndexOf('\\') + 1);
            destFileNameBox.Text = destFileName;
            destFileInfoBox.Text = destFile;

            destFilePath = destFile;
            destFileNameBox.BackColor = Color.White;

            TryLoadShields(destFilePath);
        }

        private void TryLoadShields(string aXlsFilePath) {
            try {
                if (string.IsNullOrEmpty(aXlsFilePath) || !File.Exists(aXlsFilePath)) {
                    return;
                }

                selectedShieldBox.Items.Clear();
                selectedShieldBox.Text = string.Empty;

                var shields = ASExporter.GetShields(destFilePath);
                if (shields != null && shields.Length > 0) {
                    selectedShieldBox.Items.AddRange(shields);
                }

                selectedShieldBox.Enabled = true;
            } catch (Exception ex) {
                MessageBox.Show("При загрузке листов из файла произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                selectedShieldBox.Enabled = false;
            }
        }

        private void exportButton_Click(object sender, EventArgs e) {
            exportButton.Enabled = false;
            exportAndOpenButton.Enabled = false;
            if (TryDoExport()) {
                SelectDestFile(destFilePath);
            }
            exportButton.Enabled = true;
            exportAndOpenButton.Enabled = true;
        }

        private void exportAndOpenButton_Click(object sender, EventArgs e) {
            exportButton.Enabled = false;
            exportAndOpenButton.Enabled = false;
            var result = TryDoExport();
            exportButton.Enabled = true;
            exportAndOpenButton.Enabled = true;
            if (File.Exists(destFilePath) && result) {
                try {
                    SelectDestFile(destFilePath);
                    System.Diagnostics.Process.Start(destFilePath);
                } catch (Exception ex) {
                    MessageBox.Show("При открытии файла произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }            
        }        

        private void новыйОтчетToolStripMenuItem_Click_1(object sender, EventArgs e) {
            var saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файл Excel 1997-2003 (*.xls)|*.xls";            
            if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                var fileName = saveFileDialog.FileName;                
                try {
                    ASExporter.Generate(fileName);
                    SelectDestFile(fileName);
                } catch (Exception ex) {
                    MessageBox.Show("При создании отчета произошла ошибка: " + ex.Message);
                }
            }
        }

        private void оПрограммеToolStripMenuItem1_Click(object sender, EventArgs e) {
            MessageBox.Show("Версия 2.3");
        }

        private void button1_Click(object sender, EventArgs e) {
            this.Close();
        }

        private bool TryDoExport() {
            if (string.IsNullOrEmpty(sourceFilePath)) {
                MessageBox.Show("Выберите файл с операциями для экспорта.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return false;
            }

            if (string.IsNullOrEmpty(destFilePath)) {
                MessageBox.Show("Выберите файл назначения для импорта.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return false;
            }

            if (!File.Exists(sourceFilePath)) {
                MessageBox.Show("Выбранный файл с операциями для экспорта не существует.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (!File.Exists(destFilePath)) {
                MessageBox.Show("Выбранный файл назначения для импорта не существует.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrEmpty(selectedShieldBox.Text)) {
                MessageBox.Show("Выберите существующий или введите имя нового листа для добавления операций.", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return false;
            }

            AccountSection startAmount;
            AccountSection endAmount;
            var orders = TryLoadOrders(sourceFilePath, out startAmount, out endAmount);
            if (orders == null || orders.Length == 0) {
                return false;
            }            

            if (TryCheckCompanyCorrection(destFilePath, orders[0]) == false) {
                var dialogResult = MessageBox.Show(
                    string.Format("Нельзя добавить операции компании {0} в файл, содержащий операции другой компании. \nУбедитель в правильности выбора файла назначения.", orders[0].OwnerName), 
                    "Оповещение", 
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;                
            }

            if (TryCheckShieldCorrection(destFilePath, orders[0], startAmount, endAmount) == false) {
                return false;
            }
            
            try {                    
                var message = string.Empty;
                var count = ASExporter.Export(destFilePath, selectedShieldBox.Text, orders, startAmount, endAmount, out message);                
                MessageBox.Show("Успешно добавлено операций: " + count + "\n" + message);
                return true;
            } catch (Exception ex) {
                MessageBox.Show("При экспорте данных произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private bool TryCheckShieldCorrection(string aXlsFile, OrderEntity orderEntity, AccountSection aStartAmount, AccountSection aEndAmount) {
            try {
                var destShield = selectedShieldBox.Text;
                var shields = ASExporter.GetShields(aXlsFile);
                if (shields != null) {
                    foreach (var shield in shields) {
                        if (destShield.Equals(shield)) {
                            return true;
                        }
                    }
                }

                ASExporter.CreateNewShield(aXlsFile, destShield, orderEntity, aStartAmount, aEndAmount);
                return true;
            } catch (Exception ex) {
                MessageBox.Show("Ошибка при создании листа: " + ex.Message);
                return false;
            }            
        }

        private OrderEntity[] TryLoadOrders(string aTextFile, out AccountSection aStartAmount, out AccountSection aEndAmount) {
            try {
                var orders = OrdersProvider.LoadOrders(aTextFile, out aStartAmount, out aEndAmount);
                return orders.ToArray();
            } catch (Exception ex) {
                MessageBox.Show("При загрузке операций из файла произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                aStartAmount = null;
                aEndAmount = null;
                return null;
            }
        }

        private bool TryCheckCompanyCorrection(string aXlsFile, OrderEntity aOrder) {
            if (string.IsNullOrEmpty(aXlsFile) || !File.Exists(aXlsFile)) {
                return false;
            }            
            try {
                var companyImm = ASExporter.GetCompanyInn(aXlsFile);

                if (string.IsNullOrEmpty(companyImm) || companyImm.Equals(aOrder.OwnerInn)) {
                    return true;
                }
            } catch (Exception ex) {
                MessageBox.Show("Ошика при проверке корректности компании: " + ex.Message);
            }            

            return false;
        }

        private void destFileInfoBox_TextChanged(object sender, EventArgs e) {
            openButton.Enabled = File.Exists(destFileInfoBox.Text);
        }

        private void openButton_Click(object sender, EventArgs e) {
            try {
                System.Diagnostics.Process.Start(destFilePath);
            } catch (Exception ex) {
                MessageBox.Show("При открытии файла произошла ошибка: " + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }        
    }
}
