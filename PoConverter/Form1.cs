using System;
using System.IO;
using System.Windows.Forms;

namespace PoConverter
{
    public partial class Form1 : Form
    {
        private PoSheet _currentFile = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnAbrirArquivo_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void btnLerArquivo_Click(object sender, EventArgs e)
        {
            var filePath = textBox1.Text;

            if (!File.Exists(filePath))
            {
                MessageHelper.ShowWarning("File does not exist");
                return;
            }
            if (!Path.GetExtension(filePath).IsInList(".po", ".xlsx"))
            {
                MessageHelper.ShowWarning("File extension not recognized");
                return;
            }

            try
            {
                var file = PoSheetReader.Read(filePath);
                _currentFile = file;
                dataGridView1.DataSource = file.Records;
                MessageHelper.ShowInfo("Leitura concluída");
            }
            catch (Exception ex)
            {
                MessageHelper.ShowError("Erro ao ler o arquivo: " + ex.Message);
            }
        }
    }
}
