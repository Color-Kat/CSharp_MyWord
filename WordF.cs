using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyWord
{
    public partial class WordF : Form
    {
        WordService WordService = new WordService();

        public WordF()
        {
            InitializeComponent();
        }

        private void createWordDocumentButton_Click(object sender, EventArgs e)
        {
            WordService.CreateWordDocument(textBox.Text);
        }
    }
}
