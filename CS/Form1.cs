using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraRichEdit.Export;
using DevExpress.XtraRichEdit.Import;

namespace Eml {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();

            EmlDocumentExporter.Register(richEdit);
            EmlDocumentImporter.Register(richEdit);

            richEdit.LoadDocument("Hey, look at this!.eml");
        }

    }
}