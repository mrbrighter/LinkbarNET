using System;
using System.Drawing;
using System.Windows.Forms;

namespace LinkbarNET
{
    public partial class ConfigForm : Form
    {
        // Eigenschaft zum Übernehmen des ausgewählten Shortcut-Pfades
        public string ShortcutPath { get; private set; }

        private TextBox txtShortcutPath;
        private Button btnBrowse;
        private Button btnOK;
        private Button btnCancel;

        public ConfigForm()
        {
            InitializeComponentt();
        }

        private void InitializeComponentt()
        {
            this.Text = "Einstellungen";
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Width = 420;
            this.Height = 150;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ShowInTaskbar = false;

            // Label
            Label lbl = new Label();
            lbl.Text = "Pfad für .lnk Dateien:";
            lbl.AutoSize = true;
            lbl.Location = new Point(10, 10);
            this.Controls.Add(lbl);

            // TextBox für den Pfad
            txtShortcutPath = new TextBox();
            txtShortcutPath.Location = new Point(10, 30);
            txtShortcutPath.Width = 300;
            this.Controls.Add(txtShortcutPath);

            // "Durchsuchen..." Button
            btnBrowse = new Button();
            btnBrowse.Text = "Durchsuchen...";
            btnBrowse.Location = new Point(320, 28);
            btnBrowse.Click += BtnBrowse_Click;
            this.Controls.Add(btnBrowse);

            // OK-Button
            btnOK = new Button();
            btnOK.Text = "OK";
            btnOK.Location = new Point(220, 70);
            btnOK.DialogResult = DialogResult.OK;
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            // Abbrechen-Button
            btnCancel = new Button();
            btnCancel.Text = "Abbrechen";
            btnCancel.Location = new Point(300, 70);
            btnCancel.DialogResult = DialogResult.Cancel;
            this.Controls.Add(btnCancel);

            // Festlegen der Standard-Buttons
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog fbd = new FolderBrowserDialog())
            {
                fbd.Description = "Bitte wählen Sie den Ordner, in dem sich die .lnk-Dateien befinden.";
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtShortcutPath.Text = fbd.SelectedPath;
                }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            // Überprüfen, ob ein Pfad angegeben wurde
            if (string.IsNullOrWhiteSpace(txtShortcutPath.Text))
            {
                MessageBox.Show("Bitte einen gültigen Pfad auswählen.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None; // Verhindert das Schließen des Formulars
            }
            else
            {
                ShortcutPath = txtShortcutPath.Text;
            }
        }
    }
}
