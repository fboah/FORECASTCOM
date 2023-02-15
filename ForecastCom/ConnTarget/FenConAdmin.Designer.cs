namespace ForecastCom.ConnTarget
{
    partial class FenConAdmin
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FenConAdmin));
            this.CmbAuthentification = new DevExpress.XtraEditors.LookUpEdit();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.CmbBaseDonnees = new System.Windows.Forms.ComboBox();
            this.CmbServeur = new System.Windows.Forms.ComboBox();
            this.labelControl6 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl5 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl4 = new DevExpress.XtraEditors.LabelControl();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.labelControl3 = new DevExpress.XtraEditors.LabelControl();
            this.labelControl2 = new DevExpress.XtraEditors.LabelControl();
            this.txtMotPasse = new DevExpress.XtraEditors.TextEdit();
            this.txtUser = new DevExpress.XtraEditors.TextEdit();
            this.txtVille = new DevExpress.XtraEditors.TextEdit();
            this.sBtnQuitter = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnValider = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnTesterCon = new DevExpress.XtraEditors.SimpleButton();
            ((System.ComponentModel.ISupportInitialize)(this.CmbAuthentification.Properties)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtMotPasse.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtUser.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtVille.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // CmbAuthentification
            // 
            this.CmbAuthentification.EditValue = ((short)(0));
            this.CmbAuthentification.Location = new System.Drawing.Point(143, 36);
            this.CmbAuthentification.Name = "CmbAuthentification";
            this.CmbAuthentification.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.CmbAuthentification.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Id", "Id", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("Libelle", "Libelle")});
            this.CmbAuthentification.Size = new System.Drawing.Size(327, 20);
            this.CmbAuthentification.TabIndex = 31;
            this.CmbAuthentification.EditValueChanged += new System.EventHandler(this.CmbAuthentification_EditValueChanged);
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(42, 14);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(18, 13);
            this.labelControl1.TabIndex = 23;
            this.labelControl1.Text = "Ville";
            // 
            // CmbBaseDonnees
            // 
            this.CmbBaseDonnees.FormattingEnabled = true;
            this.CmbBaseDonnees.Location = new System.Drawing.Point(144, 155);
            this.CmbBaseDonnees.Name = "CmbBaseDonnees";
            this.CmbBaseDonnees.Size = new System.Drawing.Size(327, 21);
            this.CmbBaseDonnees.TabIndex = 30;
            this.CmbBaseDonnees.Click += new System.EventHandler(this.CmbBaseDonnees_Click);
            // 
            // CmbServeur
            // 
            this.CmbServeur.FormattingEnabled = true;
            this.CmbServeur.Location = new System.Drawing.Point(143, 123);
            this.CmbServeur.Name = "CmbServeur";
            this.CmbServeur.Size = new System.Drawing.Size(327, 21);
            this.CmbServeur.TabIndex = 29;
            this.CmbServeur.Click += new System.EventHandler(this.CmbServeur_Click);
            // 
            // labelControl6
            // 
            this.labelControl6.Location = new System.Drawing.Point(42, 163);
            this.labelControl6.Name = "labelControl6";
            this.labelControl6.Size = new System.Drawing.Size(82, 13);
            this.labelControl6.TabIndex = 28;
            this.labelControl6.Text = "Base de données";
            // 
            // labelControl5
            // 
            this.labelControl5.Location = new System.Drawing.Point(42, 131);
            this.labelControl5.Name = "labelControl5";
            this.labelControl5.Size = new System.Drawing.Size(38, 13);
            this.labelControl5.TabIndex = 27;
            this.labelControl5.Text = "Serveur";
            // 
            // labelControl4
            // 
            this.labelControl4.Location = new System.Drawing.Point(42, 102);
            this.labelControl4.Name = "labelControl4";
            this.labelControl4.Size = new System.Drawing.Size(64, 13);
            this.labelControl4.TabIndex = 26;
            this.labelControl4.Text = "Mot de Passe";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.panelControl1, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(1, 3);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 1;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(518, 237);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.CmbAuthentification);
            this.panelControl1.Controls.Add(this.labelControl1);
            this.panelControl1.Controls.Add(this.CmbBaseDonnees);
            this.panelControl1.Controls.Add(this.CmbServeur);
            this.panelControl1.Controls.Add(this.labelControl6);
            this.panelControl1.Controls.Add(this.labelControl5);
            this.panelControl1.Controls.Add(this.labelControl4);
            this.panelControl1.Controls.Add(this.labelControl3);
            this.panelControl1.Controls.Add(this.labelControl2);
            this.panelControl1.Controls.Add(this.txtMotPasse);
            this.panelControl1.Controls.Add(this.txtUser);
            this.panelControl1.Controls.Add(this.txtVille);
            this.panelControl1.Controls.Add(this.sBtnQuitter);
            this.panelControl1.Controls.Add(this.sBtnValider);
            this.panelControl1.Controls.Add(this.sBtnTesterCon);
            this.panelControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelControl1.Location = new System.Drawing.Point(3, 3);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(512, 231);
            this.panelControl1.TabIndex = 0;
            // 
            // labelControl3
            // 
            this.labelControl3.Location = new System.Drawing.Point(42, 72);
            this.labelControl3.Name = "labelControl3";
            this.labelControl3.Size = new System.Drawing.Size(48, 13);
            this.labelControl3.TabIndex = 25;
            this.labelControl3.Text = "Utilisateur";
            // 
            // labelControl2
            // 
            this.labelControl2.Location = new System.Drawing.Point(42, 43);
            this.labelControl2.Name = "labelControl2";
            this.labelControl2.Size = new System.Drawing.Size(76, 13);
            this.labelControl2.TabIndex = 24;
            this.labelControl2.Text = "Authentification";
            // 
            // txtMotPasse
            // 
            this.txtMotPasse.Enabled = false;
            this.txtMotPasse.Location = new System.Drawing.Point(143, 95);
            this.txtMotPasse.Name = "txtMotPasse";
            this.txtMotPasse.Properties.PasswordChar = '*';
            this.txtMotPasse.Properties.UseSystemPasswordChar = true;
            this.txtMotPasse.Size = new System.Drawing.Size(327, 20);
            this.txtMotPasse.TabIndex = 22;
            // 
            // txtUser
            // 
            this.txtUser.Enabled = false;
            this.txtUser.Location = new System.Drawing.Point(143, 69);
            this.txtUser.Name = "txtUser";
            this.txtUser.Size = new System.Drawing.Size(327, 20);
            this.txtUser.TabIndex = 21;
            // 
            // txtVille
            // 
            this.txtVille.Location = new System.Drawing.Point(143, 7);
            this.txtVille.Name = "txtVille";
            this.txtVille.Size = new System.Drawing.Size(327, 20);
            this.txtVille.TabIndex = 20;
            // 
            // sBtnQuitter
            // 
            this.sBtnQuitter.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.sBtnQuitter.Image = ((System.Drawing.Image)(resources.GetObject("sBtnQuitter.Image")));
            this.sBtnQuitter.Location = new System.Drawing.Point(370, 197);
            this.sBtnQuitter.Name = "sBtnQuitter";
            this.sBtnQuitter.Size = new System.Drawing.Size(100, 26);
            this.sBtnQuitter.TabIndex = 19;
            this.sBtnQuitter.Text = "Quitter";
            this.sBtnQuitter.Click += new System.EventHandler(this.sBtnQuitter_Click);
            // 
            // sBtnValider
            // 
            this.sBtnValider.Image = ((System.Drawing.Image)(resources.GetObject("sBtnValider.Image")));
            this.sBtnValider.Location = new System.Drawing.Point(264, 197);
            this.sBtnValider.Name = "sBtnValider";
            this.sBtnValider.Size = new System.Drawing.Size(100, 26);
            this.sBtnValider.TabIndex = 18;
            this.sBtnValider.Text = "Valider";
            this.sBtnValider.Click += new System.EventHandler(this.sBtnValider_Click);
            // 
            // sBtnTesterCon
            // 
            this.sBtnTesterCon.Image = ((System.Drawing.Image)(resources.GetObject("sBtnTesterCon.Image")));
            this.sBtnTesterCon.Location = new System.Drawing.Point(144, 197);
            this.sBtnTesterCon.Name = "sBtnTesterCon";
            this.sBtnTesterCon.Size = new System.Drawing.Size(114, 26);
            this.sBtnTesterCon.TabIndex = 17;
            this.sBtnTesterCon.Text = "Tester Connexion";
            this.sBtnTesterCon.Click += new System.EventHandler(this.sBtnTesterCon_Click);
            // 
            // FenConAdmin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 242);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FenConAdmin";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FenConnAdmin";
            this.Load += new System.EventHandler(this.FenConTarget_Load);
            ((System.ComponentModel.ISupportInitialize)(this.CmbAuthentification.Properties)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.panelControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.txtMotPasse.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtUser.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.txtVille.Properties)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraEditors.LookUpEdit CmbAuthentification;
        private DevExpress.XtraEditors.LabelControl labelControl1;
        private System.Windows.Forms.ComboBox CmbBaseDonnees;
        private System.Windows.Forms.ComboBox CmbServeur;
        private DevExpress.XtraEditors.LabelControl labelControl6;
        private DevExpress.XtraEditors.LabelControl labelControl5;
        private DevExpress.XtraEditors.LabelControl labelControl4;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraEditors.LabelControl labelControl3;
        private DevExpress.XtraEditors.LabelControl labelControl2;
        private DevExpress.XtraEditors.TextEdit txtMotPasse;
        private DevExpress.XtraEditors.TextEdit txtUser;
        private DevExpress.XtraEditors.TextEdit txtVille;
        private DevExpress.XtraEditors.SimpleButton sBtnQuitter;
        private DevExpress.XtraEditors.SimpleButton sBtnValider;
        private DevExpress.XtraEditors.SimpleButton sBtnTesterCon;
    }
}