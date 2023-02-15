namespace ForecastCom
{
    partial class Connexion
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Connexion));
            this.tabConnexion = new System.Windows.Forms.TableLayoutPanel();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.sBtnSupprCon = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnModifCon = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnAjoutCon = new DevExpress.XtraEditors.SimpleButton();
            this.gridConnexion = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colVille = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colAliasName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colBD = new DevExpress.XtraGrid.Columns.GridColumn();
            this.tabConnexion.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridConnexion)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabConnexion
            // 
            this.tabConnexion.ColumnCount = 1;
            this.tabConnexion.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tabConnexion.Controls.Add(this.panelControl1, 0, 2);
            this.tabConnexion.Controls.Add(this.gridConnexion, 0, 1);
            this.tabConnexion.Location = new System.Drawing.Point(1, 4);
            this.tabConnexion.Name = "tabConnexion";
            this.tabConnexion.RowCount = 3;
            this.tabConnexion.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 2.266289F));
            this.tabConnexion.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 97.73371F));
            this.tabConnexion.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tabConnexion.Size = new System.Drawing.Size(691, 339);
            this.tabConnexion.TabIndex = 0;
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.sBtnSupprCon);
            this.panelControl1.Controls.Add(this.sBtnModifCon);
            this.panelControl1.Controls.Add(this.sBtnAjoutCon);
            this.panelControl1.Location = new System.Drawing.Point(3, 305);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(685, 31);
            this.panelControl1.TabIndex = 0;
            // 
            // sBtnSupprCon
            // 
            this.sBtnSupprCon.Image = ((System.Drawing.Image)(resources.GetObject("sBtnSupprCon.Image")));
            this.sBtnSupprCon.Location = new System.Drawing.Point(553, 3);
            this.sBtnSupprCon.Name = "sBtnSupprCon";
            this.sBtnSupprCon.Size = new System.Drawing.Size(132, 26);
            this.sBtnSupprCon.TabIndex = 2;
            this.sBtnSupprCon.Text = "Supprimer Connexion";
            this.sBtnSupprCon.Click += new System.EventHandler(this.sBtnSupprCon_Click);
            // 
            // sBtnModifCon
            // 
            this.sBtnModifCon.Image = ((System.Drawing.Image)(resources.GetObject("sBtnModifCon.Image")));
            this.sBtnModifCon.Location = new System.Drawing.Point(419, 3);
            this.sBtnModifCon.Name = "sBtnModifCon";
            this.sBtnModifCon.Size = new System.Drawing.Size(132, 26);
            this.sBtnModifCon.TabIndex = 1;
            this.sBtnModifCon.Text = "Modifier Connexion";
            this.sBtnModifCon.Click += new System.EventHandler(this.sBtnModifCon_Click);
            // 
            // sBtnAjoutCon
            // 
            this.sBtnAjoutCon.Image = ((System.Drawing.Image)(resources.GetObject("sBtnAjoutCon.Image")));
            this.sBtnAjoutCon.Location = new System.Drawing.Point(284, 3);
            this.sBtnAjoutCon.Name = "sBtnAjoutCon";
            this.sBtnAjoutCon.Size = new System.Drawing.Size(132, 26);
            this.sBtnAjoutCon.TabIndex = 0;
            this.sBtnAjoutCon.Text = "Ajouter Connexion";
            this.sBtnAjoutCon.Click += new System.EventHandler(this.sBtnAjoutCon_Click);
            // 
            // gridConnexion
            // 
            this.gridConnexion.Location = new System.Drawing.Point(3, 9);
            this.gridConnexion.MainView = this.gridView1;
            this.gridConnexion.Name = "gridConnexion";
            this.gridConnexion.Size = new System.Drawing.Size(685, 290);
            this.gridConnexion.TabIndex = 1;
            this.gridConnexion.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colId,
            this.colVille,
            this.colAliasName,
            this.colBD});
            this.gridView1.GridControl = this.gridConnexion;
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsBehavior.Editable = false;
            this.gridView1.OptionsCustomization.AllowColumnMoving = false;
            this.gridView1.OptionsCustomization.AllowGroup = false;
            this.gridView1.OptionsView.ShowGroupPanel = false;
            // 
            // colId
            // 
            this.colId.Caption = "Id";
            this.colId.FieldName = "mId";
            this.colId.Name = "colId";
            // 
            // colVille
            // 
            this.colVille.Caption = "Ville";
            this.colVille.FieldName = "mVille";
            this.colVille.Name = "colVille";
            this.colVille.Visible = true;
            this.colVille.VisibleIndex = 0;
            // 
            // colAliasName
            // 
            this.colAliasName.Caption = "Alias";
            this.colAliasName.FieldName = "mAliasName";
            this.colAliasName.Name = "colAliasName";
            this.colAliasName.Visible = true;
            this.colAliasName.VisibleIndex = 1;
            // 
            // colBD
            // 
            this.colBD.Caption = "Base Données";
            this.colBD.FieldName = "mBDName";
            this.colBD.Name = "colBD";
            this.colBD.Visible = true;
            this.colBD.VisibleIndex = 2;
            // 
            // Connexion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(694, 341);
            this.Controls.Add(this.tabConnexion);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Connexion";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Connexion";
            this.Load += new System.EventHandler(this.Connexion_Load);
            this.tabConnexion.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridConnexion)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tabConnexion;
        private DevExpress.XtraEditors.SimpleButton sBtnSupprCon;
        private DevExpress.XtraEditors.SimpleButton sBtnModifCon;
        private DevExpress.XtraEditors.SimpleButton sBtnAjoutCon;
        private DevExpress.XtraGrid.GridControl gridConnexion;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraGrid.Columns.GridColumn colId;
        private DevExpress.XtraGrid.Columns.GridColumn colVille;
        private DevExpress.XtraGrid.Columns.GridColumn colAliasName;
        private DevExpress.XtraGrid.Columns.GridColumn colBD;
    }
}