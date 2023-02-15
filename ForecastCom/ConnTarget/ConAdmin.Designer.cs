namespace ForecastCom
{
    partial class ConAdmin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConAdmin));
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panelControlBoutons = new DevExpress.XtraEditors.PanelControl();
            this.sBtnSupprConTarg = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnAjoutConTarg = new DevExpress.XtraEditors.SimpleButton();
            this.gridControlConTarget = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colId = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colVille = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colAliasName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colBDName = new DevExpress.XtraGrid.Columns.GridColumn();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControlBoutons)).BeginInit();
            this.panelControlBoutons.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridControlConTarget)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.panelControlBoutons, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.gridControlConTarget, 0, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(1, 1);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 2.830189F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 97.16982F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(691, 339);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panelControlBoutons
            // 
            this.panelControlBoutons.Controls.Add(this.sBtnSupprConTarg);
            this.panelControlBoutons.Controls.Add(this.sBtnAjoutConTarg);
            this.panelControlBoutons.Location = new System.Drawing.Point(3, 305);
            this.panelControlBoutons.Name = "panelControlBoutons";
            this.panelControlBoutons.Size = new System.Drawing.Size(685, 31);
            this.panelControlBoutons.TabIndex = 0;
            // 
            // sBtnSupprConTarg
            // 
            this.sBtnSupprConTarg.Image = ((System.Drawing.Image)(resources.GetObject("sBtnSupprConTarg.Image")));
            this.sBtnSupprConTarg.Location = new System.Drawing.Point(551, 2);
            this.sBtnSupprConTarg.Name = "sBtnSupprConTarg";
            this.sBtnSupprConTarg.Size = new System.Drawing.Size(132, 26);
            this.sBtnSupprConTarg.TabIndex = 4;
            this.sBtnSupprConTarg.Text = "Supprimer Connexion";
            this.sBtnSupprConTarg.Click += new System.EventHandler(this.sBtnSupprConTarg_Click);
            // 
            // sBtnAjoutConTarg
            // 
            this.sBtnAjoutConTarg.Image = ((System.Drawing.Image)(resources.GetObject("sBtnAjoutConTarg.Image")));
            this.sBtnAjoutConTarg.Location = new System.Drawing.Point(413, 2);
            this.sBtnAjoutConTarg.Name = "sBtnAjoutConTarg";
            this.sBtnAjoutConTarg.Size = new System.Drawing.Size(132, 26);
            this.sBtnAjoutConTarg.TabIndex = 3;
            this.sBtnAjoutConTarg.Text = "Ajouter Connexion";
            this.sBtnAjoutConTarg.Click += new System.EventHandler(this.sBtnAjoutConTarg_Click);
            // 
            // gridControlConTarget
            // 
            this.gridControlConTarget.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gridControlConTarget.Location = new System.Drawing.Point(3, 11);
            this.gridControlConTarget.MainView = this.gridView1;
            this.gridControlConTarget.Name = "gridControlConTarget";
            this.gridControlConTarget.Size = new System.Drawing.Size(685, 288);
            this.gridControlConTarget.TabIndex = 1;
            this.gridControlConTarget.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colId,
            this.colVille,
            this.colAliasName,
            this.colBDName});
            this.gridView1.GridControl = this.gridControlConTarget;
            this.gridView1.Name = "gridView1";
            // 
            // colId
            // 
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
            // colBDName
            // 
            this.colBDName.Caption = "Base Données";
            this.colBDName.FieldName = "mBDName";
            this.colBDName.Name = "colBDName";
            this.colBDName.Visible = true;
            this.colBDName.VisibleIndex = 2;
            // 
            // ConnTarget
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(694, 341);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "ConnTarget";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ConnTarget";
            this.Load += new System.EventHandler(this.ConnTarget_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControlBoutons)).EndInit();
            this.panelControlBoutons.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridControlConTarget)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private DevExpress.XtraEditors.PanelControl panelControlBoutons;
        private DevExpress.XtraEditors.SimpleButton sBtnSupprConTarg;
        private DevExpress.XtraEditors.SimpleButton sBtnAjoutConTarg;
        private DevExpress.XtraGrid.GridControl gridControlConTarget;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraGrid.Columns.GridColumn colId;
        private DevExpress.XtraGrid.Columns.GridColumn colVille;
        private DevExpress.XtraGrid.Columns.GridColumn colAliasName;
        private DevExpress.XtraGrid.Columns.GridColumn colBDName;
    }
}