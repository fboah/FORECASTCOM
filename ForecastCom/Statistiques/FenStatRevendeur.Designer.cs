namespace ForecastCom.Statistiques
{
    partial class FenStatRevendeur
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FenStatRevendeur));
            this.layoutControlStatRev = new DevExpress.XtraLayout.LayoutControl();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.BtnExportExcelStatRev = new DevExpress.XtraEditors.SimpleButton();
            this.xtraTabControlStatRev = new DevExpress.XtraTab.XtraTabControl();
            this.xtraTabPageTypeDocMtantRev = new DevExpress.XtraTab.XtraTabPage();
            this.chartControlNbreTypeDocRev = new DevExpress.XtraCharts.ChartControl();
            this.xtraTabPageComRev = new DevExpress.XtraTab.XtraTabPage();
            this.chartControlTopComRev = new DevExpress.XtraCharts.ChartControl();
            this.xtraTabPageTopRevendeur = new DevExpress.XtraTab.XtraTabPage();
            this.chartControlTopRev = new DevExpress.XtraCharts.ChartControl();
            this.xtraTabPageBadRevendeur = new DevExpress.XtraTab.XtraTabPage();
            this.chartControlBadRevendeur = new DevExpress.XtraCharts.ChartControl();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.sFDExcelStatRev = new System.Windows.Forms.SaveFileDialog();
            this.splashScreenManager1 = new DevExpress.XtraSplashScreen.SplashScreenManager(this, typeof(global::ForecastCom.WaitForm2), true, true);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlStatRev)).BeginInit();
            this.layoutControlStatRev.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControlStatRev)).BeginInit();
            this.xtraTabControlStatRev.SuspendLayout();
            this.xtraTabPageTypeDocMtantRev.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartControlNbreTypeDocRev)).BeginInit();
            this.xtraTabPageComRev.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartControlTopComRev)).BeginInit();
            this.xtraTabPageTopRevendeur.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartControlTopRev)).BeginInit();
            this.xtraTabPageBadRevendeur.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chartControlBadRevendeur)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControlStatRev
            // 
            this.layoutControlStatRev.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControlStatRev.Controls.Add(this.panelControl1);
            this.layoutControlStatRev.Controls.Add(this.xtraTabControlStatRev);
            this.layoutControlStatRev.Location = new System.Drawing.Point(0, 1);
            this.layoutControlStatRev.Name = "layoutControlStatRev";
            this.layoutControlStatRev.Root = this.layoutControlGroup1;
            this.layoutControlStatRev.Size = new System.Drawing.Size(774, 384);
            this.layoutControlStatRev.TabIndex = 0;
            this.layoutControlStatRev.Text = "layoutControl1";
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.BtnExportExcelStatRev);
            this.panelControl1.Location = new System.Drawing.Point(676, 12);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(86, 360);
            this.panelControl1.TabIndex = 5;
            // 
            // BtnExportExcelStatRev
            // 
            this.BtnExportExcelStatRev.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.BtnExportExcelStatRev.Image = ((System.Drawing.Image)(resources.GetObject("BtnExportExcelStatRev.Image")));
            this.BtnExportExcelStatRev.Location = new System.Drawing.Point(0, 0);
            this.BtnExportExcelStatRev.Name = "BtnExportExcelStatRev";
            this.BtnExportExcelStatRev.Size = new System.Drawing.Size(88, 35);
            this.BtnExportExcelStatRev.TabIndex = 1;
            this.BtnExportExcelStatRev.Text = "Export Excel";
            this.BtnExportExcelStatRev.Click += new System.EventHandler(this.BtnExportExcelStatRev_Click);
            // 
            // xtraTabControlStatRev
            // 
            this.xtraTabControlStatRev.Location = new System.Drawing.Point(12, 12);
            this.xtraTabControlStatRev.Name = "xtraTabControlStatRev";
            this.xtraTabControlStatRev.SelectedTabPage = this.xtraTabPageTypeDocMtantRev;
            this.xtraTabControlStatRev.Size = new System.Drawing.Size(660, 360);
            this.xtraTabControlStatRev.TabIndex = 4;
            this.xtraTabControlStatRev.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtraTabPageTypeDocMtantRev,
            this.xtraTabPageComRev,
            this.xtraTabPageTopRevendeur,
            this.xtraTabPageBadRevendeur});
            // 
            // xtraTabPageTypeDocMtantRev
            // 
            this.xtraTabPageTypeDocMtantRev.Controls.Add(this.chartControlNbreTypeDocRev);
            this.xtraTabPageTypeDocMtantRev.Name = "xtraTabPageTypeDocMtantRev";
            this.xtraTabPageTypeDocMtantRev.Size = new System.Drawing.Size(654, 332);
            this.xtraTabPageTypeDocMtantRev.Text = "Stat. MtantType doc/Rev";
            // 
            // chartControlNbreTypeDocRev
            // 
            this.chartControlNbreTypeDocRev.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chartControlNbreTypeDocRev.Location = new System.Drawing.Point(0, 0);
            this.chartControlNbreTypeDocRev.Name = "chartControlNbreTypeDocRev";
            this.chartControlNbreTypeDocRev.SeriesSerializable = new DevExpress.XtraCharts.Series[0];
            this.chartControlNbreTypeDocRev.Size = new System.Drawing.Size(654, 332);
            this.chartControlNbreTypeDocRev.TabIndex = 0;
            // 
            // xtraTabPageComRev
            // 
            this.xtraTabPageComRev.Controls.Add(this.chartControlTopComRev);
            this.xtraTabPageComRev.Name = "xtraTabPageComRev";
            this.xtraTabPageComRev.Size = new System.Drawing.Size(654, 332);
            this.xtraTabPageComRev.Text = "Top 10 Transactions Commerciales";
            // 
            // chartControlTopComRev
            // 
            this.chartControlTopComRev.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chartControlTopComRev.Location = new System.Drawing.Point(0, 0);
            this.chartControlTopComRev.Name = "chartControlTopComRev";
            this.chartControlTopComRev.SeriesSerializable = new DevExpress.XtraCharts.Series[0];
            this.chartControlTopComRev.Size = new System.Drawing.Size(654, 332);
            this.chartControlTopComRev.TabIndex = 0;
            // 
            // xtraTabPageTopRevendeur
            // 
            this.xtraTabPageTopRevendeur.Controls.Add(this.chartControlTopRev);
            this.xtraTabPageTopRevendeur.Name = "xtraTabPageTopRevendeur";
            this.xtraTabPageTopRevendeur.Size = new System.Drawing.Size(654, 332);
            this.xtraTabPageTopRevendeur.Text = "Top 10 Revendeurs";
            // 
            // chartControlTopRev
            // 
            this.chartControlTopRev.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chartControlTopRev.Location = new System.Drawing.Point(0, 0);
            this.chartControlTopRev.Name = "chartControlTopRev";
            this.chartControlTopRev.SeriesSerializable = new DevExpress.XtraCharts.Series[0];
            this.chartControlTopRev.Size = new System.Drawing.Size(654, 332);
            this.chartControlTopRev.TabIndex = 0;
            // 
            // xtraTabPageBadRevendeur
            // 
            this.xtraTabPageBadRevendeur.Controls.Add(this.chartControlBadRevendeur);
            this.xtraTabPageBadRevendeur.Name = "xtraTabPageBadRevendeur";
            this.xtraTabPageBadRevendeur.Size = new System.Drawing.Size(654, 332);
            this.xtraTabPageBadRevendeur.Text = "10 Mauvais Revendeurs";
            // 
            // chartControlBadRevendeur
            // 
            this.chartControlBadRevendeur.Dock = System.Windows.Forms.DockStyle.Fill;
            this.chartControlBadRevendeur.Location = new System.Drawing.Point(0, 0);
            this.chartControlBadRevendeur.Name = "chartControlBadRevendeur";
            this.chartControlBadRevendeur.SeriesSerializable = new DevExpress.XtraCharts.Series[0];
            this.chartControlBadRevendeur.Size = new System.Drawing.Size(654, 332);
            this.chartControlBadRevendeur.TabIndex = 0;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItem2});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(774, 384);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            this.layoutControlItem1.Control = this.xtraTabControlStatRev;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(664, 364);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.panelControl1;
            this.layoutControlItem2.Location = new System.Drawing.Point(664, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(90, 364);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // sFDExcelStatRev
            // 
            this.sFDExcelStatRev.Filter = "Fichiers (*.xlsx) | *.xlsx";
            // 
            // splashScreenManager1
            // 
            this.splashScreenManager1.ClosingDelay = 500;
            // 
            // FenStatRevendeur
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(767, 376);
            this.Controls.Add(this.layoutControlStatRev);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FenStatRevendeur";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Statistiques Revendeurs";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.FenStatRevendeur_Load);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlStatRev)).EndInit();
            this.layoutControlStatRev.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControlStatRev)).EndInit();
            this.xtraTabControlStatRev.ResumeLayout(false);
            this.xtraTabPageTypeDocMtantRev.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chartControlNbreTypeDocRev)).EndInit();
            this.xtraTabPageComRev.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chartControlTopComRev)).EndInit();
            this.xtraTabPageTopRevendeur.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chartControlTopRev)).EndInit();
            this.xtraTabPageBadRevendeur.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chartControlBadRevendeur)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControlStatRev;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraTab.XtraTabControl xtraTabControlStatRev;
        private DevExpress.XtraTab.XtraTabPage xtraTabPageTypeDocMtantRev;
        private DevExpress.XtraTab.XtraTabPage xtraTabPageComRev;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraCharts.ChartControl chartControlNbreTypeDocRev;
        private DevExpress.XtraCharts.ChartControl chartControlTopComRev;
        private DevExpress.XtraTab.XtraTabPage xtraTabPageTopRevendeur;
        private DevExpress.XtraCharts.ChartControl chartControlTopRev;
        private DevExpress.XtraTab.XtraTabPage xtraTabPageBadRevendeur;
        private DevExpress.XtraCharts.ChartControl chartControlBadRevendeur;
        private DevExpress.XtraEditors.PanelControl panelControl1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraEditors.SimpleButton BtnExportExcelStatRev;
        private System.Windows.Forms.SaveFileDialog sFDExcelStatRev;
        private DevExpress.XtraSplashScreen.SplashScreenManager splashScreenManager1;
    }
}