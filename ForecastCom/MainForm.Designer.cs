using DevExpress.XtraEditors;
using DevExpress.XtraTab;

namespace ForecastCom
{
    partial class MainForm
    {
        /// <summary>
        /// Variable nécessaire au concepteur.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Nettoyage des ressources utilisées.
        /// </summary>
        /// <param name="disposing">true si les ressources managées doivent être supprimées ; sinon, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.connexionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gestionTargetToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.connexionsToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.AdminToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.connexionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.gestionTargetToolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.gestionBUToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.groupControl2 = new DevExpress.XtraEditors.GroupControl();
            this.pictureEdit1 = new DevExpress.XtraEditors.PictureEdit();
            this.ChkListBoxControlFiliale = new DevExpress.XtraEditors.CheckedListBoxControl();
            this.gridControlForecast = new DevExpress.XtraGrid.GridControl();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.colPays = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colNom = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPrenom = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFamilleCentr = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colFamille = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colTypeDoc = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colStatutDoc = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDate = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMontantHT = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colMarge = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colRevendeur = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPaysRevendeur = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colPiece = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colDateLivraison = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colReference = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colmDO_Coord01 = new DevExpress.XtraGrid.Columns.GridColumn();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem2 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItem3 = new DevExpress.XtraLayout.LayoutControlItem();
            this.panelControl2 = new DevExpress.XtraEditors.PanelControl();
            this.sBtnAnalyseComp = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnStatRev = new DevExpress.XtraEditors.SimpleButton();
            this.BtnStat = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnGenererExcel = new DevExpress.XtraEditors.SimpleButton();
            this.sBtnVisualiserArticle = new DevExpress.XtraEditors.SimpleButton();
            this.xtraTabControl1 = new DevExpress.XtraTab.XtraTabControl();
            this.xtraTabPage1 = new DevExpress.XtraTab.XtraTabPage();
            this.panelControl1 = new DevExpress.XtraEditors.PanelControl();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkFacture = new DevExpress.XtraEditors.CheckEdit();
            this.chkSaisieDevis = new DevExpress.XtraEditors.CheckEdit();
            this.chkAccepteDevis = new DevExpress.XtraEditors.CheckEdit();
            this.groupControl5 = new DevExpress.XtraEditors.GroupControl();
            this.lblAccurency = new DevExpress.XtraEditors.LabelControl();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lblControlTotalHT = new DevExpress.XtraEditors.LabelControl();
            this.label11 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblControlRealiseHT = new DevExpress.XtraEditors.LabelControl();
            this.label10 = new System.Windows.Forms.Label();
            this.SMulCom = new DevExpress.XtraEditors.CheckEdit();
            this.chkTousCom = new DevExpress.XtraEditors.CheckEdit();
            this.LECommerciauxDE = new DevExpress.XtraEditors.LookUpEdit();
            this.LECommerciauxA = new DevExpress.XtraEditors.LookUpEdit();
            this.lblACom = new System.Windows.Forms.Label();
            this.lblDeCom = new System.Windows.Forms.Label();
            this.chkCmbMultCom = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.gCPeriode = new DevExpress.XtraEditors.GroupControl();
            this.dateEditDateDeb = new DevExpress.XtraEditors.DateEdit();
            this.label2 = new System.Windows.Forms.Label();
            this.dateEditDateFin = new DevExpress.XtraEditors.DateEdit();
            this.label1 = new System.Windows.Forms.Label();
            this.xtraTabPage2 = new DevExpress.XtraTab.XtraTabPage();
            this.panelControl3 = new DevExpress.XtraEditors.PanelControl();
            this.groupControl1 = new DevExpress.XtraEditors.GroupControl();
            this.panTypeDoc = new DevExpress.XtraEditors.PanelControl();
            this.chkCmbTypeDoc = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.chkTousTypeDoc = new DevExpress.XtraEditors.CheckEdit();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblControlTotalHT2 = new DevExpress.XtraEditors.LabelControl();
            this.gCFamille = new DevExpress.XtraEditors.GroupControl();
            this.panMarque = new DevExpress.XtraEditors.PanelControl();
            this.LEFamilleCentralDE = new DevExpress.XtraEditors.LookUpEdit();
            this.chkCmbMultFamCentral = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.lblFamCentralDe = new System.Windows.Forms.Label();
            this.lblFamCentralA = new System.Windows.Forms.Label();
            this.SMulFamilleCentral = new DevExpress.XtraEditors.CheckEdit();
            this.LEFamilleCentralA = new DevExpress.XtraEditors.LookUpEdit();
            this.chkTousFamilleCentral = new DevExpress.XtraEditors.CheckEdit();
            this.chkConfigFamilles = new DevExpress.XtraEditors.CheckEdit();
            this.panFamille = new DevExpress.XtraEditors.PanelControl();
            this.lblFamDe = new System.Windows.Forms.Label();
            this.lblFamA = new System.Windows.Forms.Label();
            this.LEFamilleA = new DevExpress.XtraEditors.LookUpEdit();
            this.LEFamilleDE = new DevExpress.XtraEditors.LookUpEdit();
            this.chkTousFamille = new DevExpress.XtraEditors.CheckEdit();
            this.SMulFamille = new DevExpress.XtraEditors.CheckEdit();
            this.chkCmbMultFam = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.label7 = new System.Windows.Forms.Label();
            this.xtraTabPageRevendeurs = new DevExpress.XtraTab.XtraTabPage();
            this.panelControl4 = new DevExpress.XtraEditors.PanelControl();
            this.gCRevendeur = new DevExpress.XtraEditors.GroupControl();
            this.lblAccurency2 = new DevExpress.XtraEditors.LabelControl();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.lblControlRealiseHT2 = new DevExpress.XtraEditors.LabelControl();
            this.label17 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblControlTotalHT3 = new DevExpress.XtraEditors.LabelControl();
            this.SMulRevendeur = new DevExpress.XtraEditors.CheckEdit();
            this.chkTousRevendeur = new DevExpress.XtraEditors.CheckEdit();
            this.LERevendeurDE = new DevExpress.XtraEditors.LookUpEdit();
            this.LERevendeurA = new DevExpress.XtraEditors.LookUpEdit();
            this.lblRevA = new System.Windows.Forms.Label();
            this.lblRevDe = new System.Windows.Forms.Label();
            this.chkCmbMultRev = new DevExpress.XtraEditors.CheckedComboBoxEdit();
            this.sFDExcel = new System.Windows.Forms.SaveFileDialog();
            this.splashScreenManager2 = new DevExpress.XtraSplashScreen.SplashScreenManager(this, typeof(global::ForecastCom.WaitForm1), true, true);
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl2)).BeginInit();
            this.groupControl2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureEdit1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ChkListBoxControlFiliale)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControlForecast)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl2)).BeginInit();
            this.panelControl2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).BeginInit();
            this.xtraTabControl1.SuspendLayout();
            this.xtraTabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).BeginInit();
            this.panelControl1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chkFacture.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkSaisieDevis.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkAccepteDevis.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl5)).BeginInit();
            this.groupControl5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SMulCom.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousCom.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LECommerciauxDE.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LECommerciauxA.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultCom.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gCPeriode)).BeginInit();
            this.gCPeriode.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateDeb.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateDeb.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateFin.Properties.CalendarTimeProperties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateFin.Properties)).BeginInit();
            this.xtraTabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl3)).BeginInit();
            this.panelControl3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).BeginInit();
            this.groupControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panTypeDoc)).BeginInit();
            this.panTypeDoc.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbTypeDoc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousTypeDoc.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gCFamille)).BeginInit();
            this.gCFamille.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panMarque)).BeginInit();
            this.panMarque.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleCentralDE.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultFamCentral.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.SMulFamilleCentral.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleCentralA.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousFamilleCentral.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkConfigFamilles.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.panFamille)).BeginInit();
            this.panFamille.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleA.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleDE.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousFamille.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.SMulFamille.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultFam.Properties)).BeginInit();
            this.xtraTabPageRevendeurs.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl4)).BeginInit();
            this.panelControl4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gCRevendeur)).BeginInit();
            this.gCRevendeur.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SMulRevendeur.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousRevendeur.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LERevendeurDE.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LERevendeurA.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultRev.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.connexionsToolStripMenuItem,
            this.gestionTargetToolStripMenuItem,
            this.menuToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(944, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // connexionsToolStripMenuItem
            // 
            this.connexionsToolStripMenuItem.Name = "connexionsToolStripMenuItem";
            this.connexionsToolStripMenuItem.Size = new System.Drawing.Size(82, 20);
            this.connexionsToolStripMenuItem.Text = "Connexions";
            this.connexionsToolStripMenuItem.Visible = false;
            this.connexionsToolStripMenuItem.Click += new System.EventHandler(this.connexionsToolStripMenuItem_Click);
            // 
            // gestionTargetToolStripMenuItem
            // 
            this.gestionTargetToolStripMenuItem.Name = "gestionTargetToolStripMenuItem";
            this.gestionTargetToolStripMenuItem.Size = new System.Drawing.Size(94, 20);
            this.gestionTargetToolStripMenuItem.Text = "Gestion Target";
            this.gestionTargetToolStripMenuItem.Visible = false;
            this.gestionTargetToolStripMenuItem.Click += new System.EventHandler(this.gestionTargetToolStripMenuItem_Click);
            // 
            // menuToolStripMenuItem
            // 
            this.menuToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.connexionsToolStripMenuItem1,
            this.AdminToolStripMenuItem});
            this.menuToolStripMenuItem.Name = "menuToolStripMenuItem";
            this.menuToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.menuToolStripMenuItem.Text = "Menu";
            // 
            // connexionsToolStripMenuItem1
            // 
            this.connexionsToolStripMenuItem1.Name = "connexionsToolStripMenuItem1";
            this.connexionsToolStripMenuItem1.Size = new System.Drawing.Size(192, 22);
            this.connexionsToolStripMenuItem1.Text = "Connexions Reporting";
            this.connexionsToolStripMenuItem1.Click += new System.EventHandler(this.connexionsToolStripMenuItem1_Click);
            // 
            // AdminToolStripMenuItem
            // 
            this.AdminToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.connexionToolStripMenuItem,
            this.gestionTargetToolStripMenuItem2,
            this.gestionBUToolStripMenuItem});
            this.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem";
            this.AdminToolStripMenuItem.Size = new System.Drawing.Size(192, 22);
            this.AdminToolStripMenuItem.Text = "Administration";
            this.AdminToolStripMenuItem.Visible = false;
            // 
            // connexionToolStripMenuItem
            // 
            this.connexionToolStripMenuItem.Name = "connexionToolStripMenuItem";
            this.connexionToolStripMenuItem.Size = new System.Drawing.Size(149, 22);
            this.connexionToolStripMenuItem.Text = "Connexion";
            this.connexionToolStripMenuItem.Click += new System.EventHandler(this.connexionToolStripMenuItem_Click);
            // 
            // gestionTargetToolStripMenuItem2
            // 
            this.gestionTargetToolStripMenuItem2.Name = "gestionTargetToolStripMenuItem2";
            this.gestionTargetToolStripMenuItem2.Size = new System.Drawing.Size(149, 22);
            this.gestionTargetToolStripMenuItem2.Text = "Gestion Target";
            this.gestionTargetToolStripMenuItem2.Click += new System.EventHandler(this.gestionTargetToolStripMenuItem2_Click);
            // 
            // gestionBUToolStripMenuItem
            // 
            this.gestionBUToolStripMenuItem.Name = "gestionBUToolStripMenuItem";
            this.gestionBUToolStripMenuItem.Size = new System.Drawing.Size(149, 22);
            this.gestionBUToolStripMenuItem.Text = "Gestion B.U";
            // 
            // layoutControl1
            // 
            this.layoutControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.layoutControl1.Controls.Add(this.groupControl2);
            this.layoutControl1.Controls.Add(this.gridControlForecast);
            this.layoutControl1.Location = new System.Drawing.Point(-9, 182);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.Root = this.layoutControlGroup1;
            this.layoutControl1.Size = new System.Drawing.Size(959, 228);
            this.layoutControl1.TabIndex = 1;
            this.layoutControl1.Text = "layoutControl1";
            // 
            // groupControl2
            // 
            this.groupControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupControl2.Controls.Add(this.pictureEdit1);
            this.groupControl2.Controls.Add(this.ChkListBoxControlFiliale);
            this.groupControl2.Location = new System.Drawing.Point(768, 12);
            this.groupControl2.Name = "groupControl2";
            this.groupControl2.Size = new System.Drawing.Size(179, 204);
            this.groupControl2.TabIndex = 6;
            this.groupControl2.Text = "Sites";
            // 
            // pictureEdit1
            // 
            this.pictureEdit1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureEdit1.EditValue = global::ForecastCom.Properties.Resources.Aitek_logo;
            this.pictureEdit1.Location = new System.Drawing.Point(1, 21);
            this.pictureEdit1.Name = "pictureEdit1";
            this.pictureEdit1.Properties.ShowCameraMenuItem = DevExpress.XtraEditors.Controls.CameraMenuItemVisibility.Auto;
            this.pictureEdit1.Properties.SizeMode = DevExpress.XtraEditors.Controls.PictureSizeMode.Squeeze;
            this.pictureEdit1.Size = new System.Drawing.Size(177, 83);
            this.pictureEdit1.TabIndex = 1;
            // 
            // ChkListBoxControlFiliale
            // 
            this.ChkListBoxControlFiliale.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ChkListBoxControlFiliale.Appearance.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.ChkListBoxControlFiliale.Appearance.Options.UseFont = true;
            this.ChkListBoxControlFiliale.CheckOnClick = true;
            this.ChkListBoxControlFiliale.Location = new System.Drawing.Point(2, 104);
            this.ChkListBoxControlFiliale.Name = "ChkListBoxControlFiliale";
            this.ChkListBoxControlFiliale.Size = new System.Drawing.Size(175, 98);
            this.ChkListBoxControlFiliale.TabIndex = 0;
            this.ChkListBoxControlFiliale.ItemCheck += new DevExpress.XtraEditors.Controls.ItemCheckEventHandler(this.chkListBoxControlFiliale_ItemCheck);
            // 
            // gridControlForecast
            // 
            this.gridControlForecast.Location = new System.Drawing.Point(12, 12);
            this.gridControlForecast.MainView = this.gridView1;
            this.gridControlForecast.Name = "gridControlForecast";
            this.gridControlForecast.Size = new System.Drawing.Size(752, 204);
            this.gridControlForecast.TabIndex = 5;
            this.gridControlForecast.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.colPays,
            this.colNom,
            this.colPrenom,
            this.colFamilleCentr,
            this.colFamille,
            this.colTypeDoc,
            this.colStatutDoc,
            this.colDate,
            this.colMontantHT,
            this.colMarge,
            this.colRevendeur,
            this.colPaysRevendeur,
            this.colPiece,
            this.colDateLivraison,
            this.colReference,
            this.colmDO_Coord01});
            this.gridView1.GridControl = this.gridControlForecast;
            this.gridView1.GroupSummary.AddRange(new DevExpress.XtraGrid.GridSummaryItem[] {
            new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "mMontantHT", this.colMontantHT, ""),
            new DevExpress.XtraGrid.GridGroupSummaryItem(DevExpress.Data.SummaryItemType.Sum, "mMargeBrute", this.colMarge, "")});
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsBehavior.Editable = false;
            this.gridView1.OptionsFind.AlwaysVisible = true;
            this.gridView1.OptionsPrint.AutoWidth = false;
            this.gridView1.OptionsView.ColumnAutoWidth = false;
            // 
            // colPays
            // 
            this.colPays.Caption = "Site AITEK";
            this.colPays.FieldName = "mPays";
            this.colPays.Name = "colPays";
            this.colPays.Visible = true;
            this.colPays.VisibleIndex = 0;
            // 
            // colNom
            // 
            this.colNom.Caption = "Nom";
            this.colNom.FieldName = "mNomCommercial";
            this.colNom.Name = "colNom";
            this.colNom.Visible = true;
            this.colNom.VisibleIndex = 1;
            this.colNom.Width = 90;
            // 
            // colPrenom
            // 
            this.colPrenom.Caption = "Prénoms";
            this.colPrenom.FieldName = "mPrenomCommercial";
            this.colPrenom.Name = "colPrenom";
            this.colPrenom.Visible = true;
            this.colPrenom.VisibleIndex = 2;
            this.colPrenom.Width = 90;
            // 
            // colFamilleCentr
            // 
            this.colFamilleCentr.Caption = "Marque";
            this.colFamilleCentr.FieldName = "mFamilleCentral";
            this.colFamilleCentr.Name = "colFamilleCentr";
            this.colFamilleCentr.Visible = true;
            this.colFamilleCentr.VisibleIndex = 3;
            this.colFamilleCentr.Width = 80;
            // 
            // colFamille
            // 
            this.colFamille.Caption = "Famille";
            this.colFamille.FieldName = "mFamille";
            this.colFamille.Name = "colFamille";
            this.colFamille.Visible = true;
            this.colFamille.VisibleIndex = 4;
            this.colFamille.Width = 90;
            // 
            // colTypeDoc
            // 
            this.colTypeDoc.Caption = "TypeDoc";
            this.colTypeDoc.FieldName = "mTypeDoc";
            this.colTypeDoc.Name = "colTypeDoc";
            this.colTypeDoc.Visible = true;
            this.colTypeDoc.VisibleIndex = 5;
            this.colTypeDoc.Width = 50;
            // 
            // colStatutDoc
            // 
            this.colStatutDoc.Caption = "Statut Doc";
            this.colStatutDoc.FieldName = "mStatut";
            this.colStatutDoc.Name = "colStatutDoc";
            this.colStatutDoc.Visible = true;
            this.colStatutDoc.VisibleIndex = 6;
            this.colStatutDoc.Width = 150;
            // 
            // colDate
            // 
            this.colDate.Caption = "Date";
            this.colDate.FieldName = "mDatePiece";
            this.colDate.Name = "colDate";
            this.colDate.Visible = true;
            this.colDate.VisibleIndex = 7;
            this.colDate.Width = 70;
            // 
            // colMontantHT
            // 
            this.colMontantHT.Caption = "MontantHT";
            this.colMontantHT.DisplayFormat.FormatString = "n0";
            this.colMontantHT.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.colMontantHT.FieldName = "mMontantHT";
            this.colMontantHT.Name = "colMontantHT";
            this.colMontantHT.Visible = true;
            this.colMontantHT.VisibleIndex = 8;
            this.colMontantHT.Width = 80;
            // 
            // colMarge
            // 
            this.colMarge.Caption = "Marge";
            this.colMarge.DisplayFormat.FormatString = "n2";
            this.colMarge.DisplayFormat.FormatType = DevExpress.Utils.FormatType.Numeric;
            this.colMarge.FieldName = "mMargeBrute";
            this.colMarge.Name = "colMarge";
            this.colMarge.Visible = true;
            this.colMarge.VisibleIndex = 9;
            this.colMarge.Width = 80;
            // 
            // colRevendeur
            // 
            this.colRevendeur.Caption = "Revendeur";
            this.colRevendeur.FieldName = "mRevendeur";
            this.colRevendeur.Name = "colRevendeur";
            this.colRevendeur.Visible = true;
            this.colRevendeur.VisibleIndex = 10;
            this.colRevendeur.Width = 110;
            // 
            // colPaysRevendeur
            // 
            this.colPaysRevendeur.Caption = "Pays Revendeur";
            this.colPaysRevendeur.FieldName = "mPaysRevendeur";
            this.colPaysRevendeur.Name = "colPaysRevendeur";
            this.colPaysRevendeur.Visible = true;
            this.colPaysRevendeur.VisibleIndex = 11;
            // 
            // colPiece
            // 
            this.colPiece.Caption = "Pièce";
            this.colPiece.FieldName = "mNumPiece";
            this.colPiece.Name = "colPiece";
            this.colPiece.Visible = true;
            this.colPiece.VisibleIndex = 12;
            this.colPiece.Width = 70;
            // 
            // colDateLivraison
            // 
            this.colDateLivraison.Caption = "Date Livraison";
            this.colDateLivraison.FieldName = "mDateLivraison";
            this.colDateLivraison.Name = "colDateLivraison";
            this.colDateLivraison.Visible = true;
            this.colDateLivraison.VisibleIndex = 13;
            // 
            // colReference
            // 
            this.colReference.Caption = "Reference";
            this.colReference.FieldName = "mReference";
            this.colReference.Name = "colReference";
            this.colReference.Visible = true;
            this.colReference.VisibleIndex = 14;
            // 
            // colmDO_Coord01
            // 
            this.colmDO_Coord01.Caption = "Entete1";
            this.colmDO_Coord01.FieldName = "mDO_Coord01";
            this.colmDO_Coord01.Name = "colmDO_Coord01";
            this.colmDO_Coord01.Visible = true;
            this.colmDO_Coord01.VisibleIndex = 15;
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem2,
            this.layoutControlItem3});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(959, 228);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem2
            // 
            this.layoutControlItem2.Control = this.gridControlForecast;
            this.layoutControlItem2.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem2.Name = "layoutControlItem2";
            this.layoutControlItem2.Size = new System.Drawing.Size(756, 208);
            this.layoutControlItem2.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem2.TextVisible = false;
            // 
            // layoutControlItem3
            // 
            this.layoutControlItem3.Control = this.groupControl2;
            this.layoutControlItem3.Location = new System.Drawing.Point(756, 0);
            this.layoutControlItem3.Name = "layoutControlItem3";
            this.layoutControlItem3.Size = new System.Drawing.Size(183, 208);
            this.layoutControlItem3.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem3.TextVisible = false;
            // 
            // panelControl2
            // 
            this.panelControl2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelControl2.Controls.Add(this.sBtnAnalyseComp);
            this.panelControl2.Controls.Add(this.sBtnStatRev);
            this.panelControl2.Controls.Add(this.BtnStat);
            this.panelControl2.Controls.Add(this.sBtnGenererExcel);
            this.panelControl2.Controls.Add(this.sBtnVisualiserArticle);
            this.panelControl2.Location = new System.Drawing.Point(1, 152);
            this.panelControl2.Name = "panelControl2";
            this.panelControl2.Size = new System.Drawing.Size(935, 28);
            this.panelControl2.TabIndex = 8;
            // 
            // sBtnAnalyseComp
            // 
            this.sBtnAnalyseComp.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.sBtnAnalyseComp.Image = ((System.Drawing.Image)(resources.GetObject("sBtnAnalyseComp.Image")));
            this.sBtnAnalyseComp.Location = new System.Drawing.Point(574, 1);
            this.sBtnAnalyseComp.Name = "sBtnAnalyseComp";
            this.sBtnAnalyseComp.Size = new System.Drawing.Size(137, 25);
            this.sBtnAnalyseComp.TabIndex = 12;
            this.sBtnAnalyseComp.Text = "Analyse Comparative";
            this.sBtnAnalyseComp.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // sBtnStatRev
            // 
            this.sBtnStatRev.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.sBtnStatRev.Image = ((System.Drawing.Image)(resources.GetObject("sBtnStatRev.Image")));
            this.sBtnStatRev.Location = new System.Drawing.Point(418, 1);
            this.sBtnStatRev.Name = "sBtnStatRev";
            this.sBtnStatRev.Size = new System.Drawing.Size(154, 25);
            this.sBtnStatRev.TabIndex = 11;
            this.sBtnStatRev.Text = "Statistiques Revendeurs";
            this.sBtnStatRev.Click += new System.EventHandler(this.sBtnStatRev_Click);
            // 
            // BtnStat
            // 
            this.BtnStat.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.BtnStat.Image = ((System.Drawing.Image)(resources.GetObject("BtnStat.Image")));
            this.BtnStat.Location = new System.Drawing.Point(259, 1);
            this.BtnStat.Name = "BtnStat";
            this.BtnStat.Size = new System.Drawing.Size(157, 25);
            this.BtnStat.TabIndex = 10;
            this.BtnStat.Text = "Statistiques Commerciaux";
            this.BtnStat.Click += new System.EventHandler(this.BtnStat_Click);
            // 
            // sBtnGenererExcel
            // 
            this.sBtnGenererExcel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.sBtnGenererExcel.Image = ((System.Drawing.Image)(resources.GetObject("sBtnGenererExcel.Image")));
            this.sBtnGenererExcel.Location = new System.Drawing.Point(130, 1);
            this.sBtnGenererExcel.Name = "sBtnGenererExcel";
            this.sBtnGenererExcel.Size = new System.Drawing.Size(127, 25);
            this.sBtnGenererExcel.TabIndex = 8;
            this.sBtnGenererExcel.Text = "Générer  Excel";
            this.sBtnGenererExcel.Click += new System.EventHandler(this.sBtnGenererExcel_Click);
            // 
            // sBtnVisualiserArticle
            // 
            this.sBtnVisualiserArticle.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.sBtnVisualiserArticle.Appearance.ForeColor = System.Drawing.Color.Black;
            this.sBtnVisualiserArticle.Appearance.Options.UseForeColor = true;
            this.sBtnVisualiserArticle.Image = ((System.Drawing.Image)(resources.GetObject("sBtnVisualiserArticle.Image")));
            this.sBtnVisualiserArticle.Location = new System.Drawing.Point(1, 1);
            this.sBtnVisualiserArticle.Name = "sBtnVisualiserArticle";
            this.sBtnVisualiserArticle.Size = new System.Drawing.Size(127, 25);
            this.sBtnVisualiserArticle.TabIndex = 7;
            this.sBtnVisualiserArticle.Text = "Aperçu";
            this.sBtnVisualiserArticle.Click += new System.EventHandler(this.sBtnVisualiserArticle_Click);
            // 
            // xtraTabControl1
            // 
            this.xtraTabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.xtraTabControl1.Location = new System.Drawing.Point(0, 27);
            this.xtraTabControl1.Name = "xtraTabControl1";
            this.xtraTabControl1.SelectedTabPage = this.xtraTabPage1;
            this.xtraTabControl1.Size = new System.Drawing.Size(935, 125);
            this.xtraTabControl1.TabIndex = 7;
            this.xtraTabControl1.TabPages.AddRange(new DevExpress.XtraTab.XtraTabPage[] {
            this.xtraTabPage1,
            this.xtraTabPage2,
            this.xtraTabPageRevendeurs});
            this.xtraTabControl1.Click += new System.EventHandler(this.xtraTabControl1_Click);
            // 
            // xtraTabPage1
            // 
            this.xtraTabPage1.Controls.Add(this.panelControl1);
            this.xtraTabPage1.Name = "xtraTabPage1";
            this.xtraTabPage1.Size = new System.Drawing.Size(929, 97);
            this.xtraTabPage1.Text = "Période - Commerciaux";
            // 
            // panelControl1
            // 
            this.panelControl1.Controls.Add(this.groupBox1);
            this.panelControl1.Controls.Add(this.groupControl5);
            this.panelControl1.Controls.Add(this.gCPeriode);
            this.panelControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelControl1.Location = new System.Drawing.Point(0, 0);
            this.panelControl1.Name = "panelControl1";
            this.panelControl1.Size = new System.Drawing.Size(929, 97);
            this.panelControl1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.groupBox1.Controls.Add(this.chkFacture);
            this.groupBox1.Controls.Add(this.chkSaisieDevis);
            this.groupBox1.Controls.Add(this.chkAccepteDevis);
            this.groupBox1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.groupBox1.Location = new System.Drawing.Point(315, 48);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(390, 47);
            this.groupBox1.TabIndex = 32;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Pipe";
            // 
            // chkFacture
            // 
            this.chkFacture.Location = new System.Drawing.Point(287, 20);
            this.chkFacture.Name = "chkFacture";
            this.chkFacture.Properties.Appearance.ForeColor = System.Drawing.Color.Black;
            this.chkFacture.Properties.Appearance.Options.UseForeColor = true;
            this.chkFacture.Properties.Caption = "100% :Facturé";
            this.chkFacture.Size = new System.Drawing.Size(97, 19);
            this.chkFacture.TabIndex = 30;
            this.chkFacture.CheckedChanged += new System.EventHandler(this.chkFacture_CheckedChanged);
            // 
            // chkSaisieDevis
            // 
            this.chkSaisieDevis.Location = new System.Drawing.Point(4, 20);
            this.chkSaisieDevis.Name = "chkSaisieDevis";
            this.chkSaisieDevis.Properties.Appearance.ForeColor = System.Drawing.Color.Black;
            this.chkSaisieDevis.Properties.Appearance.Options.UseForeColor = true;
            this.chkSaisieDevis.Properties.Caption = "60% :Devis envoyé";
            this.chkSaisieDevis.Size = new System.Drawing.Size(120, 19);
            this.chkSaisieDevis.TabIndex = 27;
            this.chkSaisieDevis.CheckedChanged += new System.EventHandler(this.chkSaisieDevis_CheckedChanged);
            // 
            // chkAccepteDevis
            // 
            this.chkAccepteDevis.Location = new System.Drawing.Point(128, 20);
            this.chkAccepteDevis.Name = "chkAccepteDevis";
            this.chkAccepteDevis.Properties.Appearance.ForeColor = System.Drawing.Color.Black;
            this.chkAccepteDevis.Properties.Appearance.Options.UseForeColor = true;
            this.chkAccepteDevis.Properties.Caption = "75% :Négociation en cours";
            this.chkAccepteDevis.Size = new System.Drawing.Size(149, 19);
            this.chkAccepteDevis.TabIndex = 29;
            this.chkAccepteDevis.CheckedChanged += new System.EventHandler(this.chkAccepteDevis_CheckedChanged);
            // 
            // groupControl5
            // 
            this.groupControl5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.groupControl5.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.groupControl5.AppearanceCaption.Options.UseFont = true;
            this.groupControl5.CaptionImage = ((System.Drawing.Image)(resources.GetObject("groupControl5.CaptionImage")));
            this.groupControl5.Controls.Add(this.lblAccurency);
            this.groupControl5.Controls.Add(this.label13);
            this.groupControl5.Controls.Add(this.label12);
            this.groupControl5.Controls.Add(this.label3);
            this.groupControl5.Controls.Add(this.lblControlTotalHT);
            this.groupControl5.Controls.Add(this.label11);
            this.groupControl5.Controls.Add(this.label4);
            this.groupControl5.Controls.Add(this.lblControlRealiseHT);
            this.groupControl5.Controls.Add(this.label10);
            this.groupControl5.Controls.Add(this.SMulCom);
            this.groupControl5.Controls.Add(this.chkTousCom);
            this.groupControl5.Controls.Add(this.LECommerciauxDE);
            this.groupControl5.Controls.Add(this.LECommerciauxA);
            this.groupControl5.Controls.Add(this.lblACom);
            this.groupControl5.Controls.Add(this.lblDeCom);
            this.groupControl5.Controls.Add(this.chkCmbMultCom);
            this.groupControl5.Location = new System.Drawing.Point(312, -2);
            this.groupControl5.Name = "groupControl5";
            this.groupControl5.Size = new System.Drawing.Size(616, 101);
            this.groupControl5.TabIndex = 19;
            this.groupControl5.Text = "Commerciaux";
            // 
            // lblAccurency
            // 
            this.lblAccurency.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblAccurency.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblAccurency.Appearance.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblAccurency.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblAccurency.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblAccurency.Location = new System.Drawing.Point(473, 82);
            this.lblAccurency.Name = "lblAccurency";
            this.lblAccurency.Size = new System.Drawing.Size(105, 13);
            this.lblAccurency.TabIndex = 34;
            this.lblAccurency.Text = "labTotal";
            // 
            // label13
            // 
            this.label13.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label13.Location = new System.Drawing.Point(428, 82);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(49, 13);
            this.label13.TabIndex = 33;
            this.label13.Text = "Marge :";
            // 
            // label12
            // 
            this.label12.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(578, 82);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(36, 13);
            this.label12.TabIndex = 32;
            this.label12.Text = "F CFA";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(392, 47);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(85, 13);
            this.label3.TabIndex = 27;
            this.label3.Text = "Montant Pipe:";
            // 
            // lblControlTotalHT
            // 
            this.lblControlTotalHT.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblControlTotalHT.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblControlTotalHT.Appearance.ForeColor = System.Drawing.Color.Red;
            this.lblControlTotalHT.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblControlTotalHT.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblControlTotalHT.Location = new System.Drawing.Point(473, 47);
            this.lblControlTotalHT.Name = "lblControlTotalHT";
            this.lblControlTotalHT.Size = new System.Drawing.Size(105, 13);
            this.lblControlTotalHT.TabIndex = 10;
            this.lblControlTotalHT.Text = "labTotal";
            // 
            // label11
            // 
            this.label11.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(578, 65);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(36, 13);
            this.label11.TabIndex = 31;
            this.label11.Text = "F CFA";
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(578, 47);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(36, 13);
            this.label4.TabIndex = 28;
            this.label4.Text = "F CFA";
            // 
            // lblControlRealiseHT
            // 
            this.lblControlRealiseHT.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblControlRealiseHT.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblControlRealiseHT.Appearance.ForeColor = System.Drawing.Color.DarkGreen;
            this.lblControlRealiseHT.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblControlRealiseHT.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblControlRealiseHT.Location = new System.Drawing.Point(473, 65);
            this.lblControlRealiseHT.Name = "lblControlRealiseHT";
            this.lblControlRealiseHT.Size = new System.Drawing.Size(105, 13);
            this.lblControlRealiseHT.TabIndex = 30;
            this.lblControlRealiseHT.Text = "labTotal";
            // 
            // label10
            // 
            this.label10.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label10.Location = new System.Drawing.Point(426, 64);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(51, 13);
            this.label10.TabIndex = 29;
            this.label10.Text = "Réalisé:";
            // 
            // SMulCom
            // 
            this.SMulCom.Location = new System.Drawing.Point(444, 27);
            this.SMulCom.Name = "SMulCom";
            this.SMulCom.Properties.Caption = "Choix Multiple";
            this.SMulCom.Size = new System.Drawing.Size(88, 19);
            this.SMulCom.TabIndex = 24;
            this.SMulCom.CheckedChanged += new System.EventHandler(this.chkMulCom_CheckedChanged);
            // 
            // chkTousCom
            // 
            this.chkTousCom.EditValue = true;
            this.chkTousCom.Location = new System.Drawing.Point(400, 27);
            this.chkTousCom.Name = "chkTousCom";
            this.chkTousCom.Properties.Caption = "Tous";
            this.chkTousCom.Size = new System.Drawing.Size(48, 19);
            this.chkTousCom.TabIndex = 23;
            this.chkTousCom.CheckStateChanged += new System.EventHandler(this.chkTousCom_CheckStateChanged);
            // 
            // LECommerciauxDE
            // 
            this.LECommerciauxDE.EditValue = "";
            this.LECommerciauxDE.Enabled = false;
            this.LECommerciauxDE.Location = new System.Drawing.Point(35, 26);
            this.LECommerciauxDE.Name = "LECommerciauxDE";
            this.LECommerciauxDE.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LECommerciauxDE.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mNomCommercial", "Nom"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPrenomCommercial", "Prénoms"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LECommerciauxDE.Size = new System.Drawing.Size(171, 20);
            this.LECommerciauxDE.TabIndex = 13;
            this.LECommerciauxDE.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.LECommerciauxDE_Closed);
            this.LECommerciauxDE.TextChanged += new System.EventHandler(this.LECommerciauxDE_TextChanged);
            // 
            // LECommerciauxA
            // 
            this.LECommerciauxA.EditValue = "";
            this.LECommerciauxA.Enabled = false;
            this.LECommerciauxA.Location = new System.Drawing.Point(222, 26);
            this.LECommerciauxA.Name = "LECommerciauxA";
            this.LECommerciauxA.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LECommerciauxA.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mNomCommercial", "Nom"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPrenomCommercial", "Prénoms"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LECommerciauxA.Size = new System.Drawing.Size(171, 20);
            this.LECommerciauxA.TabIndex = 17;
            this.LECommerciauxA.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.LECommerciauxA_Closed);
            this.LECommerciauxA.TextChanged += new System.EventHandler(this.LECommerciauxA_TextChanged);
            // 
            // lblACom
            // 
            this.lblACom.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblACom.AutoSize = true;
            this.lblACom.Location = new System.Drawing.Point(207, 33);
            this.lblACom.Name = "lblACom";
            this.lblACom.Size = new System.Drawing.Size(14, 13);
            this.lblACom.TabIndex = 16;
            this.lblACom.Text = "A";
            // 
            // lblDeCom
            // 
            this.lblDeCom.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblDeCom.AutoSize = true;
            this.lblDeCom.Location = new System.Drawing.Point(9, 33);
            this.lblDeCom.Name = "lblDeCom";
            this.lblDeCom.Size = new System.Drawing.Size(20, 13);
            this.lblDeCom.TabIndex = 15;
            this.lblDeCom.Text = "De";
            // 
            // chkCmbMultCom
            // 
            this.chkCmbMultCom.EditValue = "";
            this.chkCmbMultCom.Location = new System.Drawing.Point(34, 26);
            this.chkCmbMultCom.Name = "chkCmbMultCom";
            this.chkCmbMultCom.Properties.AllowMultiSelect = true;
            this.chkCmbMultCom.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.chkCmbMultCom.Size = new System.Drawing.Size(359, 20);
            this.chkCmbMultCom.TabIndex = 26;
            this.chkCmbMultCom.Visible = false;
            this.chkCmbMultCom.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.chkCmbMultCom_Closed);
            // 
            // gCPeriode
            // 
            this.gCPeriode.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.gCPeriode.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gCPeriode.AppearanceCaption.Options.UseFont = true;
            this.gCPeriode.CaptionImage = ((System.Drawing.Image)(resources.GetObject("gCPeriode.CaptionImage")));
            this.gCPeriode.Controls.Add(this.dateEditDateDeb);
            this.gCPeriode.Controls.Add(this.label2);
            this.gCPeriode.Controls.Add(this.dateEditDateFin);
            this.gCPeriode.Controls.Add(this.label1);
            this.gCPeriode.Location = new System.Drawing.Point(3, -2);
            this.gCPeriode.Name = "gCPeriode";
            this.gCPeriode.Size = new System.Drawing.Size(309, 101);
            this.gCPeriode.TabIndex = 18;
            this.gCPeriode.Text = "Période";
            // 
            // dateEditDateDeb
            // 
            this.dateEditDateDeb.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.dateEditDateDeb.EditValue = null;
            this.dateEditDateDeb.Location = new System.Drawing.Point(26, 48);
            this.dateEditDateDeb.Name = "dateEditDateDeb";
            this.dateEditDateDeb.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dateEditDateDeb.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dateEditDateDeb.Size = new System.Drawing.Size(121, 20);
            this.dateEditDateDeb.TabIndex = 10;
            this.dateEditDateDeb.EditValueChanged += new System.EventHandler(this.dateEditDateDeb_EditValueChanged);
            this.dateEditDateDeb.TextChanged += new System.EventHandler(this.dateEditDateDeb_TextChanged);
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(153, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(20, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Au";
            // 
            // dateEditDateFin
            // 
            this.dateEditDateFin.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.dateEditDateFin.EditValue = null;
            this.dateEditDateFin.Location = new System.Drawing.Point(179, 48);
            this.dateEditDateFin.Name = "dateEditDateFin";
            this.dateEditDateFin.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dateEditDateFin.Properties.CalendarTimeProperties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.dateEditDateFin.Size = new System.Drawing.Size(121, 20);
            this.dateEditDateFin.TabIndex = 12;
            this.dateEditDateFin.EditValueChanged += new System.EventHandler(this.dateEditDateFin_EditValueChanged);
            this.dateEditDateFin.TextChanged += new System.EventHandler(this.dateEditDateFin_TextChanged);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(20, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "Du";
            // 
            // xtraTabPage2
            // 
            this.xtraTabPage2.Controls.Add(this.panelControl3);
            this.xtraTabPage2.Name = "xtraTabPage2";
            this.xtraTabPage2.Size = new System.Drawing.Size(929, 97);
            this.xtraTabPage2.Text = "Famille-Type Document";
            // 
            // panelControl3
            // 
            this.panelControl3.Controls.Add(this.groupControl1);
            this.panelControl3.Controls.Add(this.gCFamille);
            this.panelControl3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelControl3.Location = new System.Drawing.Point(0, 0);
            this.panelControl3.Name = "panelControl3";
            this.panelControl3.Size = new System.Drawing.Size(929, 97);
            this.panelControl3.TabIndex = 0;
            // 
            // groupControl1
            // 
            this.groupControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.groupControl1.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.groupControl1.AppearanceCaption.Options.UseFont = true;
            this.groupControl1.CaptionImage = ((System.Drawing.Image)(resources.GetObject("groupControl1.CaptionImage")));
            this.groupControl1.Controls.Add(this.panTypeDoc);
            this.groupControl1.Controls.Add(this.label6);
            this.groupControl1.Controls.Add(this.label5);
            this.groupControl1.Controls.Add(this.lblControlTotalHT2);
            this.groupControl1.Location = new System.Drawing.Point(707, 0);
            this.groupControl1.Name = "groupControl1";
            this.groupControl1.Size = new System.Drawing.Size(222, 96);
            this.groupControl1.TabIndex = 2;
            this.groupControl1.Text = "Type Document";
            // 
            // panTypeDoc
            // 
            this.panTypeDoc.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panTypeDoc.Controls.Add(this.chkCmbTypeDoc);
            this.panTypeDoc.Controls.Add(this.chkTousTypeDoc);
            this.panTypeDoc.Location = new System.Drawing.Point(10, 24);
            this.panTypeDoc.Name = "panTypeDoc";
            this.panTypeDoc.Size = new System.Drawing.Size(200, 30);
            this.panTypeDoc.TabIndex = 32;
            // 
            // chkCmbTypeDoc
            // 
            this.chkCmbTypeDoc.EditValue = "";
            this.chkCmbTypeDoc.Enabled = false;
            this.chkCmbTypeDoc.Location = new System.Drawing.Point(5, 5);
            this.chkCmbTypeDoc.Name = "chkCmbTypeDoc";
            this.chkCmbTypeDoc.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.chkCmbTypeDoc.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.CheckedListBoxItem[] {
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(8, "Facture (FC et FA)"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(0, "DE"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(1, "BC"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(2, "PL"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(3, "BL"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(4, "BR"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(5, "BA")});
            this.chkCmbTypeDoc.Size = new System.Drawing.Size(131, 20);
            this.chkCmbTypeDoc.TabIndex = 24;
            this.chkCmbTypeDoc.CloseUp += new DevExpress.XtraEditors.Controls.CloseUpEventHandler(this.chkCmbTypeDoc_CloseUp);
            // 
            // chkTousTypeDoc
            // 
            this.chkTousTypeDoc.EditValue = true;
            this.chkTousTypeDoc.Location = new System.Drawing.Point(143, 6);
            this.chkTousTypeDoc.Name = "chkTousTypeDoc";
            this.chkTousTypeDoc.Properties.Caption = "Tous";
            this.chkTousTypeDoc.Size = new System.Drawing.Size(49, 19);
            this.chkTousTypeDoc.TabIndex = 23;
            this.chkTousTypeDoc.CheckStateChanged += new System.EventHandler(this.chkTousTypeDoc_CheckStateChanged);
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label6.Location = new System.Drawing.Point(5, 76);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 13);
            this.label6.TabIndex = 30;
            this.label6.Text = "Montant :";
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(181, 75);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(36, 13);
            this.label5.TabIndex = 31;
            this.label5.Text = "F CFA";
            // 
            // lblControlTotalHT2
            // 
            this.lblControlTotalHT2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblControlTotalHT2.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblControlTotalHT2.Appearance.ForeColor = System.Drawing.Color.Red;
            this.lblControlTotalHT2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblControlTotalHT2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblControlTotalHT2.Location = new System.Drawing.Point(72, 75);
            this.lblControlTotalHT2.Name = "lblControlTotalHT2";
            this.lblControlTotalHT2.Size = new System.Drawing.Size(107, 13);
            this.lblControlTotalHT2.TabIndex = 29;
            this.lblControlTotalHT2.Text = "labTotal";
            // 
            // gCFamille
            // 
            this.gCFamille.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.gCFamille.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gCFamille.AppearanceCaption.Options.UseFont = true;
            this.gCFamille.CaptionImage = ((System.Drawing.Image)(resources.GetObject("gCFamille.CaptionImage")));
            this.gCFamille.Controls.Add(this.panMarque);
            this.gCFamille.Controls.Add(this.chkConfigFamilles);
            this.gCFamille.Controls.Add(this.panFamille);
            this.gCFamille.Controls.Add(this.label7);
            this.gCFamille.Location = new System.Drawing.Point(0, 0);
            this.gCFamille.Name = "gCFamille";
            this.gCFamille.Size = new System.Drawing.Size(706, 96);
            this.gCFamille.TabIndex = 0;
            this.gCFamille.Text = "Marque-Famille";
            // 
            // panMarque
            // 
            this.panMarque.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panMarque.Controls.Add(this.LEFamilleCentralDE);
            this.panMarque.Controls.Add(this.chkCmbMultFamCentral);
            this.panMarque.Controls.Add(this.lblFamCentralDe);
            this.panMarque.Controls.Add(this.lblFamCentralA);
            this.panMarque.Controls.Add(this.SMulFamilleCentral);
            this.panMarque.Controls.Add(this.LEFamilleCentralA);
            this.panMarque.Controls.Add(this.chkTousFamilleCentral);
            this.panMarque.Location = new System.Drawing.Point(132, 23);
            this.panMarque.Name = "panMarque";
            this.panMarque.Size = new System.Drawing.Size(559, 31);
            this.panMarque.TabIndex = 36;
            // 
            // LEFamilleCentralDE
            // 
            this.LEFamilleCentralDE.EditValue = "";
            this.LEFamilleCentralDE.Enabled = false;
            this.LEFamilleCentralDE.Location = new System.Drawing.Point(49, 9);
            this.LEFamilleCentralDE.Name = "LEFamilleCentralDE";
            this.LEFamilleCentralDE.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LEFamilleCentralDE.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mFa_CodeFamille", "Nom"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LEFamilleCentralDE.Size = new System.Drawing.Size(171, 20);
            this.LEFamilleCentralDE.TabIndex = 26;
            this.LEFamilleCentralDE.EditValueChanged += new System.EventHandler(this.LEFamilleCentralDE_EditValueChanged);
            this.LEFamilleCentralDE.TextChanged += new System.EventHandler(this.LEFamilleCentralDE_TextChanged);
            // 
            // chkCmbMultFamCentral
            // 
            this.chkCmbMultFamCentral.EditValue = "";
            this.chkCmbMultFamCentral.Location = new System.Drawing.Point(49, 9);
            this.chkCmbMultFamCentral.Name = "chkCmbMultFamCentral";
            this.chkCmbMultFamCentral.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.chkCmbMultFamCentral.Size = new System.Drawing.Size(359, 20);
            this.chkCmbMultFamCentral.TabIndex = 32;
            this.chkCmbMultFamCentral.Visible = false;
            this.chkCmbMultFamCentral.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.chkCmbMultFamCentral_Closed);
            // 
            // lblFamCentralDe
            // 
            this.lblFamCentralDe.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblFamCentralDe.AutoSize = true;
            this.lblFamCentralDe.Location = new System.Drawing.Point(14, 14);
            this.lblFamCentralDe.Name = "lblFamCentralDe";
            this.lblFamCentralDe.Size = new System.Drawing.Size(20, 13);
            this.lblFamCentralDe.TabIndex = 27;
            this.lblFamCentralDe.Text = "De";
            // 
            // lblFamCentralA
            // 
            this.lblFamCentralA.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblFamCentralA.AutoSize = true;
            this.lblFamCentralA.Location = new System.Drawing.Point(222, 14);
            this.lblFamCentralA.Name = "lblFamCentralA";
            this.lblFamCentralA.Size = new System.Drawing.Size(14, 13);
            this.lblFamCentralA.TabIndex = 28;
            this.lblFamCentralA.Text = "A";
            // 
            // SMulFamilleCentral
            // 
            this.SMulFamilleCentral.Location = new System.Drawing.Point(467, 10);
            this.SMulFamilleCentral.Name = "SMulFamilleCentral";
            this.SMulFamilleCentral.Properties.Caption = "Choix Multiple";
            this.SMulFamilleCentral.Size = new System.Drawing.Size(86, 19);
            this.SMulFamilleCentral.TabIndex = 31;
            this.SMulFamilleCentral.CheckedChanged += new System.EventHandler(this.SMulFamilleCentral_CheckedChanged);
            // 
            // LEFamilleCentralA
            // 
            this.LEFamilleCentralA.EditValue = "";
            this.LEFamilleCentralA.Enabled = false;
            this.LEFamilleCentralA.Location = new System.Drawing.Point(237, 9);
            this.LEFamilleCentralA.Name = "LEFamilleCentralA";
            this.LEFamilleCentralA.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LEFamilleCentralA.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mFa_CodeFamille", "Nom"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LEFamilleCentralA.Size = new System.Drawing.Size(171, 20);
            this.LEFamilleCentralA.TabIndex = 29;
            this.LEFamilleCentralA.EditValueChanged += new System.EventHandler(this.LEFamilleCentralA_EditValueChanged);
            this.LEFamilleCentralA.TextChanged += new System.EventHandler(this.LEFamilleCentralA_TextChanged);
            // 
            // chkTousFamilleCentral
            // 
            this.chkTousFamilleCentral.EditValue = true;
            this.chkTousFamilleCentral.Location = new System.Drawing.Point(414, 10);
            this.chkTousFamilleCentral.Name = "chkTousFamilleCentral";
            this.chkTousFamilleCentral.Properties.Caption = "Tous";
            this.chkTousFamilleCentral.Size = new System.Drawing.Size(48, 19);
            this.chkTousFamilleCentral.TabIndex = 30;
            this.chkTousFamilleCentral.CheckedChanged += new System.EventHandler(this.chkTousFamilleCentral_CheckedChanged);
            // 
            // chkConfigFamilles
            // 
            this.chkConfigFamilles.Location = new System.Drawing.Point(14, 69);
            this.chkConfigFamilles.Name = "chkConfigFamilles";
            this.chkConfigFamilles.Properties.Caption = "Configurer Familles";
            this.chkConfigFamilles.Size = new System.Drawing.Size(114, 19);
            this.chkConfigFamilles.TabIndex = 35;
            this.chkConfigFamilles.CheckedChanged += new System.EventHandler(this.chkConfigFamilles_CheckedChanged);
            // 
            // panFamille
            // 
            this.panFamille.BorderStyle = DevExpress.XtraEditors.Controls.BorderStyles.NoBorder;
            this.panFamille.Controls.Add(this.lblFamDe);
            this.panFamille.Controls.Add(this.lblFamA);
            this.panFamille.Controls.Add(this.LEFamilleA);
            this.panFamille.Controls.Add(this.LEFamilleDE);
            this.panFamille.Controls.Add(this.chkTousFamille);
            this.panFamille.Controls.Add(this.SMulFamille);
            this.panFamille.Controls.Add(this.chkCmbMultFam);
            this.panFamille.Location = new System.Drawing.Point(132, 59);
            this.panFamille.Name = "panFamille";
            this.panFamille.Size = new System.Drawing.Size(559, 35);
            this.panFamille.TabIndex = 34;
            this.panFamille.Visible = false;
            // 
            // lblFamDe
            // 
            this.lblFamDe.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblFamDe.AutoSize = true;
            this.lblFamDe.Location = new System.Drawing.Point(14, 13);
            this.lblFamDe.Name = "lblFamDe";
            this.lblFamDe.Size = new System.Drawing.Size(20, 13);
            this.lblFamDe.TabIndex = 19;
            this.lblFamDe.Text = "De";
            // 
            // lblFamA
            // 
            this.lblFamA.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblFamA.AutoSize = true;
            this.lblFamA.Location = new System.Drawing.Point(222, 13);
            this.lblFamA.Name = "lblFamA";
            this.lblFamA.Size = new System.Drawing.Size(14, 13);
            this.lblFamA.TabIndex = 20;
            this.lblFamA.Text = "A";
            // 
            // LEFamilleA
            // 
            this.LEFamilleA.EditValue = "";
            this.LEFamilleA.Enabled = false;
            this.LEFamilleA.Location = new System.Drawing.Point(237, 8);
            this.LEFamilleA.Name = "LEFamilleA";
            this.LEFamilleA.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LEFamilleA.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mFa_CodeFamille", "Nom"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LEFamilleA.Size = new System.Drawing.Size(171, 20);
            this.LEFamilleA.TabIndex = 21;
            this.LEFamilleA.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.LEFamilleA_Closed);
            this.LEFamilleA.TextChanged += new System.EventHandler(this.LEFamilleA_TextChanged);
            // 
            // LEFamilleDE
            // 
            this.LEFamilleDE.EditValue = "";
            this.LEFamilleDE.Enabled = false;
            this.LEFamilleDE.Location = new System.Drawing.Point(49, 8);
            this.LEFamilleDE.Name = "LEFamilleDE";
            this.LEFamilleDE.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LEFamilleDE.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mFa_CodeFamille", "Nom"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LEFamilleDE.Size = new System.Drawing.Size(171, 20);
            this.LEFamilleDE.TabIndex = 18;
            this.LEFamilleDE.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.LEFamilleDE_Closed);
            this.LEFamilleDE.TextChanged += new System.EventHandler(this.LEFamilleDE_TextChanged);
            // 
            // chkTousFamille
            // 
            this.chkTousFamille.EditValue = true;
            this.chkTousFamille.Location = new System.Drawing.Point(414, 9);
            this.chkTousFamille.Name = "chkTousFamille";
            this.chkTousFamille.Properties.Caption = "Tous";
            this.chkTousFamille.Size = new System.Drawing.Size(48, 19);
            this.chkTousFamille.TabIndex = 22;
            this.chkTousFamille.CheckedChanged += new System.EventHandler(this.chkTousFamille_CheckedChanged);
            // 
            // SMulFamille
            // 
            this.SMulFamille.Location = new System.Drawing.Point(467, 9);
            this.SMulFamille.Name = "SMulFamille";
            this.SMulFamille.Properties.Caption = "Choix Multiple";
            this.SMulFamille.Size = new System.Drawing.Size(86, 19);
            this.SMulFamille.TabIndex = 23;
            this.SMulFamille.CheckedChanged += new System.EventHandler(this.chkMulFamille_CheckedChanged);
            // 
            // chkCmbMultFam
            // 
            this.chkCmbMultFam.EditValue = "";
            this.chkCmbMultFam.Location = new System.Drawing.Point(49, 8);
            this.chkCmbMultFam.Name = "chkCmbMultFam";
            this.chkCmbMultFam.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.chkCmbMultFam.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.CheckedListBoxItem[] {
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(0, "DE"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(1, "BC"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(2, "PL"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(3, "BL"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(4, "BR"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(5, "BA"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(6, "FAC"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(7, "FC")});
            this.chkCmbMultFam.Size = new System.Drawing.Size(359, 20);
            this.chkCmbMultFam.TabIndex = 25;
            this.chkCmbMultFam.Visible = false;
            this.chkCmbMultFam.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.chkCmbMultFam_Closed);
            // 
            // label7
            // 
            this.label7.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(11, 34);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 33;
            this.label7.Text = "Marque";
            // 
            // xtraTabPageRevendeurs
            // 
            this.xtraTabPageRevendeurs.Controls.Add(this.panelControl4);
            this.xtraTabPageRevendeurs.Name = "xtraTabPageRevendeurs";
            this.xtraTabPageRevendeurs.Size = new System.Drawing.Size(929, 97);
            this.xtraTabPageRevendeurs.Text = "Revendeurs";
            // 
            // panelControl4
            // 
            this.panelControl4.Controls.Add(this.gCRevendeur);
            this.panelControl4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelControl4.Location = new System.Drawing.Point(0, 0);
            this.panelControl4.Name = "panelControl4";
            this.panelControl4.Size = new System.Drawing.Size(929, 97);
            this.panelControl4.TabIndex = 0;
            // 
            // gCRevendeur
            // 
            this.gCRevendeur.AppearanceCaption.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.gCRevendeur.AppearanceCaption.Options.UseFont = true;
            this.gCRevendeur.CaptionImage = ((System.Drawing.Image)(resources.GetObject("gCRevendeur.CaptionImage")));
            this.gCRevendeur.Controls.Add(this.lblAccurency2);
            this.gCRevendeur.Controls.Add(this.label14);
            this.gCRevendeur.Controls.Add(this.label15);
            this.gCRevendeur.Controls.Add(this.label16);
            this.gCRevendeur.Controls.Add(this.lblControlRealiseHT2);
            this.gCRevendeur.Controls.Add(this.label17);
            this.gCRevendeur.Controls.Add(this.label8);
            this.gCRevendeur.Controls.Add(this.label9);
            this.gCRevendeur.Controls.Add(this.lblControlTotalHT3);
            this.gCRevendeur.Controls.Add(this.SMulRevendeur);
            this.gCRevendeur.Controls.Add(this.chkTousRevendeur);
            this.gCRevendeur.Controls.Add(this.LERevendeurDE);
            this.gCRevendeur.Controls.Add(this.LERevendeurA);
            this.gCRevendeur.Controls.Add(this.lblRevA);
            this.gCRevendeur.Controls.Add(this.lblRevDe);
            this.gCRevendeur.Controls.Add(this.chkCmbMultRev);
            this.gCRevendeur.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gCRevendeur.Location = new System.Drawing.Point(2, 2);
            this.gCRevendeur.Name = "gCRevendeur";
            this.gCRevendeur.Size = new System.Drawing.Size(925, 93);
            this.gCRevendeur.TabIndex = 0;
            this.gCRevendeur.Text = "Revendeurs";
            // 
            // lblAccurency2
            // 
            this.lblAccurency2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblAccurency2.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblAccurency2.Appearance.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblAccurency2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblAccurency2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblAccurency2.Location = new System.Drawing.Point(778, 69);
            this.lblAccurency2.Name = "lblAccurency2";
            this.lblAccurency2.Size = new System.Drawing.Size(105, 13);
            this.lblAccurency2.TabIndex = 41;
            this.lblAccurency2.Text = "labTotal";
            // 
            // label14
            // 
            this.label14.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label14.Location = new System.Drawing.Point(732, 69);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(46, 13);
            this.label14.TabIndex = 40;
            this.label14.Text = "Marge:";
            // 
            // label15
            // 
            this.label15.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(884, 69);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(36, 13);
            this.label15.TabIndex = 39;
            this.label15.Text = "F CFA";
            // 
            // label16
            // 
            this.label16.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(884, 52);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(36, 13);
            this.label16.TabIndex = 38;
            this.label16.Text = "F CFA";
            // 
            // lblControlRealiseHT2
            // 
            this.lblControlRealiseHT2.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblControlRealiseHT2.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblControlRealiseHT2.Appearance.ForeColor = System.Drawing.Color.DarkGreen;
            this.lblControlRealiseHT2.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblControlRealiseHT2.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblControlRealiseHT2.Location = new System.Drawing.Point(788, 52);
            this.lblControlRealiseHT2.Name = "lblControlRealiseHT2";
            this.lblControlRealiseHT2.Size = new System.Drawing.Size(95, 13);
            this.lblControlRealiseHT2.TabIndex = 37;
            this.lblControlRealiseHT2.Text = "labTotal";
            // 
            // label17
            // 
            this.label17.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label17.Location = new System.Drawing.Point(727, 51);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(51, 13);
            this.label17.TabIndex = 36;
            this.label17.Text = "Réalisé:";
            // 
            // label8
            // 
            this.label8.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold);
            this.label8.Location = new System.Drawing.Point(693, 32);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(85, 13);
            this.label8.TabIndex = 34;
            this.label8.Text = "Montant Pipe:";
            // 
            // label9
            // 
            this.label9.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(885, 31);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(36, 13);
            this.label9.TabIndex = 35;
            this.label9.Text = "F CFA";
            // 
            // lblControlTotalHT3
            // 
            this.lblControlTotalHT3.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.lblControlTotalHT3.Appearance.Font = new System.Drawing.Font("Tahoma", 10F, System.Drawing.FontStyle.Bold);
            this.lblControlTotalHT3.Appearance.ForeColor = System.Drawing.Color.Red;
            this.lblControlTotalHT3.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
            this.lblControlTotalHT3.AutoSizeMode = DevExpress.XtraEditors.LabelAutoSizeMode.None;
            this.lblControlTotalHT3.Location = new System.Drawing.Point(788, 31);
            this.lblControlTotalHT3.Name = "lblControlTotalHT3";
            this.lblControlTotalHT3.Size = new System.Drawing.Size(95, 13);
            this.lblControlTotalHT3.TabIndex = 33;
            this.lblControlTotalHT3.Text = "labTotal";
            // 
            // SMulRevendeur
            // 
            this.SMulRevendeur.Location = new System.Drawing.Point(629, 42);
            this.SMulRevendeur.Name = "SMulRevendeur";
            this.SMulRevendeur.Properties.Caption = "Choix Multiple";
            this.SMulRevendeur.Size = new System.Drawing.Size(86, 19);
            this.SMulRevendeur.TabIndex = 31;
            this.SMulRevendeur.CheckedChanged += new System.EventHandler(this.SMulRevendeur_CheckedChanged);
            // 
            // chkTousRevendeur
            // 
            this.chkTousRevendeur.EditValue = true;
            this.chkTousRevendeur.Location = new System.Drawing.Point(576, 42);
            this.chkTousRevendeur.Name = "chkTousRevendeur";
            this.chkTousRevendeur.Properties.Caption = "Tous";
            this.chkTousRevendeur.Size = new System.Drawing.Size(48, 19);
            this.chkTousRevendeur.TabIndex = 30;
            this.chkTousRevendeur.CheckedChanged += new System.EventHandler(this.chkTousRevendeur_CheckedChanged);
            // 
            // LERevendeurDE
            // 
            this.LERevendeurDE.EditValue = "";
            this.LERevendeurDE.Enabled = false;
            this.LERevendeurDE.Location = new System.Drawing.Point(52, 41);
            this.LERevendeurDE.Name = "LERevendeurDE";
            this.LERevendeurDE.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LERevendeurDE.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mCT_Num", 30, "Code Revendeur"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mCT_Intitule", 30, "Nom Revendeur"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LERevendeurDE.Size = new System.Drawing.Size(248, 20);
            this.LERevendeurDE.TabIndex = 26;
            this.LERevendeurDE.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.LERevendeurDE_Closed);
            // 
            // LERevendeurA
            // 
            this.LERevendeurA.EditValue = "";
            this.LERevendeurA.Enabled = false;
            this.LERevendeurA.Location = new System.Drawing.Point(322, 41);
            this.LERevendeurA.Name = "LERevendeurA";
            this.LERevendeurA.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.LERevendeurA.Properties.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mCT_Num", 30, "Code Revendeur"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mCT_Intitule", 30, "Nom Revendeur"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("mPays", "Pays")});
            this.LERevendeurA.Size = new System.Drawing.Size(248, 20);
            this.LERevendeurA.TabIndex = 29;
            this.LERevendeurA.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.LERevendeurA_Closed);
            // 
            // lblRevA
            // 
            this.lblRevA.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblRevA.AutoSize = true;
            this.lblRevA.Location = new System.Drawing.Point(304, 49);
            this.lblRevA.Name = "lblRevA";
            this.lblRevA.Size = new System.Drawing.Size(14, 13);
            this.lblRevA.TabIndex = 28;
            this.lblRevA.Text = "A";
            // 
            // lblRevDe
            // 
            this.lblRevDe.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.lblRevDe.AutoSize = true;
            this.lblRevDe.Location = new System.Drawing.Point(4, 49);
            this.lblRevDe.Name = "lblRevDe";
            this.lblRevDe.Size = new System.Drawing.Size(20, 13);
            this.lblRevDe.TabIndex = 27;
            this.lblRevDe.Text = "De";
            // 
            // chkCmbMultRev
            // 
            this.chkCmbMultRev.EditValue = "";
            this.chkCmbMultRev.Location = new System.Drawing.Point(52, 41);
            this.chkCmbMultRev.Name = "chkCmbMultRev";
            this.chkCmbMultRev.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.chkCmbMultRev.Properties.Items.AddRange(new DevExpress.XtraEditors.Controls.CheckedListBoxItem[] {
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(0, "DE"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(1, "BC"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(2, "PL"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(3, "BL"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(4, "BR"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(5, "BA"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(6, "FAC"),
            new DevExpress.XtraEditors.Controls.CheckedListBoxItem(7, "FC")});
            this.chkCmbMultRev.Size = new System.Drawing.Size(518, 20);
            this.chkCmbMultRev.TabIndex = 32;
            this.chkCmbMultRev.Visible = false;
            this.chkCmbMultRev.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.chkCmbMultRev_Closed);
            // 
            // sFDExcel
            // 
            this.sFDExcel.Filter = "Fichiers (*.xlsx) | *.xlsx";
            // 
            // splashScreenManager2
            // 
            this.splashScreenManager2.ClosingDelay = 500;
            // 
            // timer1
            // 
            this.timer1.Interval = 300;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(944, 401);
            this.Controls.Add(this.layoutControl1);
            this.Controls.Add(this.menuStrip1);
            this.Controls.Add(this.xtraTabControl1);
            this.Controls.Add(this.panelControl2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ForecastCOM";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.groupControl2)).EndInit();
            this.groupControl2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureEdit1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ChkListBoxControlFiliale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridControlForecast)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.panelControl2)).EndInit();
            this.panelControl2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.xtraTabControl1)).EndInit();
            this.xtraTabControl1.ResumeLayout(false);
            this.xtraTabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl1)).EndInit();
            this.panelControl1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chkFacture.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkSaisieDevis.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkAccepteDevis.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.groupControl5)).EndInit();
            this.groupControl5.ResumeLayout(false);
            this.groupControl5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SMulCom.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousCom.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LECommerciauxDE.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LECommerciauxA.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultCom.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gCPeriode)).EndInit();
            this.gCPeriode.ResumeLayout(false);
            this.gCPeriode.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateDeb.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateDeb.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateFin.Properties.CalendarTimeProperties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dateEditDateFin.Properties)).EndInit();
            this.xtraTabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl3)).EndInit();
            this.panelControl3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.groupControl1)).EndInit();
            this.groupControl1.ResumeLayout(false);
            this.groupControl1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panTypeDoc)).EndInit();
            this.panTypeDoc.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbTypeDoc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousTypeDoc.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gCFamille)).EndInit();
            this.gCFamille.ResumeLayout(false);
            this.gCFamille.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panMarque)).EndInit();
            this.panMarque.ResumeLayout(false);
            this.panMarque.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleCentralDE.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultFamCentral.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.SMulFamilleCentral.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleCentralA.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousFamilleCentral.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkConfigFamilles.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.panFamille)).EndInit();
            this.panFamille.ResumeLayout(false);
            this.panFamille.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleA.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LEFamilleDE.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousFamille.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.SMulFamille.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultFam.Properties)).EndInit();
            this.xtraTabPageRevendeurs.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelControl4)).EndInit();
            this.panelControl4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gCRevendeur)).EndInit();
            this.gCRevendeur.ResumeLayout(false);
            this.gCRevendeur.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SMulRevendeur.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkTousRevendeur.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LERevendeurDE.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LERevendeurA.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chkCmbMultRev.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem connexionsToolStripMenuItem;
        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraEditors.GroupControl groupControl2;
        private DevExpress.XtraGrid.GridControl gridControlForecast;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private DevExpress.XtraEditors.SimpleButton sBtnGenererExcel;
        private DevExpress.XtraEditors.SimpleButton sBtnVisualiserArticle;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem2;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem3;
        private DevExpress.XtraEditors.DateEdit dateEditDateFin;
        private System.Windows.Forms.Label label2;
        private DevExpress.XtraEditors.DateEdit dateEditDateDeb;
        private System.Windows.Forms.Label label1;
        private DevExpress.XtraEditors.CheckedListBoxControl ChkListBoxControlFiliale;
        private DevExpress.XtraGrid.Columns.GridColumn colNom;
        private DevExpress.XtraGrid.Columns.GridColumn colPrenom;
        private DevExpress.XtraGrid.Columns.GridColumn colFamille;
        private DevExpress.XtraGrid.Columns.GridColumn colTypeDoc;
        private DevExpress.XtraGrid.Columns.GridColumn colDate;
        private DevExpress.XtraGrid.Columns.GridColumn colMontantHT;
        private DevExpress.XtraGrid.Columns.GridColumn colMarge;
        private DevExpress.XtraGrid.Columns.GridColumn colPiece;
        private DevExpress.XtraGrid.Columns.GridColumn colRevendeur;
        private DevExpress.XtraGrid.Columns.GridColumn colPays;
        private System.Windows.Forms.SaveFileDialog sFDExcel;
        private DevExpress.XtraSplashScreen.SplashScreenManager splashScreenManager1;
        private DevExpress.XtraEditors.PictureEdit pictureEdit1;
        private DevExpress.XtraEditors.LookUpEdit LECommerciauxDE;
        private DevExpress.XtraGrid.Columns.GridColumn colStatutDoc;
        private DevExpress.XtraGrid.Columns.GridColumn colDateLivraison;
        private PanelControl panelControl2;
        private XtraTabControl xtraTabControl1;
        private XtraTabPage xtraTabPage1;
        private PanelControl panelControl1;
        private System.Windows.Forms.Label lblDeCom;
        private XtraTabPage xtraTabPage2;
        private GroupControl gCPeriode;
        private LookUpEdit LECommerciauxA;
        private System.Windows.Forms.Label lblACom;
        private PanelControl panelControl3;
        private GroupControl gCFamille;
        private GroupControl groupControl5;
        private GroupControl groupControl1;
        private LookUpEdit LEFamilleDE;
        private LookUpEdit LEFamilleA;
        private System.Windows.Forms.Label lblFamA;
        private System.Windows.Forms.Label lblFamDe;
        private CheckEdit chkTousTypeDoc;
        private CheckEdit chkTousFamille;
        private CheckEdit SMulCom;
        private CheckEdit chkTousCom;
        private CheckedComboBoxEdit chkCmbTypeDoc;
        private CheckEdit SMulFamille;
        private CheckedComboBoxEdit chkCmbMultCom;
        private CheckedComboBoxEdit chkCmbMultFam;
        private CheckEdit chkAccepteDevis;
        private CheckEdit chkSaisieDevis;
        private System.Windows.Forms.GroupBox groupBox1;
        private LabelControl lblControlTotalHT;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private DevExpress.XtraSplashScreen.SplashScreenManager splashScreenManager2;
        private System.Windows.Forms.Timer timer1;
        private DevExpress.XtraGrid.Columns.GridColumn colReference;
        private SimpleButton BtnStat;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private LabelControl lblControlTotalHT2;
        private System.Windows.Forms.ToolStripMenuItem gestionTargetToolStripMenuItem;
        private XtraTabPage xtraTabPageRevendeurs;
        private PanelControl panelControl4;
        private GroupControl gCRevendeur;
        private CheckEdit SMulRevendeur;
        private CheckEdit chkTousRevendeur;
        private LookUpEdit LERevendeurDE;
        private LookUpEdit LERevendeurA;
        private System.Windows.Forms.Label lblRevA;
        private System.Windows.Forms.Label lblRevDe;
        private CheckedComboBoxEdit chkCmbMultRev;
        private SimpleButton sBtnStatRev;
        private System.Windows.Forms.ToolStripMenuItem menuToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem connexionsToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem AdminToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem connexionToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem gestionTargetToolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem gestionBUToolStripMenuItem;
        private SimpleButton sBtnAnalyseComp;
        private CheckEdit SMulFamilleCentral;
        private CheckEdit chkTousFamilleCentral;
        private LookUpEdit LEFamilleCentralDE;
        private LookUpEdit LEFamilleCentralA;
        private System.Windows.Forms.Label lblFamCentralA;
        private System.Windows.Forms.Label lblFamCentralDe;
        private CheckedComboBoxEdit chkCmbMultFamCentral;
        private CheckEdit chkConfigFamilles;
        private PanelControl panFamille;
        private System.Windows.Forms.Label label7;
        private DevExpress.XtraGrid.Columns.GridColumn colFamilleCentr;
        private CheckEdit chkFacture;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private LabelControl lblControlTotalHT3;
        private System.Windows.Forms.Label label11;
        private LabelControl lblControlRealiseHT;
        private System.Windows.Forms.Label label10;
        private LabelControl lblAccurency;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private LabelControl lblAccurency2;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private LabelControl lblControlRealiseHT2;
        private System.Windows.Forms.Label label17;
        private PanelControl panMarque;
        private PanelControl panTypeDoc;
        private DevExpress.XtraGrid.Columns.GridColumn colPaysRevendeur;
        private DevExpress.XtraGrid.Columns.GridColumn colmDO_Coord01;
    }
}

