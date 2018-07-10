namespace ExcelUsedRange
{
	partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// 필수 디자이너 변수입니다.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public Ribbon()
			: base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();

			
		}

		/// <summary> 
		/// 사용 중인 모든 리소스를 정리합니다.
		/// </summary>
		/// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region 구성 요소 디자이너에서 생성한 코드

		/// <summary>
		/// 디자이너 지원에 필요한 메서드입니다. 
		/// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
		/// </summary>
		private void InitializeComponent()
		{
			this.ExcelUsedRangeTab = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.UsedRangeLabel = this.Factory.CreateRibbonLabel();
			this.UsedRangeRowCol = this.Factory.CreateRibbonLabel();
			this.UsedRangeE4 = this.Factory.CreateRibbonLabel();
			this.UsedRangeFromBeginLabel = this.Factory.CreateRibbonLabel();
			this.group2 = this.Factory.CreateRibbonGroup();
			this.group3 = this.Factory.CreateRibbonGroup();
			this.SelectionLabel = this.Factory.CreateRibbonLabel();
			this.SelectAsUsedRange = this.Factory.CreateRibbonButton();
			this.UsedRangeFromBeginRowCol = this.Factory.CreateRibbonLabel();
			this.ExcelUsedRangeTab.SuspendLayout();
			this.group1.SuspendLayout();
			this.group2.SuspendLayout();
			this.group3.SuspendLayout();
			this.SuspendLayout();
			// 
			// ExcelUsedRangeTab
			// 
			this.ExcelUsedRangeTab.Groups.Add(this.group1);
			this.ExcelUsedRangeTab.Groups.Add(this.group2);
			this.ExcelUsedRangeTab.Groups.Add(this.group3);
			this.ExcelUsedRangeTab.Label = "Used range";
			this.ExcelUsedRangeTab.Name = "ExcelUsedRangeTab";
			// 
			// group1
			// 
			this.group1.Items.Add(this.UsedRangeLabel);
			this.group1.Items.Add(this.UsedRangeRowCol);
			this.group1.Items.Add(this.UsedRangeE4);
			this.group1.Label = "Used range display";
			this.group1.Name = "group1";
			// 
			// UsedRangeLabel
			// 
			this.UsedRangeLabel.Label = "UsedRangeLabel";
			this.UsedRangeLabel.Name = "UsedRangeLabel";
			// 
			// UsedRangeRowCol
			// 
			this.UsedRangeRowCol.Label = "UsedRangeRowCol";
			this.UsedRangeRowCol.Name = "UsedRangeRowCol";
			// 
			// UsedRangeE4
			// 
			this.UsedRangeE4.Label = "UsedRange Excel4Macro";
			this.UsedRangeE4.Name = "UsedRangeE4";
			// 
			// UsedRangeFromBeginLabel
			// 
			this.UsedRangeFromBeginLabel.Label = "UsedRangeFromBeginLabel";
			this.UsedRangeFromBeginLabel.Name = "UsedRangeFromBeginLabel";
			// 
			// group2
			// 
			this.group2.Items.Add(this.UsedRangeFromBeginLabel);
			this.group2.Items.Add(this.UsedRangeFromBeginRowCol);
			this.group2.Items.Add(this.SelectAsUsedRange);
			this.group2.Label = "Used range from A1";
			this.group2.Name = "group2";
			// 
			// group3
			// 
			this.group3.Items.Add(this.SelectionLabel);
			this.group3.Label = "Current Selection";
			this.group3.Name = "group3";
			// 
			// SelectionLabel
			// 
			this.SelectionLabel.Label = "SelectionLabel";
			this.SelectionLabel.Name = "SelectionLabel";
			// 
			// SelectAsUsedRange
			// 
			this.SelectAsUsedRange.Label = "SelectAsUsedRange";
			this.SelectAsUsedRange.Name = "SelectAsUsedRange";
			this.SelectAsUsedRange.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SelectAsUsedRange_Click);
			// 
			// UsedRangeFromBeginRowCol
			// 
			this.UsedRangeFromBeginRowCol.Label = "UsedRangeFromBeginRowCol";
			this.UsedRangeFromBeginRowCol.Name = "UsedRangeFromBeginRowCol";
			// 
			// Ribbon
			// 
			this.Name = "Ribbon";
			this.RibbonType = "Microsoft.Excel.Workbook";
			this.Tabs.Add(this.ExcelUsedRangeTab);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
			this.ExcelUsedRangeTab.ResumeLayout(false);
			this.ExcelUsedRangeTab.PerformLayout();
			this.group1.ResumeLayout(false);
			this.group1.PerformLayout();
			this.group2.ResumeLayout(false);
			this.group2.PerformLayout();
			this.group3.ResumeLayout(false);
			this.group3.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab ExcelUsedRangeTab;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel UsedRangeLabel;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel UsedRangeRowCol;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel UsedRangeE4;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel UsedRangeFromBeginLabel;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel SelectionLabel;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton SelectAsUsedRange;
		internal Microsoft.Office.Tools.Ribbon.RibbonLabel UsedRangeFromBeginRowCol;
	}

	partial class ThisRibbonCollection
	{
		internal Ribbon Ribbon
		{
			get { return this.GetRibbon<Ribbon>(); }
		}
	}
}
