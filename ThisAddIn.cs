using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelUsedRange
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
			Globals.ThisAddIn.Application.SheetSelectionChange += Application_SheetSelectionChange;
		}

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

		private void Application_SheetSelectionChange(object Sh, Microsoft.Office.Interop.Excel.Range Target)
		{
			Excel.Worksheet sheet = ThisSheet;
			if (sheet == null) { return; }

			Globals.Ribbons.Ribbon.UsedRangeLabel.Label = "Raw UsedRange: " + sheet.UsedRange.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
			Globals.Ribbons.Ribbon.UsedRangeRowCol.Label = $"Raw UsedRange: rows: {sheet.UsedRange.Rows.Count}, cols: {sheet.UsedRange.Columns.Count}";

			int E4usedRow = (int)Globals.ThisAddIn.Application.ExecuteExcel4Macro("Get.Document(10)");
			int E4usedCol = (int)Globals.ThisAddIn.Application.ExecuteExcel4Macro("Get.Document(12)");
			Globals.Ribbons.Ribbon.UsedRangeE4.Label = $"Excel 4 Macro rows: {E4usedRow}, cols: {E4usedCol}";

			var fromA1Range = sheet.Range[sheet.Cells[1, 1], sheet.UsedRange] as Excel.Range;
			Globals.Ribbons.Ribbon.UsedRangeFromBeginLabel.Label = "From A1: " + fromA1Range.Address[true, true, Excel.XlReferenceStyle.xlA1, true] + ", Area count:" + fromA1Range.Areas.Count;
			Globals.Ribbons.Ribbon.UsedRangeFromBeginRowCol.Label = $"From A1: rows: {fromA1Range.Rows.Count}, cols: {fromA1Range.Columns.Count}";

			Globals.Ribbons.Ribbon.SelectionLabel.Label = "Selection: " + Target.Address[true, true, Excel.XlReferenceStyle.xlA1, true] + ", Area count:" + Target.Areas.Count;
		}

		public static Excel.Worksheet ThisSheet
		{
			get
			{
				dynamic shit = Globals.ThisAddIn.Application.ActiveSheet;
				if (!(shit is Microsoft.Office.Interop.Excel.Worksheet)) { return null; }

				var sheet = shit as Microsoft.Office.Interop.Excel.Worksheet;
				return sheet;
			}
		}

		#region VSTO에서 생성한 코드

		/// <summary>
		/// 디자이너 지원에 필요한 메서드입니다. 
		/// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
		/// </summary>
		private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
