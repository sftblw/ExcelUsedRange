using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelUsedRange
{
	public partial class Ribbon
	{
		private void Ribbon_Load(object sender, RibbonUIEventArgs e)
		{

		}

		private void SelectAsUsedRange_Click(object sender, RibbonControlEventArgs e)
		{
			Excel.Worksheet sheet = ThisAddIn.ThisSheet;
			if (sheet == null) { return; }

			var range = sheet.Range[sheet.Cells[1, 1], sheet.UsedRange] as Excel.Range;
			range.Select();
		}
	}
}
