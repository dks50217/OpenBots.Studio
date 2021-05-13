﻿using Microsoft.Office.Interop.Excel;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Commands.Microsoft.Library;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Threading.Tasks;
using OpenBots.Core.Model.ApplicationModel;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command deletes a specific cell or range in an Excel Worksheet.")]
	public class ExcelDeleteRangeCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Excel Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Application** command.")]
		[SampleUsage("MyExcelInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Application** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("Range")]
		[Description("Enter the location of the cell or range to delete.")]
		[SampleUsage("\"A1\" || \"A1:B10\" || \"A1:\" || vRange || vStart + \":\" + vEnd || vStart + \":\"")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_Range { get; set; }

		[Required]
		[DisplayName("Shift Cells Up")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("'Yes' removes the entire range. 'No' only clears the range of its cell values.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_ShiftUp { get; set; }

		public ExcelDeleteRangeCommand()
		{
			CommandName = "ExcelDeleteRangeCommand";
			SelectionName = "Delete Range";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_InstanceName = "DefaultExcel";
			v_ShiftUp = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;
			Worksheet excelSheet = excelInstance.ActiveSheet;

			string vRange = (string)await v_Range.EvaluateCode(engine);
			var splitRange = vRange.Split(':');
			Range cellRange;
            Range sourceRange = excelSheet.UsedRange;

			//Delete a range of cells
			try
			{
				var last = excelInstance.GetLastIndexOfNonEmptyCell(sourceRange, sourceRange.Range["A1"]);
				if (splitRange[1] == "")
					cellRange = excelSheet.Range[splitRange[0], last];
				else
					cellRange = excelSheet.Range[splitRange[0], splitRange[1]];
			}
			//Delete a cell
			catch (Exception)
			{
				cellRange = excelSheet.Range[splitRange[0], Type.Missing];
			}

			if (v_ShiftUp == "Yes")
				cellRange.Delete();          
			else
				cellRange.Clear();
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Range", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_ShiftUp", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Delete Cells '{v_Range}' - Instance Name '{v_InstanceName}']";
		}
	}
}