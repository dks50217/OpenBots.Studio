﻿using Microsoft.Office.Interop.Excel;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command deletes a specific row in an Excel Worksheet.")]

	public class ExcelDeleteRowCommand : ScriptCommand
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
		[DisplayName("Row Number")]
		[Description("Enter the number of the row to be deleted.")]
		[SampleUsage("1 || vRowNumber")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_RowNumber { get; set; }

		[Required]
		[DisplayName("Shift Cells Up")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("'Yes' removes the entire row. 'No' only clears the row of its cell values.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_ShiftUp { get; set; }

		public ExcelDeleteRowCommand()
		{
			CommandName = "ExcelDeleteRowCommand";
			SelectionName = "Delete Row";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_ShiftUp = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;
			Worksheet workSheet = excelInstance.ActiveSheet;
			var vRowToDelete = (int)await v_RowNumber.EvaluateCode(engine);

			var cells = workSheet.Range["A" + vRowToDelete.ToString(), Type.Missing];
			var entireRow = cells.EntireRow;
			if (v_ShiftUp == "Yes")            
				entireRow.Delete();            
			else            
				entireRow.Clear();      
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_RowNumber", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_ShiftUp", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Delete Row '{v_RowNumber}' - Instance Name '{v_InstanceName}']";
		}
	}
}