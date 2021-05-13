﻿using Microsoft.Office.Interop.Excel;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command writes a DataRow to an Excel Worksheet starting from a specific cell address.")]
	public class ExcelWriteRowCommand : ScriptCommand
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
		[DisplayName("Row")]
		[Description("Enter the row value to set at the selected cell.")]
		[SampleUsage("new List<string>() { \"Hello\", \"World\" } || vList || vDataRow")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(DataRow), typeof(List<string>) })]
		public string v_RowToSet { get; set; }

		[Required]
		[DisplayName("Cell Location")]
		[Description("Enter the location of the cell to write the row to.")]
		[SampleUsage("\"A1\" || vCellLocation")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_CellLocation { get; set; }

		public ExcelWriteRowCommand()
		{
			CommandName = "ExcelWriteRowCommand";
			SelectionName = "Write Row";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_InstanceName = "DefaultExcel";
			v_CellLocation = "A1";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;          
			var vTargetAddress = (string)await v_CellLocation.EvaluateCode(engine);
			dynamic vRow = await v_RowToSet.EvaluateCode(engine);

			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var excelInstance = (Application)excelObject;
			var excelSheet = (Worksheet)excelInstance.ActiveSheet;
			
			if (string.IsNullOrEmpty(vTargetAddress)) 
				throw new ArgumentNullException("columnName");

			var numberOfRow = int.Parse(Regex.Match(vTargetAddress, @"\d+").Value);
			vTargetAddress = Regex.Replace(vTargetAddress, @"[\d-]", string.Empty);
			vTargetAddress = vTargetAddress.ToUpperInvariant();

			int sum = 0;
			for (int i = 0; i < vTargetAddress.Length; i++)
			{
				sum *= 26;
				sum += (vTargetAddress[i] - 'A' + 1);
			}

			//Write row
			DataRow row;
			if (vRow != null && vRow is DataRow)
			{
				row = (DataRow)vRow;

				string cellValue;
				for (int j = 0; j < row.ItemArray.Length; j++)
				{
					if (row.ItemArray[j] == null)
						cellValue = string.Empty;
					else
						cellValue = row.ItemArray[j].ToString();

					excelSheet.Cells[numberOfRow, j + sum] = cellValue;
				}
			}
			else
			{
				var vRowList = (List<string>)vRow;
				string cellValue;
				for (int j = 0; j < vRowList.Count; j++)
				{
					cellValue = vRowList[j];
					if (cellValue == "null")
					{
						cellValue = string.Empty;
					}
					excelSheet.Cells[numberOfRow, j + sum] = cellValue;
				}
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_RowToSet", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_CellLocation", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Write '{v_RowToSet}' to Row '{v_CellLocation}' - Instance Name '{v_InstanceName}']";
		}       
	}
}