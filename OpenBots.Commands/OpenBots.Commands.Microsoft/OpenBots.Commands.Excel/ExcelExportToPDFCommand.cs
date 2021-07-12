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
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace OpenBots.Commands.Excel
{
	[Serializable]
	[Category("Excel Commands")]
	[Description("This command exports a Excel Worksheet to a PDF file.")]
	public class ExcelExportToPDFCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Excel Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Application** command.")]
		[SampleUsage("MyExcelInstance || vExcelInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Application** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("PDF Location")]
		[Description("Enter or Select the path of the folder to export the PDF to.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath || ProjectPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_FolderPath { get; set; }

		[Required]
		[DisplayName("PDF File Name")]
		[Description("Enter or Select the name of the PDF file.")]
		[SampleUsage("@\"myFile.pdf\" || vFilename")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_FileName { get; set; }

		[Required]
		[DisplayName("AutoFit Cells")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Indicate whether to autofit cell sizes to fit their contents.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_AutoFitCells { get; set; }

		[Required]
		[DisplayName("Display Gridlines")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Indicate whether to display Worksheet gridlines.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_DisplayGridlines { get; set; }

		public ExcelExportToPDFCommand()
		{
			CommandName = "ExcelExportToPDFCommand";
			SelectionName = "Export To PDF";
			CommandEnabled = true;
			CommandIcon = Resources.command_excel;

			v_AutoFitCells = "Yes";
			v_DisplayGridlines = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vFileName = (string)await v_FileName.EvaluateCode(engine);
			var vFolderPath = (string)await v_FolderPath.EvaluateCode(engine);

			//get excel app object
			var excelObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			//convert object
			Application excelInstance = (Application)excelObject;
			Worksheet excelWorksheet = excelInstance.ActiveSheet;

			var fileFormat = XlFixedFormatType.xlTypePDF;
			string pdfPath = Path.Combine(vFolderPath, vFileName);
			
			if (v_AutoFitCells == "Yes")
			{
				excelWorksheet.Columns.AutoFit();
				excelWorksheet.Rows.AutoFit();
			}

			if (v_DisplayGridlines == "Yes")
			{
				Range last = excelWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
				Range range = excelWorksheet.Range["A1", last];
				range.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
				range.Borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
				range.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
				range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
				range.Borders.Color = Color.Black;
			}
				
			excelWorksheet.ExportAsFixedFormat(fileFormat, pdfPath);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_FolderPath", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_FileName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_AutoFitCells", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_DisplayGridlines", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Export to '{v_FolderPath}\\{v_FileName}' - Instance Name '{v_InstanceName}']";
		}
	}
}