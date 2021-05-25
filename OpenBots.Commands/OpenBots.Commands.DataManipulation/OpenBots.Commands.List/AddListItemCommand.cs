﻿using Microsoft.Office.Interop.Outlook;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;
using OBDataTable = System.Data.DataTable;

namespace OpenBots.Commands.List
{
    [Serializable]
	[Category("List Commands")]
	[Description("This command adds an item to an existing List variable.")]
	public class AddListItemCommand : ScriptCommand
	{
		[Required]
		[DisplayName("List")]
		[Description("Provide a List variable.")]
		[SampleUsage("vList || new List<string>() { \"hello\", \"world\" }")]
		[Remarks("Any type of variable other than List will cause error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(List<>) })]
		public string v_ListName { get; set; }

		[Required]
		[DisplayName("List Item")]
		[Description("Enter the item to add to the List.")]
		[SampleUsage("\"Hello\" || 1 || vItem")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(object) })]
		public OBDataTable v_ListItemsDataTable { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output List Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(List<>) })]
		public string v_OutputUserVariableName { get; set; }

		public AddListItemCommand()
		{
			CommandName = "AddListItemCommand";
			SelectionName = "Add List Item";
			CommandEnabled = true;
			CommandIcon = Resources.command_function;

			//initialize Datatable
			v_ListItemsDataTable = new OBDataTable
			{
				TableName = "ListItemsDataTable" + DateTime.Now.ToString("MMddyy.hhmmss")
			};

			v_ListItemsDataTable.Columns.Add("Items");
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;

			dynamic dynamicList = await v_ListName.EvaluateCode(engine);

			foreach (DataRow rwColumnName in v_ListItemsDataTable.Rows)
			{
				dynamic dynamicItem = await rwColumnName.Field<string>("Items").EvaluateCode(engine);
				dynamicList.Add(dynamicItem);
			}

			((object)dynamicList).SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_ListName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDataGridViewGroupFor("v_ListItemsDataTable", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Add {v_ListItemsDataTable.Rows.Count} Item(s) to List '{v_ListName}' - Store List in '{v_OutputUserVariableName}']";
		}
	}
}
