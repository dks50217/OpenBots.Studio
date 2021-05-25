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
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Outlook
{
	[Serializable]
	[Category("Outlook Commands")]
	[Description("This command deletes a selected email in Outlook.")]

	public class DeleteOutlookEmailCommand : ScriptCommand
	{

		[Required]
		[DisplayName("MailItem")]
		[Description("Enter the MailItem to delete.")]
		[SampleUsage("vMailItem")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(MailItem) })]
		public string v_MailItem { get; set; }

		[Required]
		[DisplayName("Delete Read Emails Only")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to delete read email messages only.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_DeleteReadOnly { get; set; }

		public DeleteOutlookEmailCommand()
		{
			CommandName = "DeleteOutlookEmailCommand";
			SelectionName = "Delete Outlook Email";
			CommandEnabled = true;
			CommandIcon = Resources.command_outlook;

			v_DeleteReadOnly = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			MailItem vMailItem = (MailItem)await v_MailItem.EvaluateCode(engine);

			if (v_DeleteReadOnly == "Yes")
			{
				if (vMailItem.UnRead == false)
					vMailItem.Delete();
			}
			else
				vMailItem.Delete();
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_MailItem", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_DeleteReadOnly", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [MailItem '{v_MailItem}']";
		}
	}
}