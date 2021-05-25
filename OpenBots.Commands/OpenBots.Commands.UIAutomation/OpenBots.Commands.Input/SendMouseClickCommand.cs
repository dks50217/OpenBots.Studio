﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.User32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Input
{
	[Serializable]
	[Category("Input Commands")]
	[Description("This command simulates a mouse click of a given type.")]
	public class SendMouseClickCommand : ScriptCommand
	{

		[Required]
		[DisplayName("Click Type")]
		[PropertyUISelectionOption("Left Click")]
		[PropertyUISelectionOption("Middle Click")]
		[PropertyUISelectionOption("Right Click")]
		[PropertyUISelectionOption("Left Down")]
		[PropertyUISelectionOption("Middle Down")]
		[PropertyUISelectionOption("Right Down")]
		[PropertyUISelectionOption("Left Up")]
		[PropertyUISelectionOption("Middle Up")]
		[PropertyUISelectionOption("Right Up")]
		[PropertyUISelectionOption("Double Left Click")]
		[Description("Indicate the type of click required.")]
		[SampleUsage("")]
		[Remarks("You can simulate a custom click by using multiple mouse click commands in succession, adding **Pause Command** in between where required.")]
		public string v_MouseClick { get; set; }

		public SendMouseClickCommand()
		{
			CommandName = "SendMouseClickCommand";
			SelectionName = "Send Mouse Click";
			CommandEnabled = true;
			CommandIcon = Resources.command_input;

			v_MouseClick = "Left Click";
		}

		public async override Task RunCommand(object sender)
		{
			var mousePosition = Cursor.Position;
			User32Functions.SendMouseClick(v_MouseClick, mousePosition.X, mousePosition.Y);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_MouseClick", this, editor));
	   
			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Click Type '{v_MouseClick}']";
		}
	}
}