﻿using OpenBots.Commands.Terminal.Forms;
using OpenBots.Commands.Terminal.Library;
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

namespace OpenBots.Commands.Terminal
{
    [Serializable]
	[Category("Terminal Commands")]
	[Description("This command sets a stored username in a targeted terminal screen.")]
	public class TerminalSetUsernameCommand : ScriptCommand
	{
		[Required]
		[DisplayName(" Terminal Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Terminal Session** command.")]
		[SampleUsage("MyTerminalInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Terminal Session** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[DisplayName("Row Position (Optional)")]
		[Description("Input the new vertical position of the terminal. Starts from 0 at the top and increases going down.")]
		[SampleUsage("0 || vRowPosition")]
		[Remarks("This number is the pixel location on screen. Maximum value should be the maximum value allowed by the terminal.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_YMousePosition { get; set; }

		[DisplayName("Column Position (Optional)")]
		[Description("Input the new horizontal position of the terminal. Starts from 1 on the left and increases going right.")]
		[SampleUsage("1 || vColPosition")]
		[Remarks("This number is the pixel location on screen. Maximum value should be the maximum value allowed by the terminal.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_XMousePosition { get; set; }

		[Required]
		[DisplayName("Timeout (Seconds)")]
		[Description("Specify how many seconds to wait before throwing an exception.")]
		[SampleUsage("30 || vSeconds")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_Timeout { get; set; }

		public TerminalSetUsernameCommand()
		{
			CommandName = "TerminalSetUsernameCommand";
			SelectionName = "Set Username";
			CommandEnabled = true;
			CommandIcon = Resources.command_system;
			v_Timeout = "30";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;

			int mouseX = 0, mouseY = 0;
			if (!string.IsNullOrEmpty(v_XMousePosition))
				mouseX = (int)await v_XMousePosition.EvaluateCode(engine);

			if (!string.IsNullOrEmpty(v_YMousePosition))
				mouseY = (int)await v_YMousePosition.EvaluateCode(engine);

			var timeout = ((int)await v_Timeout.EvaluateCode(engine)) * 1000;
			var terminalObject = (OpenEmulator)((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			if (terminalObject.TN3270 == null || !terminalObject.TN3270.IsConnected)
				throw new Exception($"Terminal Instance {v_InstanceName} is not connected.");

			if (!string.IsNullOrEmpty(v_XMousePosition) && !string.IsNullOrEmpty(v_YMousePosition))
				terminalObject.TN3270.SetCursor(mouseY, mouseX);

			terminalObject.TN3270.SetText(terminalObject.Username);
			int result = terminalObject.TN3270.WaitForTextOnScreen(timeout, terminalObject.Username);

			if (result == -1)
				throw new TimeoutException($"Unable to find '{terminalObject.Username}' within the allotted time of {timeout / 1000} seconds.");

			terminalObject.Redraw();
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_YMousePosition", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_XMousePosition", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Row/Col '{{{v_YMousePosition}, {v_XMousePosition}}}' - Instance Name '{v_InstanceName}']";
		}     
	}
}