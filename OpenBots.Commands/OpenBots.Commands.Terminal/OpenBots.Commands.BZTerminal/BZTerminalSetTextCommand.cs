﻿using OpenBots.Commands.Terminal.Library;
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
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.BZTerminal
{
    [Serializable]
	[Category("BlueZone Terminal Commands")]
	[Description("This command sets text in a targeted terminal screen.")]
	public class BZTerminalSetTextCommand : ScriptCommand
	{
		[Required]
		[DisplayName("BZ Terminal Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create BZ Terminal Session** command.")]
		[SampleUsage("MyBZTerminalInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create BZ Terminal Session** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[DisplayName("Row Position (Optional)")]
		[Description("Input the new vertical position of the terminal. Starts from 1 at the top and increases going down.")]
		[SampleUsage("1 || vRowPosition")]
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
		[DisplayName("Text to Set")]
		[Description("Enter the text to be sent to the specified terminal.")]
		[SampleUsage("\"Hello, World!\" || vText")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_TextToSet { get; set; }

		[Required]
		[DisplayName("Timeout (Seconds)")]
		[Description("Specify how many seconds to wait before throwing an exception.")]
		[SampleUsage("30 || vSeconds")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_Timeout { get; set; }

		public BZTerminalSetTextCommand()
		{
			CommandName = "BZTerminalSetTextCommand";
			SelectionName = "BZ Set Text";
			CommandEnabled = true;
			CommandIcon = Resources.command_system;
			v_InstanceName = "DefaultBZTerminal";
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

			var textToSend = (string)await v_TextToSet.EvaluateCode(engine);
			var timeout = (int)await v_Timeout.EvaluateCode(engine);
			var terminalContext = (BZTerminalContext)((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			if (terminalContext.BZTerminalObj == null || !terminalContext.BZTerminalObj.Connected)
				throw new Exception($"Terminal Instance {v_InstanceName} is not connected.");

			if (!string.IsNullOrEmpty(v_XMousePosition) && !string.IsNullOrEmpty(v_YMousePosition))
				terminalContext.BZTerminalObj.SetCursor(mouseY, mouseX);

			terminalContext.BZTerminalObj.SendKey(textToSend);
			int result = terminalContext.BZTerminalObj.WaitForText(textToSend, 1, 1, timeout);

			if (result != 0)
				throw new TimeoutException($"Unable to find '{textToSend}' within the allotted time of {timeout} seconds.");
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_YMousePosition", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_XMousePosition", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_TextToSet", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Text '{v_TextToSet}' at Row/Col '{{{v_YMousePosition}, {v_XMousePosition}}}' - Instance Name '{v_InstanceName}']";
		}     
	}
}