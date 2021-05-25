﻿using OpenBots.Commands.Terminal.Forms;
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

namespace OpenBots.Commands.Terminal
{
    [Serializable]
	[Category("Terminal Commands")]
	[Description("This command waits for text to appear on a targeted terminal screen.")]
	public class TerminalWaitForTextCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Terminal Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Terminal Session** command.")]
		[SampleUsage("MyTerminalInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Terminal Session** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("Text to Wait for")]
		[Description("Enter the text to wait for on the terminal.")]
		[SampleUsage("\"Hello, World!\" || vText")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_TextToWaitFor { get; set; }

		[Required]
		[DisplayName("Timeout (Seconds)")]
		[Description("Specify how many seconds to wait before throwing an exception.")]
		[SampleUsage("30 || vSeconds")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_Timeout { get; set; }

		public TerminalWaitForTextCommand()
		{
			CommandName = "TerminalWaitForTextCommand";
			SelectionName = "Wait For Text";
			CommandEnabled = true;
			CommandIcon = Resources.command_system;
			v_InstanceName = "DefaultTerminal";
			v_Timeout = "30";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			string textToWaitFor = (string)await v_TextToWaitFor.EvaluateCode(engine);
			var timeout = ((int)await v_Timeout.EvaluateCode(engine)) * 1000;
			var terminalObject = (OpenEmulator)((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			if (terminalObject.TN3270 == null || !terminalObject.TN3270.IsConnected)
				throw new Exception($"Terminal Instance {v_InstanceName} is not connected.");

			int result = terminalObject.TN3270.WaitForTextOnScreen(timeout, textToWaitFor);

			if (result == -1)
				throw new TimeoutException($"Unable to find '{textToWaitFor}' within the allotted time of {timeout} seconds.");
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_TextToWaitFor", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Text '{v_TextToWaitFor}' - Instance Name '{v_InstanceName}']";
		}     
	}
}