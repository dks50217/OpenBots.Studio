﻿using OpenBots.Commands.Terminal.Forms;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows.Forms;

namespace OpenBots.Commands.Terminal
{
    [Serializable]
	[Category("Terminal Commands")]
	[Description("This command retrieves and loads the text from a targeted terminal screen.")]
	public class DumpTerminalScreenXMLCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Terminal Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Terminal Session** command.")]
		[SampleUsage("MyTerminalInstance")]
		[Remarks("Failure to enter the correct instance or failure to first call the **Create Terminal Session** command will cause an error.")]
		public string v_InstanceName { get; set; }		

		public DumpTerminalScreenXMLCommand()
		{
			CommandName = "DumpTerminalScreenXMLCommand";
			SelectionName = "Dump Terminal Screen XML";
			CommandEnabled = true;
			CommandIcon = Resources.command_system;

			v_InstanceName = "DefaultTerminal";
		}

		public override void RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var terminalObject = (OpenEmulator)v_InstanceName.GetAppInstance(engine);

			if (terminalObject.TN3270 == null || !terminalObject.TN3270.IsConnected)
				throw new Exception($"Terminal Instance {v_InstanceName} is not connected.");

			terminalObject.Redraw();
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Instance Name '{v_InstanceName}']";
		}     
	}
}