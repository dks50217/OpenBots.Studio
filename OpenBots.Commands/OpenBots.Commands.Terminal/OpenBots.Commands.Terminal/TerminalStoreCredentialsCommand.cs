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
using System.Security;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Terminal
{
    [Serializable]
	[Category("Terminal Commands")]
	[Description("This command stores credentials to use in the *Set Username* and *Set Password* commands.")]
	public class TerminalStoreCredentialsCommand : ScriptCommand
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
		[DisplayName("Username")]
		[Description("Define the username to use when connecting to the terminal.")]
		[SampleUsage("\"myRobot\" || vUsername")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_Username { get; set; }

		[Required]
		[DisplayName("Password")]
		[Description("Define the password to use when connecting to the terminal.")]
		[SampleUsage("vPassword")]
		[Remarks("Password input must be a SecureString variable.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(SecureString) })]
		public string v_Password { get; set; }

		public TerminalStoreCredentialsCommand()
		{
			CommandName = "TerminalStoreCredentialsCommand";
			SelectionName = "Store Credentials";
			CommandEnabled = true;
			CommandIcon = Resources.command_system;
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vUserName = (string)await v_Username.EvaluateCode(engine);
			var vPassword = (SecureString)await v_Password.EvaluateCode(engine);
			var terminalObject = (OpenEmulator)((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;

			if (terminalObject.TN3270 == null || !terminalObject.TN3270.IsConnected)
				throw new Exception($"Terminal Instance {v_InstanceName} is not connected.");

			terminalObject.Username = vUserName;
			terminalObject.Password = vPassword;
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Username", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Password", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Username '{v_Username}' - Password '{v_Password}' - Instance Name '{v_InstanceName}']";
		}
	}
}
