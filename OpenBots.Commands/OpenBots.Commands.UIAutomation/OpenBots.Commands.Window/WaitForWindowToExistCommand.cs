﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.User32;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Window
{
	[Serializable]
	[Category("Window Commands")]
	[Description("This command waits for a window to exist.")]
	public class WaitForWindowToExistCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Window Name")]
		[Description("Select the name of the window to wait for.")]
		[SampleUsage("\"Untitled - Notepad\" || \"Current Window\" || vWindow")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("CaptureWindowHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_WindowName { get; set; }

		[Required]
		[DisplayName("Timeout (Seconds)")]
		[Description("Specify how many seconds to wait before throwing an exception.")]
		[SampleUsage("30 || vSeconds")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_Timeout { get; set; }

		public WaitForWindowToExistCommand()
		{
			CommandName = "WaitForWindowToExistCommand";
			SelectionName = "Wait For Window To Exist";
			CommandEnabled = true;
			CommandIcon = Resources.command_window;


			v_WindowName = "\"Current Window\"";
			v_Timeout = "30";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			string windowName = (string)await v_WindowName.EvaluateCode(engine);
			var timeout = (int)await v_Timeout.EvaluateCode(engine);

			var timeToEnd = DateTime.Now.AddSeconds(timeout);
			IntPtr hWnd = IntPtr.Zero;

			while (DateTime.Now < timeToEnd)
			{
				if (engine.IsCancellationPending)
					break;
				hWnd = User32Functions.FindWindow(windowName);

				if (hWnd != IntPtr.Zero) //If found
					break;
				engine.ReportProgress($"Window '{windowName}' Not Yet Found... {(timeToEnd - DateTime.Now).Minutes}m, {(timeToEnd - DateTime.Now).Seconds}s remain");
				Thread.Sleep(1000);
			}

			if (hWnd == IntPtr.Zero)
				throw new Exception($"Window '{windowName}' was not found in the allowed time!");
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultWindowControlGroupFor("v_WindowName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Window '{v_WindowName}' - Timeout '{v_Timeout}']";
		}
	}
}
