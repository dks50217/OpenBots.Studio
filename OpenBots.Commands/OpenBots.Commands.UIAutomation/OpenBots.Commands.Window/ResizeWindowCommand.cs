﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
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
	[Description("This command resizes an open window to a specified size.")]
	public class ResizeWindowCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Window Name")]
		[Description("Select the name of the window to resize.")]
		[SampleUsage("\"Untitled - Notepad\" || \"Current Window\" || vWindow")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("CaptureWindowHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_WindowName { get; set; }

		[Required]
		[DisplayName("Width (Pixels)")]
		[Description("Input the new width size of the window.")]
		[SampleUsage("800 || vWidth")]
		[Remarks("Maximum value should be the maximum value allowed by your resolution. For 1920x1080, the valid width range would be 0-1920.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_XWindowSize { get; set; }

		[Required]
		[DisplayName("Height (Pixels)")]
		[Description("Input the new height size of the window.")]
		[SampleUsage("500 || vHeight")]
		[Remarks("Maximum value should be the maximum value allowed by your resolution. For 1920x1080, the valid height range would be 0-1080.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_YWindowSize { get; set; }

		[Required]
		[DisplayName("Timeout (Seconds)")]
		[Description("Specify how many seconds to wait before throwing an exception.")]
		[SampleUsage("30 || vSeconds")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(int) })]
		public string v_Timeout { get; set; }

		public ResizeWindowCommand()
		{
			CommandName = "ResizeWindowCommand";
			SelectionName = "Resize Window";
			CommandEnabled = true;
			CommandIcon = Resources.command_window;

			v_WindowName = "\"Current Window\"";
			v_Timeout = "30";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			string windowName = (string)await v_WindowName.EvaluateCode(engine);
			var xSize = (int)await v_XWindowSize.EvaluateCode(engine);
			var ySize = (int)await v_YWindowSize.EvaluateCode(engine);
			int timeout = (int)await v_Timeout.EvaluateCode(engine);
			DateTime timeToEnd = DateTime.Now.AddSeconds(timeout);
			List<IntPtr> targetWindows;

			while (timeToEnd >= DateTime.Now)
			{
				try
				{
					if (engine.IsCancellationPending)
						break;

					targetWindows = User32Functions.FindTargetWindows(windowName);

					if (targetWindows.Count == 0)
						throw new Exception($"Window '{windowName}' Not Yet Found... ");
					break;
				}
				catch (Exception)
				{
					engine.ReportProgress($"Window '{windowName}' Not Yet Found... {(timeToEnd - DateTime.Now).Minutes}m, {(timeToEnd - DateTime.Now).Seconds}s remain");
					Thread.Sleep(500);
				}
			}

			targetWindows = User32Functions.FindTargetWindows(windowName);

			//loop each window and set the window state
			foreach (var targetedWindow in targetWindows)
			{
				User32Functions.SetWindowState(targetedWindow, WindowState.SwShowNormal);
				User32Functions.SetWindowSize(targetedWindow, xSize, ySize);
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultWindowControlGroupFor("v_WindowName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_XWindowSize", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_YWindowSize", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Timeout", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Window '{v_WindowName}' - Target Size '({v_XWindowSize},{v_YWindowSize})']";
		}
	}
}