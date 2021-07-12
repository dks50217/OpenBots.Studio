﻿using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.WebBrowser
{
	[Serializable]
	[Category("Web Browser Commands")]
	[Description("This command retrieves information from a Selenium web browser session.")]
	public class SeleniumGetBrowserInfoCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Browser Instance Name")]
		[Description("Enter the unique instance that was specified in the **Create Browser** command.")]
		[SampleUsage("MyBrowserInstance")]
		[Remarks("Failure to enter the correct instance name or failure to first call the **Create Browser** command will cause an error.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
		public string v_InstanceName { get; set; }

		[Required]
		[DisplayName("Info Property")]
		[PropertyUISelectionOption("Window Title")]
		[PropertyUISelectionOption("Window URL")]
		[PropertyUISelectionOption("Current Handle ID")]
		[PropertyUISelectionOption("HTML Page Source")]
		[PropertyUISelectionOption("Handle ID List")]
		[Description("Indicate which info property to retrieve.")]
		[SampleUsage("")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		public string v_InfoType { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Info Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_OutputUserVariableName { get; set; }

		public SeleniumGetBrowserInfoCommand()
		{
			CommandName = "SeleniumGetBrowserInfoCommand";
			SelectionName = "Selenium Get Browser Info";
			CommandEnabled = true;
			CommandIcon = Resources.command_web;
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var browserObject = ((OBAppInstance)await v_InstanceName.EvaluateCode(engine)).Value;
			var seleniumInstance = (IWebDriver)browserObject;
			string info;

			switch (v_InfoType)
			{
				case "Window Title":
					info = seleniumInstance.Title;
					break;
				case "Window URL":
					info = seleniumInstance.Url;
					break;
				case "Current Handle ID":
					info = seleniumInstance.CurrentWindowHandle;
					break;
				case "HTML Page Source":
					info = seleniumInstance.PageSource;
					break;
				case "Handle ID List":
					info = JsonConvert.SerializeObject(seleniumInstance.WindowHandles);
					break;
				default:
					throw new NotImplementedException($"{v_InfoType} is not implemented for lookup.");
			}
			//store data
			info.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_InfoType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Get {v_InfoType} - Store Info in '{v_OutputUserVariableName}' - Instance Name '{v_InstanceName}']";
		}
	}
}