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
using System.Data;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace OpenBots.Commands.Input
{
    [Serializable]
	[Category("Input Commands")]
	[Description("This command sends advanced keystrokes to a targeted window.")]
	public class SendAdvancedKeystrokesCommand : ScriptCommand, ISendAdvancedKeystrokesCommand
	{

		[Required]
		[DisplayName("Window Name")]
		[Description("Select the name of the window to send advanced keystrokes to.")]
		[SampleUsage("\"Untitled - Notepad\" || \"Current Window\" || vWindow")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("CaptureWindowHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_WindowName { get; set; }

		[Required]
		[DisplayName("Keystroke Parameters")]
		[Description("Define the parameters for the keystroke actions.")]
		[SampleUsage("[Enter [Return] | Key Press (Down + Up)]")]
		[Remarks("")]
		public DataTable v_KeyActions { get; set; }

		[Required]
		[DisplayName("Return All Keys to 'UP' Position")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]      
		[Description("Select whether to return all keys to the 'UP' position after execution.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_KeyUpDefault { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private DataGridView _keystrokeGridHelper;

		public SendAdvancedKeystrokesCommand()
		{
			CommandName = "SendAdvancedKeystrokesCommand";
			SelectionName = "Send Advanced Keystrokes";
			CommandEnabled = true;
			CommandIcon = Resources.command_input;

			v_KeyActions = new DataTable();
			v_KeyActions.Columns.Add("Key");
			v_KeyActions.Columns.Add("Action");
			v_KeyActions.TableName = "SendAdvancedKeyStrokesCommand" + DateTime.Now.ToString("MMddyy.hhmmss");

			v_WindowName = "\"Current Window\"";
			v_KeyUpDefault = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var variableWindowName = (string)await v_WindowName.EvaluateCode(engine);

			//activate anything except current window
			if (variableWindowName != "Current Window")
				User32Functions.ActivateWindow(variableWindowName);

			//track all keys down
			var keysDown = new List<Keys>();

			//run each selected item
			foreach (DataRow rw in v_KeyActions.Rows)
			{
				//get key name
				var keyName = rw.Field<string>("Key");

				//get key action
				var action = rw.Field<string>("Action");

				//parse OEM key name
				string oemKeyString = keyName.Split('[', ']')[1];

				var oemKeyName = (Keys)Enum.Parse(typeof(Keys), oemKeyString);
		   
				//"Key Press (Down + Up)", "Key Down", "Key Up"
				switch (action)
				{
					case "Key Press (Down + Up)":
						//simulate press
						User32Functions.KeyDown(oemKeyName);
						User32Functions.KeyUp(oemKeyName);
						
						//key returned to UP position so remove if we added it to the keys down list
						if (keysDown.Contains(oemKeyName))
							keysDown.Remove(oemKeyName);
						break;
					case "Key Down":
						//simulate down
						User32Functions.KeyDown(oemKeyName);

						//track via keys down list
						if (!keysDown.Contains(oemKeyName))
							keysDown.Add(oemKeyName);
						break;
					case "Key Up":
						//simulate up
						User32Functions.KeyUp(oemKeyName);

						//remove from key down
						if (keysDown.Contains(oemKeyName))
							keysDown.Remove(oemKeyName);
						break;
					default:
						break;
				}
			}

			//return key to up position if requested
			if (v_KeyUpDefault == "Yes")
			{
				foreach (var key in keysDown)
					User32Functions.KeyUp(key);
			}       
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultWindowControlGroupFor("v_WindowName", this, editor));

			RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_KeyActions", this));

			_keystrokeGridHelper = commandControls.CreateDefaultDataGridViewFor("v_KeyActions", this);
			_keystrokeGridHelper.AutoGenerateColumns = false;

			DataGridViewComboBoxColumn propertyName = new DataGridViewComboBoxColumn();
			propertyName.DataSource = CommonMethods.GetAvailableKeys();
			propertyName.HeaderText = "Selected Key";
			propertyName.DataPropertyName = "Key";
			_keystrokeGridHelper.Columns.Add(propertyName);

			DataGridViewComboBoxColumn propertyValue = new DataGridViewComboBoxColumn();
			propertyValue.DataSource = new List<string> { "Key Press (Down + Up)", "Key Down", "Key Up" };
			propertyValue.HeaderText = "Selected Action";
			propertyValue.DataPropertyName = "Action";
			_keystrokeGridHelper.Columns.Add(propertyValue);

			RenderedControls.Add(_keystrokeGridHelper);

			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_KeyUpDefault", this, editor));

			return RenderedControls;
		}
	 
		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [Window '{v_WindowName}']";
		}
	}
}