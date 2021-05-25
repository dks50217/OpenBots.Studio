﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Folder
{
	[Serializable]
	[Category("Folder Operation Commands")]
	[Description("This command returns a list of folder directories from a specified location.")]
	public class GetFoldersCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Root Folder Path")]
		[Description("Enter or Select the path to the root folder to get its subdirectories.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_SourceFolderPath { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output Folder Path(s) Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(List<string>) })]
		public string v_OutputUserVariableName { get; set; }

		public GetFoldersCommand()
		{
			CommandName = "GetFoldersCommand";
			SelectionName = "Get Folders";
			CommandEnabled = true;
			CommandIcon = Resources.command_folders;

		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			//apply variable logic
			var sourceFolder = (string)await v_SourceFolderPath.EvaluateCode(engine);

            if (!Directory.Exists(sourceFolder))
				throw new DirectoryNotFoundException($"Directory {sourceFolder} does not exist");

			//Get Subdirectories List
			var directoriesList = Directory.GetDirectories(sourceFolder).ToList();

			directoriesList.SetVariableValue(engine, v_OutputUserVariableName);
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_SourceFolderPath", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));
			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [From '{v_SourceFolderPath}' - Store Folder Path(s) in '{v_OutputUserVariableName}']";
		}
	}
}