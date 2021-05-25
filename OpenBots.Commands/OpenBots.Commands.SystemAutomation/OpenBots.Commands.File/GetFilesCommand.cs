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

namespace OpenBots.Commands.File
{
	[Serializable]
	[Category("File Operation Commands")]
	[Description("This command returns a list of file paths from a specified location.")]
	public class GetFilesCommand : ScriptCommand
	{
		[Required]
		[DisplayName("Source Folder Path")]
		[Description("Enter or Select the path to the folder.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_SourceFolderPath { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output File Path(s) List Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(List<string>) })]
		public string v_OutputUserVariableName { get; set; }

		public GetFilesCommand()
		{
			CommandName = "GetFilesCommand";
			SelectionName = "Get Files";
			CommandEnabled = true;
			CommandIcon = Resources.command_files;

		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			//apply variable logic
			var sourceFolder = (string)await v_SourceFolderPath.EvaluateCode(engine);

			if (!Directory.Exists(sourceFolder))
				throw new DirectoryNotFoundException($"{sourceFolder} is not a valid directory");

			//Get File Paths from the folder
			var filesList = Directory.GetFiles(sourceFolder, ".", SearchOption.AllDirectories).ToList();

			//Add File Paths to the output variable
			filesList.SetVariableValue(engine, v_OutputUserVariableName);
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
			return base.GetDisplayValue() + $" [From Folder '{v_SourceFolderPath}' - Store File Path(s) List in '{v_OutputUserVariableName}']";
		}
	}
}