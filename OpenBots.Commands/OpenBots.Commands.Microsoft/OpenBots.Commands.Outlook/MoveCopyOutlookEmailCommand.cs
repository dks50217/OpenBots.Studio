﻿using Microsoft.Office.Interop.Outlook;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace OpenBots.Commands.Outlook
{
	[Serializable]
	[Category("Outlook Commands")]
	[Description("This command moves or copies a selected email in Outlook.")]
	public class MoveCopyOutlookEmailCommand : ScriptCommand
	{
		[Required]
		[DisplayName("MailItem")]
		[Description("Enter the MailItem to move or copy.")]
		[SampleUsage("vMailItem")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(MailItem) })]
		public string v_MailItem { get; set; }

		[Required]
		[DisplayName("Destination Mail Folder Name")]
		[Description("Enter the name of the Outlook mail folder the emails are being moved/copied to.")]
		[SampleUsage("\"New Folder\" || vFolderName")]
		[Remarks("Destination folder cannot be a subfolder.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_DestinationFolder { get; set; }

		[Required]
		[DisplayName("Mail Operation")]
		[PropertyUISelectionOption("Move MailItem")]
		[PropertyUISelectionOption("Copy MailItem")]
		[Description("Specify whether to move or copy the selected emails.")]
		[SampleUsage("")]
		[Remarks("Moving will remove the emails from the original folder while copying will not.")]
		public string v_OperationType { get; set; }

		[Required]
		[DisplayName("Unread Only")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to move/copy unread email messages only.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_MoveCopyUnreadOnly { get; set; }

		public MoveCopyOutlookEmailCommand()
		{
			CommandName = "MoveCopyOutlookEmailCommand";
			SelectionName = "Move/Copy Outlook Email";
			CommandEnabled = true;
			CommandIcon = Resources.command_outlook;

			v_OperationType = "Move MailItem";
			v_MoveCopyUnreadOnly = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			MailItem vMailItem = (MailItem)await v_MailItem.EvaluateCode(engine);
			var vDestinationFolder = (string)await v_DestinationFolder.EvaluateCode(engine);

			Application outlookApp = new Application();
			NameSpace test = outlookApp.GetNamespace("MAPI");
			test.Logon("", "", false, true);
			AddressEntry currentUser = outlookApp.Session.CurrentUser.AddressEntry;

			if (currentUser.Type == "EX")
			{
				MAPIFolder inboxFolder = (MAPIFolder)test.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Parent;
				MAPIFolder destinationFolder = inboxFolder.Folders[vDestinationFolder];

				if(v_OperationType == "Move MailItem")
				{
					if (v_MoveCopyUnreadOnly == "Yes")
					{
						if (vMailItem.UnRead == true)
							vMailItem.Move(destinationFolder);
					}
					else
						vMailItem.Move(destinationFolder);
				}
				else if (v_OperationType == "Copy MailItem")
				{
                    MailItem copyMail;

                    if (v_MoveCopyUnreadOnly == "Yes")
					{
						if (vMailItem.UnRead == true)
						{
							copyMail = (MailItem)vMailItem.Copy();
							copyMail.Move(destinationFolder);
						}
					}
					else
					{
						copyMail = (MailItem)vMailItem.Copy();
						copyMail.Move(destinationFolder);
					}                       
				}               
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_MailItem", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_DestinationFolder", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_OperationType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_MoveCopyUnreadOnly", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [{v_OperationType} '{v_MailItem}' to '{v_DestinationFolder}']";
		}
	}
}