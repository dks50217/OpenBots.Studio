﻿using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
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
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Outlook.Application;

namespace OpenBots.Commands.Outlook
{
	[Serializable]
	[Category("Outlook Commands")]
	[Description("This command gets selected emails and their attachments from Outlook.")]

	public class GetOutlookEmailsCommand : ScriptCommand
	{

		[Required]
		[DisplayName("Source Mail Folder Name")]
		[Description("Enter the name of the Outlook mail folder the emails are located in.")]
		[SampleUsage("\"Inbox\" || vFolderName")]
		[Remarks("Source folder cannot be a subfolder.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_SourceFolder { get; set; }

		[Required]
		[DisplayName("Filter")]
		[Description("Enter a valid Outlook filter string.")]
		[SampleUsage("\"[Subject] = 'Hello'\" || \"[Subject] = 'Hello' and [SenderName] = 'Jane Doe'\" || vFilter || \"None\"")]
		[Remarks("*Warning* Using 'None' as the Filter will return every email in the selected Mail Folder.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_Filter { get; set; }

		[Required]
		[DisplayName("Unread Only")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to retrieve unread email messages only.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_GetUnreadOnly { get; set; }

		[Required]
		[DisplayName("Mark As Read")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to mark retrieved emails as read.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_MarkAsRead { get; set; }

		[Required]
		[DisplayName("Save MailItems and Attachments")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to save the email attachments to a local directory.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_SaveMessagesAndAttachments { get; set; }

		[Required]
		[DisplayName("Include Embedded Images")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to consider images in body as attachments.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_IncludeEmbeddedImagesAsAttachments { get; set; }

		[Required]
		[DisplayName("Output MailItem Directory")]
		[Description("Enter or Select the path of the directory to store the messages in.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("This input is optional and will only be used if *Save MailItems and Attachments* is set to **Yes**.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_MessageDirectory { get; set; }

		[Required]
		[DisplayName("Output Attachment Directory")]
		[Description("Enter or Select the path to the directory to store the attachments in.")]
		[SampleUsage("@\"C:\\temp\" || ProjectPath + @\"\\temp\" || vDirectoryPath")]
		[Remarks("This input is optional and will only be used if *Save MailItems and Attachments* is set to **Yes**.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[Editor("ShowFolderSelectionHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_AttachmentDirectory { get; set; }

		[Required]
		[Editable(false)]
		[DisplayName("Output MailItem List Variable")]
		[Description("Create a new variable or select a variable from the list.")]
		[SampleUsage("vUserVariable")]
		[Remarks("New variables/arguments may be instantiated by utilizing the Ctrl+K/Ctrl+J shortcuts.")]
		[CompatibleTypes(new Type[] { typeof(List<MailItem>) })]
		public string v_OutputUserVariableName { get; set; }

		[JsonIgnore]
		[Browsable(false)]
		private List<Control> _savingControls;

		[JsonIgnore]
		[Browsable(false)]
		private bool _hasRendered;

		public GetOutlookEmailsCommand()
		{
			CommandName = "GetOutlookEmailsCommand";
			SelectionName = "Get Outlook Emails";
			CommandEnabled = true;
			CommandIcon = Resources.command_smtp;

			v_SourceFolder = "\"Inbox\"";
			v_GetUnreadOnly = "No";
			v_MarkAsRead = "Yes";
			v_SaveMessagesAndAttachments = "No";
			v_IncludeEmbeddedImagesAsAttachments = "No";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			var vFolder = (string)await v_SourceFolder.EvaluateCode(engine);
			var vFilter = (string)await v_Filter.EvaluateCode(engine);
			var vAttachmentDirectory = (string)await v_AttachmentDirectory.EvaluateCode(engine);
			var vMessageDirectory = (string)await v_MessageDirectory.EvaluateCode(engine);

			if (vFolder == "") 
				vFolder = "Inbox";

			Application outlookApp = new Application();
			AddressEntry currentUser = outlookApp.Session.CurrentUser.AddressEntry;
			NameSpace test = outlookApp.GetNamespace("MAPI");

			if (currentUser.Type == "EX")
			{
				MAPIFolder inboxFolder = (MAPIFolder)test.GetDefaultFolder(OlDefaultFolders.olFolderInbox).Parent;
				MAPIFolder userFolder = inboxFolder.Folders[vFolder];
				Items filteredItems = null;

				if (string.IsNullOrEmpty(vFilter.Trim()))
					throw new NullReferenceException("Outlook Filter not specified");
				else if (vFilter != "None")
				{
					try
					{
						filteredItems = userFolder.Items.Restrict(vFilter);
					}
					catch (System.Exception)
					{
						throw new InvalidDataException("Outlook Filter is not valid");
					}
				}
				else
					filteredItems = userFolder.Items;

				List<MailItem> outMail = new List<MailItem>();

				foreach (object _obj in filteredItems)
				{
					if (_obj is MailItem)
					{

						MailItem tempMail = (MailItem)_obj;
						if (v_GetUnreadOnly == "Yes")
						{
							if (tempMail.UnRead == true)
							{
								ProcessEmail(tempMail, vMessageDirectory, vAttachmentDirectory);
								outMail.Add(tempMail);
							}
						}
						else {
							ProcessEmail(tempMail, vMessageDirectory, vAttachmentDirectory);
							outMail.Add(tempMail);
						}   
					}
				}

				outMail.SetVariableValue(engine, v_OutputUserVariableName);
			}
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_SourceFolder", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_Filter", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_GetUnreadOnly", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_MarkAsRead", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_SaveMessagesAndAttachments", this, editor));
			((ComboBox)RenderedControls[11]).SelectedIndexChanged += SaveMailItemsComboBox_SelectedIndexChanged;

			_savingControls = new List<Control>();
			_savingControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_IncludeEmbeddedImagesAsAttachments", this, editor));
			_savingControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_MessageDirectory", this, editor));
			_savingControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_AttachmentDirectory", this, editor));

			RenderedControls.AddRange(_savingControls);

			RenderedControls.AddRange(commandControls.CreateDefaultOutputGroupFor("v_OutputUserVariableName", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [From '{v_SourceFolder}' - Filter by '{v_Filter}' - Store MailItem List in '{v_OutputUserVariableName}']";
		}

		public override void Shown()
		{
			base.Shown();
			_hasRendered = true;
			SaveMailItemsComboBox_SelectedIndexChanged(null, null);
		}

		private void ProcessEmail(MailItem mail, string msgDirectory, string attDirectory)
		{
			if (v_MarkAsRead == "Yes")
			{
				mail.UnRead = false;
			}
			if (v_SaveMessagesAndAttachments == "Yes")
			{
				if (Directory.Exists(msgDirectory))
				{
					if (string.IsNullOrEmpty(mail.Subject))
						mail.Subject = "(no subject)";
					string mailFileName = string.Join("_", mail.Subject.Split(Path.GetInvalidFileNameChars()));
					mail.SaveAs(Path.Combine(msgDirectory, mailFileName + ".msg"));
				}

				if (Directory.Exists(attDirectory))
				{
					if (v_IncludeEmbeddedImagesAsAttachments.Equals("Yes")) {
						foreach (Attachment attachment in mail.Attachments)
						{
							attachment.SaveAsFile(Path.Combine(attDirectory, attachment.FileName));
						}
					}
					else
					{
						foreach (Attachment attachment in mail.Attachments)
						{
							var flags = mail.HTMLBody.Contains(attachment.FileName);

							if (!flags)
							{
								attachment.SaveAsFile(Path.Combine(attDirectory, attachment.FileName));
							}
						}
					}
				}
			}
		}

		private void SaveMailItemsComboBox_SelectedIndexChanged(object sender, EventArgs e)
		{
			if (((ComboBox)RenderedControls[11]).Text == "Yes" && _hasRendered)
			{
				foreach (var ctrl in _savingControls)
					ctrl.Visible = true;
			}
			else if(_hasRendered)
			{
				foreach (var ctrl in _savingControls)
				{
					ctrl.Visible = false;
					if (ctrl is TextBox)
						((TextBox)ctrl).Clear();
				}
			}
		}
	}
}
