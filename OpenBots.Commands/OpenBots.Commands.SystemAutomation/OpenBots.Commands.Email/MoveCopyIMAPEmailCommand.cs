﻿using MailKit;
using MailKit.Net.Imap;
using MimeKit;
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
using System.Linq;
using System.Security;
using System.Security.Authentication;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.Email
{
	[Serializable]
	[Category("Email Commands")]
	[Description("This command moves or copies a selected email using IMAP protocol.")]
	public class MoveCopyIMAPEmailCommand : ScriptCommand
	{
		[Required]
		[DisplayName("MimeMessage")]
		[Description("Enter the MimeMessage to move or copy.")]
		[SampleUsage("vMimeMessage")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(MimeMessage) })]
		public string v_IMAPMimeMessage { get; set; }

		[Required]
		[DisplayName("Host")]
		[Description("Define the host/service name that the script should use.")]
		[SampleUsage("\"imap.gmail.com\" || vHost")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_IMAPHost { get; set; }

		[Required]
		[DisplayName("Port")]
		[Description("Define the port number that should be used when contacting the IMAP service.")]
		[SampleUsage("\"993\" || vPort")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_IMAPPort { get; set; }

		[Required]
		[DisplayName("Username")]
		[Description("Define the username to use when contacting the IMAP service.")]
		[SampleUsage("\"myRobot\" || vUsername")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_IMAPUserName { get; set; }

		[Required]
		[DisplayName("Password")]
		[Description("Define the password to use when contacting the IMAP service.")]
		[SampleUsage("vPassword")]
		[Remarks("Password input must be a SecureString variable.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(SecureString) })]
		public string v_IMAPPassword { get; set; }

		[Required]
		[DisplayName("Destination Mail Folder Name")]
		[Description("Enter the name of the mail folder the emails are being moved/copied to.")]
		[SampleUsage("\"New Folder\" || vFolderName")]
		[Remarks("")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
		[CompatibleTypes(new Type[] { typeof(string) })]
		public string v_IMAPDestinationFolder { get; set; }

		[Required]
		[DisplayName("Mail Operation")]
		[PropertyUISelectionOption("Move MimeMessage")]
		[PropertyUISelectionOption("Copy MimeMessage")]
		[Description("Specify whether to move or copy the selected emails.")]
		[SampleUsage("")]
		[Remarks("Moving will remove the emails from the original folder while copying will not.")]
		public string v_IMAPOperationType { get; set; }

		[Required]
		[DisplayName("Unread Only")]
		[PropertyUISelectionOption("Yes")]
		[PropertyUISelectionOption("No")]
		[Description("Specify whether to move/copy unread email messages only.")]
		[SampleUsage("")]
		[Remarks("")]
		public string v_IMAPMoveCopyUnreadOnly { get; set; }

		public MoveCopyIMAPEmailCommand()
		{
			CommandName = "MoveCopyIMAPEmailCommand";
			SelectionName = "Move/Copy IMAP Email";
			CommandEnabled = true;
			CommandIcon = Resources.command_smtp;

			v_IMAPOperationType = "Move MimeMessage";
			v_IMAPMoveCopyUnreadOnly = "Yes";
		}

		public async override Task RunCommand(object sender)
		{
			var engine = (IAutomationEngineInstance)sender;
			MimeMessage vMimeMessage = (MimeMessage)await v_IMAPMimeMessage.EvaluateCode(engine);
			string vIMAPHost = (string)await v_IMAPHost.EvaluateCode(engine);
			string vIMAPPort = (string)await v_IMAPPort.EvaluateCode(engine);
			string vIMAPUserName = (string)await v_IMAPUserName.EvaluateCode(engine);
			string vIMAPPassword = ((SecureString)await v_IMAPPassword.EvaluateCode(engine)).ConvertSecureStringToString();
			var vIMAPDestinationFolder = (string)await v_IMAPDestinationFolder.EvaluateCode(engine);

			using (var client = new ImapClient())
			{
				client.ServerCertificateValidationCallback = (sndr, certificate, chain, sslPolicyErrors) => true;
				client.SslProtocols = SslProtocols.None;

				using (var cancel = new CancellationTokenSource())
				{
					try
					{
						client.Connect(vIMAPHost, int.Parse(vIMAPPort), true, cancel.Token); //SSL
					}
					catch (Exception)
					{
						client.Connect(vIMAPHost, int.Parse(vIMAPPort)); //TLS
					}

					client.AuthenticationMechanisms.Remove("XOAUTH2");
					client.Authenticate(vIMAPUserName, vIMAPPassword, cancel.Token);

					var splitId = vMimeMessage.MessageId.Split('#').ToList();
					UniqueId messageId = UniqueId.Parse(splitId.Last());
					splitId.RemoveAt(splitId.Count - 1);
					string messageFolder = string.Join("", splitId);

					IMailFolder toplevel = client.GetFolder(client.PersonalNamespaces[0]);
					IMailFolder foundSourceFolder = GetIMAPEmailsCommand.FindFolder(toplevel, messageFolder);
					IMailFolder foundDestinationFolder = GetIMAPEmailsCommand.FindFolder(toplevel, vIMAPDestinationFolder);

					if (foundSourceFolder != null)
						foundSourceFolder.Open(FolderAccess.ReadWrite, cancel.Token);
					else
						throw new Exception("Source Folder not found");

					if (foundDestinationFolder == null)
						throw new Exception("Destination Folder not found");

					var messageSummary = foundSourceFolder.Fetch(new[] { messageId }, MessageSummaryItems.Flags);

					if (v_IMAPOperationType == "Move MimeMessage")
					{
						if (v_IMAPMoveCopyUnreadOnly == "Yes")
						{
							if (!messageSummary[0].Flags.Value.HasFlag(MessageFlags.Seen))
								foundSourceFolder.MoveTo(messageId, foundDestinationFolder, cancel.Token);
						}
						else
							foundSourceFolder.MoveTo(messageId, foundDestinationFolder, cancel.Token);
					}
					else if (v_IMAPOperationType == "Copy MimeMessage")
					{
						if (v_IMAPMoveCopyUnreadOnly == "Yes")
						{
							if (!messageSummary[0].Flags.Value.HasFlag(MessageFlags.Seen))
								foundSourceFolder.CopyTo(messageId, foundDestinationFolder, cancel.Token);
						}
						else
							foundSourceFolder.CopyTo(messageId, foundDestinationFolder, cancel.Token);
					}

					client.Disconnect(true, cancel.Token);
					client.ServerCertificateValidationCallback = null;
				}
			} 
		}

		public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
		{
			base.Render(editor, commandControls);

			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_IMAPMimeMessage", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_IMAPHost", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_IMAPPort", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_IMAPUserName", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_IMAPPassword", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_IMAPDestinationFolder", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_IMAPOperationType", this, editor));
			RenderedControls.AddRange(commandControls.CreateDefaultDropdownGroupFor("v_IMAPMoveCopyUnreadOnly", this, editor));

			return RenderedControls;
		}

		public override string GetDisplayValue()
		{
			return base.GetDisplayValue() + $" [{v_IMAPOperationType} '{v_IMAPMimeMessage}' to '{v_IMAPDestinationFolder}']";
		}
	}
}