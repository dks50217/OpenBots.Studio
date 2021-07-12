﻿using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Interfaces;
using OpenBots.Core.Model.ApplicationModel;
using OpenBots.Core.Properties;
using OpenBots.Core.Utilities.CommonUtilities;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OpenBots.Commands.IEBrowser
{
    [Serializable]
    [Category("IE Browser Commands")]
    [Description("This command creates a new IE Web Browser Session.")]
    public class IECreateBrowserCommand : ScriptCommand
    {
        [Required]
        [DisplayName("IE Browser Instance Name")]
        [Description("Enter a unique name that will represent the application instance.")]
        [SampleUsage("MyIEBrowserInstance")]
        [Remarks("This unique name allows you to refer to the instance by name in future commands, " +
                 "ensuring that the commands you specify run against the correct application.")]
		[Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
        [CompatibleTypes(new Type[] { typeof(OBAppInstance) })]
        public string v_InstanceName { get; set; }

        [Required]
        [DisplayName("URL")]
        [Description("Enter a Web URL to navigate to.")]
        [SampleUsage("\"https://example.com/\" || vURL")]
        [Remarks("This input is optional.")]
        [Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
        [CompatibleTypes(new Type[] { typeof(string) })]
        public string v_URL { get; set; }

        public IECreateBrowserCommand()
        {
            CommandName = "IECreateBrowserCommand";
            SelectionName = "Create IE Browser";           
            CommandEnabled = true;
            CommandIcon = Resources.command_web;
        }

        public async override Task RunCommand(object sender)
        {
            var engine = (IAutomationEngineInstance)sender;
            var webURL = (string)await v_URL.EvaluateCode(engine);

            InternetExplorer newBrowserSession = new InternetExplorer();

            if (!string.IsNullOrEmpty(webURL.Trim()))
            {
                try
                {
                    newBrowserSession.Navigate(webURL);
                    WaitForReadyState(newBrowserSession);
                    newBrowserSession.Visible = true;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
                
            //add app instance
            new OBAppInstance(v_InstanceName, newBrowserSession).SetVariableValue(engine, v_InstanceName);
        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));
            RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_URL", this, editor));

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + $" [New Instance Name '{v_InstanceName}']";
        }

        public static void WaitForReadyState(InternetExplorer ieInstance)
        {
            try
            {
                DateTime waitExpires = DateTime.Now.AddSeconds(15);

                do
                {
                    Thread.Sleep(500);
                }

                while ((ieInstance.ReadyState != tagREADYSTATE.READYSTATE_COMPLETE) && (waitExpires > DateTime.Now));
            }
            catch (Exception ex) 
            {
                throw ex;
            }
        }
    }
}