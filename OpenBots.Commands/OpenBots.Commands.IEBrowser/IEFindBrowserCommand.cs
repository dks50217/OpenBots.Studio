﻿using mshtml;
using Newtonsoft.Json;
using OpenBots.Core.Attributes.PropertyAttributes;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using SHDocVw;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Windows.Forms;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Core.Properties;
using System.Threading.Tasks;
using OpenBots.Core.Model.ApplicationModel;

namespace OpenBots.Commands.IEBrowser
{
    [Serializable]
    [Category("IE Browser Commands")]
    [Description("This command finds and attaches to an existing IE Web Browser session.")]
    public class IEFindBrowserCommand : ScriptCommand
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
        [DisplayName("Browser Name (Title)")]
        [Description("Select the Name (Title) of the IE Browser Instance to get attached to.")]
        [SampleUsage("\"OpenBots\"")]
        [Remarks("")]
        [Editor("ShowVariableHelper", typeof(UIAdditionalHelperType))]
        [CompatibleTypes(new Type[] { typeof(string) })]
        public string v_IEBrowserName { get; set; }

        [JsonIgnore]
        [Browsable(false)]
        private ComboBox _ieBrowerNameDropdown;

        public IEFindBrowserCommand()
        {
            CommandName = "IEFindBrowserCommand";
            SelectionName = "Find IE Browser";          
            CommandEnabled = true;
            CommandIcon = Resources.command_web;

            v_InstanceName = "DefaultIEBrowser";
        }

        public async override Task RunCommand(object sender)
        {
            var engine = (IAutomationEngineInstance)sender;

            string IEBrowserName = (string)await v_IEBrowserName.EvaluateCode(engine);
            bool browserFound = false;
            var shellWindows = new ShellWindows();
            foreach (IWebBrowser2 shellWindow in shellWindows)
            {
                if ((shellWindow.Document is HTMLDocument) && (IEBrowserName == null || shellWindow.Document.Title == IEBrowserName))
                {
                    new OBAppInstance(v_InstanceName, (object)shellWindow.Application).SetVariableValue(engine, v_InstanceName);

                    browserFound = true;
                    break;
                }
            }

            //try partial match
            if (!browserFound)
            {
                foreach (IWebBrowser2 shellWindow in shellWindows)
                {
                    if ((shellWindow.Document is HTMLDocument) && 
                        ((shellWindow.Document.Title.Contains(IEBrowserName) || 
                        shellWindow.Document.Url.Contains(IEBrowserName))))
                    {
                        new OBAppInstance(v_InstanceName, (object)shellWindow.Application).SetVariableValue(engine, v_InstanceName);

                        browserFound = true;
                        break;
                    }
                }
            }

            if (!browserFound)
            {
                throw new Exception("Browser was not found!");
            }
        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            RenderedControls.AddRange(commandControls.CreateDefaultInputGroupFor("v_InstanceName", this, editor));

            _ieBrowerNameDropdown = commandControls.CreateDropdownFor("v_IEBrowserName", this);
            var shellWindows = new ShellWindows();
            foreach (IWebBrowser2 shellWindow in shellWindows)
            {
                if (shellWindow.Document is HTMLDocument)
                    _ieBrowerNameDropdown.Items.Add($"\"{shellWindow.Document.Title}\"");
            }
            RenderedControls.Add(commandControls.CreateDefaultLabelFor("v_IEBrowserName", this));
            RenderedControls.AddRange(commandControls.CreateUIHelpersFor("v_IEBrowserName", this, new Control[] { _ieBrowerNameDropdown }, editor));
            RenderedControls.Add(_ieBrowerNameDropdown);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue() + $" [Having Title '{v_IEBrowserName}' - Instance Name '{v_InstanceName}']";
        }
    }

}