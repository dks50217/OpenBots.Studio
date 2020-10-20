﻿using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.ComponentModel;
using OpenBots.Core.Command;
using OpenBots.Core.Infrastructure;

namespace OpenBots.Commands
{
    [Serializable]
    [Category("Error Handling Commands")]
    [Description("This command rethrows an exception caught in a catch block.")]
    public class RethrowCommand : ScriptCommand
    {
        public RethrowCommand()
        {
            CommandName = "RethrowCommand";
            SelectionName = "Rethrow";
            CommandEnabled = true;          
        }

        public override void RunCommand(object sender)
        {
            throw new Exception("Rethrowing Original Exception");
        }

        public override List<Control> Render(IfrmCommandEditor editor, ICommandControls commandControls)
        {
            base.Render(editor, commandControls);

            return RenderedControls;
        }

        public override string GetDisplayValue()
        {
            return base.GetDisplayValue();
        }
    }
}