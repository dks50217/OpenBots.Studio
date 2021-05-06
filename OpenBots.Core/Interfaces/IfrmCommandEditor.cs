﻿using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Model.EngineModel;
using OpenBots.Core.Script;
using OpenBots.Core.UI.Controls;
using System.Collections.Generic;

namespace OpenBots.Core.Infrastructure
{
    public interface IfrmCommandEditor
    {
        List<AutomationCommand> CommandList { get; set; }
        EngineContext ScriptEngineContext { get; set; }       
        ScriptCommand SelectedCommand { get; set; }
        ScriptCommand OriginalCommand { get; set; }
        CreationMode CreationModeInstance { get; set; }
        string DefaultStartupCommand { get; set; }
        ScriptCommand EditingCommand { get; set; }
        List<ScriptCommand> ConfiguredCommands { get; set; }
        string HTMLElementRecorderURL { get; set; }
        TypeContext TypeContext { get; set; }

        void ShowMessage(string message, string title, DialogType dialogType, int closeAfter);
    }
}
