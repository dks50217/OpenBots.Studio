﻿using Autofac;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;
using Newtonsoft.Json;
using OpenBots.Core.App;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.IO;
using OpenBots.Core.Model.EngineModel;
using OpenBots.Core.Script;
using OpenBots.Core.Settings;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Engine.Enums;
using RestSharp;
using Serilog;
using Serilog.Core;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using OBScript = OpenBots.Core.Script.Script;
using OBScriptVariable = OpenBots.Core.Script.ScriptVariable;

namespace OpenBots.Engine
{
    public class AutomationEngineInstance : IAutomationEngineInstance
    {
        //engine variables
        public EngineContext AutomationEngineContext { get; set; } = new EngineContext();        
        public List<ScriptError> ErrorsOccured { get; set; }
        public string ErrorHandlingAction { get; set; }
        public bool ChildScriptFailed { get; set; }
        public bool ChildScriptErrorCaught { get; set; }
        public ScriptCommand LastExecutedCommand { get; set; }
        public bool IsCancellationPending { get; set; }
        public bool CurrentLoopCancelled { get; set; }
        public bool CurrentLoopContinuing { get; set; }
        public bool _isScriptPaused { get; private set; }
        private bool _isScriptSteppedOver { get; set; }
        private bool _isScriptSteppedInto { get; set; }
        private bool _isScriptSteppedOverBeforeException { get; set; }
        private bool _isScriptSteppedIntoBeforeException { get; set; }
        [JsonIgnore]
        private Stopwatch _stopWatch { get; set; }
        private EngineStatus _currentStatus { get; set; }
        public EngineSettings EngineSettings { get; set; }
        public string PrivateCommandLog { get; set; }
        public List<DataTable> DataTables { get; set; }
        public string FileName { get; set; }
        public bool IsServerExecution { get; set; }
        public bool IsServerChildExecution { get; set; }
        public List<IRestResponse> ServiceResponses { get; set; }
        public bool AutoCalculateVariables { get; set; }
        public string TaskResult { get; set; } = "";
        //events
        public event EventHandler<ReportProgressEventArgs> ReportProgressEvent;
        public event EventHandler<ScriptFinishedEventArgs> ScriptFinishedEvent;
        public event EventHandler<LineNumberChangedEventArgs> LineNumberChangedEvent;

        public AutomationEngineInstance(EngineContext engineContext)
        {
            if (engineContext != null)
                AutomationEngineContext = engineContext;

            //initialize logger
            if (AutomationEngineContext.EngineLogger != null)
            {
                Log.Logger = AutomationEngineContext.EngineLogger;
                Log.Information("Engine Class has been initialized");
            }
            
            PrivateCommandLog = "Can't log display value as the command contains sensitive data";

            //initialize error tracking list
            ErrorsOccured = new List<ScriptError>();

            //set to initialized
            _currentStatus = EngineStatus.Loaded;

            //get engine settings
            var settings = new ApplicationSettings().GetOrCreateApplicationSettings();
            EngineSettings = settings.EngineSettings;

            if (AutomationEngineContext.Variables == null)
                AutomationEngineContext.Variables = new List<OBScriptVariable>();

            if (AutomationEngineContext.Arguments == null)
                AutomationEngineContext.Arguments = new List<ScriptArgument>();

            if (AutomationEngineContext.Elements == null)
                AutomationEngineContext.Elements = new List<ScriptElement>();

            if (AutomationEngineContext.AppInstances == null)
                AutomationEngineContext.AppInstances = new Dictionary<string, object>();

            if (AutomationEngineContext.ImportedNamespaces == null)
                AutomationEngineContext.ImportedNamespaces = new Dictionary<string, AssemblyReference>(ScriptDefaultNamespaces.DefaultNamespaces);

            ServiceResponses = new List<IRestResponse>();
            DataTables = new List<DataTable>();

            //this value can be later overriden by script
            AutoCalculateVariables = EngineSettings.AutoCalcVariables;

            ErrorHandlingAction = string.Empty;

            //initialize roslyn instance
            AutomationEngineContext.AssembliesList = NamespaceMethods.GetAssemblies(this);
            AutomationEngineContext.NamespacesList = NamespaceMethods.GetNamespacesList(this);

            AutomationEngineContext.EngineScript = CSharpScript.Create("", ScriptOptions.Default.WithReferences(AutomationEngineContext.AssembliesList)
                                                                                                .WithImports(AutomationEngineContext.NamespacesList));
            AutomationEngineContext.EngineScriptState = null;
        }

        public IAutomationEngineInstance CreateAutomationEngineInstance(EngineContext engineContext)
        {
            return new AutomationEngineInstance(engineContext);
        }

        public void ExecuteScriptSync()
        {
            Log.Information("Client requesting to execute script independently");
            IsServerExecution = true;
            ExecuteScript(true);
        }

        public void ExecuteScriptAsync()
        {
            Log.Information("Client requesting to execute script using frmEngine");

            AutomationEngineContext.ScriptEngine = AutomationEngineContext.ScriptEngine;

            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                ExecuteScript(true);
            }).Start();
        }

        public void ExecuteScriptAsync(string filePath, string projectPath)
        {
            Log.Information("Client requesting to execute script independently");

            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;

                AutomationEngineContext.FilePath = filePath;
                AutomationEngineContext.ProjectPath = projectPath;

                ExecuteScript(true);
            }).Start();
        }

        public void ExecuteScriptJson()
        {
            Log.Information("Client requesting to execute script independently");

            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                ExecuteScript(false);
            }).Start();
        }

        private async void ExecuteScript(bool dataIsFile)
        {
            try
            {
                _currentStatus = EngineStatus.Running;

                //create stopwatch for metrics tracking
                _stopWatch = new Stopwatch();
                _stopWatch.Start();

                //log starting
                ReportProgress("Bot Engine Started: " + DateTime.Now.ToString());

                //get automation script
                OBScript automationScript;
                if (dataIsFile)
                {
                    ReportProgress("Deserializing File");
                    Log.Information("Script Path: " + AutomationEngineContext.FilePath);
                    FileName = AutomationEngineContext.FilePath;
                    automationScript = OBScript.DeserializeFile(AutomationEngineContext);
                }
                else
                {
                    ReportProgress("Deserializing JSON");
                    automationScript = OBScript.DeserializeJsonString(AutomationEngineContext.FilePath);
                }
                
                ReportProgress("Creating Variable List");

                //set variables if they were passed in
                if (AutomationEngineContext.Variables != null)
                {
                    foreach (var var in AutomationEngineContext.Variables)
                    {
                        var variableFound = automationScript.Variables.Where(f => f.VariableName == var.VariableName).FirstOrDefault();

                        if (variableFound != null)
                        {
                            variableFound.VariableValue = var.VariableValue;
                        }
                    }
                }

                AutomationEngineContext.Variables = automationScript.Variables;
                
                //update ProjectPath variable
                var projectPathVariable = AutomationEngineContext.Variables.Where(v => v.VariableName == "ProjectPath").SingleOrDefault();
                if (projectPathVariable != null)
                    projectPathVariable.VariableValue = "@\"" + AutomationEngineContext.ProjectPath + '"';
                else
                {
                    projectPathVariable = new OBScriptVariable
                    {
                        VariableName = "ProjectPath",
                        VariableType = typeof(string),
                        VariableValue = "@\"" + AutomationEngineContext.ProjectPath + '"'
                    };
                    AutomationEngineContext.Variables.Add(projectPathVariable);
                }

                foreach (OBScriptVariable var in AutomationEngineContext.Variables)
                {
                    dynamic evaluatedValue = await VariableMethods.InstantiateVariable(var.VariableName, (string)var.VariableValue, var.VariableType, this);
                    VariableMethods.SetVariableValue(evaluatedValue, this, var.VariableName);
                }           

                ReportProgress("Creating Argument List");

                //set arguments if they were passed in
                if (AutomationEngineContext.Arguments != null)
                {
                    foreach (var arg in AutomationEngineContext.Arguments)
                    {
                        var argumentFound = automationScript.Arguments.Where(f => f.ArgumentName == arg.ArgumentName).FirstOrDefault();

                        if (argumentFound != null)
                        {
                            argumentFound.ArgumentValue = arg.ArgumentValue;
                        }
                    }
                }

                AutomationEngineContext.Arguments = automationScript.Arguments;
                
                //used by RunTaskCommand to assign parent values to child arguments 
                if(AutomationEngineContext.IsChildEngine)
                {
                    foreach (ScriptArgument arg in AutomationEngineContext.Arguments)
                    {
                        await VariableMethods.InstantiateVariable(arg.ArgumentName, "", arg.ArgumentType, this);
                        VariableMethods.SetVariableValue(arg.ArgumentValue, this, arg.ArgumentName);
                    }
                }
                else
                {
                    foreach (ScriptArgument arg in AutomationEngineContext.Arguments)
                    {
                        dynamic evaluatedValue = await VariableMethods.InstantiateVariable(arg.ArgumentName, (string)arg.ArgumentValue, arg.ArgumentType, this);
                        VariableMethods.SetVariableValue(evaluatedValue, this, arg.ArgumentName);
                    }
                }

                ReportProgress("Creating Element List");

                //set elements if they were passed in
                if (AutomationEngineContext.Elements != null)
                {
                    foreach (var elem in AutomationEngineContext.Elements)
                    {
                        var elementFound = automationScript.Elements.Where(f => f.ElementName == elem.ElementName).FirstOrDefault();

                        if (elementFound != null)
                            elementFound.ElementValue = elem.ElementValue;
                    }
                }

                AutomationEngineContext.Elements = automationScript.Elements;

                ReportProgress("Creating App Instance Tracking List");
                //create app instances and merge in global instances
                if (AutomationEngineContext.AppInstances == null)
                {
                    AutomationEngineContext.AppInstances = new Dictionary<string, object>();
                }
                var GlobalInstances = GlobalAppInstances.GetInstances();
                foreach (var instance in GlobalInstances)
                {
                    AutomationEngineContext.AppInstances[instance.Key] = instance.Value;
                }

                if (AutomationEngineContext.ImportedNamespaces == null)
                {
                    AutomationEngineContext.ImportedNamespaces = new Dictionary<string, AssemblyReference>(ScriptDefaultNamespaces.DefaultNamespaces);
                }

                //execute commands
                ScriptAction startCommand = automationScript.Commands.Where(x => x.ScriptCommand.LineNumber <= AutomationEngineContext.StartFromLineNumber)
                                                                         .Last();

                int startCommandIndex = automationScript.Commands.FindIndex(x => x.ScriptCommand.LineNumber == startCommand.ScriptCommand.LineNumber);

                while (startCommandIndex < automationScript.Commands.Count)
                {
                    if (IsCancellationPending)
                    {
                        ReportProgress("Cancelling Script");
                        ScriptFinished(ScriptFinishedResult.Cancelled);
                        return;
                    }

                    await ExecuteCommand(automationScript.Commands[startCommandIndex]);
                    startCommandIndex++;
                }

                if (IsCancellationPending)
                {
                    //mark cancelled - handles when cancelling and user defines 1 parent command or else it will show successful
                    ScriptFinished(ScriptFinishedResult.Cancelled);
                }
                else
                {
                    //mark finished
                    ScriptFinished(ScriptFinishedResult.Successful);
                }
            }
            catch (Exception ex)
            {
                ScriptFinished(ScriptFinishedResult.Error, ex.ToString());
            }
            if((AutomationEngineContext.ScriptEngine != null && !AutomationEngineContext.ScriptEngine.IsChildEngine) || (IsServerExecution && !IsServerChildExecution))
                AutomationEngineContext.EngineLogger.Dispose();
        }

        public async Task ExecuteCommand(ScriptAction command)
        {
            //get command
            ScriptCommand parentCommand = command.ScriptCommand;

            if (parentCommand == null)
                return;

            //in RunFromThisCommand exection, determine if/loop logic. If logic returns true, skip until reaching the selected command
            if (!parentCommand.ScopeStartCommand && parentCommand.LineNumber < AutomationEngineContext.StartFromLineNumber)
                return;
            //if the selected command is within a loop/retry, reset starting line number so that previous commands within the scope run in the following iteration
            else if (!parentCommand.ScopeStartCommand && parentCommand.LineNumber == AutomationEngineContext.StartFromLineNumber)
                AutomationEngineContext.StartFromLineNumber = 1;

            if (AutomationEngineContext.ScriptEngine != null && (parentCommand.CommandName == "RunTaskCommand" || parentCommand.CommandName == "ShowMessageCommand"))
                parentCommand.CurrentScriptBuilder = AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder;

            //set LastCommandExecuted
            LastExecutedCommand = command.ScriptCommand;

            //update execution line numbers
            LineNumberChanged(parentCommand.LineNumber);

            //handle pause request
            if (AutomationEngineContext.ScriptEngine != null && parentCommand.PauseBeforeExecution && AutomationEngineContext.ScriptEngine.IsDebugMode && !ChildScriptFailed)
            {
                ReportProgress("Pausing Before Execution");
                _isScriptPaused = true;
                AutomationEngineContext.ScriptEngine.IsHiddenTaskEngine = false;
            }

            //handle pause
            bool isFirstWait = true;
            while (_isScriptPaused)
            {
                //only show pause first loop
                if (isFirstWait)
                {
                    _currentStatus = EngineStatus.Paused;
                    ReportProgress("Paused on Line " + parentCommand.LineNumber + ": "
                        + (parentCommand.v_IsPrivate ? PrivateCommandLog : parentCommand.GetDisplayValue()));
                    ReportProgress("[Please select 'Resume' when ready]");
                    isFirstWait = false;
                }

                if (_isScriptSteppedInto && parentCommand.CommandName == "RunTaskCommand")
                {
                    parentCommand.IsSteppedInto = true;
                    parentCommand.CurrentScriptBuilder = AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder;
                    _isScriptSteppedInto = false;
                    AutomationEngineContext.ScriptEngine.IsHiddenTaskEngine = true;
                    
                    break;
                }
                else if (_isScriptSteppedOver || _isScriptSteppedInto)
                {
                    _isScriptSteppedOver = false;
                    _isScriptSteppedInto = false;
                    break;
                }

                if (((Form)AutomationEngineContext.ScriptEngine).IsDisposed)
                {
                    IsCancellationPending = true;
                    break;
                }
                                  
                //wait
                Thread.Sleep(1000);
            }

            _currentStatus = EngineStatus.Running;

            //handle if cancellation was requested
            if (IsCancellationPending)
            {
                return;
            }

            //If Child Script Failed and Child Script Error not Caught, next command should not be executed
            if (ChildScriptFailed && !ChildScriptErrorCaught)
                throw new Exception("Child Script Failed");

            //bypass comments
            if (parentCommand.CommandName == "AddCodeCommentCommand" || parentCommand.CommandName == "BrokenCodeCommentCommand" || parentCommand.IsCommented)             
                return;

            //report intended execution
            if (parentCommand.CommandName != "LogMessageCommand")
                ReportProgress($"Running Line {parentCommand.LineNumber}: {(parentCommand.v_IsPrivate ? PrivateCommandLog : parentCommand.GetDisplayValue())}");

            //handle any errors
            try
            {
                //determine type of command
                if ((parentCommand.CommandName == "LoopNumberOfTimesCommand") || (parentCommand.CommandName == "LoopContinuouslyCommand") ||
                    (parentCommand.CommandName == "LoopCollectionCommand") || (parentCommand.CommandName == "BeginIfCommand") ||
                    (parentCommand.CommandName == "BeginMultiIfCommand") || (parentCommand.CommandName == "BeginTryCommand") ||
                    (parentCommand.CommandName == "BeginLoopCommand") || (parentCommand.CommandName == "BeginMultiLoopCommand") ||
                    (parentCommand.CommandName == "BeginRetryCommand" || (parentCommand.CommandName == "BeginSwitchCommand")))
                {
                    //run the command and pass bgw/command as this command will recursively call this method for sub commands
                    //TODO: Make sure that removing these lines doesn't create any other issues
                    //command.IsExceptionIgnored = true;
                    await parentCommand.RunCommand(this, command);
                }
                else if (parentCommand.CommandName == "SequenceCommand")
                {
                    //command.IsExceptionIgnored = true;
                    await parentCommand.RunCommand(this, command);
                }
                else if (parentCommand.CommandName == "StopCurrentTaskCommand")
                {
                    if (AutomationEngineContext.ScriptEngine != null && AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder != null)
                        AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.IsScriptRunning = false;

                    IsCancellationPending = true;
                    return;
                }
                else if (parentCommand.CommandName == "ExitLoopCommand")
                {
                    CurrentLoopCancelled = true;
                }
                else if (parentCommand.CommandName == "NextLoopCommand")
                {
                    CurrentLoopContinuing = true;
                }
                else
                {
                    //sleep required time
                    Thread.Sleep(EngineSettings.DelayBetweenCommands);

                    if (!parentCommand.v_ErrorHandling.Equals("None"))
                        ErrorHandlingAction = parentCommand.v_ErrorHandling;
                    else
                        ErrorHandlingAction = string.Empty;

                    //run the command
                    try
                    {
                        await parentCommand.RunCommand(this);
                    }
                    catch (Exception ex)
                    {
                        switch (ErrorHandlingAction)
                        {
                            case "Ignore Error":
                                ReportProgress("Error Occured at Line " + parentCommand.LineNumber + ":" + ex.ToString(), Enum.GetName(typeof(LogEventLevel), LogEventLevel.Error));
                                ReportProgress("Ignoring Per Error Handling");
                                break;
                            case "Report Error":
                                ReportProgress("Error Occured at Line " + parentCommand.LineNumber + ":" + ex.ToString(), Enum.GetName(typeof(LogEventLevel), LogEventLevel.Error));
                                ReportProgress("Handling Error and Attempting to Continue");
                                throw ex;
                            default:
                                throw ex;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                if (!(LastExecutedCommand.CommandName == "RethrowCommand"))
                {
                    if (ChildScriptFailed)
                    {
                        ChildScriptFailed = false;
                        ErrorsOccured.Clear();
                    }

                    ErrorsOccured.Add(new ScriptError()
                    {
                        SourceFile = FileName,
                        LineNumber = parentCommand.LineNumber,
                        StackTrace = ex.ToString(),
                        ErrorType = ex.GetType().Name,
                        ErrorMessage = ex.Message
                    });
                }

                var error = ErrorsOccured.OrderByDescending(x => x.LineNumber).FirstOrDefault();
                string errorMessage = $"Source: {error.SourceFile}, Line: {error.LineNumber} {parentCommand.GetDisplayValue()}, " +
                        $"Exception Type: {error.ErrorType}, Exception Message: {error.ErrorMessage}";


                if (AutomationEngineContext.ScriptEngine != null && !command.IsExceptionIgnored && AutomationEngineContext.ScriptEngine.IsDebugMode)
                {
                    //load error form if exception is not handled
                    AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.IsUnhandledException = true;
                    AutomationEngineContext.ScriptEngine.AddStatus("Pausing Before Exception");

                    DialogResult result = DialogResult.OK;
                    if (ErrorHandlingAction != "Ignore Error")
                        result = AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.LoadErrorForm(errorMessage);
                    ReportProgress("Error Occured at Line " + parentCommand.LineNumber + ":" + ex.ToString(), Enum.GetName(typeof(LogEventLevel), LogEventLevel.Debug));
                    AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.IsUnhandledException = false;

                    if (result == DialogResult.OK)
                    {                           
                        ReportProgress("Ignoring Per User Choice");
                        ErrorsOccured.Clear();

                        if (_isScriptSteppedIntoBeforeException)
                        {
                            AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.IsScriptSteppedInto = true;
                            _isScriptSteppedIntoBeforeException = false;
                        }
                        else if (_isScriptSteppedOverBeforeException)
                        {
                            AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.IsScriptSteppedOver = true;
                            _isScriptSteppedOverBeforeException = false;
                        }

                        AutomationEngineContext.ScriptEngine.uiBtnPause_Click(null, null);
                    }
                    else if (result == DialogResult.Abort || result == DialogResult.Cancel)
                    {
                        ReportProgress("Continuing Per User Choice");
                        AutomationEngineContext.ScriptEngine.ScriptEngineContext.ScriptBuilder.RemoveDebugTab();
                        AutomationEngineContext.ScriptEngine.uiBtnPause_Click(null, null);                           
                        throw ex;
                    }
                    //TODO: Add Break Option
                }
                else
                    throw ex;
                
            }
        }     

        public void CancelScript()
        {
            IsCancellationPending = true;
        }

        public void PauseScript()
        {
            _isScriptPaused = true;
        }

        public void ResumeScript()
        {
            _isScriptPaused = false;
        }

        public void StepOverScript()
        {
            _isScriptSteppedOver = true;
            _isScriptSteppedOverBeforeException = true;
        }

        public void StepIntoScript()
        {
            _isScriptSteppedInto = true;
            _isScriptSteppedIntoBeforeException = true;
        }

        public virtual void ReportProgress(string progress, string eventLevel = "Information")
        {
            ReportProgressEventArgs args = new ReportProgressEventArgs();
            LogEventLevel logEventLevel = (LogEventLevel)Enum.Parse(typeof(LogEventLevel), eventLevel);

            switch (logEventLevel)
            {
                case LogEventLevel.Verbose:
                    Log.Verbose(progress);
                    args.LoggerColor = Color.Purple;
                    break;
                case LogEventLevel.Debug:
                    Log.Debug(progress);
                    args.LoggerColor = Color.Green;
                    break;
                case LogEventLevel.Information:
                    Log.Information(progress);
                    args.LoggerColor = SystemColors.Highlight;
                    break;
                case LogEventLevel.Warning:
                    Log.Warning(progress);
                    args.LoggerColor = Color.Goldenrod;
                    break;
                case LogEventLevel.Error:
                    Log.Error(progress);
                    args.LoggerColor = Color.Red;
                    break;
                case LogEventLevel.Fatal:
                    Log.Fatal(progress);
                    args.LoggerColor = Color.Black;
                    break;
            }

            if (progress.StartsWith("Skipping"))
                args.LoggerColor = Color.Green;
             
            args.ProgressUpdate = progress;

            //invoke event
            ReportProgressEvent?.Invoke(this, args);
        }

        public virtual void ScriptFinished(ScriptFinishedResult result, string error = null)
        {
            if (ChildScriptFailed && !ChildScriptErrorCaught)
            {
                error = "Terminate with failure";
                result = ScriptFinishedResult.Error;
                Log.Fatal("Result Code: " + result.ToString());
            }
            else
            {
                Log.Information("Result Code: " + result.ToString());
            }

            //add result variable if missing
            var resultVar = AutomationEngineContext.Variables.Where(f => f.VariableName == "OpenBots.Result").FirstOrDefault();

            //handle if variable is missing
            if (resultVar == null)
            {
                resultVar = new OBScriptVariable() { VariableName = "OpenBots.Result", VariableValue = "" };
            }

            //check value
            var resultValue = resultVar.VariableValue.ToString();

            if (error == null)
            {
                Log.Information("Error: None");

                if (string.IsNullOrEmpty(resultValue))
                {
                    TaskResult = "Successfully Completed Script";
                }
                else
                {
                    TaskResult = resultValue;
                }
            }

            else
            {
                if (ErrorsOccured.Count > 0)
                    error = ErrorsOccured.OrderByDescending(x => x.LineNumber).FirstOrDefault().StackTrace;

                Log.Error("Error: " + error);
                TaskResult = error;
            }

            if (AutomationEngineContext.ScriptEngine != null && !AutomationEngineContext.ScriptEngine.IsChildEngine)
                Log.CloseAndFlush();

            if (IsServerExecution && !IsServerChildExecution)
                Log.CloseAndFlush();

            _currentStatus = EngineStatus.Finished;
            ScriptFinishedEventArgs args = new ScriptFinishedEventArgs
            {
                LoggedOn = DateTime.Now,
                Result = result,
                Error = error,
                ExecutionTime = _stopWatch.Elapsed,
                FileName = FileName
            };

            //convert to json
            var serializedArguments = JsonConvert.SerializeObject(args);

            //write execution metrics
            if (EngineSettings.TrackExecutionMetrics && (FileName != null))
            {
                string summaryLoggerFilePath = Path.Combine(Folders.GetFolder(FolderType.LogFolder), "OpenBots Execution Summary Logs.txt");
                Logger summaryLogger = new Logging().CreateJsonFileLogger(summaryLoggerFilePath, Serilog.RollingInterval.Infinite);
                summaryLogger.Information(serializedArguments);
                if (AutomationEngineContext.ScriptEngine != null && !AutomationEngineContext.ScriptEngine.IsChildEngine)
                    summaryLogger.Dispose();
            }

            ScriptFinishedEvent?.Invoke(this, args);
        }

        public virtual void LineNumberChanged(int lineNumber)
        {
            LineNumberChangedEventArgs args = new LineNumberChangedEventArgs
            {
                CurrentLineNumber = lineNumber
            };
            LineNumberChangedEvent?.Invoke(this, args);
        }

        public string GetEngineContext()
        {
            //set json settings
            JsonSerializerSettings settings = new JsonSerializerSettings
            {
                Error = (serializer, err) =>
                {
                    err.ErrorContext.Handled = true;                    
                },
                Formatting = Formatting.Indented
            };

            return  JsonConvert.SerializeObject(this, settings);
        }

        public string GetProjectPath()
        {
            string projectPath = string.Empty;
            var projectPathVariable = AutomationEngineContext.Variables.Where(v => v.VariableName == "ProjectPath").SingleOrDefault();
            if (projectPathVariable != null)
                projectPath = projectPathVariable.VariableValue.ToString();

            return projectPath;
        }
    }
}
