﻿using Autofac;
using OpenBots.Core.Command;
using OpenBots.Core.Enums;
using OpenBots.Core.Infrastructure;
using OpenBots.Core.Settings;
using OpenBots.Core.UI.Forms;
using OpenBots.Core.User32;
using OpenBots.Core.Utilities.CommonUtilities;
using OpenBots.Studio.Utilities;
using OpenBots.UI.Forms.ScriptBuilder_Forms;
using OpenBots.UI.Supplement_Forms;
using OpenBots.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using DPoint = System.Drawing.Point;

namespace OpenBots.UI.Forms.Supplement_Forms
{
    public partial class frmAdvancedUIElementRecorder : UIForm, IfrmAdvancedUIElementRecorder
    {
        public DataTable SearchParameters { get; set; }
        public string LastItemClicked { get; set; }
        public bool IsRecordingSequence { get; set; }
        public bool IsCommandItemSelected { get; set; }
        public string WindowName { get; set; }
        public frmScriptBuilder CallBackForm { get; set; }
        private List<ScriptCommand> _sequenceCommandList;

        private DataTable _parameterSettings;      
        private bool _isFirstRecordClick;
        private bool _isRecording;
        private bool _isHookStopped;
        private ApplicationSettings _appSettings;
        private Stopwatch _stopwatch;

        private string _recordingMessage = "Recording. Press F2 to save and close.";
        private string _errorMessage = "Error cloning element. Please Try Again.";

        private Point _lastClickedMouseCoordinates;
        private IContainer _container;

        public frmAdvancedUIElementRecorder(IContainer container)
        {
            _container = container;
            _appSettings = new ApplicationSettings();
            _appSettings = _appSettings.GetOrCreateApplicationSettings();
            _isFirstRecordClick = true;

            InitializeComponent();

            _stopwatch = new Stopwatch();
            _stopwatch.Start();
        }

        private void frmThickAppElementRecorder_Load(object sender, EventArgs e)
        {
            //create data source from windows
            cboWindowTitle.DataSource = CommonMethods.GetAvailableWindowNames();
            if (!string.IsNullOrEmpty(WindowName))
                cboWindowTitle.SelectedItem = WindowName;
        }

        private void pbRecord_Click(object sender, EventArgs e)
        {
            // this.WindowState = FormWindowState.Minimized;
            if (!_isRecording)
            {
                _isRecording = true;

                SearchParameters = NewSearchParameterDataTable();

                //clear all
                SearchParameters.Rows.Clear();

                //get window name and find window
                WindowName = cboWindowTitle.Text;
                IntPtr hWnd = User32Functions.FindWindow(WindowName);

                if (IsRecordingSequence && _isFirstRecordClick)
                {
                    _isFirstRecordClick = false;
                    _sequenceCommandList = new List<ScriptCommand>();

                    frmThickAppElementRecorderSettings settingsForm = new frmThickAppElementRecorderSettings();
                    settingsForm.ShowDialog();

                    if (settingsForm.DialogResult == DialogResult.OK)
                    {
                        _parameterSettings = settingsForm.ParameterSettingsDT;
                        settingsForm.Dispose();
                    }
                    else
                    {
                        _isRecording = false;
                        _isFirstRecordClick = true;

                        lblDescription.Text = "Instructions: Select the target window name from the drop-down " +
                                              "list and click the record button. Once recording has started, click " +
                                              "the element in the target application that you want to capture.";

                        //remove wait for left mouse down event
                        GlobalHook.MouseEvent -= GlobalHook_MouseEvent;
                        GlobalHook.KeyDownEvent -= GlobalHook_KeyDownEvent;
                        GlobalHook.HookStopped -= GlobalHook_HookStopped;

                        settingsForm.Dispose();
                        return;
                    }

                    dynamic activateWindowCommand = TypeMethods.CreateTypeInstance(_container, "ActivateWindowCommand");
                    activateWindowCommand.v_WindowName = $"\"{WindowName}\"";
                    _sequenceCommandList.Add(activateWindowCommand);
                }

                //check if window is found
                if (hWnd != IntPtr.Zero)
                {
                    //set window state and move to 0,0
                    User32Functions.ActivateWindow(WindowName);
                    User32Functions.SetWindowPosition(hWnd, 0, 0);

                    //start global hook and wait for left mouse down event
                    GlobalHook.StartEngineCancellationHook(Keys.F2);
                    GlobalHook.HookStopped += GlobalHook_HookStopped;
                    GlobalHook.StartElementCaptureHook(chkStopOnClick.Checked, _container);
                    GlobalHook.MouseEvent += GlobalHook_MouseEvent;
                    GlobalHook.KeyDownEvent += GlobalHook_KeyDownEvent;
                }

                if (!chkStopOnClick.Checked)
                {
                    lblDescription.Text = _recordingMessage;
                    MoveFormToBottomRight(this);
                }
                else
                    WindowState = FormWindowState.Minimized;
            }
            else
            {
                _isRecording = false;
                if (!chkStopOnClick.Checked)
                    lblDescription.Text = "Recording has stopped. Press F2 to save and close.";
            }
        }

        private void GlobalHook_HookStopped(object sender, EventArgs e)
        {
            _isHookStopped = true;
            Close();
        }

        private void GlobalHook_MouseEvent(object sender, MouseCoordinateEventArgs e)
        {
            //mouse down has occured
            if (e != null)
            {
                //invoke UIA
                try
                {                    
                    if (_isRecording)
                    {
                        var windowName = GetWindowName((int)e.MouseCoordinates.X, (int)e.MouseCoordinates.Y);

                        if (WindowName != windowName && !string.IsNullOrEmpty(windowName))
                        {
                            WindowName = windowName;
                            cboWindowTitle.SelectedItem = WindowName;

                            if (IsRecordingSequence)
                                BuildActivateWindowCommand();
                        }

                        LoadSearchParameters(e.MouseCoordinates);
                        lblDescription.Text = _recordingMessage;
                    }
                   
                    string clickType;
                    switch (e.MouseMessage)
                    {
                        case MouseMessages.WmLButtonDown:
                            if (e.MouseCoordinates.X == _lastClickedMouseCoordinates.X &&
                                e.MouseCoordinates.Y == _lastClickedMouseCoordinates.Y &&
                                _stopwatch.ElapsedMilliseconds <= 500)
                            {
                                _stopwatch.Stop();

                                //remove previous commands generated from single click
                                for (int i = 0; i < 2; i++)
                                    _sequenceCommandList.RemoveAt(_sequenceCommandList.Count - 1);

                                clickType = "Double Left Click";
                            }
                            else
                                clickType = "Left Click";
                            break;
                        case MouseMessages.WmMButtonDown:
                            clickType = "Middle Click";
                            break;
                        case MouseMessages.WmRButtonDown:
                            clickType = "Right Click";
                            break;
                        default:
                            clickType = "Left Click";
                            break;
                    }

                    if (IsRecordingSequence && _isRecording)
                        BuildElementClickActionCommand(clickType);

                    _lastClickedMouseCoordinates = e.MouseCoordinates;
                }
                catch (Exception)
                {
                    lblDescription.Text = _errorMessage;
                }
            }

            if (chkStopOnClick.Checked)
                Close();
        }

        private void GlobalHook_KeyDownEvent(object sender, KeyDownEventArgs e)
        {
            //key down has occured
            if (e != null)
            {
                //invoke UIA
                try
                {
                    if (_isRecording && e.Key.ToString() != GlobalHook.StopHookKey)
                    {
                        var windowName = GetWindowName((int)e.MouseCoordinates.X, (int)e.MouseCoordinates.Y);

                        if (WindowName != windowName && !string.IsNullOrEmpty(windowName))
                        {
                            WindowName = windowName;
                            cboWindowTitle.SelectedItem = WindowName;

                            if (IsRecordingSequence)
                                BuildActivateWindowCommand();
                        }
                     
                        LoadSearchParameters(e.MouseCoordinates);
                        lblDescription.Text = _recordingMessage;
                    }

                    if (IsRecordingSequence && _isRecording && e.Key.ToString() != GlobalHook.StopHookKey)
                    {
                        BuildElementSetTextActionCommand(e.Key);
                    }
                    else if (e.Key.ToString() == GlobalHook.StopHookKey)
                    {
                        //STOP HOOK
                        GlobalHook.StopHook();
                        GlobalHook.HookStopped -= GlobalHook_HookStopped;
                        return;
                    }
                }
                catch (Exception)
                {
                    lblDescription.Text = _errorMessage;
                }
            }
        }

        private string GetWindowName(int XCoordinate, int YCoordinate)
        {
            DPoint point = new DPoint(XCoordinate, YCoordinate);
            var window = User32Functions.WindowFromPoint(point);

            User32Functions.GetWindowThreadProcessId(window, out uint processId);
            Process process = Process.GetProcessById((int)processId);

            var windowName = process.MainWindowTitle;
            return windowName;
        }

        private void LoadSearchParameters(Point point)
        {
            AutomationElement element;

            try
            {
                element = AutomationElement.FromPoint(point);
            }
            catch (Exception)
            {
                Thread.Sleep(500);
                element = AutomationElement.FromPoint(point);
            }
          
            AutomationElement.AutomationElementInformation elementProperties = element.Current;

            LastItemClicked = $"[Automation ID:{elementProperties.AutomationId}].[Name:{elementProperties.Name}].[Class:{elementProperties.ClassName}]";
            lblSubHeader.Text = LastItemClicked;

            //get properties from class via reflection
            PropertyInfo[] properties = typeof(AutomationElement.AutomationElementInformation).GetProperties();
            Array.Sort(properties, (x, y) => string.Compare(x.Name, y.Name));

            SearchParameters = NewSearchParameterDataTable();

            if (IsRecordingSequence)
            {
                //loop through each property and get value from the element
                foreach (DataRow row in _parameterSettings.Rows)
                {
                    foreach (PropertyInfo property in properties)
                    {
                        try
                        {
                            var propName = property.Name;
                            var propValue = property.GetValue(elementProperties, null);

                            //if property is a basic type then display
                            if ((propName == row[1].ToString()) && ((propValue is string) || (propValue is bool) || (propValue is int) || (propValue is double)))
                                SearchParameters.Rows.Add(row[0], propName, $"\"{propValue}\"");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error Iterating over properties in window: " + ex.ToString());
                        }
                    }
                }
            }
            else
            {
                //loop through each property and get value from the element
                foreach (PropertyInfo property in properties)
                {
                    try
                    {
                        var propName = property.Name;
                        var propValue = property.GetValue(elementProperties, null);

                        //if property is a basic type then display
                        if ((propValue is string) || (propValue is bool) || (propValue is int) || (propValue is double))
                            SearchParameters.Rows.Add(false, propName, $"\"{propValue}\"");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error Iterating over properties in window: " + ex.ToString());
                    }
                }
            }
        }

        private void pbRefresh_Click(object sender, EventArgs e)
        {
            //handle window refresh requests
            cboWindowTitle.DataSource = CommonMethods.GetAvailableWindowNames();
        }

        private void uiBtnOk_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }

        private void uiBtnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void BuildActivateWindowCommand()
        {
            dynamic activateWindowCommand = TypeMethods.CreateTypeInstance(_container, "ActivateWindowCommand");
            activateWindowCommand.v_WindowName = $"\"{WindowName}\"";

            _sequenceCommandList.Add(activateWindowCommand);
        }

        private void BuildElementClickActionCommand(string clickType)
        {
            dynamic clickElementActionCommand = TypeMethods.CreateTypeInstance(_container, "UIAutomationCommand");
            clickElementActionCommand.v_WindowName = $"\"{WindowName}\"";
            clickElementActionCommand.v_UIASearchParameters = SearchParameters;
            clickElementActionCommand.v_AutomationType = "Click Element";

            DataTable webActionDT = clickElementActionCommand.v_UIAActionParameters;
            DataRow clickTypeRow = webActionDT.NewRow();
            clickTypeRow["Parameter Name"] = "Click Type";
            clickTypeRow["Parameter Value"] = clickType;
            webActionDT.Rows.Add(clickTypeRow);

            _sequenceCommandList.Add(clickElementActionCommand);

            _stopwatch.Restart();
        }

        private void BuildElementSetTextActionCommand(Keys key)
        {
            bool toUpperCase = false;

            //determine if casing is needed
            if (GlobalHook.IsKeyDown(Keys.ShiftKey) && GlobalHook.IsKeyToggled(Keys.Capital))
                toUpperCase = false;
            else if (!GlobalHook.IsKeyDown(Keys.ShiftKey) && GlobalHook.IsKeyToggled(Keys.Capital))
                toUpperCase = true;
            else if (GlobalHook.IsKeyDown(Keys.ShiftKey) && !GlobalHook.IsKeyToggled(Keys.Capital))
                toUpperCase = true;
            else if (!GlobalHook.IsKeyDown(Keys.ShiftKey) && !GlobalHook.IsKeyToggled(Keys.Capital))
                toUpperCase = false;

            var buf = new StringBuilder(256);
            var keyboardState = new byte[256];

            if (toUpperCase)
                keyboardState[(int)Keys.ShiftKey] = 0xff;

            GlobalHook.ToUnicode((uint)key, 0, keyboardState, buf, 256, 0);
            var selectedKey = buf.ToString();

            //translate key press to sendkeys identifier
            if (key.ToString() == GlobalHook.StopHookKey)
            {
                //STOP HOOK
                GlobalHook.StopHook();
                GlobalHook.HookStopped -= GlobalHook_HookStopped;
                GlobalHook.MouseEvent -= GlobalHook_MouseEvent;
                GlobalHook.KeyDownEvent -= GlobalHook_KeyDownEvent;
                return;
            }
            else
            {
                bool result = GlobalHook.BuildSendAdvancedKeystrokesCommand(key, _sequenceCommandList, WindowName);
                if (result) return;
            }

            //generate sendkeys together
            if ((_sequenceCommandList.Count > 1) && (_sequenceCommandList[_sequenceCommandList.Count - 1] is IUIAutomationCommand)
                && (_sequenceCommandList[_sequenceCommandList.Count - 1] as IUIAutomationCommand).v_AutomationType == "Set Text")
            {
                var lastCreatedSendKeysCommand = (IUIAutomationCommand)_sequenceCommandList[_sequenceCommandList.Count - 1];

                //append chars to previously created command
                //this makes editing easier for the user because only 1 command is issued rather than multiples
                var previouslyInputChars = lastCreatedSendKeysCommand.v_UIAActionParameters.Rows[0][1].ToString();
                lastCreatedSendKeysCommand.v_UIAActionParameters.Rows[0][1] = previouslyInputChars.Insert(previouslyInputChars.Length - 1, selectedKey);
            }
            else
            {
                //build keyboard command
                dynamic setTextElementActionCommand = TypeMethods.CreateTypeInstance(_container, "UIAutomationCommand");
                setTextElementActionCommand.v_WindowName = $"\"{WindowName}\"";
                setTextElementActionCommand.v_UIASearchParameters = SearchParameters;
                setTextElementActionCommand.v_AutomationType = "Set Text";
           
                DataTable webActionDT = setTextElementActionCommand.v_UIAActionParameters;
                DataRow textToSetRow = webActionDT.NewRow();
                textToSetRow["Parameter Name"] = "Text To Set";
                textToSetRow["Parameter Value"] = $"\"{selectedKey}\"";
                webActionDT.Rows.Add(textToSetRow);

                _sequenceCommandList.Add(setTextElementActionCommand);
            }
        }

        private void FinalizeRecording()
        {
            string sequenceComment = $"Advanced UI Sequence Recorded {DateTime.Now}";

            dynamic commentCommand = TypeMethods.CreateTypeInstance(_container, "AddCodeCommentCommand");
            commentCommand.v_Comment = sequenceComment;
            
            dynamic sequenceCommand = TypeMethods.CreateTypeInstance(_container, "SequenceCommand");
            sequenceCommand.ScriptActions = _sequenceCommandList;
            sequenceCommand.v_Comment = sequenceComment;

            CallBackForm.AddCommandToListView(commentCommand);
            CallBackForm.AddCommandToListView(sequenceCommand);
        }

        private void frmThickAppElementRecorder_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsRecordingSequence && _isHookStopped)
            {
                GlobalHook.HookStopped -= GlobalHook_HookStopped;               
                FinalizeRecording();
            }

            GlobalHook.MouseEvent -= GlobalHook_MouseEvent;
            GlobalHook.KeyDownEvent -= GlobalHook_KeyDownEvent;
            DialogResult = DialogResult.Cancel;
        }

        private DataTable NewSearchParameterDataTable()
        {
            DataTable searchParameters = new DataTable();
            searchParameters.Columns.Add("Enabled");
            searchParameters.Columns.Add("Parameter Name");
            searchParameters.Columns.Add("Parameter Value");
            searchParameters.TableName = DateTime.Now.ToString("UIASearchParamTable" + DateTime.Now.ToString("MMddyy.hhmmss"));
            return searchParameters;
        }
    }
}
