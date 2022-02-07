//Copyright (c) 2019 Jason Bayldon
//Modifications - Copyright (c) 2020 OpenBots Inc.
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//   http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
using OpenBots.Core.Enums;
using OpenBots.Core.IO;
using OpenBots.Core.Metrics;
using OpenBots.Core.Settings;
using OpenBots.Core.UI.Forms;
using OpenBots.Studio.Utilities.Documentation;
using OpenBots.UI.Forms.Supplement_Forms;
using OpenBots.Utilities;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using AContainer = Autofac.IContainer;

namespace OpenBots.UI.Forms
{
    public partial class frmSettings : UIForm
    {
        private ApplicationSettings _newAppSettings;
        private AContainer _container;
        private string startingRootFolder;

        public frmSettings(AContainer container)
        {
            InitializeComponent();
            _container = container;
        }

        private void frmSettings_Load(object sender, EventArgs e)
        {
            if (_container == null)
                btnGenerateWikiDocs.Enabled = false;

            _newAppSettings = new ApplicationSettings().GetOrCreateApplicationSettings();

            var engineSettings = _newAppSettings.EngineSettings;
            chkShowDebug.DataBindings.Add("Checked", engineSettings, "ShowDebugWindow", false, DataSourceUpdateMode.OnPropertyChanged);
            chkAutoCloseWindow.DataBindings.Add("Checked", engineSettings, "AutoCloseDebugWindow", false, DataSourceUpdateMode.OnPropertyChanged);
            chkAdvancedDebug.DataBindings.Add("Checked", engineSettings, "ShowAdvancedDebugOutput", false, DataSourceUpdateMode.OnPropertyChanged);
            chkTrackMetrics.DataBindings.Add("Checked", engineSettings, "TrackExecutionMetrics", false, DataSourceUpdateMode.OnPropertyChanged);
            txtCommandDelay.DataBindings.Add("Text", engineSettings, "DelayBetweenCommands", false, DataSourceUpdateMode.OnPropertyChanged);

            Keys cancellationKey = engineSettings.CancellationKey;
            cbxCancellationKey.DataSource = Enum.GetValues(typeof(Keys));
            cbxCancellationKey.SelectedIndex = cbxCancellationKey.Items.IndexOf(cancellationKey);

            SinkType loggingSinkType = engineSettings.LoggingSinkType;
            cbxSinkType.DataSource = Enum.GetValues(typeof(SinkType));           
            cbxSinkType.SelectedIndex = cbxSinkType.Items.IndexOf(loggingSinkType);

            LogEventLevel minLogLevel = engineSettings.MinLogLevel;
            cbxMinLogLevel.DataSource = Enum.GetValues(typeof(LogEventLevel));
            cbxMinLogLevel.SelectedIndex = cbxMinLogLevel.Items.IndexOf(minLogLevel);

            txtLoggingValue.DataBindings.Add("Text", engineSettings, "LoggingValue", false, DataSourceUpdateMode.OnPropertyChanged);

            var clientSettings = _newAppSettings.ClientSettings;
            chkAntiIdle.DataBindings.Add("Checked", clientSettings, "AntiIdleWhileOpen", false, DataSourceUpdateMode.OnPropertyChanged);
            txtAppFolderPath.DataBindings.Add("Text", clientSettings, "RootFolder2", false, DataSourceUpdateMode.OnPropertyChanged);
            txtScriptsFolder.DataBindings.Add("Text", clientSettings, "ScriptsFolder", false, DataSourceUpdateMode.OnPropertyChanged);
            chkInsertCommandsInline.DataBindings.Add("Checked", clientSettings, "InsertCommandsInline", false, DataSourceUpdateMode.OnPropertyChanged);
            chkSequenceDragDrop.DataBindings.Add("Checked", clientSettings, "EnableSequenceDragDrop", false, DataSourceUpdateMode.OnPropertyChanged);
            chkMinimizeToTray.DataBindings.Add("Checked", clientSettings, "MinimizeToTray", false, DataSourceUpdateMode.OnPropertyChanged);
            chkCloseToTray.DataBindings.Add("Checked", clientSettings, "CloseToTray", false, DataSourceUpdateMode.OnPropertyChanged);
            cboStartUpMode.DataBindings.Add("Text", clientSettings, "StartupMode", false, DataSourceUpdateMode.OnPropertyChanged);
            chkSlimActionBar.DataBindings.Add("Checked", clientSettings, "UseSlimActionBar", false, DataSourceUpdateMode.OnPropertyChanged);

            //get metrics
            bgwMetrics.RunWorkerAsync();
            startingRootFolder = _newAppSettings.ClientSettings.RootFolder2;
        }

        private void uiBtnOk_Click(object sender, EventArgs e)
        {
            //create references to old and new root folders
            var oldRootFolder = startingRootFolder;
            var newRootFolder = txtAppFolderPath.Text;

            if (oldRootFolder != newRootFolder)
            {
                //ask user to confirm
                var confirmNewFolderSelection = MessageBox.Show("Please confirm the changes below:" +
                    Environment.NewLine + Environment.NewLine + "Old Root Folder: " + oldRootFolder +
                    Environment.NewLine + Environment.NewLine + "New Root Folder: " + newRootFolder,
                    "Change Default Root Folder", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

                //handle if user decides to cancel
                if (confirmNewFolderSelection == DialogResult.Cancel)
                {
                    txtAppFolderPath.Text = oldRootFolder;
                    return;
                }
                //ask if we should migrate the data
                var migrateCopyData = MessageBox.Show("Would you like to attempt to move the data from" +
                    " the old folder to the new folder?  Please note, depending on how many files you have," +
                    " this could take a few minutes.", "Migrate Data?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                //check if user wants to migrate data
                if (migrateCopyData == DialogResult.Yes)
                {
                    try
                    {
                        //find and copy files
                        foreach (string dirPath in Directory.GetDirectories(oldRootFolder, "*", SearchOption.AllDirectories))
                        {
                            Directory.CreateDirectory(dirPath.Replace(oldRootFolder, newRootFolder));
                        }
                        foreach (string newPath in Directory.GetFiles(oldRootFolder, "*.*", SearchOption.AllDirectories))
                        {
                            File.Copy(newPath, newPath.Replace(oldRootFolder, newRootFolder), true);
                        }

                        MessageBox.Show("Data Migration Complete", "Data Migration Complete", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        //handle any unexpected errors
                        MessageBox.Show("An Error Occurred during Data Migration Copy: " + ex.ToString());
                    }
                }
            }

            _newAppSettings.EngineSettings.CancellationKey = (Keys)cbxCancellationKey.SelectedValue;

            if ((SinkType)cbxSinkType.SelectedItem == SinkType.File && string.IsNullOrEmpty(txtLoggingValue.Text.Trim()))
                _newAppSettings.EngineSettings.LoggingValue = Path.Combine(Folders.GetFolder(FolderType.LogFolder), "OpenBots Engine Logs.txt");

            _newAppSettings.Save();

            Close();
        }

        private void btnUpdateCheck_Click(object sender, EventArgs e)
        {
            ManifestUpdate manifest = new ManifestUpdate();
            try
            {
                manifest = ManifestUpdate.GetManifest();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting manifest: " + ex.ToString());
                return;
            }

            if (manifest.RemoteVersionNewer)
            {
                frmUpdate frmUpdate = new frmUpdate(manifest);
                if (frmUpdate.ShowDialog() == DialogResult.OK)
                {
                    frmUpdating frmUpdating = new frmUpdating(manifest.PackageURL);
                    frmUpdating.ShowDialog();
                    frmUpdating.Dispose();
                }
                frmUpdate.Dispose();
            }
            else
            {
                MessageBox.Show("The application is currently up-to-date!", "No Updates Available", MessageBoxButtons.OK);
            }
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            //prompt user to confirm they want to select a new folder
            var updateFolderRequest = MessageBox.Show("Would you like to change the default root folder that OpenBots uses" +
                " to store tasks and information? " + Environment.NewLine + Environment.NewLine + "Current Root Folder: " +
                txtAppFolderPath.Text, "Change Default Root Folder", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            //if user does not want to update folder then exit
            if (updateFolderRequest == DialogResult.No)
                return;

            //user folder browser to let user select top level folder
            using (var fbd = new FolderBrowserDialog())
            {
                //check if user selected a folder
                if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                { 
                    //update textbox which will be updated once user selects "Ok"
                    txtAppFolderPath.Text = Path.Combine(fbd.SelectedPath, "OpenBots Studio"); ;
                }
            }
        }

        private void bgwMetrics_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = Metric.ExecutionMetricsSummary();
        }

        private void bgwMetrics_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                if (e.Error is FileNotFoundException)
                {
                    lblGettingMetrics.Text = "Metrics Unavailable - Metrics are only available after running" +
                        " tasks which will generate metrics logs";
                }
                else
                {
                    lblGettingMetrics.Text = "Metrics Unavailable: " + e.Error.ToString();
                }
            }
            else
            {
                var metricsSummary = (List<ExecutionMetric>)(e.Result);

                if (metricsSummary.Count == 0)
                {
                    lblGettingMetrics.Text = "No Metrics Found";
                    lblGettingMetrics.Show();
                    tvExecutionTimes.Hide();
                    btnClearMetrics.Hide();
                }
                else
                {
                    lblGettingMetrics.Hide();
                    tvExecutionTimes.Show();
                    btnClearMetrics.Show();
                }

                foreach (var metric in metricsSummary)
                {
                    var rootNode = new TreeNode
                    {
                        Text = metric.FileName + " [" + metric.AverageExecutionTime + " avg.]"
                    };

                    foreach (var metricItem in metric.ExecutionData)
                    {
                        var subNode = new TreeNode
                        {
                            Text = string.Join(" - ", metricItem.LoggedOn.ToString("MM/dd/yy hh:mm"), metricItem.ExecutionTime)
                        };
                        rootNode.Nodes.Add(subNode);
                    }
                    tvExecutionTimes.Nodes.Add(rootNode);
                }
            }
        }

        private void btnClearMetrics_Click(object sender, EventArgs e)
        {
            Metric.ClearExecutionMetrics();
            bgwMetrics.RunWorkerAsync();
        }

        private void btnGenerateWikiDocs_Click(object sender, EventArgs e)
        {
            DocumentationGeneration docGeneration = new DocumentationGeneration();
            if (Directory.Exists(txtAppFolderPath.Text))
            {
                var docsRoot = docGeneration.GenerateMarkdownFiles(_container, txtAppFolderPath.Text);
                Process.Start(docsRoot);
            }
            else
            {
                MessageBox.Show("Root path is not valid.");
            }
        }

        /// <summary>
        /// 建立RPA WEB用JSON結構
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnGenerateRPAJson_Click(object sender, EventArgs e)
        {
            DocumentationGeneration docGeneration = new DocumentationGeneration();
            if (Directory.Exists(txtAppFolderPath.Text))
            {
                var docsRoot = docGeneration.GenerateJsonFiles(_container, txtAppFolderPath.Text);
                Process.Start(docsRoot);
            }
            else
            {
                MessageBox.Show("Root path is not valid.");
            }
        }


        private void btnLaunchAttendedMode_Click(object sender, EventArgs e)
        {
            var frmAttended = new frmAttendedMode();
            frmAttended.Show();
            Close();
        }

        private void btnSelectScriptsFolder_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    var newScriptsFolder = Path.Combine(fbd.SelectedPath);
                    txtScriptsFolder.Text = newScriptsFolder;
                }
            }
        }

        private void cbxSinkType_SelectedIndexChanged(object sender, EventArgs e)
        {
            _newAppSettings.EngineSettings.LoggingSinkType = (SinkType)cbxSinkType.SelectedItem;
            switch (_newAppSettings.EngineSettings.LoggingSinkType)
            {
                case SinkType.File:
                    LoadFileLoggingSettings();
                    break;
                case SinkType.HTTP:
                    LoadHTTPLoggingSettings();
                    break;                
            }
        }

        private void LoadFileLoggingSettings()
        {
            lblLoggingValue.Text = "File Path: ";
            txtLoggingValue.Clear();
            txtLoggingValue.Text = Path.Combine(Folders.GetFolder(FolderType.LogFolder), "OpenBots Engine Logs.txt");
            btnFileManager.Visible = true;        
        }

        private void LoadHTTPLoggingSettings()
        {
            lblLoggingValue.Text = "URI: ";
            txtLoggingValue.Clear();
            btnFileManager.Visible = false;
        }

        private void cbxMinLogLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            _newAppSettings.EngineSettings.MinLogLevel = (LogEventLevel)cbxMinLogLevel.SelectedItem;
        }

        private void btnFileManager_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtLoggingValue.Text = ofd.FileName;
            }
        }

        //prevents tab control from flickering
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x02000000;  // Turn on WS_EX_COMPOSITED
                return cp;
            }
        }
    }
}