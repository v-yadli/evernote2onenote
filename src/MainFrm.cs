// Evernote2Onenote - imports Evernote notes to Onenote
// Copyright (C) 2014 - Stefan Kueng

// This program is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.

// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.

// You should have received a copy of the GNU General Public License
// along with this program.  If not, see <http://www.gnu.org/licenses/>.

using Evernote2Onenote.Enums;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using System.Xml;
using OneNote = Microsoft.Office.Interop.OneNote;

namespace Evernote2Onenote
{
    /// <summary>
    /// main dialog
    /// </summary>
    public partial class MainFrm : Form
    {
        internal const string c_AllNotebooks = "__ALL__";

        private delegate void StringDelegate(string foo);
        private string enscriptpath;
        private string m_EvernoteNotebookPath;
        private SynchronizationContext synchronizationContext;
        internal static bool cancelled = false;
        internal static SyncStep syncStep = SyncStep.Start;
        private string m_enexfile = "";

        private string cmdNoteBook = "";

        public MainFrm(string cmdNotebook, string cmdDate)
        {
            InitializeComponent();
            this.synchronizationContext = SynchronizationContext.Current;
            string version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();
            versionLabel.Text = string.Format("Version: {0}", version);

            if (cmdNotebook.Length > 0)
                cmdNoteBook = cmdNotebook;
            if (cmdDate.Length > 0)
            {
                try
                {
                    OnenoteWrapper.cmdDate = DateTime.Parse(cmdDate);
                }
                catch (Exception)
                {
                    MessageBox.Show(string.Format("The Datestring\n{0}\nis not valid!", cmdDate));
                }
            }
            try
            {
                importDatePicker.Value = OnenoteWrapper.cmdDate;
            }
            catch (Exception)
            {
                importDatePicker.Value = importDatePicker.MinDate;
            }

            var notebooklist = EvernoteWrapper.GetNotebooks();

            //Yadli: add option "__ALL__"
            notebooklist.Add(c_AllNotebooks);
            //Yadli: set event callback
            OnenoteWrapper.ImportingNote += SetInfo;

            foreach (string s in notebooklist)
                this.notebookCombo.Items.Add(s);
            if (notebooklist.Count == 0)
            {
                MessageBox.Show("No Notebooks found in Evernote!\nMake sure you have at least one locally synched notebook.", "Evernote2Onenote");
                startsync.Enabled = false;
            }
            else
                this.notebookCombo.SelectedIndex = 0;

            if (cmdNotebook.Length > 0)
            {
                Startsync_Click(null, null);
            }
        }

        private void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        private void SetInfo(string line1, string line2, int pos, int max)
        {
            //return;

            int fullpos = 0;

            switch (MainFrm.syncStep)
            {
                // full progress is from 0 - 100'000
                case SyncStep.ExtractNotes:      // 0- 10%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.1) : 0;
                    break;
                case SyncStep.ParseNotes:        // 10-20%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.1) + 10000 : 10000;
                    break;
                case SyncStep.CalculateWhatToDo: // 30-35%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.05) + 30000 : 30000;
                    break;
                case SyncStep.ImportNotes:       // 35-100%
                    fullpos = max != 0 ? (int)(pos * 100000.0 / max * 0.65) + 35000 : 35000;
                    break;
            }

            synchronizationContext.Send(new SendOrPostCallback(delegate (object state)
            {
                if (line1 != null)
                    this.infoText1.Text = line1;
                if (line2 != null)
                    this.infoText2.Text = line2;
                this.progressIndicator.Minimum = 0;
                this.progressIndicator.Maximum = 100000;
                this.progressIndicator.Value = fullpos;
            }), null);

            if (max == 0)
                syncStep++;
        }

        private void btnENEXImport_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Evernote exports|*.enex";
            openFileDialog1.Title = "Select the ENEX file";
            openFileDialog1.CheckPathExists = true;

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                m_enexfile = openFileDialog1.FileName;
                Startsync_Click(sender, e);
            }
        }

        private async void Startsync_Click(object sender, EventArgs e)
        {
            var ENNotebookName = this.notebookCombo.SelectedItem as string;

            if (!Import_init(ENNotebookName)) return;

            if (ENNotebookName == c_AllNotebooks)
            {
                List<string> notebooks = new List<string>();
                foreach (var item in (this.notebookCombo.Items as IList))
                {
                    string nb = item as string;
                    if (nb != c_AllNotebooks) notebooks.Add(nb);
                }

                foreach (var nb in notebooks)
                {
                    ENNotebookName = nb;
                    await Import_impl(ENNotebookName);
                }
            }
            else
            {
                Import_impl(ENNotebookName);
            }
        }

        /// <returns> True on success </returns>
        private bool Import_init(string ENNotebookName)
        {
            if (m_enexfile != null && m_enexfile.Length > 0)
            {
                ENNotebookName = Path.GetFileNameWithoutExtension(m_enexfile);
            }
            if (cmdNoteBook.Length > 0)
                ENNotebookName = cmdNoteBook;
            if (ENNotebookName.Length == 0)
            {
                MessageBox.Show("Please enter a notebook in EverNote to import the notes from", "Evernote2Onenote");
                return false;
            }

            return true;
        }

        private async Task Import_impl(string section)
        {
            if (importDatePicker.Value > OnenoteWrapper.cmdDate)
                OnenoteWrapper.cmdDate = importDatePicker.Value;

            if (startsync.Text == "Start Import")
            {
                startsync.Text = "Cancel";
                //ImportNotesToOnenote();
                MethodInvoker syncDelegate = new MethodInvoker(() => ImportNotesToOnenote(section));
                var async_result = syncDelegate.BeginInvoke(null, null);
                await Task.Factory.FromAsync(async_result, (_) => { });
            }
            else
            {
                cancelled = true;
                return;
            }
        }

        private void ImportNotesToOnenote(string section)
        {
            syncStep = SyncStep.Start;
            if (m_enexfile != null && m_enexfile.Length > 0)
            {
                List<Note> notesEvernote = new List<Note>();
                if (m_enexfile != string.Empty)
                {
                    SetInfo("Parsing notes from Evernote", "", 0, 0);
                    notesEvernote = EvernoteWrapper.ParseNotes(m_enexfile, section);
                }
                if (m_enexfile != string.Empty)
                {
                    SetInfo("importing notes to Onenote", "", 0, 0);
                    OnenoteWrapper.ImportNotesToOnenote(section, notesEvernote, m_enexfile);
                }
            }
            else
            {
                SetInfo("Extracting notes from Evernote", "", 0, 0);
                string exportFile = ExtractNotes(section);
                if (exportFile != null)
                {
                    List<Note> notesEvernote = new List<Note>();
                    if (exportFile != string.Empty)
                    {
                        SetInfo("Parsing notes from Evernote", "", 0, 0);
                        notesEvernote = EvernoteWrapper.ParseNotes(exportFile, section);
                    }
                    if (exportFile != string.Empty)
                    {
                        SetInfo("importing notes to Onenote", "", 0, 0);
                        OnenoteWrapper.ImportNotesToOnenote(section, notesEvernote, exportFile);
                    }
                }
                else
                {
                    MessageBox.Show(string.Format("The notebook \"{0}\" either does not exist or isn't accessible!", section));
                }
            }

            m_enexfile = "";
            if (cancelled)
            {
                SetInfo(null, "Operation cancelled", 0, 0);
            }
            else
                SetInfo("", "", 0, 0);

            synchronizationContext.Send(new SendOrPostCallback(delegate (object state)
            {
                startsync.Text = "Start Import";
                this.infoText1.Text = "Finished";
                this.progressIndicator.Minimum = 0;
                this.progressIndicator.Maximum = 100000;
                this.progressIndicator.Value = 0;
            }), null);
            if (cmdNoteBook.Length > 0)
            {
                synchronizationContext.Send(new SendOrPostCallback(delegate (object state)
                {
                    this.Close();
                }), null);
            }
        }

        private void homeLink_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://stefanstools.sourceforge.net/Evernote2Onenote.html");
        }

        private string ExtractNotes(string notebook)
        {
            if (cancelled)
            {
                return null;
            }
            syncStep = SyncStep.ExtractNotes;

            string exportFile = Path.GetTempFileName();
#if DEBUG
            exportFile = @"D:\temp\evimsync\" + notebook + ".xml";
#endif
            if (EvernoteWrapper.ExportNotebook(notebook, exportFile))
            {
                return exportFile;
            }

            // in case the selected notebook is empty, we don't get
            // an exportFile. But just to make sure the notebook
            // exists anyway, we check that here before giving up
            if (EvernoteWrapper.GetNotebooks().Contains(notebook))
                return string.Empty;

            return null;
        }

        private void modifiedDateCheckbox_CheckedChanged(object sender, EventArgs e)
        {
            OnenoteWrapper.modifiedDate = modifiedDateCheckbox.Checked;
        }
    }
}
