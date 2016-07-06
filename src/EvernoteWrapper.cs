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
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using System.Xml;

namespace Evernote2Onenote
{
    /// <summary>
    /// Wrapper for the ENScript.exe tool
    /// </summary>
    public class EvernoteWrapper
    {
        /// <summary>
        /// path to the ENScript.exe
        /// </summary>
        private static string exePath;

        /// <summary>
        /// the full path to ENScript.exe
        /// </summary>
        public static string ENScriptPath
        {
            get
            {
                return exePath;
            }

            set
            {
                exePath = value;
            }
        }

        static EvernoteWrapper()
        {
            string enscriptpath = ProgramFilesx86() + "\\Evernote\\Evernote\\ENScript.exe";
            if (!File.Exists(enscriptpath))
            {
                MessageBox.Show("Could not find the ENScript.exe file from Evernote!\nPlease select this file in the next dialog.", "Evernote2Onenote");
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Applications|*.exe";
                openFileDialog1.Title = "Select the ENScript.exe file";
                openFileDialog1.CheckPathExists = true;

                // Show the Dialog.
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    enscriptpath = openFileDialog1.FileName;
                }
                if (!File.Exists(enscriptpath))
                    throw new Exception();
            }
            EvernoteWrapper.ENScriptPath = enscriptpath;
        }
        internal static List<Note> ParseNotes(string exportFile, string notebookName)
        {
            MainFrm.syncStep = SyncStep.ParseNotes;
            List<Note> noteList = new List<Note>();
            if (MainFrm.cancelled)
            {
                return noteList;
            }

            XmlTextReader xtrInput;
            XmlDocument xmlDocItem;

            xtrInput = new XmlTextReader(exportFile);
            string xmltext = "";
            try
            {
                while (xtrInput.Read())
                {
                    while ((xtrInput.NodeType == XmlNodeType.Element) && (xtrInput.Name.ToLower() == "note"))
                    {
                        if (MainFrm.cancelled)
                        {
                            break;
                        }

                        xmlDocItem = new XmlDocument();
                        xmltext = OnenoteWrapper.SanitizeXml(xtrInput.ReadOuterXml());
                        xmlDocItem.LoadXml(xmltext);
                        XmlNode node = xmlDocItem.FirstChild;

                        // node is <note> element
                        // node.FirstChild.InnerText is <title>
                        node = node.FirstChild;

                        Note note = new Note();
                        note.Title = HttpUtility.HtmlDecode(node.InnerText);

                        noteList.Add(note);
                    }
                }

                xtrInput.Close();
            }
            catch (System.Xml.XmlException ex)
            {
                // happens if the notebook was empty or does not exist.
                // Or due to a parsing error if a note isn't properly xml encoded
                // 
                // try to find the name of the note that's causing the problems
                string notename = "";
                if (xmltext.Length > 0)
                {
                    Regex rxnote = new Regex("<title>(.+)</title>", RegexOptions.IgnoreCase);
                    var notematch = rxnote.Match(xmltext);
                    if (notematch.Groups.Count == 2)
                    {
                        notename = notematch.Groups[1].ToString();
                    }
                }
                if (notename.Length > 0)
                    MessageBox.Show(string.Format("Error parsing the note \"{2}\" in notebook \"{0}\",\n{1}", notebookName, ex.ToString(), notename));
                else
                    MessageBox.Show(string.Format("Error parsing the notebook \"{0}\"\n{1}", notebookName, ex.ToString()));
            }

            return noteList;
        }


        static string ProgramFilesx86()
        {
            if (8 == IntPtr.Size
                || (!String.IsNullOrEmpty(Environment.GetEnvironmentVariable("PROCESSOR_ARCHITEW6432"))))
            {
                return Environment.GetEnvironmentVariable("ProgramFiles(x86)");
            }

            return Environment.GetEnvironmentVariable("ProgramFiles");
        }

        /// <summary>
        /// Lists all the notebooks in Evernote
        /// </summary>
        /// <returns>List of notebook names</returns>
        public static List<string> GetNotebooks()
        {
            List<string> notebooks = new List<string>();

            ProcessStartInfo processStartInfo = new ProcessStartInfo(exePath, "listNotebooks");
            processStartInfo.UseShellExecute = false;
            processStartInfo.ErrorDialog = false;
            processStartInfo.RedirectStandardError = true;
            processStartInfo.RedirectStandardInput = true;
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.StandardOutputEncoding = Encoding.UTF8;
            processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            processStartInfo.CreateNoWindow = true;
            Process process = new Process();
            process.StartInfo = processStartInfo;
            bool processStarted = process.Start();
            if (processStarted)
            {
                StreamWriter inputWriter = process.StandardInput;
                StreamReader outputReader = process.StandardOutput;
                StreamReader errorReader = process.StandardError;
                process.WaitForExit();
                while (outputReader.Peek() >= 0)
                {
                    notebooks.Add(outputReader.ReadLine());
                }
            }

            return notebooks;
        }

        /// <summary>
        /// Exports the specified notebook to the specified file
        /// </summary>
        /// <param name="notebook">the notebook to export</param>
        /// <param name="exportFile">the file to export the notebook to</param>
        /// <returns>true if successful, false in case of an error</returns>
        public static bool ExportNotebook(string notebook, string exportFile)
        {
            bool ret = false;
            if (!File.Exists(exePath))
                return ret;

            ProcessStartInfo processStartInfo = new ProcessStartInfo(exePath, "exportNotes /q notebook:\"\\\"" + notebook + "\\\"\" /f \"" + exportFile + "\"");
            processStartInfo.UseShellExecute = false;
            processStartInfo.ErrorDialog = false;
            processStartInfo.RedirectStandardError = true;
            processStartInfo.RedirectStandardInput = true;
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.StandardOutputEncoding = Encoding.UTF8;
            processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            processStartInfo.CreateNoWindow = true;
            Process process = new Process();
            process.StartInfo = processStartInfo;
            bool processStarted = process.Start();
            if (processStarted)
            {
                StreamWriter inputWriter = process.StandardInput;
                StreamReader outputReader = process.StandardOutput;
                StreamReader errorReader = process.StandardError;
                process.WaitForExit();
                ret = (process.ExitCode == 0) && File.Exists(exportFile);
            }

            return ret;
        }

        /// <summary>
        /// Imports all the notes in the export file to the 
        /// specified notebook in Evernote
        /// </summary>
        /// <param name="notesPath">the path to the export file</param>
        /// <param name="notebook">the notebook where the export file should be imported to</param>
        /// <returns>true if successful, false in case of an error</returns>
        public static bool ImportNotes(string notesPath, string notebook)
        {
            bool ret = false;

            ProcessStartInfo processStartInfo = new ProcessStartInfo(exePath, "importNotes /n \"" + notebook + "\" /s \"" + notesPath + "\"");
            processStartInfo.UseShellExecute = false;
            processStartInfo.ErrorDialog = false;
            processStartInfo.RedirectStandardError = true;
            processStartInfo.RedirectStandardInput = true;
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.StandardOutputEncoding = Encoding.UTF8;
            processStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            processStartInfo.CreateNoWindow = true;
            Process process = new Process();
            process.StartInfo = processStartInfo;
            bool processStarted = process.Start();
            if (processStarted)
            {
                StreamWriter inputWriter = process.StandardInput;
                StreamReader outputReader = process.StandardOutput;
                StreamReader errorReader = process.StandardError;
                process.WaitForExit();
                ret = process.ExitCode == 0;
            }

            return ret;
        }
    }
}
