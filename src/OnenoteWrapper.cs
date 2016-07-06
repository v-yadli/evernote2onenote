using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using Evernote2Onenote.Enums;
using System.Web;
using System.Globalization;
using OneNote = Microsoft.Office.Interop.OneNote;
using System.Threading;

namespace Evernote2Onenote
{
    static class OnenoteWrapper
    {
        internal const string c_NotebookName = "Evernote";
        private static Microsoft.Office.Interop.OneNote.Application s_onenote_app = null;
        internal static string m_EvernoteNotebookPath;

        private static readonly Regex rxStyle = new Regex("(?<text>\\<div.)style=\\\"[^\\\"]*\\\"", RegexOptions.IgnoreCase);
        private static readonly Regex rxCDATA = new Regex(@"<!\[CDATA\[<\?xml version=[""']1.0[""'][^?]*\?>", RegexOptions.IgnoreCase);
        private static readonly Regex rxCDATAInner = new Regex(@"\<\!\[CDATA\[(?<text>.*)\]\]\>", RegexOptions.IgnoreCase|RegexOptions.Singleline);
        private static readonly Regex rxBodyStart = new Regex(@"<en-note[^>/]*>", RegexOptions.IgnoreCase);
        private static readonly Regex rxBodyEnd = new Regex(@"</en-note\s*>\s*]]>", RegexOptions.IgnoreCase);
        private static readonly Regex rxBodyEmpty = new Regex(@"<en-note[^>/]*/>\s*]]>", RegexOptions.IgnoreCase);
        private static readonly Regex rxDate = new Regex(@"^date:(.*)$", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        private static readonly Regex rxNote = new Regex("<title>(.+)</title>", RegexOptions.IgnoreCase);
        private static readonly Regex rxComment = new Regex("<!--(.+)-->", RegexOptions.IgnoreCase);
        private static readonly Regex rxDtd = new Regex(@"<!DOCTYPE en-note SYSTEM \""http:\/\/xml\.evernote\.com\/pub\/enml\d*\.dtd\"">", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        internal static DateTime cmdDate = new DateTime(0);
        internal static string newnbID = "";
        internal static event Action<string, string, int, int> ImportingNote = delegate { };

        internal static string m_PageID;
        internal static string m_xmlNewOutlineContent =
            "<one:Meta name=\"{2}\" content=\"{1}\"/>" +
            "<one:OEChildren><one:HTMLBlock><one:Data><![CDATA[{0}]]></one:Data></one:HTMLBlock>{3}</one:OEChildren>";

        internal static string m_xmlSourceUrl = "<one:OE alignment=\"left\" quickStyleIndex=\"2\"><one:T><![CDATA[From &lt;<a href=\"{0}\">{0}</a>&gt; ]]></one:T></one:OE>";
        internal static string m_xmlNewOutline =
            "<?xml version=\"1.0\"?>" +
            "<one:Page xmlns:one=\"{2}\" ID=\"{1}\" dateTime=\"{5}\">" +
            "<one:Title selected=\"partial\" lang=\"en-US\">" +
                        "<one:OE creationTime=\"{5}\" lastModifiedTime=\"{5}\">" +
                            "<one:T><![CDATA[{3}]]></one:T> " +
                        "</one:OE>" +
                        "</one:Title>{4}" +
            "<one:Outline>{0}</one:Outline></one:Page>";
        internal static string m_xmlns = "http://schemas.microsoft.com/office/onenote/2013/onenote";
        internal static string ENNotebookName = "";
        internal static bool m_bUseUnfiledSection = false;
        internal static bool modifiedDate = true;

        static OnenoteWrapper()
        {
            CreateOnenoteCOMObject();

            CreateNotebook();
        }

        private static void CreateOnenoteCOMObject()
        {
            if (s_onenote_app != null)
            {
                s_onenote_app = null;
            }

            try
            {
                s_onenote_app = new OneNote.Application();
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Could not connect to Onenote!\nReasons for this might be:\n* The desktop version of onenote is not installed\n* Onenote is not installed properly\n* Onenote is already running but with a different user account\n\n{0}", ex.ToString()));
                throw new Exception();
            }
            if (s_onenote_app == null)
            {
                MessageBox.Show(string.Format("Could not connect to Onenote!\nReasons for this might be:\n* The desktop version of onenote is not installed\n* Onenote is not installed properly\n* Onenote is already running but with a different user account\n\n{0}"));
                throw new Exception();
            }
        }

        public static string FilenameMatchEvaluator(Match m)
        {
            string filename = m.Groups[1].ToString();
            filename = filename.Replace("&nbsp;", " ");
            // remove illegal path chars
            string invalid = new string(Path.GetInvalidFileNameChars());
            foreach (char c in invalid)
            {
                filename = filename.Replace(c.ToString(), "");
            }
            filename = System.Security.SecurityElement.Escape(filename);
            return "<file-name>" + filename + "</file-name>";
        }

        internal static string SanitizeXml(string text)
        {
            //text = HttpUtility.HtmlDecode(text);
            Regex rxtitle = new Regex("<note><title>(.+)</title>", RegexOptions.IgnoreCase);
            var match = rxtitle.Match(text);
            if (match.Groups.Count == 2)
            {
                string title = match.Groups[1].ToString();
                title = title.Replace("&", "&amp;");
                title = title.Replace("\"", "&quot;");
                title = title.Replace("'", "&apos;");
                title = title.Replace("’", "&apos;");
                title = title.Replace("<", "&lt;");
                title = title.Replace(">", "&gt;");
                title = title.Replace("@", "&#64;");
                text = rxtitle.Replace(text, "<note><title>" + title + "</title>");
            }

            Regex rxauthor = new Regex("<author>(.+)</author>", RegexOptions.IgnoreCase);
            var authormatch = rxauthor.Match(text);
            if (match.Groups.Count == 2)
            {
                string author = authormatch.Groups[1].ToString();
                author = author.Replace("&", "&amp;");
                author = author.Replace("\"", "&quot;");
                author = author.Replace("'", "&apos;");
                author = author.Replace("’", "&apos;");
                author = author.Replace("<", "&lt;");
                author = author.Replace(">", "&gt;");
                author = author.Replace("@", "&#64;");
                text = rxauthor.Replace(text, "<author>" + author + "</author>");
            }

            Regex rxfilename = new Regex("<file-name>(.+)</file-name>", RegexOptions.IgnoreCase);
            var filenamematch = rxfilename.Match(text);
            if (match.Groups.Count == 2)
            {
                MatchEvaluator myEvaluator = new MatchEvaluator(FilenameMatchEvaluator);
                text = rxfilename.Replace(text, myEvaluator);
            }

            return text;
        }
        internal static void ImportNotesToOnenote(string section, List<Note> notesEvernote, string exportFile)
        {
            MainFrm.syncStep = SyncStep.CalculateWhatToDo;
            int uploadcount = notesEvernote.Count;

            string temppath = Path.GetTempPath() + "\\ev2on";
            Directory.CreateDirectory(temppath);

            MainFrm.syncStep = SyncStep.ImportNotes;
            int counter = 0;


            XmlTextReader xtrInput;
            XmlDocument xmlDocItem;
            string xmltext = "";
            try
            {
                xtrInput = new XmlTextReader(exportFile);
                while (xtrInput.Read())
                {
                    while ((xtrInput.NodeType == XmlNodeType.Element) && (xtrInput.Name.ToLower() == "note"))
                    {
                        if (MainFrm.cancelled)
                        {
                            break;
                        }

                        xmlDocItem = new XmlDocument();
                        xmltext = SanitizeXml(xtrInput.ReadOuterXml());
                        xmlDocItem.LoadXml(xmltext);
                        XmlNode node = xmlDocItem.FirstChild;

                        // node is <note> element
                        // node.FirstChild.InnerText is <title>
                        node = node.FirstChild;

                        Note note = new Note();
                        note.Title = HttpUtility.HtmlDecode(node.InnerText);
                        node = node.NextSibling;
                        note.Content = HttpUtility.HtmlDecode(node.InnerXml);

                        XmlNodeList atts = xmlDocItem.GetElementsByTagName("resource");
                        foreach (XmlNode xmln in atts)
                        {
                            Attachment attachment = new Attachment();
                            attachment.Base64Data = xmln.FirstChild.InnerText;
                            byte[] data = Convert.FromBase64String(xmln.FirstChild.InnerText);
                            byte[] hash = new System.Security.Cryptography.MD5CryptoServiceProvider().ComputeHash(data);
                            string hashHex = BitConverter.ToString(hash).Replace("-", string.Empty).ToLower();

                            attachment.Hash = hashHex;

                            XmlNodeList fns = xmlDocItem.GetElementsByTagName("file-name");
                            if (fns.Count > note.Attachments.Count)
                            {
                                attachment.FileName = HttpUtility.HtmlDecode(fns.Item(note.Attachments.Count).InnerText);
                                string invalid = new string(Path.GetInvalidFileNameChars());
                                foreach (char c in invalid)
                                {
                                    attachment.FileName = attachment.FileName.Replace(c.ToString(), "");
                                }
                                attachment.FileName = System.Security.SecurityElement.Escape(attachment.FileName);
                            }

                            XmlNodeList mimes = xmlDocItem.GetElementsByTagName("mime");
                            if (mimes.Count > note.Attachments.Count)
                            {
                                attachment.ContentType = HttpUtility.HtmlDecode(mimes.Item(note.Attachments.Count).InnerText);
                            }

                            note.Attachments.Add(attachment);
                        }

                        XmlNodeList tagslist = xmlDocItem.GetElementsByTagName("tag");
                        foreach (XmlNode n in tagslist)
                        {
                            note.Tags.Add(HttpUtility.HtmlDecode(n.InnerText));
                        }

                        XmlNodeList datelist = xmlDocItem.GetElementsByTagName("created");
                        foreach (XmlNode n in datelist)
                        {
                            DateTime dateCreated;

                            if (DateTime.TryParseExact(n.InnerText, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture, DateTimeStyles.AdjustToUniversal, out dateCreated))
                            {
                                note.Date = dateCreated;
                            }
                        }
                        if (modifiedDate)
                        {
                            XmlNodeList datelist2 = xmlDocItem.GetElementsByTagName("updated");
                            foreach (XmlNode n in datelist2)
                            {
                                DateTime dateUpdated;

                                if (DateTime.TryParseExact(n.InnerText, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture, DateTimeStyles.AdjustToUniversal, out dateUpdated))
                                {
                                    note.Date = dateUpdated;
                                }
                            }
                        }

                        XmlNodeList sourceurl = xmlDocItem.GetElementsByTagName("source-url");
                        note.SourceUrl = "";
                        foreach (XmlNode n in sourceurl)
                        {
                            try
                            {
                                note.SourceUrl = n.InnerText;
                                if (n.InnerText.StartsWith("file://"))
                                    note.SourceUrl = "";
                            }
                            catch (System.FormatException)
                            {
                            }
                        }

                        if (cmdDate > note.Date)
                            continue;

                        ImportingNote(null, string.Format("importing note ({0} of {1}) : \"{2}\"", counter + 1, uploadcount, note.Title), counter++, uploadcount);

                        string htmlBody = note.Content;

                        if (note.Tags.Count != 0)
                        {
                            htmlBody = htmlBody.Replace("</en-note>", "<p>Tags:" + string.Join(",", note.Tags) + "</p></en-note>");
                        }

                        List<string> tempfiles = new List<string>();
                        string xmlAttachments = "";
                        foreach (Attachment attachment in note.Attachments)
                        {
                            // save the attached file
                            string tempfilepath = temppath + "\\";
                            byte[] data = Convert.FromBase64String(attachment.Base64Data);
                            tempfilepath += attachment.Hash;
                            Stream fs = new FileStream(tempfilepath, FileMode.Create);
                            fs.Write(data, 0, data.Length);
                            fs.Close();
                            tempfiles.Add(tempfilepath);

                            Regex rx = new Regex(@"<en-media\b[^>]*?hash=""" + attachment.Hash + @"""[^>]*/>", RegexOptions.IgnoreCase);
                            if ((attachment.ContentType != null) && (attachment.ContentType.Contains("image") && rx.Match(htmlBody).Success))
                            {
                                // replace the <en-media /> tag with an <img /> tag
                                htmlBody = rx.Replace(htmlBody, @"<img src=""file:///" + tempfilepath + @"""/>");
                            }
                            else
                            {
                                rx = new Regex(@"<en-media\b[^>]*?hash=""" + attachment.Hash + @"""[^>]*></en-media>", RegexOptions.IgnoreCase);
                                if ((attachment.ContentType != null) && (attachment.ContentType.Contains("image") && rx.Match(htmlBody).Success))
                                {
                                    // replace the <en-media /> tag with an <img /> tag
                                    htmlBody = rx.Replace(htmlBody, @"<img src=""file:///" + tempfilepath + @"""/>");
                                }
                                else
                                {
                                    if ((attachment.FileName != null) && (attachment.FileName.Length > 0))
                                        xmlAttachments += string.Format("<one:InsertedFile pathSource=\"{0}\" preferredName=\"{1}\" />", tempfilepath, attachment.FileName);
                                    else
                                        xmlAttachments += string.Format("<one:InsertedFile pathSource=\"{0}\" preferredName=\"{1}\" />", tempfilepath, attachment.Hash);
                                }
                            }
                        }
                        note.Attachments.Clear();

                        htmlBody = rxStyle.Replace(htmlBody, "${text}");
                        htmlBody = rxComment.Replace(htmlBody, string.Empty);
                        htmlBody = rxCDATA.Replace(htmlBody, string.Empty);
                        htmlBody = rxDtd.Replace(htmlBody, string.Empty);
                        htmlBody = rxBodyStart.Replace(htmlBody, "<body>");
                        htmlBody = rxBodyEnd.Replace(htmlBody, "</body>");
                        htmlBody = rxBodyEmpty.Replace(htmlBody, "<body></body>");
                        htmlBody = htmlBody.Trim();
                        htmlBody = @"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN""><head></head>" + htmlBody;

                        string emailBody = htmlBody;
                        emailBody = rxDate.Replace(emailBody, "Date: " + note.Date.ToString("ddd, dd MMM yyyy HH:mm:ss K"));
                        emailBody = emailBody.Replace("&apos;", "'");
                        emailBody = emailBody.Replace("’", "'");
                        emailBody = rxCDATAInner.Replace(emailBody, "&lt;![CDATA[${text}]]&gt;");
                        emailBody = emailBody.Replace("‘", "'");

                        InsertIntoOnenote_impl(section, note, xmlAttachments, emailBody);

                        foreach (string p in tempfiles)
                        {
                            File.Delete(p);
                        }
                    }
                }

                xtrInput.Close();
            }
            catch (System.Xml.XmlException ex)
            {
                // happens if the notebook was empty or does not exist.
                // Or due to a parsing error if a note isn't properly xml encoded
                // try to find the name of the note that's causing the problems
                string notename = "";
                if (xmltext.Length > 0)
                {
                    var notematch = rxNote.Match(xmltext);
                    if (notematch.Groups.Count == 2)
                    {
                        notename = notematch.Groups[1].ToString();
                    }
                }
                if (notename.Length > 0)
                    MessageBox.Show(string.Format("Error parsing the note \"{2}\" in notebook \"{0}\",\n{1}", ENNotebookName, ex.ToString(), notename));
                else
                    MessageBox.Show(string.Format("Error parsing the notebook \"{0}\"\n{1}", ENNotebookName, ex.ToString()));
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Exception importing notes:\n{0}", ex.ToString()));
            }
        }

        private static void InsertIntoOnenote_impl(string section, Note note, string xmlAttachments, string emailBody)
        {
            int retry = 0;
        begin:
            try
            {
                // Get the hierarchy for all the notebooks
                //if ((note.Tags.Count > 0) && (!m_bUseUnfiledSection))
                //{
                //    foreach (string tag in note.Tags)
                //    {
                //        string sectionId = GetSection(tag);
                //        onApp.CreateNewPage(sectionId, out m_PageID, Microsoft.Office.Interop.OneNote.NewPageStyle.npsBlankPageWithTitle);
                //        string textToSave;
                //        onApp.GetPageContent(m_PageID, out textToSave, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
                //        //OneNote uses HTML for the xml string to pass to the UpdatePageContent, so use the
                //        //Outlook HTMLBody property.  It coerces rtf and plain text to HTML.
                //        int outlineID = new System.Random().Next();
                //        //string outlineContent = string.Format(m_xmlNewOutlineContent, emailBody, outlineID, m_outlineIDMetaName);
                //        string xmlSource = string.Format(m_xmlSourceUrl, note.SourceUrl);
                //        string outlineContent = string.Format(m_xmlNewOutlineContent, emailBody, outlineID, System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"), note.SourceUrl.Length > 0 ? xmlSource : "");
                //        string xml = string.Format(m_xmlNewOutline, outlineContent, m_PageID, m_xmlns, System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"), xmlAttachments, note.Date.ToString("yyyy'-'MM'-'ddTHH':'mm':'ss'Z'"));
                //        onApp.UpdatePageContent(xml, DateTime.MinValue, OneNote.XMLSchema.xs2013, true);
                //    }
                //}
                //else
                {
                    string sectionId = m_bUseUnfiledSection ? newnbID : GetSection(section);
                    s_onenote_app.CreateNewPage(sectionId, out m_PageID, Microsoft.Office.Interop.OneNote.NewPageStyle.npsBlankPageWithTitle);
                    string textToSave;
                    s_onenote_app.GetPageContent(m_PageID, out textToSave, Microsoft.Office.Interop.OneNote.PageInfo.piBasic);
                    //OneNote uses HTML for the xml string to pass to the UpdatePageContent, so use the
                    //Outlook HTMLBody property.  It coerces rtf and plain text to HTML.
                    int outlineID = new System.Random().Next();
                    //string outlineContent = string.Format(m_xmlNewOutlineContent, emailBody, outlineID, m_outlineIDMetaName);
                    string xmlSource = string.Format(m_xmlSourceUrl, note.SourceUrl);
                    string outlineContent = string.Format(m_xmlNewOutlineContent, emailBody, outlineID, System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"), note.SourceUrl.Length > 0 ? xmlSource : "");
                    string xml = string.Format(m_xmlNewOutline, outlineContent, m_PageID, m_xmlns, System.Security.SecurityElement.Escape(note.Title).Replace("&apos;", "'"), xmlAttachments, note.Date.ToString("yyyy'-'MM'-'ddTHH':'mm':'ss'Z'"));
                    s_onenote_app.UpdatePageContent(xml, DateTime.MinValue, OneNote.XMLSchema.xs2013, true);

                    return;
                }
            }
            catch (Exception ex)
            {
                if (retry < 5)
                {
                    ++retry;
                    CreateOnenoteCOMObject();
                    Thread.Sleep(5000);
                }
                else
                {
                    MessageBox.Show(string.Format("Note:{0}\n{1}", note.Title, ex.ToString()));
                    return;
                }
            }

            goto begin;
        }

        private static void AppendHierarchy(XmlNode xml, StringBuilder str, int level)
        {
            // The set of elements that are themselves meaningful to export:
            if (xml.Name == "one:Notebook" || xml.Name == "one:SectionGroup" || xml.Name == "one:Section" || xml.Name == "one:Page")
            {
                string ID;
                if (xml.LocalName == "Section" && xml.Attributes["path"].Value == m_EvernoteNotebookPath)
                    ID = "UnfiledNotes";
                else
                    ID = xml.Attributes["ID"].Value;
                string name = HttpUtility.HtmlEncode(xml.Attributes["name"].Value);
                if (str.Length > 0)
                    str.Append("\n");
                str.Append(string.Format("{0} {1} {2} {3}",
                    new string[] { level.ToString(), xml.LocalName, ID, name }));
            }
            // The set of elements that contain children that are meaningful to export:
            if (xml.Name == "one:Notebooks" || xml.Name == "one:Notebook" || xml.Name == "one:SectionGroup" || xml.Name == "one:Section")
            {
                foreach (XmlNode child in xml.ChildNodes)
                {
                    int nextLevel;
                    if (xml.Name == "one:Notebooks")
                        nextLevel = level;
                    else
                        nextLevel = level + 1;
                    AppendHierarchy(child, str, nextLevel);
                }
            }
        }

        private static string GetSection(string sectionName)
        {
            string newnbID = "";
            try
            {
                // remove and/or replace characters that are not allowed in Onenote section names
                sectionName = sectionName.Replace("?", "");
                sectionName = sectionName.Replace("*", "");
                sectionName = sectionName.Replace("/", "");
                sectionName = sectionName.Replace("\\", "");
                sectionName = sectionName.Replace(":", "");
                sectionName = sectionName.Replace("<", "");
                sectionName = sectionName.Replace(">", "");
                sectionName = sectionName.Replace("|", "");
                sectionName = sectionName.Replace("&", "");
                sectionName = sectionName.Replace("#", "");
                sectionName = sectionName.Replace("\"", "'");
                sectionName = sectionName.Replace("%", "");

                string xmlHierarchy;
                s_onenote_app.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out xmlHierarchy);

                s_onenote_app.OpenHierarchy(m_EvernoteNotebookPath + "\\" + sectionName + ".one", "", out newnbID, OneNote.CreateFileType.cftSection);
                string xmlSections;
                s_onenote_app.GetHierarchy(newnbID, OneNote.HierarchyScope.hsSections, out xmlSections);

                // Load and process the hierarchy
                XmlDocument docHierarchy = new XmlDocument();
                docHierarchy.LoadXml(xmlHierarchy);
                StringBuilder Hierarchy = new StringBuilder(sectionName);
                AppendHierarchy(docHierarchy.DocumentElement, Hierarchy, 0);
            }
            catch (Exception /*ex*/)
            {
                //MessageBox.Show(string.Format("Exception creating section \"{0}\":\n{1}", sectionName, ex.ToString()));
            }
            return newnbID;
        }


        internal static void CreateNotebook()
        {
            // create a new notebook 
            try
            {
                string xmlHierarchy;
                s_onenote_app.GetHierarchy("", OneNote.HierarchyScope.hsNotebooks, out xmlHierarchy);

                // Get the hierarchy for the default notebook folder
                s_onenote_app.GetSpecialLocation(OneNote.SpecialLocation.slDefaultNotebookFolder, out m_EvernoteNotebookPath);
                m_EvernoteNotebookPath += "\\" + c_NotebookName;
                string newnbID;
                s_onenote_app.OpenHierarchy(m_EvernoteNotebookPath, "", out newnbID, OneNote.CreateFileType.cftNotebook);
                string xmlUnfiledNotes;
                s_onenote_app.GetHierarchy(newnbID, OneNote.HierarchyScope.hsPages, out xmlUnfiledNotes);

                // Load and process the hierarchy
                XmlDocument docHierarchy = new XmlDocument();
                docHierarchy.LoadXml(xmlHierarchy);
                StringBuilder Hierarchy = new StringBuilder();
                AppendHierarchy(docHierarchy.DocumentElement, Hierarchy, 0);
            }
            catch (Exception)
            {
                try
                {
                    string xmlHierarchy;
                    s_onenote_app.GetHierarchy("", OneNote.HierarchyScope.hsPages, out xmlHierarchy);

                    // Get the hierarchy for the default notebook folder
                    s_onenote_app.GetSpecialLocation(OneNote.SpecialLocation.slUnfiledNotesSection, out m_EvernoteNotebookPath);
                    s_onenote_app.OpenHierarchy(m_EvernoteNotebookPath, "", out newnbID, OneNote.CreateFileType.cftNone);
                    string xmlUnfiledNotes;
                    s_onenote_app.GetHierarchy(newnbID, OneNote.HierarchyScope.hsPages, out xmlUnfiledNotes);
                    m_bUseUnfiledSection = true;
                }
                catch (Exception ex2)
                {
                    MessageBox.Show(string.Format("Could not create the target notebook in Onenote!\n{0}", ex2.ToString()));
                    return;
                }
            }
        }
    }
}
