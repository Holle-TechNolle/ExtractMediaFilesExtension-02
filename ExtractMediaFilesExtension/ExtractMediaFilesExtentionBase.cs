using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExtractMediaFilesExtension
{
    public class ExtractMediaFilesExtensionBase
    {
        protected List<string> extensions;
        protected string DefaultFolderName;

        public ExtractMediaFilesExtensionBase()
        {
            extensions = new List<string> { "avi", "mp4", "mkv", "flv", "m4v", "m4p", "mpg", "mpeg", "mpg", "mp3" };
            DefaultFolderName = "A";
        }

        public void ExecuteExtensionAction(string[] selectedItemPaths)
        {
            string newFolderName;
            List<string> filesToExtract = new List<string>();

            string startFolder = Path.GetFullPath(Directory.GetParent(selectedItemPaths.ElementAtOrDefault(0)).ToString()); // We know there will allways be at least one element
            //MessageBox.Show(startFolder, "startFolder");

            GetExtractionList(selectedItemPaths.ToList(), filesToExtract);

            newFolderName = GetNewFolderName(filesToExtract);

            MoveTheFiles(filesToExtract, startFolder, newFolderName);
        }

        private string IsDirectoryOrFile(string path)
        {
            if (Directory.Exists(path))
                return "Directory";
            else if (File.Exists(path))
                return "File";
            else
                return "";
        }

        public void GetExtractionList(List<string> selectedItemPaths, List<string> subFiles)
        {
            try
            {
                //MessageBox.Show(selectedItemPaths.Count().ToString(), "selectedItemPaths.Count");
                // Get all media subfiles
                foreach (var sip in selectedItemPaths)
                {
                    switch (IsDirectoryOrFile(sip))
                    {
                        // Get a selected File - if it's one we like
                        case "File":
                            if ((extensions.Contains(Path.GetExtension(sip))))
                            {
                                subFiles.Add(sip);
                            }
                            break;
                        case "Directory":
                            foreach (string e in extensions)
                            {
                                // Add files from the sub Directory
                                string[] sf = Directory.GetFiles(sip, $"*.{e}");
                                if (sf.Length > 0)
                                {
                                    subFiles.AddRange(sf);
                                    //foreach (string s in sf) { MessageBox.Show(s, "Selected file"); }
                                    //MessageBox.Show(subFiles.Count().ToString(), "subFiles.Count");
                                }
                            }
                            // Get Child Directories to the sub Directory
                            List<string> subDirs = Directory.GetDirectories(sip).ToList();
                            if (subDirs.Count > 0)
                                GetExtractionList(subDirs, subFiles);
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                MessageBoxes.OutputBox("Exception", e.Message);
            }
        }

        public void MoveTheFiles(List<string> subFiles, string startFolder, string newFolderName)
        {
            // Create the new subfolder (if it does not exist)
            string newFolderPath = Path.Combine(startFolder, newFolderName);
            if (!Directory.Exists(newFolderPath)) Directory.CreateDirectory(newFolderPath);

            // Move the media files
            List<string> movedFiles = new List<string>();
            List<string> killingFolders = new List<string>();

            foreach (var f in subFiles)
            {
                string newFile = Path.Combine(newFolderPath, Path.GetFileName(f));
                if (!File.Exists(newFile))
                {
                    try
                    {
                        //File.Move(f, newFile);
                        movedFiles.Add(Path.GetFileName(f));
                        var kf = Path.GetDirectoryName(f);
                        if (!killingFolders.Contains(kf)) killingFolders.Add(kf);
                    }
                    catch (Exception e)
                    {
                        MessageBoxes.OutputBox($"Moving violation: {Path.GetFileName(f)}", $"Error: {e.Message}");
                    }
                }
            }

            killingFolders = killingFolders.OrderBy(o => o.Length).ToList();
            foreach (var kf in killingFolders)
            {
                // We have a mix of folders and subfolders here but we don't care
                
                //MessageBox.Show($"{kf}","Killing folder");
                if (Directory.Exists(kf) && kf != newFolderPath)
                {
                    try
                    {
                       // Directory.Delete(kf, true);
                    }
                    catch (Exception)
                    {
                        // Do nothing
                    }
                }
            }

            // Go tell it on the mountain
            movedFiles.Sort();
            if (movedFiles.Count == 0) movedFiles.Add("No files were moved");
            // Linebreaks
            string @out = movedFiles.Aggregate(new StringBuilder(),
                (sb, a) => sb.AppendLine(String.Join(",", a)),
                sb => sb.ToString());

            MessageBoxes.OutputBox("Moved files", @out);
        }

        public string GetNewFolderName(List<string> subFiles)
        {
            // Figure out the new subfolder (candidate) names

            // Patterns for series episodes from the net
            List<string> regExPatterns = new List<string>()
            {
                @"(([sS][0-9]{2}))",
                @"((\d{4}[-_\.]\d{2}[-_\.]\d{2}))"
            };

            List<string> newFolderNames = new List<string>();
            foreach (var sf in subFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(sf);
                int indexOf = 0;
                Match match = Regex.Match("abc", "somethingwewillnotfind");

                // See if there's match in the filename
                foreach (string regExPattern in regExPatterns)
                {
                    match = Regex.Match(fileName, regExPattern);
                    indexOf = match.Index;
                    if (indexOf > 0) break;
                }

                //MessageBox.Show($"match: {match.ToString()}");
                // Trim the filename according to the match we found
                if (indexOf > 0)
                {
                    var newFName = fileName.Substring(0, indexOf).Trim();
                    Regex.Replace(newFName, @"[-_\.]*$", "");
                    if (newFName.Length > 0 && !newFolderNames.Contains(newFName))
                    {
                        newFolderNames.Add(newFName);
                        //MessageBox.Show(newFName, "newFName");
                    }
                }
                else
                {
                    newFolderNames.Add(fileName);
                }
            }

            // If we only have one candidate for the new folder name 
            string newFolderName = String.Empty;
            if (newFolderNames.Count == 1)
                newFolderName = newFolderNames[0];
            else
                newFolderName = DefaultFolderName;

            // Have the user approve the new foldername
            string value = newFolderName;
            if (MessageBoxes.InputBox("Extraction folder", "Folder name:", ref value) == DialogResult.OK)
            {
                newFolderName = value;
            }

            return newFolderName;
        }
    }
}
