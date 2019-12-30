using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExtractMediaFilesExtension;

namespace DryRun
{
    public class Program : ExtractMediaFilesExtensionBase
    {
        [STAThread]
        static void Main(string[] args)
        {
            string startFolder;
            List<string> selectedThings = new List<string>()
            {
                @"\\Synology\Video\SquidBillies"
            };

                if (selectedThings != null && selectedThings.Any())
                {
                    List<string> filesToExtract = new List<string>();
                    startFolder = Path.GetPathRoot(selectedThings[0]);

                    (new ExtractMediaFilesExtensionBase()).GetExtractionList(
                        selectedThings.Select(t => Path.Combine(startFolder, t)).ToList()
                        , filesToExtract);

                    string newFolderName = (new ExtractMediaFilesExtensionBase()).GetNewFolderName(filesToExtract);

                    (new ExtractMediaFilesExtensionBase()).MoveTheFiles(filesToExtract, startFolder, newFolderName);
                }
        }
    }
}
