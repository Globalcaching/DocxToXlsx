using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxToXlsx
{
    public class Converter
    {
        private Settings _settings;
        private List<string> _sourceFiles;
        private IWorkbook workbook;

        public Converter(Settings settings)
        {
            _settings = settings;
            _sourceFiles = new List<string>();
        }

        public void Execute()
        {
            //get source files
            if (File.Exists(_settings.ApplicationSettings.FileLocations.Source))
            {
                _sourceFiles.Add(_settings.ApplicationSettings.FileLocations.Source);
            }
            else if (Directory.Exists(_settings.ApplicationSettings.FileLocations.Source))
            {
                GetSourceFiles(_settings.ApplicationSettings.FileLocations.Source);
            }

            //set destination
            if (File.Exists(_settings.ApplicationSettings.FileLocations.Destination))
            {
                workbook = WorkbookFactory.Create(_settings.ApplicationSettings.FileLocations.Destination);
            }
            else
            {
                workbook = new XSSFWorkbook();
            }

            //create sheets
            foreach (var sheet in _settings.Mapping.Sheets)
            {
                if (workbook.GetSheet(sheet.Name) == null)
                {
                    workbook.CreateSheet(sheet.Name);
                }
            }

            //add captions
            foreach (var sheet in _settings.Mapping.Sheets)
            {
                var wbsheet = workbook.GetSheet(sheet.Name);
                foreach (var caption in sheet.Elements.Captions)
                {
                    ICell cell = wbsheet.GetOrCreateCell(caption.Row, caption.Col);
                    cell.SetCellValue(caption.Text);
                }
            }

            //run through all documents
            for (int i = 0; i < _sourceFiles.Count; i++)
            {
                Console.Write(string.Format("{0}...", Path.GetFileName(_sourceFiles[i])));
                foreach (var sheet in _settings.Mapping.Sheets)
                {
                    var wbsheet = workbook.GetSheet(sheet.Name);
                    //add filename
                    foreach (var fn in sheet.Elements.Filenames)
                    {
                        ICell cell = wbsheet.GetOrCreateCell(fn.Row, fn.Col);
                        cell.SetCellValue(Path.GetFileName(_sourceFiles[i]));
                        fn.SetNextCell();
                    }
                    //add paragraph
                    foreach (var p in sheet.Elements.Paragraphs)
                    {
                        ICell cell = wbsheet.GetOrCreateCell(p.Row, p.Col);
                        cell.SetCellValue(GetParagraphFromDocument(_sourceFiles[i], p.Heading));
                        p.SetNextCell();
                    }
                }
                Console.WriteLine("gedaan");
            }

            using (FileStream stream = new FileStream(_settings.ApplicationSettings.FileLocations.Destination, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }

        /*
        public string GetParagraphFromDocument(string filename, string displayParagraph)
        {
            List<string> headingHirachy = _settings.Mapping.HeadingStyles.ToList();
            List<string> activeHeading = new List<string>();
            string activePath = "";
            StringBuilder sb = new StringBuilder();
            bool inParagraph = false;
            using (StreamReader streamReader = new StreamReader(filename))
            {
                var doc = new XWPFDocument(streamReader.BaseStream);
                foreach (var p in doc.Paragraphs)
                {
                    int index = headingHirachy.IndexOf(p.Style);
                    if (index >= 0)
                    {
                        if (inParagraph) break;
                        while (activeHeading.Count > index)
                        {
                            activeHeading.RemoveAt(activeHeading.Count - 1);
                        }
                        while (activeHeading.Count < index)
                        {
                            activeHeading.Add("");
                        }
                        activeHeading.Add(p.Text);
                        activePath = string.Join("@", activeHeading);
                        sb.Length = 0;
                    }
                    else if (string.Compare(displayParagraph,activePath, true)==0)
                    {
                        inParagraph = true;
                        sb.AppendLine(p.Text);
                    }
                }
            }
            return sb.ToString();
        }
        */

        public string GetParagraphFromDocument(string filename, string displayParagraph)
        {
            List<string> headingHirachy = _settings.Mapping.HeadingStyles.ToList();
            string activePath = "";
            StringBuilder sb = new StringBuilder();
            bool inParagraph = false;
            using (StreamReader streamReader = new StreamReader(filename))
            {
                var doc = new XWPFDocument(streamReader.BaseStream);
                foreach (var p in doc.Paragraphs)
                {
                    int index = headingHirachy.IndexOf(p.Style);
                    if (index >= 0)
                    {
                        if (inParagraph) break;
                        activePath = p.Text;
                        sb.Length = 0;
                    }
                    else if (string.Compare(displayParagraph, activePath, true) == 0)
                    {
                        inParagraph = true;
                        sb.AppendLine(p.Text);
                    }
                }
            }
            return sb.ToString();
        }

        public void GetSourceFiles(string path)
        {
            _sourceFiles.AddRange(Directory.GetFiles(path, "*.docx"));
            var dirs = Directory.GetDirectories(path);
            foreach (var d in dirs)
            {
                GetSourceFiles(d);
            }
        }
    }
}
