using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace DocxToXlsx
{
    [XmlType("settings")]
    public class Settings
    {
        [XmlElement("application")]
        public ApplicationSettings ApplicationSettings { get; set; }

        [XmlElement("mapping")]
        public Mapping Mapping { get; set; }
    }

    public class ApplicationSettings
    {
        [XmlElement("locations")]
        public FileLocations FileLocations { get; set; }
    }

    public class FileLocations
    {
        [XmlElement("source")]
        public string Source { get; set; }

        [XmlElement("destination")]
        public string Destination { get; set; }
    }

    public class Mapping
    {
        public Mapping()
        {
            HeadingStyles = new List<string>();
            Sheets = new List<Sheet>();
        }

        [XmlArray("headingstyles")]
        [XmlArrayItem("style")]
        public List<string> HeadingStyles { get; set; }

        [XmlArray("sheets")]
        [XmlArrayItem("sheet")]
        public List<Sheet> Sheets { get; set; }
    }

    public class Sheet
    {
        public Sheet()
        {
        }

        [XmlElement("name")]
        public string Name { get; set; }

        [XmlElement("elements")]
        public Elements Elements { get; set; }
    }

    public class Elements
    {
        public Elements()
        {
            Captions = new List<Caption>();
            Filenames = new List<Filename>();
            Paragraphs = new List<Paragraph>();
        }

        [XmlArray("captions")]
        [XmlArrayItem("caption")]
        public List<Caption> Captions { get; set; }

        [XmlArray("filenames")]
        [XmlArrayItem("filename")]
        public List<Filename> Filenames { get; set; }

        [XmlArray("paragraphs")]
        [XmlArrayItem("paragraph")]
        public List<Paragraph> Paragraphs { get; set; }
    }

    public class Caption: CellInfo
    {
        [XmlElement("text")]
        public string Text { get; set; }
    }

    public class Filename: CellInfo
    {
        [XmlElement("NthWord", IsNullable = true)]
        public int? NthWord { get; set; }
    }

    public class Paragraph: CellInfo
    {
        [XmlElement("heading")]
        public string Heading { get; set; }
    }


    public class CellInfo
    {
        private static string[] _verticalOrientation = new string[] {"y", "vertical", "row", "rows" };
        private static string[] _horizontalOrientation = new string[] { "x", "horizontal", "col", "cols" };

        [XmlElement("row")]
        public int Row { get; set; }

        [XmlElement("col")]
        public int Col { get; set; }

        [XmlElement("orientation")]
        public string Orientation { get; set; }

        public void SetNextCell()
        {
            if (!string.IsNullOrEmpty(Orientation))
            {
                if (_verticalOrientation.Contains(Orientation.ToLower()))
                {
                    Row++;
                }
                else if (_horizontalOrientation.Contains(Orientation.ToLower()))
                {
                    Col++;
                }
            }
        }
    }
}
