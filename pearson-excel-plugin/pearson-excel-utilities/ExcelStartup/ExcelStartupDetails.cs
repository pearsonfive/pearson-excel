using System.Collections.Generic;
using System.Xml.Serialization;

namespace Pearson.Excel.Plugin.ExcelStartup
{
    public class ExcelStartupDetails
    {
        public int MaximumTimeAllowed { get; set; }
        public int MaximumAttempts { get; set; }
        public Addins Addins { get; set; } = new Addins();
    }

    public class Addins
    {
        public List<StandardAddin> StandardAddins { get; set; } = new List<StandardAddin>();
    }

    public class StandardAddin
    {
        [XmlAttribute("name")]
        public string Name { get; set; }

        [XmlAttribute("path")]
        public string Path { get; set; }

        [XmlAttribute("check")]
        public bool Check { get; set; }
    }
}