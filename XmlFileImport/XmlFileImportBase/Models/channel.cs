using System.Collections.Generic;
using System.Xml.Serialization;

namespace XmlFileImportBase.Models
{
    [XmlRoot("channel")]
    public class channel
    {
        public channel() { Items = new List<item>(); }
        [XmlElement("item")]
        public List<item> Items { get; set; }
    }
}
