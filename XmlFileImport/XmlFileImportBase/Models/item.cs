using System.Xml.Serialization;

namespace XmlFileImportBase.Models
{
    public class item
    {
        [XmlElement("title")]
        public string title { get; set; }
        [XmlElement("link")]
        public string link { get; set; }
        [XmlElement("description")]
        public string description { get; set; }
        [XmlElement("category")]
        public string category { get; set; }
        [XmlElement("pubDate")]
        public string pubDate { get; set; }

        public item(string Title, string Link, string Description, string Category, string PubDate)
        {
            title = Title;
            link = Link;
            description = Description;
            category = Category;
            pubDate = PubDate;
        }
        public item() { }
    }
}
