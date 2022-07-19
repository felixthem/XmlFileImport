using System;
using System.IO;
using System.Xml.Serialization;
using XmlFileImportBase.Models;

namespace XmlFileImportBase.Helpers
{
    public static class XmlHelper
    {
        public static channel ParseXML(string filePath)
        {
            try
            {
                channel ch = new channel();

                if (!File.Exists($@"data.xml"))
                    throw new Exception($@"Файл для считывания не существует!");

                XmlSerializer formatter = new XmlSerializer(typeof(channel));

                using (FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate))
                {
                    ch = formatter.Deserialize(fs) as channel;
                }

                return ch;
            }
            catch (Exception e)
            {
                throw new Exception($@"Возникли трудности при считывании данных из xml. {e.Message}", e);
            }
        }
    }
}
