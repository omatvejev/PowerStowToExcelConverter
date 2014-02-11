using System;
using System.IO;
using System.Collections.Generic;
using System.Drawing;
using System.Xml;
namespace PowerStowToExcelConverter.Core
{
    class Translator
    {
        private Dictionary<string, string> dictionary;

        public Translator(string path)
        {
            FileInfo fileInfo = new FileInfo(path);

            // Check if the file exists
            if (fileInfo.Exists)
            {
                try
                {
                    dictionary = new Dictionary<string, string>();
                    generateDictionary(path);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            else
                throw new Exception("Warning! Could not open Translator.xml in the program directory. The port names will not be translated to their symbol representation");
        }
    
        private void generateDictionary(string path)
        {

            XmlDocument xmlDoc= new XmlDocument(); // Create an XML document object
            xmlDoc.Load(path); // Load the XML document from the specified file

            XmlNodeList port = xmlDoc.GetElementsByTagName(@"port");
            XmlNodeList symbol = xmlDoc.GetElementsByTagName(@"symbol");

            if (port.Count != symbol.Count)
                throw new Exception("Error parsing XML file. Port and symbol size miss match");

            for (int i = 0; i < port.Count; i++)
                dictionary.Add(port[i].InnerText, symbol[i].InnerText);
        }

        public string translate(string name)
        {
            if (dictionary.ContainsKey(name))
                return dictionary[name];
            else
                return name;
        }
    }
}
