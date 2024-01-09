using System;
using System.IO;
using System.Security;
using System.Xml;
using Xunit;

namespace ConverterToXml.Test
{
    
    public class DocToXmlTest
    {
        [Fact]
        public void DocToDocxConvertToXml_NotNull()
        {
            //Converters.DocToXml converter = new Converters.DocToXml();
            //string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            //string path = curDir + @"/Files/doc1.doc";

            //using (FileStream fs = new FileStream(path, FileMode.Open))
            //{
            //    string result = converter.Convert(fs);
            //    Assert.NotNull(result);
            //}

            ConverterToXml.Converters.DocToXml converter = new ConverterToXml.Converters.DocToXml();
            string curDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string path = @"C:\01\test.doc"; // curDir +

            using (FileStream fs = new FileStream(path, FileMode.Open))
            {
                string result = converter.Convert(fs);

                var xmlWithEscapedCharacters = SecurityElement.Escape(result);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(xmlWithEscapedCharacters.Substring(1)); // xmlWithEscapedCharacters

                string filePath = @"C:\01\test.xml";
                xmlDoc.Save(filePath);
                //Assert.NotNull(result);
            }

        }
    }
}
