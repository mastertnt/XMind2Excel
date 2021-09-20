using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using ICSharpCode.SharpZipLib.Zip;
using Newtonsoft.Json;

namespace XMind2Xls.XMind
{
    class MindFile
    {
        public IList<MindSheet> Sheets
        {
            get;
            set;
        }

        public int Version
        {
            get;
            private set;
        }

        public void ReadXMind(string pFilename)
        {
            ZipFile lFile = null;
            try
            {
                FileStream lFileStream = File.OpenRead(pFilename);
                lFile = new ZipFile(lFileStream);
                this.Version = 8;
                foreach (ZipEntry lZipEntry in lFile)
                {
                    if (lZipEntry.Name == "content.json")
                    {
                        this.Version = 11;
                    }
                }

                if (this.Version == 11)
                {
                    ReadXMind11(lFile);
                }
                else
                {
                    ReadXMind8(lFile);
                }

            }
            finally
            {
                if (lFile != null)
                {
                    lFile.IsStreamOwner = true; // Makes close also shut the underlying stream
                    lFile.Close(); // Ensure we release resources
                }
            }
        }

        private bool ReadXMind11(ZipFile pFile)
        {
            foreach (ZipEntry lZipEntry in pFile)
            {
                if (lZipEntry.Name == "content.json")
                {
                    Stream lZipStream = pFile.GetInputStream(lZipEntry);
                    using var sr = new StreamReader(lZipStream, Encoding.UTF8);
                    string lAllContent = sr.ReadToEnd();
                    this.Sheets = JsonConvert.DeserializeObject<IList<MindSheet>>(lAllContent);
                    return true;
                }
            }

            return false;
        }

        private bool ReadXMind8(ZipFile pFile)
        {
            XNamespace lNamespace = "urn:xmind:xmap:xmlns:content:2.0";
            foreach (ZipEntry lZipEntry in pFile)
            {
                if (lZipEntry.Name == "content.xml")
                {
                    Stream lZipStream = pFile.GetInputStream(lZipEntry);
                    XElement lRoot = XElement.Load(lZipStream);
                    this.Sheets = new List<MindSheet>();
                    foreach (var lSheetElement in lRoot.Elements(lNamespace + "sheet"))
                    {
                        this.Sheets.Add(new MindSheet { id = lSheetElement.Attribute("id").Value, title = lSheetElement.Element(lNamespace +  "title").Value} );
                        this.Sheets.Last().rootTopic = ReadXMind8Topic(lSheetElement.Element(lNamespace + "topic"));
                    }
                }
            }

            return true;
        }

        private MindTopic ReadXMind8Topic(XElement pRootElement)
        {
            XNamespace lNamespace = "urn:xmind:xmap:xmlns:content:2.0";
            MindTopic lTopic = new MindTopic() { id= pRootElement.Attribute("id").Value, title = pRootElement.Element(lNamespace + "title").Value, structureClass = pRootElement.Attribute("structure-class")?.Value };
            
            XElement lChildrenElement = pRootElement.Element(lNamespace + "children");
            if (lChildrenElement != null)
            {
                XElement lAttachedElement = lChildrenElement.Element(lNamespace + "topics");
                if (lAttachedElement != null)
                {
                    lTopic.children = new MindAttached {attached = new List<MindTopic>()};
                    foreach (var lTopicElement in lAttachedElement.Elements(lNamespace + "topic"))
                    {
                        lTopic.children.attached.Add(ReadXMind8Topic(lTopicElement));
                    }
                }
            }
            return lTopic;
        }
    }
}
