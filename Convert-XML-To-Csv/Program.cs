using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.IO;

namespace Convert_XML_To_Csv
{
    class Program
    {
        static string[] GetXml()
        {
            string[] dir = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.xml", SearchOption.AllDirectories);

            return dir;
        }

        static void ConvertToCsv(string filePath)
        {
            XmlTextReader reader = new XmlTextReader(filePath);
            XmlNodeType type;
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string outputFile = Directory.GetCurrentDirectory() + @"\CSV\" + fileName + ".csv";

            //Column Header
            string output = "BulletinID,KBID,Title,Severity,BulletinURL,InformationURL,DownloadURL,IsInstalled";
            output += Environment.NewLine;

            //Column
            string BulletinID = null,
                   KBID = null,
                   Title = null,
                   Severity = null,
                   BulletinURL = null,
                   InformationURL = null,
                   DownloadURL = null,
                   IsInstalled = null;

            //Switch
            bool hasBulletin = false;

            while (reader.Read())
            {
                type = reader.NodeType;

                if (type == XmlNodeType.Element)
                {
                    if (reader.Name == "UpdateData")
                    {
                        hasBulletin = false;
                        reader.ReadAttributeValue();
                        BulletinID = "\"" + reader.GetAttribute("BulletinID") + "\"";
                        KBID = "\"" + reader.GetAttribute("KBID") + "\"";
                        Severity = "\"" + reader.GetAttribute("Severity") + "\"";
                        IsInstalled = "\"" + reader.GetAttribute("IsInstalled") + "\"";
                    }

                    if (reader.Name == "Title")
                    {
                        reader.Read();
                        Title = "\"" + reader.Value + "\"";
                    }

                    if (reader.Name == "BulletinURL")
                    {
                        hasBulletin = true;
                        reader.Read();
                        BulletinURL = "\"" + reader.Value + "\"";
                    }

                    if (reader.Name == "InformationURL")
                    {
                        reader.Read();
                        InformationURL = "\"" + reader.Value + "\"";
                    }

                    if (reader.Name == "DownloadURL")
                    {
                        reader.Read();
                        DownloadURL = "\"" + reader.Value + "\"";

                        if (hasBulletin == true)
                        {
                            output += BulletinID + "," + KBID + "," + Title + "," + Severity + "," + BulletinURL + "," + InformationURL + "," + DownloadURL + "," + IsInstalled + Environment.NewLine;
                        }
                        else
                        {
                            output += "," + KBID + "," + Title + "," + Severity + "," + "," + InformationURL + "," + DownloadURL + "," + IsInstalled + Environment.NewLine;
                        }
                    }
                }
            }
            reader.Close();

            File.WriteAllText(outputFile, output);
        }

        static void Main(string[] args)
        {
            string[] xmlFiles = GetXml();

            foreach (var item in xmlFiles)
            {
                ConvertToCsv(item);
            }
        }
    }
}
