using HtmlAgilityPack;
using Newtonsoft.Json;
using QuickType;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Nirsoft_check
{
    class Program
    {
        static Dictionary<string, string> dict_of_links = new Dictionary<string, string>();
        static readonly string txt = "sites.txt";
        static void Main(string[] args)
        {
            Load();

            //CheckOneFile(@"E:\\hdd\\batexe\\nirsoft\\iehv.exe");
            CheckOneFile(@"E:\\hdd\\batexe\\passreccommandline\\mailpv.exe");

            //CheckOneFolder(@"E:\hdd\batexe\nirsoft");
            //CheckOneFolder(@"E:\hdd\batexe\nirsoft64");
            //CheckOneFolder(@"E:\hdd\batexe\passreccommandline");

            Save();
            Console.WriteLine("done");
            Console.ReadLine();
        }
        public static void Load()
        {
            if (File.Exists(txt))
            {
                dict_of_links = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(txt));
            }

            Console.WriteLine("loaded {0} records", dict_of_links.Count());
        }
        public static void Save()
        {
            File.WriteAllText(txt, JsonConvert.SerializeObject(dict_of_links));
            Console.WriteLine("saved {0} records", dict_of_links.Count());
        }

        public static void CheckOneFolder(string fullfolderPath)
        {
            var files = System.IO.Directory.GetFiles(fullfolderPath);
            Console.WriteLine("\nFolder: {0}: total {1} files", fullfolderPath, files.Length);
            int counter = 0;
            foreach (var item_fullPath in files)
            {
                if (FileVersionInfo.GetVersionInfo(item_fullPath).CompanyName == "NirSoft")
                {
                    CheckOneFile(item_fullPath);
                    counter++;
                }
            }
            Console.WriteLine("\nFolder: {0} done: checked {1} files", fullfolderPath, counter);
        }
        public static void CheckOneFile(string fullName)
        {
            var clearName = Path.GetFileName(fullName);

            string firstLink = "";
            if (dict_of_links.ContainsKey(clearName)) firstLink = dict_of_links[clearName];
            else
            {
                //Google Part ---/ Link from Name
                string apikey = "AIzaSyA0gCLSi7El_3RedtU4HKzniO5FxfyU2v0";
                string cx = "90eec8eaf4dcb8dc9";
                var url = "https://www.googleapis.com/customsearch/v1?key=" + apikey + "&cx=" + cx + "&q=" + "site:nirsoft.net " + clearName;
                //Console.WriteLine(url);
                using (var wc = new System.Net.WebClient())
                {
                    var json = wc.DownloadString(url);
                    var jsonobj = V1.FromJson(json);
                    firstLink = jsonobj.Items.First(c => c.FormattedUrl.Contains("https://www.nirsoft.net/utils/")).FormattedUrl;
                    if (firstLink == null) return;
                    //Console.WriteLine(firstLink);
                    dict_of_links.Add(clearName, firstLink);
                    Save();
                    Load();
                }
            }

            //Link
            var web = new HtmlWeb();
            var doc = web.Load(firstLink);

            //Version From Title
            var node = doc.DocumentNode.SelectNodes("//*[@class=\"utilcaption\"]/tr/td[2]");
            string titleText = "";
            string version = "";
            if (node != null)
            {
                titleText = node.FirstOrDefault().InnerText;
                //Console.WriteLine(titleText);
                //parse
                var mas = titleText.Split();
                var filtered = mas.Where(c => c.Contains("v") && c.Contains(".")).ToList();
                if (filtered.Count() == 1) version = filtered[0].Replace("v", "").Replace(":", "");
                //Console.WriteLine(version);
            }

            //Version_from_vers_history
            var node2 = doc.DocumentNode.SelectNodes("//*[contains(text(),'Version ')]");
            string titleText2 = "";
            string version2 = "";
            if (node2 != null)
            {
                var nodes = node2.FirstOrDefault(c =>
                (c.InnerHtml.IndexOf("Version ") == 0 || c.InnerHtml.IndexOf("Version ") == 1) && c.Name == "li");

                if (nodes != null)
                {
                    titleText2 = nodes.ChildNodes.FirstOrDefault(c => c.InnerText.Contains("Version ")).InnerText;

                    //parse
                    string v = "Version ";
                    var pos = titleText2.IndexOf(v);
                    if (pos > -1)
                    {
                        var end1 = titleText2.IndexOf(":", pos + v.Length);
                        var end2 = titleText2.IndexOf("\n", pos + v.Length);
                        var end3 = titleText2.IndexOf(" -", pos + v.Length);
                        var end = new List<int>() { titleText2.Length, end1, end2, end3 }.Where(k => k > 0).Min();

                        version2 = titleText2.Substring(pos + 8, end - pos - v.Length).Trim();
                        //Console.WriteLine(version2);
                    }
                }
            }

            if (version2 == "")
            {
                //https://www.nirsoft.net/utils/iehv.html
                //table
                var node3 = doc.DocumentNode.SelectNodes("//*[contains(text(),'Version')]");
                var finded_node = node3.FirstOrDefault(c => c.OriginalName == "th");
                if (finded_node != default)
                {
                    var pos = finded_node.ParentNode.ChildNodes.GetNodeIndex(finded_node);
                    var nodes = finded_node.ParentNode.NextSibling.ChildNodes.Where(i => i.Name == "td").ToArray();
                    version2 = nodes[pos].InnerText.Trim();
                }
            }

            var v3 = FileVersionInfo.GetVersionInfo(fullName).FileVersion;
            //Console.WriteLine(v3);

            if (version == "" || version2 == "")
            {
                Console.WriteLine("\nNoTitleOrVersionHistoryOrBoth");
                Console.WriteLine("{0} - {1} - {2} - {3}/{4}", clearName, firstLink, v3, version, version2);
            }
            else
            {
                if (version == version2)
                {
                    if (v3 == version)
                    {
                        //Console.WriteLine("{0} - {1} - {2} - no need update", clearName, firstLink, v3);
                    }
                    else
                    {
                        Console.WriteLine("{0} - {1} - {2} - {3}", clearName, firstLink, v3, version);
                    }
                }
                else
                {
                    Console.WriteLine("\nSite Version Mismatch");
                    Console.WriteLine("{0} - {1} - {2} - {3}/{4}", clearName, firstLink, v3, version, version2);
                }
            }
        }
    }
}
