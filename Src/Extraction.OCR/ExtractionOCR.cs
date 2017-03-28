using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace Extraction.OCR
{
    public class ExtractionOCR
    {
        #region
        private static readonly ExtractionOCR instance = new ExtractionOCR();
        public static ExtractionOCR Instance { get { return instance; } }
        public static string section_path { get; set; }
        public static int waitTime = 3 * 1000;
        #endregion
        /// <summary>
        /// office2007 MODI组件OCR识别
        /// </summary>
        /// <param name="imgPath"></param>
        /// <returns></returns>
        public string Ocr_2007(string imgPath)
        {
            try
            {
                var imgType = imgPath.Substring(imgPath.Length - 3);
                var data = File.ReadAllBytes(imgPath);
                string imgInfo = "";
                int i = 0;
                var localimgFile = AppDomain.CurrentDomain.BaseDirectory + @"\" + Guid.NewGuid().ToString() + "." + imgType;
                while (!imgInfo.Equals("转换成功") && i < 3)
                {
                    ++i;
                    imgInfo = this.GetBase64(data, imgType, localimgFile);
                }
                MODI.Document doc = new MODI.Document();
                doc.Create(localimgFile);
                MODI.Image image;
                MODI.Layout layout;
                doc.OCR(MODI.MiLANGUAGES.miLANG_CHINESE_SIMPLIFIED, true, true);
                StringBuilder sb = new StringBuilder();
                image = (MODI.Image)doc.Images[0];
                layout = image.Layout;
                sb.Append(layout.Text);
                doc = null;
                var localFilePath = AppDomain.CurrentDomain.BaseDirectory + @"\" + Guid.NewGuid().ToString() + ".txt";
                File.WriteAllText(localFilePath, sb.ToString());
                Console.WriteLine(sb.ToString());
                return localFilePath;
            }
            catch (Exception e)
            {
                File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + @"\log.txt", e.ToString());
                return "";
            }
            finally
            {
                GC.Collect();
            }
        }
        /// <summary>
        /// onenote 2010，注意需要先在onenote中创建笔记本，并且将至转换为onenote2007格式
        /// 推荐使用onenote2016（个人版即可），API与2010类似，（去掉XMLSchema.xs2007参数即可）其他可参考API参数命名。
        /// 注意1：一定要将dll属性中的“嵌入互操作类型”属性关闭
        /// </summary>
        /// <param name="imgPath"></param>
        /// <returns></returns>
        public string Ocr_2010(string imgPath)
        {
            try
            {
                #region 确定section_path存在
                section_path = @"C:\Users\zhensheng\Desktop\打杂\ocr\ocr.one";
                if (string.IsNullOrEmpty(section_path))
                {
                    Console.WriteLine("请先建立笔记本");
                    File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + @"\log.txt", "需要先在onenote中创建笔记本，并且将至转换为onenote2007格式，且将.one文件得路径赋值给section_path");
                    return "";
                }
                #endregion

                #region 准备数据
                //后缀
                var imgType = Path.GetExtension(imgPath);
                imgPath = imgPath.Replace(".", "");

                var data = File.ReadAllBytes(imgPath);
                //根据大小确定重试次数
                int fileSize = Convert.ToInt32(data.Length / 1024 / 1024); // 文件大小 单位M

                string guid = Guid.NewGuid().ToString();
                string pageID = "{" + guid + "}{1}{B0}";  // 格式 {guid}{tab}{??}

                string pageXml;
                XNamespace ns;

                var onenoteApp = new Microsoft.Office.Interop.OneNote.Application();  //onenote提供的API
                if (onenoteApp == null)
                {
                    File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + @"\log.txt", "Microsoft.Office.Interop.OneNote.Application()创建失败");
                    return "";
                }

                //重试使用
                XmlNode xmlNode;
                int retry = 0;

                #endregion

                do
                {
                    #region 创建页面并返回pageID
                    string sectionID;
                    onenoteApp.OpenHierarchy(section_path, null, out sectionID, CreateFileType.cftSection);
                    onenoteApp.CreateNewPage(sectionID, out pageID);
                    #endregion

                    #region 获取onenote页面xml结构格式
                    string notebookXml;
                    onenoteApp.GetHierarchy(null, HierarchyScope.hsPages, out notebookXml, XMLSchema.xs2007);
                    var doc = XDocument.Parse(notebookXml);
                    ns = doc.Root.Name.Namespace;

                    #endregion

                    #region 将图片插入页面
                    Tuple<string, int, int> imgInfo = this.GetBase64(data, imgType);
                    var page = new XDocument(new XElement(ns + "Page",
                                                    new XElement(ns + "Outline",
                                                    new XElement(ns + "OEChildren",
                                                        new XElement(ns + "OE",
                                                        new XElement(ns + "Image",
                                                            new XAttribute("format", imgType), new XAttribute("originalPageNumber", "0"),
                                                            new XElement(ns + "Position",
                                                                new XAttribute("x", "0"), new XAttribute("y", "0"), new XAttribute("z", "0")),
                                                            new XElement(ns + "Size",
                                                                new XAttribute("width", imgInfo.Item2), new XAttribute("height", imgInfo.Item3)),
                                                            new XElement(ns + "Data", imgInfo.Item1)))))));

                    page.Root.SetAttributeValue("ID", pageID);
                    onenoteApp.UpdatePageContent(page.ToString(), DateTime.MinValue, XMLSchema.xs2007);
                    #endregion

                    #region 通过轮询访问获取OCR识别的结果,轮询超时次数为6次
                    int count = 0;
                    do
                    {
                        System.Threading.Thread.Sleep(waitTime * (fileSize > 1 ? fileSize : 1)); // 小于1M的都默认1M
                        onenoteApp.GetPageContent(pageID, out pageXml, PageInfo.piBinaryData, XMLSchema.xs2007);
                    }
                    while (pageXml == "" && count++ < 6);
                    #endregion

                    #region 删除页面
                    onenoteApp.DeleteHierarchy(pageID, DateTime.MinValue);
                    //onenoteApp = null;
                    #endregion

                    #region 从xml中提取OCR识别后的文档信息，然后输出到string中
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(pageXml);
                    XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                    nsmgr.AddNamespace("one", ns.ToString());
                    xmlNode = xmlDoc.SelectSingleNode("//one:Image//one:OCRText", nsmgr);
                    #endregion
                }
                //如果OCR没有识别出信息，则重试三次（个人测试2010失败率为0.2~0.3）
                while (xmlNode == null && retry++ < 3);
                if (xmlNode == null)
                {
                    File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + @"\log.txt", "OCR没有识别出值");
                    return "";
                }
                var localFilePath = AppDomain.CurrentDomain.BaseDirectory + @"\" + Guid.NewGuid().ToString() + ".txt";
                File.WriteAllText(localFilePath, xmlNode.InnerText.ToString());
                Console.WriteLine(xmlNode.InnerText.ToString());

                return localFilePath;
            }
            catch (Exception e)
            {
                File.AppendAllText(AppDomain.CurrentDomain.BaseDirectory + @"\log.txt", e.ToString());
                return "";
            }
        }
        private string GetBase64(byte[] data, string imgType, string filePath)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                MemoryStream ms1 = new MemoryStream(data);
                Bitmap bp = (Bitmap)Image.FromStream(ms1);
                switch (imgType.ToLower())
                {
                    case "jpg":
                        bp.Save(ms, ImageFormat.Jpeg);
                        break;

                    case "jpeg":
                        bp.Save(ms, ImageFormat.Jpeg);
                        break;

                    case "gif":
                        bp.Save(ms, ImageFormat.Gif);
                        break;

                    case "bmp":
                        bp.Save(ms, ImageFormat.Bmp);
                        break;

                    case "tiff":
                        bp.Save(ms, ImageFormat.Tiff);
                        break;

                    case "png":
                        bp.Save(ms, ImageFormat.Png);
                        break;

                    case "emf":
                        bp.Save(ms, ImageFormat.Emf);
                        break;

                    default:
                        return "不支持的图片格式。";
                }
                byte[] buffer = ms.ToArray();
                File.WriteAllBytes(filePath, buffer);
                ms1.Close();
                ms.Close();
                return "转换成功";
                //return new Tuple<string, int, int>(Convert.ToBase64String(buffer), bp.Width, bp.Height);
            }
        }
        private Tuple<string, int, int> GetBase64(byte[] data, string imgType)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                MemoryStream ms1 = new MemoryStream(data);
                Bitmap bp = (Bitmap)Image.FromStream(ms1);
                switch (imgType.ToLower())
                {
                    case "jpg":
                        bp.Save(ms, ImageFormat.Jpeg);
                        break;

                    case "jpeg":
                        bp.Save(ms, ImageFormat.Jpeg);
                        break;

                    case "gif":
                        bp.Save(ms, ImageFormat.Gif);
                        break;

                    case "bmp":
                        bp.Save(ms, ImageFormat.Bmp);
                        break;

                    case "tiff":
                        bp.Save(ms, ImageFormat.Tiff);
                        break;

                    case "png":
                        bp.Save(ms, ImageFormat.Png);
                        break;

                    case "emf":
                        bp.Save(ms, ImageFormat.Emf);
                        break;

                    default:
                        return new Tuple<string, int, int>("不支持的图片格式。", 0, 0);
                }
                byte[] buffer = ms.ToArray();
                ms1.Close();
                ms.Close();
                return new Tuple<string, int, int>(Convert.ToBase64String(buffer), bp.Width, bp.Height);
            }
        }

        // 多个图片合并的方法参考 Graphics 类：先获取图片大小设置画布，然后使用 Graphics 类合并。同时需要使用 Image 类
    }
}
