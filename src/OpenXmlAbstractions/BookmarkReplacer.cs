using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OpenXmlAbstractions
{
    public class BookmarkReplacer
    {
        private static XNamespace BookmarkReplacerCustomNamespace = "http://powertools.codeplex.com/2011/bookmarkreplacer";

        public static byte[] GenerateDocument(string documentFilePath, string bookmarkDataXML, string tempFilePath)
        {
            var documentBookmarks = GetDocumentBookmarks(documentFilePath);

            var matchedData = CompareDocumentBookmarksToXML(documentBookmarks, bookmarkDataXML);

            if (File.Exists(tempFilePath))
            {
                File.Delete(tempFilePath);
            }

            //Copies the Template document to a specified temporary directory
            File.Copy(documentFilePath, tempFilePath);
            var htmlList = ReplaceDocumentBookmarks(tempFilePath, matchedData);

            ReplaceHTMLContent(tempFilePath, htmlList);

            var fileByteArray = ConvertFileToByteArray(tempFilePath);

            //Deletes the temporary file created for document generation
            File.Delete(tempFilePath);

            return fileByteArray;
        }

        private static object FlattenParagraphsTransform(XNode node)
        {
            var element = node as XElement;

            if (element != null)
            {
                if (element.Name == W.p)
                {
                    return element
                        .Elements()
                        .Select(e => FlattenParagraphsTransform(e))
                        .Concat(
                            new[]
                    {
                        new XElement(W.p, element.Attributes(), element.Elements(W.pPr))
                    });
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => FlattenParagraphsTransform(n)));
            }

            return node;
        }

        private struct BlockLevelState
        {
            public int Index;
            public bool BlockLevelElementBefore;
        };

        private static object UnflattenParagraphsTransform(XNode node)
        {
            var element = node as XElement;

            if (element != null)
            {
                if (element.Elements().Any(e => e.Name == W.p))
                {
                    var paraIndex = element
                        .Elements()
                        .Rollup(
                            new BlockLevelState
                            {
                                Index = 0,
                                BlockLevelElementBefore = false,
                            },
                            (e2, s) =>
                            {
                                if (s.BlockLevelElementBefore)
                                    return new BlockLevelState
                                    {
                                        Index = s.Index + 1,
                                        BlockLevelElementBefore =
                                            (e2.Name == W.p ||
                                                e2.Name == W.tbl ||
                                                e2.Name == W.tcPr ||
                                            (e2.Name == W.sdt &&
                                                e2.Descendants(W.p).Any())),
                                    };

                                return new BlockLevelState
                                {
                                    Index = s.Index,
                                    BlockLevelElementBefore =
                                        (e2.Name == W.p ||
                                            e2.Name == W.tbl ||
                                            e2.Name == W.tcPr ||
                                        (e2.Name == W.sdt &&
                                            e2.Descendants(W.p).Any())),
                                };
                            });

                    var zipped = element.Elements().Zip(paraIndex, (a, b) =>
                        new
                        {
                            Element = a,
                            ParaIndex = b,
                        });

                    var grouped = zipped.GroupAdjacent(e3 => e3.ParaIndex.Index);

                    var newElements = grouped
                        .Select(g =>
                        {
                            var lastElement = g.Last().Element;

                            if (lastElement.Name != W.p)
                            {
                                return (object)g.Select(gc => UnflattenParagraphsTransform(gc.Element));
                            }

                            var newParagraph = new XElement(W.p,
                                lastElement.Attributes(),
                                g.Take(g.Count() - 1).Select(e4 => UnflattenParagraphsTransform(e4.Element)));

                            return newParagraph;
                        });

                    return new XElement(element.Name, element.Attributes(), newElements);
                }

                return new XElement(element.Name, element.Attributes(), element.Nodes().Select(n => UnflattenParagraphsTransform(n)));
            }

            return node;
        }

        private static object ReplaceInsertElement(XNode node, string replacementText)
        {
            var element = node as XElement;

            if (element != null)
            {
                if (element.Name == BookmarkReplacerCustomNamespace + "Insert")
                {
                    XName parentName = element.Parent.Name;
                    if (parentName == W.body || parentName == W.tc || parentName == W.txbxContent)
                    {
                        return new XElement(W.p, new XElement(W.r, element.Elements(), new XElement(W.t, replacementText)));
                    }

                    return new XElement(W.r, element.Elements(), new XElement(W.t, replacementText));
                }

                return new XElement(element.Name, element.Attributes(), element.Nodes().Select(n => ReplaceInsertElement(n, replacementText)));
            }

            return node;
        }

        private static object DemoteRunChildrenOfBodyTransform(XNode node)
        {
            var element = node as XElement;

            if (element != null)
            {
                if (element.Name == W.r && element.Parent.Name == W.body)
                {
                    return new XElement(W.p, element);
                }
                    
                return new XElement(element.Name, element.Attributes(), element.Nodes().Select(n => DemoteRunChildrenOfBodyTransform(n)));
            }

            return node;
        }

        private static void ReplaceBookmarkText(WordprocessingDocument doc, string bookmarkName, string replacementText)
        {
            XDocument xDoc = doc.MainDocumentPart.GetXDocument();
            XElement bookmark = xDoc.Descendants(W.bookmarkStart).FirstOrDefault(d => (string)d.Attribute(W.name) == bookmarkName);

            // Checks for illegal bookmarks
            if (bookmark == null)
                throw new Exception("noBookmark");
            if (bookmark.Parent.Name.Namespace == M.m)
                throw new Exception("noMathFormula");
            if (RevisionAccepter.HasTrackedRevisions(doc))
                throw new Exception("noReplaceTrackedChanges");
            if (xDoc.Descendants(W.sdt).Any())
                throw new Exception("noContentControls");

            XElement newRoot = (XElement)FlattenParagraphsTransform(xDoc.Root);
            var start = newRoot.Descendants(W.fldChar).Where(d => (string)d.Attribute(W.fldCharType) == "begin");
            XElement startBookmarkElement = newRoot.Descendants(W.bookmarkStart)
                .Where(d => (string)d.Attribute(W.name) == bookmarkName)
                .FirstOrDefault();
            int bookmarkId = (int)startBookmarkElement.Attribute(W.id);
            XElement endBookmarkElement = newRoot.Descendants(W.bookmarkEnd)
                .Where(d => (int)d.Attribute(W.id) == bookmarkId)
                .FirstOrDefault();

            // Checks for illegal bookmarks
            if (startBookmarkElement.Ancestors(W.hyperlink).Any() ||
                endBookmarkElement.Ancestors(W.hyperlink).Any())
                throw new Exception("noHyperlinks");
            if (startBookmarkElement.Ancestors(W.fldSimple).Any() ||
                endBookmarkElement.Ancestors(W.fldSimple).Any())
                throw new Exception("noSimpleField");
            if (startBookmarkElement.Ancestors(W.smartTag).Any() ||
                endBookmarkElement.Ancestors(W.smartTag).Any())
                throw new Exception("noSmartTag");
            if (startBookmarkElement.Parent != endBookmarkElement.Parent)
                throw new Exception("noSameLevels");

            XElement parentElement = startBookmarkElement.Parent;

            var elementsBetweenBookmarks = startBookmarkElement.ElementsAfterSelf().TakeWhile(e => e != endBookmarkElement);

            var newElements = parentElement.Elements().TakeWhile(e => e != startBookmarkElement)
                .Concat(new[]
            {
            startBookmarkElement,
            new XElement(BookmarkReplacerCustomNamespace + "Insert",
                elementsBetweenBookmarks.Where(e => e.Name == W.r).Take(1).Elements(W.rPr).FirstOrDefault()),
            })
                .Concat(elementsBetweenBookmarks.Where(e => e.Name != W.p &&
                    e.Name != W.r && e.Name != W.tbl))
                .Concat(new[]
            {
            endBookmarkElement
            })
                .Concat(endBookmarkElement.ElementsAfterSelf());
            parentElement.ReplaceNodes(newElements);


            newRoot = (XElement)UnflattenParagraphsTransform(newRoot);
            newRoot = (XElement)ReplaceInsertElement(newRoot, replacementText);
            newRoot = (XElement)DemoteRunChildrenOfBodyTransform(newRoot);

            xDoc.Elements().First().ReplaceWith(newRoot);
            doc.MainDocumentPart.PutXDocument();
        }

        private static List<string> ReplaceDocumentBookmarks(string documentFilePath, Dictionary<string, string> bookmarks)
        {
            var htmlList = new List<string>();

            using (var doc = WordprocessingDocument.Open(documentFilePath, true))
            {
                foreach (var bookmark in bookmarks)
                {
                    if (bookmark.Value.StartsWith("<html>"))
                    {
                        htmlList.Add(bookmark.Value);
                    }

                    ReplaceBookmarkText(doc, bookmark.Key, bookmark.Value);
                }
            }

            return htmlList;
        }

        /// <summary>
        /// Uses AltChunk to insert HTML content into the document, replacing the inline HTML 
        /// </summary>
        /// <param name="documentFilePath"></param>
        /// <param name="htmlContent"></param>
        private static void ReplaceHTMLContent(string documentFilePath, List<string> htmlContent)
        {
            using (var doc = WordprocessingDocument.Open(documentFilePath, true))
            {
                var i = 1;

                foreach (var html in htmlContent)
                {
                    var altChunkId = "AltChunkId" + i;
                    var mainPart = doc.MainDocumentPart;
                    var chunk = mainPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Xhtml, altChunkId);

                    using (Stream chunkStream = chunk.GetStream())
                    {
                        using (var stringWriter = new StreamWriter(chunkStream, Encoding.UTF8)) //Encoding.UTF8 is important to remove special characters
                        {
                            stringWriter.Write(html);
                        }
                    }

                    var altChunk = new AltChunk();
                    altChunk.Id = altChunkId;
                    mainPart.Document
                                .Body
                                .InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().FirstOrDefault(e => e.InnerText == html));
                    mainPart.Document.Body.Elements<Paragraph>().FirstOrDefault(e => e.InnerText == html).Remove();
                    mainPart.Document.Save();
                    i++;
                }
            }
        }

        private static string GetBookmarkText(WordprocessingDocument doc, string bookmarkName)
        {
            var xDoc = doc.MainDocumentPart.GetXDocument();
            var containsBookmark = xDoc.Descendants(W.bookmarkStart)
                .Where(d => (string)d.Attribute(W.name) == bookmarkName)
                .Any();

            if (!containsBookmark)
            {
                throw new Exception("noBookmark");
            }

            var newRoot = (XElement)FlattenParagraphsTransform(xDoc.Root);

            var startBookmarkElement = newRoot.Descendants(W.bookmarkStart)
                .Where(d => (string)d.Attribute(W.name) == bookmarkName)
                .FirstOrDefault();

            var bookmarkId = (int)startBookmarkElement.Attribute(W.id);

            var endBookmarkElement = newRoot.Descendants(W.bookmarkEnd)
                .Where(d => (int)d.Attribute(W.id) == bookmarkId)
                .FirstOrDefault();

            if (startBookmarkElement.Parent != endBookmarkElement.Parent)
            {
                throw new Exception("notSameLevels");
            }

            var parentElement = startBookmarkElement.Parent;

            var elementsBetweenBookmarks = startBookmarkElement
                .ElementsAfterSelf()
                .TakeWhile(e => e != endBookmarkElement);

            var text = elementsBetweenBookmarks
                .Select(e =>
                {
                    if (e.Name == W.r)
                    {
                        return e.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                    }

                    if (e.Name == W.p)
                    {
                        return Environment.NewLine;
                    }

                    return string.Empty;
                })
                .StringConcatenate();

            return text;
        }

        /// <summary>
        /// Enumerates over a .docx document and returns a ArrayList of the bookmarks it contains
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns>ArrayList of bookmarks</returns>
        private static ArrayList GetDocumentBookmarks(string filePath)
        {
            using (var doc = WordprocessingDocument.Open(filePath, true))
            {
                if (RevisionAccepter.HasTrackedRevisions(doc) == true)
                {
                    throw new Exception("noTrackedChanges");
                }

                var bookmarks = new ArrayList();
                var xDoc = doc.MainDocumentPart.GetXDocument();

                if (HasBookmarks(xDoc) == false)
                {
                    throw new Exception("noBookmarks");
                }

                var newRoot = (XElement)FlattenParagraphsTransform(xDoc.Root);

                foreach (var newStart in newRoot.Descendants(W.bookmarkStart))
                {
                    bookmarks.Add(newStart.LastAttribute.Value.ToString());
                }

                return bookmarks;
            }
        }

        private static bool HasBookmarks(XDocument xDoc)
        {
            return xDoc.Descendants(W.bookmarkStart).Any();
        }

        /// <summary>
        /// Strips the w:name bookmark from the XElement
        /// </summary>
        /// <param name="elementsBetweenBookmarks"></param>
        /// <returns>string</returns>
        private static string StripBookmarkTextOutOfElement(IEnumerable<XElement> elementsBetweenBookmarks)
        {
            return elementsBetweenBookmarks
                .Select(e =>
                {
                    if (e.Name == W.r)
                    {
                        return e.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                    }

                    if (e.Name == W.p)
                    {
                        return Environment.NewLine;
                    }

                    return string.Empty;
                })
                .StringConcatenate().Replace(" ", "");
        }

        /// <summary>
        /// Coverts a File to a Byte Array
        /// </summary>
        /// <param name="documentName"></param>
        /// <returns>Byte Array</returns>
        private static byte[] ConvertFileToByteArray(string documentName)
        {
            using (var docStream = new FileStream(documentName, FileMode.Open, FileAccess.Read))
            {
                var bytes = new byte[docStream.Length];
                var numBytesToRead = (int)docStream.Length;
                var numBytesRead = 0;
                while (numBytesToRead > 0)
                {
                    var num = docStream.Read(bytes, numBytesRead, numBytesToRead);

                    if (num == 0) break;

                    numBytesRead += num;
                    numBytesToRead -= num;
                }

                return bytes;
            }
        }

        /// <summary>
        /// Compares the bookmarks from GetDocumentBookmarks to the Database Bookmark XML
        /// Returns a Dictionary of the matching bookmarks
        /// </summary>
        /// <param name="bookmarks"></param>
        /// <param name="bookmarkDataXML"></param>
        /// <returns></returns>
        private static Dictionary<string, string> CompareDocumentBookmarksToXML(ArrayList bookmarks, string bookmarkDataXML)
        {
            var xmlBookmarks = new Dictionary<string, string>();
            var matchingBookmarks = new Dictionary<string, string>();

            // Parse the XML into a Dictionary
            var reader = XmlReader.Create(new StringReader(bookmarkDataXML));
            while (reader.Read())
            {
                if (reader.IsStartElement())
                {
                    if (reader.Name == "bookmark")
                    {

                        if (reader.HasAttributes)
                        {
                            var attr = reader.GetAttribute("name");
                            xmlBookmarks.Add(attr, reader.ReadString());
                        }
                    }
                }
            }

            //Compare the XML keys to the document bookmarks, return a dictionary of the matching bookmarks
            foreach (var xmlKey in xmlBookmarks)
            {
                if (bookmarks.Contains(xmlKey.Key))
                {
                    matchingBookmarks.Add(xmlKey.Key, xmlKey.Value);
                }
            }

            return matchingBookmarks;
        }
    }
}