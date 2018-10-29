using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML_Sample
{
    class ChangeFormatting
    {
        static void Main(string[] args)
        {
            
            string applicationPath = System.IO.Directory.GetParent(System.IO.Directory.GetParent(Environment.CurrentDirectory).ToString()).ToString();
            string documentPath = applicationPath + @"\Documents\Sassafras Springs Vineyard 401(k) Plan_2018 Safe Harbor Notice.doc";

            WriteToWordDoc(documentPath);

        }

        public static void WriteToWordDoc(string filepath)
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument =
                 WordprocessingDocument.Open(filepath, true))
            {
                // Get a reference to the main document part.
                MainDocumentPart docPart = wordprocessingDocument.MainDocumentPart;
                // Assign a reference to the existing document body.
                Body body = docPart.Document.Body;

                //This for References :: may be needed later
                //StyleDefinitionsPart part = wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart;
                //IEnumerable<HeaderPart> headerParts = wordprocessingDocument.MainDocumentPart.HeaderParts;
                //FontTablePart fontTablePart = wordprocessingDocument.MainDocumentPart.FontTablePart;


                //remove border and shading from first main title
                Paragraph paragraph2 = body.Elements<Paragraph>().ElementAt(1);
                //This way can get property value
                //ParagraphProperties paragraphProperties2 = paragraph2.GetFirstChild<ParagraphProperties>(); 
                ParagraphProperties paragraphProperties2 = paragraph2.ParagraphProperties;
                Run run = paragraph2.GetFirstChild<Run>();

                run.RunProperties.Bold.Remove();

                //ParagraphMarkRunProperties pmrp1 = paragraphProperties2.ParagraphMarkRunProperties;
                paragraphProperties2.ParagraphBorders.Remove();
                paragraphProperties2.Shading.Remove();

                //remove first main title font weight = bold, change font size = 14
                var runProp = new RunProperties();
                var runFont = new RunFonts { Ascii = "Calibri Light" };

                // 28 point font size : always half of size ( 28 /2 = 14)
                var size = new FontSize { Val = new StringValue("28") };
                Color color = new Color() { Val = "2f5496" };


                runProp.Append(runFont);
                runProp.Append(size);
                runProp.Append(color);

                run.PrependChild(runProp);

                IEnumerable<Run> otherRuns = paragraph2.Elements<Run>().Skip(1);
                otherRuns.ToList().ForEach(r => r.Remove());


                //paragraph dynamically bind after remove node , so this may can call last in function
                //remove  empty paragraph before  main title
                Paragraph paragraph1 = body.Elements<Paragraph>().ElementAt(0);
                if (paragraph1 != null)
                {
                    paragraph1.RemoveAllChildren();
                    paragraph1.Remove();
                }

                //remove  second main title
                Paragraph paragraph3 = body.Elements<Paragraph>().ElementAt(1);
                if (paragraph3 != null)
                {
                    paragraph3.RemoveAllChildren();
                    paragraph3.Remove();
                }

                //remove other two empty paragraph after second main title
                Paragraph paragraph4 = body.Elements<Paragraph>().ElementAt(1);

                if (paragraph4 != null)
                {
                    paragraph4.RemoveAllChildren();
                    paragraph4.Remove();
                }

                Paragraph paragraph5 = body.Elements<Paragraph>().ElementAt(1);
                if (paragraph5 != null)
                {
                    paragraph5.RemoveAllChildren();
                    paragraph5.Remove();
                }

                //getting all paragraph in document
                IEnumerable<Paragraph> allParagraph = body.Descendants<Paragraph>();

                //for adding font,font size to all paragraph excluding main title , secondary title and empty paragrpah
                IEnumerable<Paragraph> formattingNeedsToApplyOfParas = allParagraph.Where(para => !(para.InnerText.Trim().Contains("Safe Harbor Matching Contribution")) && !(para.InnerText.Trim().Equals(string.Empty)));
                foreach (var para in formattingNeedsToApplyOfParas)
                {
                    para.ParagraphProperties.ParagraphMarkRunProperties.Remove();
                    IEnumerable<Run> allInternalRuns = para.Descendants<Run>();
                    foreach (var iRun in allInternalRuns)
                    {
                        iRun.RunProperties.FontSize.Remove();
                        var irunProp = new RunProperties();
                        var irunFont = new RunFonts { Ascii = "Calibri" };
                        // 16 point font size : always half of size ( 16 /2 = 8)
                        var isize = new FontSize { Val = new StringValue("16") };
                        irunProp.Append(irunFont);
                        irunProp.Append(isize);
                        iRun.PrependChild(irunProp);
                    }

                }

                //for adding font,font size to secondary  title paragraph
                var secondaryTitleParaWithIndex = allParagraph.Select((para, index) => new { para, index }).SingleOrDefault(ano => ano.para.InnerText.Trim().Equals("Safe Harbor Matching Contribution"));
                if (secondaryTitleParaWithIndex != null && secondaryTitleParaWithIndex.para != null)
                {
                    secondaryTitleParaWithIndex.para.ParagraphProperties.ParagraphMarkRunProperties.Remove();
                    IEnumerable<Run> allInternalRuns = secondaryTitleParaWithIndex.para.Descendants<Run>();
                    foreach (var iRun in allInternalRuns)
                    {
                        iRun.RunProperties.FontSize.Remove();
                        var irunProp = new RunProperties();
                        var irunFont = new RunFonts { Ascii = "Calibri Light" };
                        // 22 point font size : always half of size ( 22 /2 = 11)
                        var isize = new FontSize { Val = new StringValue("22") };
                        irunProp.Append(irunFont);
                        irunProp.Append(isize);
                        iRun.PrependChild(irunProp);
                    }

                    //remove empty paragraph right after secondary title paragraph
                    Paragraph paragraph6 = body.Elements<Paragraph>().ElementAt(secondaryTitleParaWithIndex.index + 1);
                    if (paragraph6 != null)
                    {
                        paragraph6.RemoveAllChildren();
                        paragraph6.Remove();
                    }

                }


                //getting all tables in document
                IEnumerable<Table> allTables = body.Descendants<Table>();
                foreach (var iTable in allTables)
                {
                    TableProperties iTableProperties = iTable.Elements<TableProperties>().ElementAt(0);
                    if (iTableProperties != null)
                    {
                        iTableProperties.TableIndentation.Remove();
                        TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Center };
                        iTableProperties.Append(tableJustification1);



                        iTableProperties.TableBorders.TopBorder = new TopBorder() { Val = BorderValues.Single, Color = "B4C6E7", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                        iTableProperties.TableBorders.LeftBorder = new LeftBorder() { Val = BorderValues.Single, Color = "B4C6E7", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                        iTableProperties.TableBorders.BottomBorder = new BottomBorder() { Val = BorderValues.Single, Color = "B4C6E7", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                        iTableProperties.TableBorders.RightBorder = new RightBorder() { Val = BorderValues.Single, Color = "B4C6E7", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                        iTableProperties.TableBorders.InsideHorizontalBorder = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "B4C6E7", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                        iTableProperties.TableBorders.InsideVerticalBorder = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "B4C6E7", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
                        iTableProperties.TableLook = new TableLook() { Val = "0020", FirstRow = true, LastRow = false, FirstColumn = false, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = false };
                    }

                    //for header row of table
                    TableRow iTableRow1 = iTable.Elements<TableRow>().ElementAt(0);
                    if (iTableRow1 != null)
                    {
                        
                        TableRowProperties tableRowProperties1 = new TableRowProperties();
                        TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };
                        tableRowProperties1.Append(tableJustification2);
                        iTableRow1.Append(tableRowProperties1);

                        IEnumerable<TableCell> iTableCells = iTableRow1.Elements<TableCell>();
                        foreach (var iTableCell in iTableCells)
                        {
                            TableCellBorders tableCellBorders1 = new TableCellBorders();
                            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "8EAADB", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                            tableCellBorders1.Append(bottomBorder1);

                            iTableCell.TableCellProperties.Append(tableCellBorders1);
                            iTableCell.TableCellProperties.Shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                            Paragraph iParagraph = iTableCell.Elements<Paragraph>().ElementAt(0);

                            if (iParagraph != null)
                            {
                                ParagraphMarkRunProperties iParagraphMarkRunProperties = new ParagraphMarkRunProperties();
                                if (iParagraphMarkRunProperties != null)
                                {
                                    RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
                                    Bold bold2 = new Bold();
                                    BoldComplexScript boldComplexScript2 = new BoldComplexScript();
                                    FontSize fontSize2 = new FontSize() { Val = "16" };
                                    FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "16" };

                                    iParagraphMarkRunProperties.Append(runFonts2);
                                    iParagraphMarkRunProperties.Append(bold2);
                                    iParagraphMarkRunProperties.Append(boldComplexScript2);
                                    iParagraphMarkRunProperties.Append(fontSize2);
                                    iParagraphMarkRunProperties.Append(fontSizeComplexScript2);
                                    iParagraph.ParagraphProperties.ParagraphMarkRunProperties = iParagraphMarkRunProperties;
                                }



                                Run iRun2 = iParagraph.Elements<Run>().ElementAt(0);
                                if (iRun2 != null)
                                {
                                    RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };
                                    Bold bold3 = new Bold();
                                    BoldComplexScript boldComplexScript3 = new BoldComplexScript();
                                    iRun2.RunProperties.Append(runFonts3);
                                    iRun2.RunProperties.Append(bold3);
                                    iRun2.RunProperties.Append(boldComplexScript3);
                                    iRun2.RunProperties.FontSize = new FontSize() { Val = "16" };
                                    iRun2.RunProperties.FontSizeComplexScript = new FontSizeComplexScript() { Val = "16" };
                                }
                            }

                        }//end of foreach

                    }

                    //for rest of rows of table 
                    IEnumerable<TableRow> restOfTableRows = iTable.Elements<TableRow>().Skip(1);
                    foreach (var restOfTableRow in restOfTableRows)
                    {
                        TableRowProperties tableRowProperties1 = new TableRowProperties();
                        TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Center };
                        tableRowProperties1.Append(tableJustification2);
                        restOfTableRow.Append(tableRowProperties1);

                        IEnumerable<TableCell> iTableCells = restOfTableRow.Elements<TableCell>();
                        foreach (var iTableCell in iTableCells)
                        {
                            TableCellBorders tableCellBorders1 = new TableCellBorders();
                            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "8EAADB", Size = (UInt32Value)12U, Space = (UInt32Value)0U };

                            tableCellBorders1.Append(bottomBorder1);

                            iTableCell.TableCellProperties.Append(tableCellBorders1);
                            iTableCell.TableCellProperties.Shading = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

                            Paragraph iParagraph = iTableCell.Elements<Paragraph>().ElementAt(0);

                            if (iParagraph != null)
                            {
                                ParagraphMarkRunProperties iParagraphMarkRunProperties = new ParagraphMarkRunProperties();
                                if (iParagraphMarkRunProperties != null)
                                {
                                    RunFonts runFonts2 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

                                    FontSize fontSize2 = new FontSize() { Val = "16" };
                                    FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "16" };

                                    iParagraphMarkRunProperties.Append(runFonts2);

                                    iParagraphMarkRunProperties.Append(fontSize2);

                                    iParagraphMarkRunProperties.Append(fontSizeComplexScript2);
                                    iParagraph.ParagraphProperties.ParagraphMarkRunProperties = iParagraphMarkRunProperties;


                                }

                                Run iRun2 = iParagraph.Elements<Run>().ElementAt(0);
                                if (iRun2 != null)
                                {
                                    RunFonts runFonts3 = new RunFonts() { Ascii = "Calibri", HighAnsi = "Calibri", ComplexScript = "Calibri" };

                                    iRun2.RunProperties.Append(runFonts3);

                                    iRun2.RunProperties.FontSize = new FontSize() { Val = "16" };
                                    iRun2.RunProperties.FontSizeComplexScript = new FontSizeComplexScript() { Val = "16" };
                                }
                            }

                        }//end of foreach table cell

                    }//end of foreach rest of table rows

                }//end of foreach all tables

                // Count the header and footer parts and continue if there are any.
                if (docPart.HeaderParts.Count() > 0 || docPart.FooterParts.Count() > 0)
                {
                    // Remove the header and footer parts.
                    docPart.DeleteParts(docPart.HeaderParts);
                    docPart.DeleteParts(docPart.FooterParts);

                    // Remove all references to the headers and footers.
                    var headers = docPart.Document.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                    {
                        header.Remove();
                    }

                    var footers = docPart.Document.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Remove();
                    }
                }


                IEnumerable<SectionProperties> sections = docPart.Document.Descendants<SectionProperties>();
                foreach (SectionProperties sectPr in sections)
                {

                    PageSize pgSz = sectPr.Descendants<PageSize>().FirstOrDefault();
                    if (pgSz != null)
                    {                        
                        // change the page size.                       
                        pgSz.Width = Convert.ToUInt32(7920);
                        pgSz.Height = Convert.ToUInt32(12240); 

                        PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
                        if (pgMar != null)
                        {
                            // change the page margin.
                            pgMar.Top = Convert.ToInt32(1080); 
                            pgMar.Bottom = Convert.ToInt32(360); 
                            pgMar.Left = Convert.ToUInt32(576);
                            pgMar.Right = Convert.ToUInt32(576);
                        }

                    }
                }

                docPart.Document.Save();

            }
        }


    }
}
