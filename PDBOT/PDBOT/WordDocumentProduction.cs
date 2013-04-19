﻿using Aspose.Pdf;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Reporting;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace pdbot
{
    class Program
    {
        
        static Stopwatch stopwatch = new Stopwatch();
        static string outputFile = null;
        static void Main(string[] args)
        {
            
            stopwatch.Start();
       
            writeToLog("Application PDBOT Started...");
            Console.WriteLine("Application PDBOT Started...");            

            //license details
            Aspose.Words.License asposeWordsLicense = new Aspose.Words.License();
            asposeWordsLicense.SetLicense("Aspose.Words.lic");
            Aspose.Pdf.License asposePdfLisence = new Aspose.Pdf.License();
            asposePdfLisence.SetLicense("Aspose.Pdf.lic");

            string hflDocumentPath = null;
            string path = @"C:\Users\51210\Desktop\Customers\SG Finans\Oppgraderings Prosjekt SG Finans\BL0019\pdbot_control.xml";
            XmlDocument controlXMl = null;
            XmlNamespaceManager nsmngr = null;
            

            //Lists for saving values from control xml
            List<string> documentKeyswordsList = null;
            List<string> globalsKeywordsList = null;
            Aspose.Words.Document docTemplate = null;
            //load xml document
            try
            {
                controlXMl = new XmlDocument();                
                controlXMl.Load(path);
                nsmngr = new XmlNamespaceManager(controlXMl.NameTable);
                nsmngr.AddNamespace("pdbot", "www.canon.no/pdbot");

                writeToLog("Control xml loaded successfuly");
            }
            catch (Exception e)
            {
                Console.WriteLine("Application PDBOT failed with error " + e.StackTrace);
                writeToLog("Error loading control xml " + e.StackTrace);
                Console.ReadKey();
            }
            //----------------------------------------------------------------------------------------------------------------------------------
            
            //get information from control xml nodes
            if (controlXMl != null)
            {                          
                try
                {                    
                    //get global document keys and value
                    //select Globals/Keys/Key 
                   globalsKeywordsList = new List<string>();
                   XmlNodeList keyNodes = controlXMl.SelectNodes("//pdbot:Globals//pdbot:Keys//pdbot:Key", nsmngr);
                   
                   if (keyNodes.Count != 0)
                   {
                       foreach (XmlNode KeyNode in keyNodes)
                       {
                           string keyword = KeyNode["Keyword"].InnerText;
                           string value = KeyNode["Value"].InnerText;
                           string KeywordAndValue = keyword + "|" + value;
                           globalsKeywordsList.Add(KeywordAndValue);

                       }
                       writeToLog("Global keywords and value read");
                   }                   
                    //---------------------------------------------------------------------------------------------------------------------------
                    
                    //get Global sections
                    //select Globals/Sections/section 
                    List<string> sectionsList = new List<string>();
                    XmlNodeList sectionNodes = controlXMl.SelectNodes("//pdbot:Globals//pdbot:Sections//pdbot:Section", nsmngr);
                    if (sectionNodes.Count != 0)
                    {
                        foreach (XmlNode sectionNode in sectionNodes)
                        {
                            string name = sectionNode["Name"].InnerText;
                            sectionsList.Add(name);

                        }
                        writeToLog("Global Section names read");
                    }                    
                    //---------------------------------------------------------------------------------------------------------------------------

                    //get all documents in the control xml
                    XmlNodeList documentNodes = controlXMl.SelectNodes("//pdbot:Docs//pdbot:Doc", nsmngr);
                    
                    //get all document information
                    //get document content - Fields, Keys, paragraphkeywords, sections, copies and pagewatermarkings                  
                    foreach (XmlNode documentNode in documentNodes)
                    {
                       //get document template information
                        List<string> documentFieldsList = new List<string>();
                        string templateFormat = documentNode["TemplateFormat"].InnerText;
                        string repositoryTemplate = documentNode["RepositoryTemplate"].InnerText;
                        //---------------------------------------------------------------------------------------------------------------------------
                        
                        //get document field names and value
                        //select Docs/Doc/DocContent/Fields 
                        XmlNodeList fieldNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:Fields//pdbot:Field", nsmngr);
                        if (fieldNodes.Count != 0)
                        {
                            foreach (XmlNode fieldNode in fieldNodes)
                            {
                                string fieldName = fieldNode["FieldName"].InnerText;
                                string fieldNameValue = fieldNode["Value"].InnerText;

                                string fieldNameWithValue = fieldName + "|" + fieldNameValue;
                                documentFieldsList.Add(fieldNameWithValue);

                            }
                            writeToLog("Document archive fields read");
                        }
                        //---------------------------------------------------------------------------------------------------------------------------
                        
                        //get document keywords and value
                        //select Docs/Doc/DocContent/Keys 
                        documentKeyswordsList = new List<string>();
                        XmlNodeList docKeyNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:Keys//pdbot:Key", nsmngr);
                        if (docKeyNodes.Count != 0)
                        {
                            foreach (XmlNode docKeyNode in docKeyNodes)
                            {
                                string keyword = docKeyNode["Keyword"].InnerText;
                                string value = docKeyNode["Value"].InnerText;
                                string keywordWithValue = keyword + "|" + value;

                                documentKeyswordsList.Add(keywordWithValue);

                            }
                            writeToLog("Document keys read");
                        }
                        //---------------------------------------------------------------------------------------------------------------------------
                                               

                        //get document paragraphs and value
                        //select Docs/Doc/DocContent/ParagraphKeywords/Paragraph 
                        List<string> paragraphKeywordsList = new List<string>();
                        XmlNodeList paragraphNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:ParagraphKeywords//pdbot:Paragraph", nsmngr);
                        if (paragraphNodes.Count != 0)
                        {
                            foreach (XmlNode paragraphNode in paragraphNodes)
                            {
                                string keyword = paragraphNode["Keyword"].InnerText;
                                string value = paragraphNode["Value"].InnerText;
                                string keywordWithValue = keyword + "|" + value;
                                paragraphKeywordsList.Add(keywordWithValue);

                            }
                            writeToLog("Document paragragh keys read");
                        }
                        //---------------------------------------------------------------------------------------------------------------------------
                        
                        //get document sections 
                        //select Docs/Doc/DocContent/Sections/Section                         
                        XmlNodeList docSectionNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:Sections//pdbot:Section", nsmngr);
                        if (docSectionNodes.Count != 0)
                        {                            
                            foreach (XmlNode docSectionNode in docSectionNodes)
                            {
                                string name = docSectionNode["Name"].InnerText;

                                if (!sectionsList.Contains(name))
                                {
                                    sectionsList.Add(name);
                                }

                            }
                            writeToLog("Document sections read");
                        }
                        //---------------------------------------------------------------------------------------------------------------------------
                       
                        //get PageWatermarkings 
                        //select Docs/Doc/PageWatermarkings/PageWatermarking                    
                        XmlNode hflPageWatermarkingsNodes = documentNode.SelectSingleNode("pdbot:PageWatermarkings//pdbot:PageWatermarking", nsmngr);
                        hflDocumentPath = hflPageWatermarkingsNodes["ResourceFile"].InnerText;

                        //select Docs/Doc/PageWatermarkings/PageWatermarking/Watermark  
                        XmlNodeList pageWatermarkingsNodes = documentNode.SelectNodes("pdbot:PageWatermarkings//pdbot:PageWatermarking//pdbot:Watermark", nsmngr);
                        
                        List<string> waterMarkings = new List<string>();
                        //get watermark values
                        if (pageWatermarkingsNodes.Count != 0)
                        {  
                            foreach (XmlNode watermarkNode in pageWatermarkingsNodes)
                            {
                                string wmark = watermarkNode.InnerText;

                                waterMarkings.Add(wmark);
                            }
                            writeToLog("PageWatermarkings dokument variables read");
                        }
                        //---------------------------------------------------------------------------------------------------------------------------
                       
                        //get PageInserts 
                        //select Docs/Doc/PageInserts/Inserts                    
                        XmlNode PageInsertNodes = documentNode.SelectSingleNode("pdbot:PageInserts//pdbot:Inserts", nsmngr);
                        //---------------------------------------------------------------------------------------------------------------------------

                        //get copies information 
                        //select Docs/Doc/Copies/Copy     
                        List<string> copiesList = new List<string>();
                        XmlNodeList copyNodes = documentNode.SelectNodes("pdbot:Copies//pdbot:Copy", nsmngr);
                        if (copyNodes.Count != 0)
                        {
                            foreach (XmlNode copyNode in copyNodes)
                            {
                                string name = copyNode["Name"].InnerText;
                                string stampText = copyNode["StampText"].InnerText;
                                string flatten = copyNode["Flatten"].InnerText;
                                string outputFile = copyNode["OutputFile"].InnerText;

                                string copyValues = name + "|" + stampText + "|" + flatten + "|" + outputFile;
                                copiesList.Add(copyValues);

                            }
                            writeToLog("Copies document variables read.");
                        }
                        //---------------------------------------------------------------------------------------------------------------------------

                        //replace variables from word template with keywords and values
                        //loop through all the fields in the document and replace content with values from control xml:
                        //load word document template
                        try
                        {
                            //Read RepositoryTemplate
                            string keyword = null;
                            string value = null;
                            repositoryTemplate = repositoryTemplate.Replace("xml", "docx");
                            docTemplate = new Aspose.Words.Document(repositoryTemplate);
                            writeToLog("Document template " + docTemplate.OriginalFileName.ToString() + " loaded succesfully");
                            //--------------------------------------------------------------------------------------------------------------------------
                            
                            if (docTemplate != null)
                            {
                                //Remove content controls/sections which will not be used
                                if (sectionsList.Count != 0)
                                {
                                    var ccntrls = docTemplate.GetChildNodes(NodeType.StructuredDocumentTag, true);
                                    foreach (var ccntrl in ccntrls)
                                    {

                                        var sdt = ccntrl as StructuredDocumentTag;
                                        var section = sdt.Title;                                        
                                        if (!sectionsList.Contains(section))
                                        {
                                            sdt.Remove();
                                        }
                                    }
                                    writeToLog("Sections which will not be used in the document removed from template");
                                }
                                //--------------------------------------------------------------------------------------------------------------------------

                                //loop through all the fields in the document and replace content with values from control xml:
                                //replace word template variables/keywords with document values
                                if (documentKeyswordsList.Count != 0)
                                {
                                    foreach (var key in documentKeyswordsList)
                                    {
                                        //split keywords and value
                                        string[] keywordWithValues = key.Split('|');

                                        keyword = keywordWithValues[keywordWithValues.Length - 2];
                                        value = keywordWithValues[keywordWithValues.Length - 1];

                                        //loop through all the fields in the document and replace content with values from control xml:
                                        docTemplate.Range.Replace(keyword, value, false, false);
                                    }
                                    writeToLog("Document template variables replaced with document keywords and values");
                                    //---------------------------------------------------------------------------------------------------------------------------
                                 }
                                
                                
                                //replace word template variables/keywords with global document values
                                if (globalsKeywordsList.Count != 0)
                                {
                                    foreach (var key in globalsKeywordsList)
                                    {
                                        //split keywords and value
                                        string[] keywordWithValues = key.Split('|');

                                        keyword = keywordWithValues[keywordWithValues.Length - 2];
                                        value = keywordWithValues[keywordWithValues.Length - 1];

                                        docTemplate.Range.Replace(keyword, value, false, false);
                                    }
                                    writeToLog("Document template variables replaced with global keywords and values");
                                }
                                //---------------------------------------------------------------------------------------------------------------------------


                                //replace word template paragraphs variables/keywords with document paragraphs values from control xml
                                if (paragraphKeywordsList.Count != 0)
                                {
                                    foreach (var key in paragraphKeywordsList)
                                    {
                                        //split keywords and value
                                        string[] keywordWithValues = key.Split('|');

                                        keyword = keywordWithValues[keywordWithValues.Length - 2];
                                        value = keywordWithValues[keywordWithValues.Length - 1];

                                        //add a line break in paragraphs with line breaks
                                        if (value.Contains(@"\n"))
                                        {
                                            string[] paragraphLines = value.Split('\n');
                                            foreach (var line in paragraphLines)
                                            {
                                                string paragraphText = value.Replace(@"\n", ControlChar.LineBreak);
                                                docTemplate.Range.Replace(keyword, paragraphText, false, false);
                                            }                                         
                                     }
                                        else
                                        {
                                            docTemplate.Range.Replace(keyword, value, false, false);
                                        }
                                    }
                                    writeToLog("Document template paragraph variables replaced with paragraoh keywords and values");
                                }
                                //---------------------------------------------------------------------------------------------------------------------------
                                
                                //Table building!!!!------------------------------------

                               //NB! THIS WILL ONLY WORK WITH TEMPLATE THAT HAS ONLY ONE TABLES!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                                                               
                                //select Docs/Doc/DocContent/table/Row 
                                List<string> tableCellKeyList = new List<string>();
                                XmlNodeList rowNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:Table//pdbot:Row", nsmngr);
                                if (rowNodes.Count != 0)
                                {
                                    Aspose.Words.Tables.Table table = null;
                                    Aspose.Words.Tables.Row firstRow = null;
                                    foreach (XmlNode rowNode in rowNodes)
                                    {
                                        //select the table to insert rows. 1 represents table no.
                                         table = (Aspose.Words.Tables.Table)docTemplate.GetChild(NodeType.Table, 1, true);
                                        //clone the first table row
                                         firstRow = (Aspose.Words.Tables.Row)table.LastRow.Clone(true);
                                        XmlNodeList cellKeyNodes = rowNode.SelectNodes("pdbot:CellKey", nsmngr);                                        
                                        
                                        foreach (XmlNode cellKeyNode in cellKeyNodes)
                                        {
                                            string cellKey = cellKeyNode["Key"].InnerText;
                                            string cellValue = cellKeyNode["Value"].InnerText;

                                            docTemplate.Range.Replace(cellKey, cellValue, false, false);
                                        }
                                        //Insert a new table row after the current row (template row):
                                        table.AppendChild(firstRow);
                                    }
                                    //Delete the last template table row after the table is created
                                    table.LastRow.Remove();
                                    
                                }
                                writeToLog("Tabel cell keys and keywords read from control xml and replaced in the document template");
                                //---------------------------------------------------------------------------------------------------------------------------

                                //produce document copies
                                foreach (var copy in copiesList)
                                {
                                    //split copy values
                                    string[] copyValues = copy.Split('|');

                                    string name = copyValues[copyValues.Length - 4];
                                    string stampText = copyValues[copyValues.Length - 3];
                                    string slatten = copyValues[copyValues.Length - 2];
                                    outputFile = copyValues[copyValues.Length - 1];

                                    
                                    //manage pagebreaks problem - remove pagebreak and insert sectionbreak - new page
                                    RemovePageBreaks(docTemplate);

                                    ////save document                                    
                                    saveDoc(docTemplate, outputFile);

                                    ////watermark and stamp final document
                                    WaterMarkDocument(outputFile, hflDocumentPath, stampText, waterMarkings);

                                    
                                }                                                          
                                //---------------------------------------------------------------------------------------------------------------------------
 
                            }
                            
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine("Application failed with error " + e.Message + "." + e.StackTrace);
                            writeToLog("Error loading document template " + docTemplate.OriginalFileName.ToString() + "," + e.StackTrace);
                            Console.ReadKey();
                        }
                        
                    }    
                   
                }
                catch (Exception e)
                {
                    Console.WriteLine("Application failed with error: " + e.Message + "." + e.StackTrace);
                    writeToLog("Error reading control xml " + e.StackTrace);
                    Console.ReadKey();
                }

                //save pdbot control xml in the same temp folder as the produced documents
                string outputfilename = outputFile.Remove(outputFile.Remove(outputFile.Length - 1).LastIndexOf('\\') + 1);
                outputfilename = outputfilename + "pdbot_control.xml";

                controlXMl.Save(outputfilename);
            }
            controlXMl = null;

            //stopwatch.Stop();
            //Console.WriteLine("Time used: " + stopwatch.Elapsed.Seconds + " Seconds");
            //Console.ReadKey();
            //---------------------------------------------------------------------------------------------------------------------------------------------
        }

        //METHODS
        //method for saving document
        private static void saveDoc(Aspose.Words.Document document, string outputFile)
        {
            try
            {
                document.Save(outputFile);
                writeToLog("Temporary document saved " + outputFile.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine("Application PDBOT failed with error: " + e.StackTrace);
                writeToLog("Error saving temporary document " + e.StackTrace);
                Console.ReadKey();
            }
            finally 
            {
                document = null;
            }    

        }

        //method for pagewatermarkings and stamping
        private static void WaterMarkDocument(string pdfDocumentPath, string hflDocumentPath, string stampText, List<string> watermarkingValues)
        {
            try
            {
                Aspose.Pdf.Document document = new Aspose.Pdf.Document(pdfDocumentPath);
                Aspose.Pdf.Document hfl = new Aspose.Pdf.Document(hflDocumentPath);
                //create page stamp
                PdfPageStamp pageStamp = null;

                foreach (var value in watermarkingValues)
                {
                    string[] waterMark = value.Split('=');

                    string hflDocPageNo = waterMark[waterMark.Length - 1];
                    string docPageNo = waterMark[waterMark.Length - 2];

                    //Console.WriteLine(hflDocPageNo + " " + docPageNo);
                    //Console.ReadKey();
                    int hflDocPageNumber = Convert.ToInt32(hflDocPageNo);

                    if (docPageNo == "[ALL]")
                    {
                        //create page stamp
                        pageStamp = new PdfPageStamp(hfl.Pages[hflDocPageNumber]);

                        //add stamp to all pages
                        for (int pageCount = 1; pageCount <= document.Pages.Count; pageCount++)
                        {
                            //add stamp to particular page
                            document.Pages[pageCount].AddStamp(pageStamp);
                        }
                    }
                    else
                    {
                        int docPageNumber = Convert.ToInt32(docPageNo);

                        //create page stamp
                        pageStamp = new PdfPageStamp(hfl.Pages[hflDocPageNumber]);


                        if (docPageNumber <= document.Pages.Count)
                        {
                            //add stamp to particular page
                            document.Pages[docPageNumber].AddStamp(pageStamp);
                        }

                    }
                }

                if (stampText.Length > 0)
                {
                    //Create text stamp
                    TextStamp textStamp = new TextStamp(stampText);
                    //set whether stamp is background
                    textStamp.Background = true;
                    //set origin
                    textStamp.XIndent = 420;
                    textStamp.YIndent = 825;
                    textStamp.TextState.FontSize = 8.0F;
                    pageStamp.Background = true;
                    //add stamp to particular page
                    document.Pages[1].AddStamp(textStamp);
                }

                //finally save the pdf document
                document.Save(pdfDocumentPath);
                writeToLog("Final document " + document.FileName.ToString() + " Produced");

            }
            catch (Exception e)
            {
                Console.WriteLine("Application PDBOT failed with error: " + e.Message + "." + e.StackTrace);
                writeToLog("Error saving final document " + e.StackTrace);
                Console.ReadKey();

            }
           
        }

        //manage linebreaks problematic in Aspose.words - remove page break and insert a new sectionbreaknewpage
        private static void RemovePageBreaks(Aspose.Words.Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            foreach (Paragraph par in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                foreach (Run run in par.Runs)
                {
                    //If run contains PageBreak then remove it and insert section break
                    if (run.Text.Contains("\f"))
                    {                       
                        builder.MoveTo(run);
                        builder.InsertBreak(BreakType.SectionBreakNewPage);
                        run.Remove();
                        break;              
                    
                    }

                }
            }
           
        }        

        //method for logging
        private static void writeToLog(string toLog)
        {
            System.IO.StreamWriter sw = null;
            try
            {
                sw = System.IO.File.AppendText("logFile.txt");
                string logLine = System.String.Format("{0:G}: {1}.", System.DateTime.Now, toLog);
                sw.WriteLine(logLine);
            }
            finally
            { 
                sw.Close(); 
            }
        }
    }
}
