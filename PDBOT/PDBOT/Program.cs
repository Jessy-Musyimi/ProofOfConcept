﻿using Aspose.Pdf;
using Aspose.Words;
using Aspose.Words.Markup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ns
{
    class Program
    {
        static void Main(string[] args)
        {
            writeToLog("Application PDBOT Started.");

            string hflDocumentPath = null;
            int hflDocPageNumber = 0;
            int docPageNumber = 0;
            string path = @"C:\temp\PDBOT\Templates\pdbot_control.xml";
            XmlDocument controlXMl = null;
            XmlNamespaceManager nsmngr = null;

            //arrays for saving values from control xml
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
                Console.WriteLine(e.Message);
                writeToLog("Error loading control xml " + e.StackTrace);
            }

            //----------------------------------------------------------------------------------------------------------------------------------
            //get information from control xml nodes
            if (controlXMl != null)
            {                          
                try
                {
                    //---------------------------------------------------------------------------------------------------------------------------
                    //get global document keys and value
                    //select Globals/Keys/Key 
                   globalsKeywordsList = new List<string>();
                   XmlNodeList keyNodes = controlXMl.SelectNodes("//pdbot:Globals//pdbot:Keys//pdbot:Key", nsmngr);
                 
                   foreach (XmlNode KeyNode in keyNodes)
                    {
                        string keyword = KeyNode["Keyword"].InnerText;
                        string value = KeyNode["Value"].InnerText;
                        string KeywordAndValue = keyword + "|" + value;
                        globalsKeywordsList.Add(KeywordAndValue);
                        
                    }
                    writeToLog("Global keywords and value read");
                    //---------------------------------------------------------------------------------------------------------------------------
                    
                    //get Global sections
                    //select Globals/Sections/section 
                    List<string> sectionsList = new List<string>();
                    XmlNodeList sectionNodes = controlXMl.SelectNodes("//pdbot:Globals//pdbot:Sections//pdbot:Section", nsmngr);
                    
                    foreach (XmlNode sectionNode in sectionNodes)
                    {
                        string name = sectionNode["Name"].InnerText;
                        sectionsList.Add(name);
                       
                    }
                    writeToLog("Global Section names read");
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
                       
                        foreach (XmlNode fieldNode in fieldNodes)
                        {
                            string fieldName = fieldNode["FieldName"].InnerText;
                            string fieldNameValue = fieldNode["Value"].InnerText;
                            
                            string fieldNameWithValue = fieldName + "|" + fieldNameValue;
                            documentFieldsList.Add(fieldNameWithValue);
                            
                        }
                        writeToLog("Document archive fields read");
                        //---------------------------------------------------------------------------------------------------------------------------
                        
                        //get document keywords and value
                        //select Docs/Doc/DocContent/Keys 
                        documentKeyswordsList = new List<string>();
                        XmlNodeList docKeyNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:Keys//pdbot:Key", nsmngr);
                       
                        foreach (XmlNode docKeyNode in docKeyNodes)
                        {
                            string keyword = docKeyNode["Keyword"].InnerText;
                            string value = docKeyNode["Value"].InnerText;
                            string keywordWithValue = keyword + "|" + value;

                            documentKeyswordsList.Add(keywordWithValue);
                            
                        }
                        writeToLog("Document keys read");
                        //---------------------------------------------------------------------------------------------------------------------------
                        
                        //get document paragraphs and value
                        //select Docs/Doc/DocContent/ParagraphKeywords/Paragraph 
                        List<string> paragraphKeywordsList = new List<string>();
                        XmlNodeList paragraphNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:ParagraphKeywords//pdbot:Paragraph", nsmngr);
                        
                        foreach (XmlNode paragraphNode in paragraphNodes)
                        {
                            string keyword = paragraphNode["Keyword"].InnerText;
                            string value = paragraphNode["Value"].InnerText;
                            string keywordWithValue = keyword + "|" + value;
                            paragraphKeywordsList.Add(keywordWithValue);
                            
                        }
                        writeToLog("Document paragragh keys read");
                        //---------------------------------------------------------------------------------------------------------------------------
                        
                        //get document sections 
                        //select Docs/Doc/DocContent/Sections/Section                         
                        XmlNodeList docSectionNodes = documentNode.SelectNodes("pdbot:DocContent//pdbot:Sections//pdbot:Section", nsmngr);
                       
                        foreach (XmlNode docSectionNode in docSectionNodes)
                        {
                            string name = docSectionNode["Name"].InnerText;

                            if (!sectionsList.Contains(name))
                            {
                                sectionsList.Add(name);
                            }

                        }
                        writeToLog("Document sections read");
                        //---------------------------------------------------------------------------------------------------------------------------
                       
                        //get PageWatermarkings 
                        //select Docs/Doc/PageWatermarkings/PageWatermarking                    
                        XmlNode pageWatermarkingsNodes = documentNode.SelectSingleNode("pdbot:PageWatermarkings//pdbot:PageWatermarking", nsmngr);
                        
                        hflDocumentPath = pageWatermarkingsNodes["ResourceFile"].InnerText;
                        string watermark = pageWatermarkingsNodes["Watermark"].InnerText;
                        //split watermark values
                        string[] watermarks = watermark.Split('=');
                        string hflDocPageNo = watermarks[watermarks.Length - 1];
                        string docPageNo = watermarks[watermarks.Length - 2];

                        hflDocPageNumber = Convert.ToInt32(hflDocPageNo);
                        docPageNumber = Convert.ToInt32(docPageNo);
                        
                        writeToLog("PageWatermarkings dokument variables read");
                        //---------------------------------------------------------------------------------------------------------------------------
                       
                        //get copies information 
                        //select Docs/Doc/Copies/Copy     
                        List<string> copiesList = new List<string>();
                        XmlNodeList copyNodes = documentNode.SelectNodes("pdbot:Copies//pdbot:Copy", nsmngr);
                        
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
                        //---------------------------------------------------------------------------------------------------------------------------

                        //replace variables from word template with keywords and values
                        //loop through all the fields in the document and replace content with values from control xml:
                        //load word document template
                        try
                        {
                            string keyword = null;
                            string value = null;
                            docTemplate = new Aspose.Words.Document(@"C:\temp\PDBOT\Templates\BL5099.docx");
                            writeToLog("Document template " + docTemplate.OriginalFileName.ToString() + " loaded succesfully");
                            //--------------------------------------------------------------------------------------------------------------------------
                            
                            if (docTemplate != null)
                            {                                
                                //Remove content controls/sections which will not be used
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
                                //--------------------------------------------------------------------------------------------------------------------------

                                //loop through all the fields in the document and replace content with values from control xml:
                                //replace word template variables/keywords with document values
                                foreach (var key in documentKeyswordsList)
                                {
                                    //split keywords and value
                                    string[] keywordWithValues = key.Split('|');

                                     keyword = keywordWithValues[keywordWithValues.Length -2];
                                     value = keywordWithValues[keywordWithValues.Length - 1];

                                     //loop through all the fields in the document and replace content with values from control xml:
                                     docTemplate.Range.Replace(keyword, value, true, false);
                                }
                                writeToLog("Document template variables replaced with document keywords and values");  
                                //---------------------------------------------------------------------------------------------------------------------------

                                
                                //replace word template variables/keywords with global document values
                                foreach (var key in globalsKeywordsList)
                                {
                                    //split keywords and value
                                    string[] keywordWithValues = key.Split('|');

                                    keyword = keywordWithValues[keywordWithValues.Length - 2];
                                    value = keywordWithValues[keywordWithValues.Length - 1];    
                                    
                                    docTemplate.Range.Replace(keyword, value, true, false);
                                }
                                writeToLog("Document template variables replaced with global keywords and values"); 
                                //---------------------------------------------------------------------------------------------------------------------------


                                //replace word template paragraphs variables/keywords with document paragraphs values
                                foreach (var key in paragraphKeywordsList)
                                {
                                    //split keywords and value
                                    string[] keywordWithValues = key.Split('|');

                                    keyword = keywordWithValues[keywordWithValues.Length - 2];
                                    value = keywordWithValues[keywordWithValues.Length - 1];    

                                    
                                    docTemplate.Range.Replace(keyword, value, true, false);
                                }
                                writeToLog("Document template paragraph variables replaced with paragraoh keywords and values"); 
                                //---------------------------------------------------------------------------------------------------------------------------

                                //produce document copies
                                foreach (var copy in copiesList)
                                {
                                    //split copy values
                                    string[] copyValues = copy.Split('|');

                                    string name = copyValues[copyValues.Length - 4];
                                    string StampText = copyValues[copyValues.Length - 3];
                                    string Flatten = copyValues[copyValues.Length - 2];
                                    string OutputFile = copyValues[copyValues.Length - 1];
                                    
                                    //save document                                    
                                    save(docTemplate, OutputFile);

                                    //watermark and stamp final document
                                    WaterMarkDocument(OutputFile, hflDocumentPath, StampText, docPageNumber, hflDocPageNumber);
                                }                                                             
                                //---------------------------------------------------------------------------------------------------------------------------
 
                            }
                        }
                        catch (Exception e)
                        {
                            writeToLog("Error loading document template " + docTemplate.OriginalFileName.ToString() + "," + e.StackTrace);
                        }
                        
                    }    
                   
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.StackTrace);
                    writeToLog("Error reading control xml " + e.Message);
                }
                               

                //Console.ReadKey();
            }
            //---------------------------------------------------------------------------------------------------------------------------------------------
        }

        //method for saving document
        private static void save(Aspose.Words.Document document, string outputFile)
        {
            try
            {
                document.Save(outputFile);
                writeToLog("Temporary document saved " + outputFile.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                writeToLog("Error saving temporary document " + e.Message);
            }

        }

        //method for pagewatermarking and stamping
        private static void WaterMarkDocument(string pdfDocumentPath, string hflDocumentPath, string stamptext, int docPageNumber, int hflDocPageNumber)
        {
            try
            {                
                Aspose.Pdf.Document document = new Aspose.Pdf.Document(pdfDocumentPath);

                Aspose.Pdf.Document hfl = new Aspose.Pdf.Document(hflDocumentPath);
                //create page stamp
                PdfPageStamp pageStamp = new PdfPageStamp(hfl.Pages[1]);
                //add stamp to particular page
                document.Pages[docPageNumber].AddStamp(pageStamp);

                //Create text stamp
                TextStamp textStamp = new TextStamp(stamptext);
                //set whether stamp is background
                textStamp.Background = true;
                //set origin
                textStamp.XIndent = 420;
                textStamp.YIndent = 820;
                pageStamp.Background = true;
                //add stamp to particular page
                document.Pages[1].AddStamp(textStamp);

                document.Save(pdfDocumentPath);
                writeToLog("Final document " + document.FileName.ToString() + " Produced");
            }
            catch (Exception e)
            {
                writeToLog("Error saving final document "  + e.StackTrace);
               
            }
        }


        //method for logging
        private static void writeToLog(string toLog)
        {
            System.IO.StreamWriter sw = null;
            try
            {
                sw = System.IO.File.AppendText(@"C:\temp\PDBOT\Output\logFile.txt");
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
