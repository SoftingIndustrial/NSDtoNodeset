/* 
 * 
 * Copyright (c) 2017, Softing Industrial Automation GmbH. All rights reserved.
 * 
 * XMItoNodeset is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace NSDtoNodeset
{
    class NSDtoNodeset
    {
        XmlDocument _nodesetDoc;
        XmlNode _nodesetUANodeSetNode;
        XmlNode _nodesetAliasesNode;
        XmlNode _nodesetNamespaceUrisNode;
        XmlDocument _binaryTypesDoc;
        XmlNode _binaryTypesRootNode;
        XmlDocument _xmlTypesDoc;
        XmlNode _xmlTypesRootNode;

        String[,] _aliases;
        Dictionary<String, String> _nodesetNodeIdMap;
        Dictionary<String, String> _nsdocEnMap;
        Dictionary<String, XmlNode> _nsdObjectMap;

        List<String> _nsdFiles;
        List<String> _nsdFilesBase;
        String _nodesetFile;
        String _nodesetImportFile;
        String _nodesetURL;
        String _nodesetURLBase;
        String _nodesetTypeDictionaryName;
        String _nodeIdMapFileName;
        String _nodeIdMapFileNameBase;
        String _binaryTypesFileName;
        String _xmlTypesFileName;
        String _wordFileName;
        Int32 _nextNodeId;
        Int32 _nsIdx;
        String _nodesetModelBaseVersion;
        String _nodesetModelVersion;

        Application _wordApp = null;
        Document _wordDoc = null;
        Table _wordCurrentTable = null;
        bool _nodesetHasDT;

        string _nodeIdTextBinarySchema = "BinarySchema";
        string _nodeIdTextXmlSchema = "XmlSchema";

        static void Main(string[] args)
        {
            NSDtoNodeset prog = new NSDtoNodeset();
            prog.start(args);
        }

        bool parseCommandLineArgs(string[] args)
        {
            string command = "";

            _nsdFiles = new List<String>();
            _nsdFilesBase = new List<String>();

            // default initialization
            _nodesetTypeDictionaryName = "NSDtoNodeset";
            _nodesetURL = "http://industrial.softing.com/NSDtoNodeset";
            _nodesetURLBase = "";
            _nodesetFile = "nodeset.xml";
            _nodesetImportFile = "";
            _nextNodeId = 0;
            _nsIdx = 1;
            _nodeIdMapFileName = "NodeIdMap.txt";
            _nodeIdMapFileNameBase = "NodeIdMapBase.txt";
            _binaryTypesFileName = "BinaryTypes.xml";
            _xmlTypesFileName = "XmlTypes.xml";
            _wordFileName = "";

            foreach (string arg in args)
            {
                if (command == "")
                { // no command set -> has to be specified
                    if ((arg == "/nsd") || (arg == "/nsdBase") || (arg == "/nodeset") || (arg == "/nodesetUrl")  || (arg == "/nodesetUrlBase") || (arg == "/nodesetTypeDictionary") || (arg == "/nodesetImport") || (arg == "/nodesetStartId") || (arg == "/nodeIdMap")  || (arg == "/nodeIdMapBase") || (arg == "/binaryTypes") || (arg == "/xmlTypes") || (arg == "/word") || (arg == "/nodesetModelVersion") || (arg == "/nodesetModelBaseVersion"))
                    {
                        command = arg;
                    }
                    else
                    {
                        Console.WriteLine("Invalid command: {0}", arg);
                        return false;
                    }
                }
                else
                { // command argument
                    if (command == "/nsd")
                    {
                        _nsdFiles.Add(arg);
                    }
                    if (command == "/nsdBase")
                    {
                        _nsdFilesBase.Add(arg);
                    }
                    else if (command == "/nodeset")
                    {
                        _nodesetFile = arg;
                    }
                    else if (command == "/nodesetUrl")
                    {
                        _nodesetURL = arg;
                    }
                    else if (command == "/nodesetUrlBase")
                    {
                        _nodesetURLBase = arg;
                        _nsIdx = 2;
                        _nodeIdTextBinarySchema = "BinarySchema2";
                        _nodeIdTextXmlSchema = "XmlSchema2";
                    }
                    else if (command == "/nodesetStartId")
                    {
                        _nextNodeId = Int32.Parse(arg);
                    }
                    else if (command == "/nodesetImport")
                    {
                        _nodesetImportFile = arg;
                    }
                    else if (command == "/nodesetTypeDictionary")
                    {
                        _nodesetTypeDictionaryName = arg;
                    }
                    else if (command == "/nodeIdMap")
                    {
                        _nodeIdMapFileName = arg;
                    }
                    else if (command == "/nodeIdMapBase")
                    {
                        _nodeIdMapFileNameBase = arg;
                    }
                    else if (command == "/binaryTypes")
                    {
                        _binaryTypesFileName = arg;
                    }
                    else if (command == "/xmlTypes")
                    {
                        _xmlTypesFileName = arg;
                    }
                    else if (command == "/word")
                    {
                        _wordFileName = arg;
                    }
                    else if (command == "/nodesetModelVersion")
                    {
                        _nodesetModelVersion = arg;
                    }
                    else if (command == "/nodesetModelBaseVersion")
                    {
                        _nodesetModelBaseVersion = arg;
                    }

                    command = "";
                }
            }
            return true;
        }

        void initOutputXmlDocuments()
        {
            // nodeset document 
            _nodesetDoc = new XmlDocument();
            XmlNode nodesetDocNode = _nodesetDoc.CreateXmlDeclaration("1.0", "UTF-8", null);
            _nodesetDoc.AppendChild(nodesetDocNode);
            _nodesetUANodeSetNode = addXmlElement(_nodesetDoc, _nodesetDoc, "UANodeSet");
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "xmlns:xsd", "http://www.w3.org/2001/XMLSchema");
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "LastModified", String.Format("{0:yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fffffff'Z'}", DateTime.Now));
            addXmlAttribute(_nodesetDoc, _nodesetUANodeSetNode, "xmlns", "http://opcfoundation.org/UA/2011/03/UANodeSet.xsd");
            _nodesetNamespaceUrisNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "NamespaceUris");
            if (_nodesetURLBase != "")
            {
                addXmlElement(_nodesetDoc, _nodesetNamespaceUrisNode, "Uri", _nodesetURLBase);
            }
            addXmlElement(_nodesetDoc, _nodesetNamespaceUrisNode, "Uri", _nodesetURL);

            XmlNode nodesetModelsNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "Models");
            XmlNode nodesetModelNode = addXmlElement(_nodesetDoc, nodesetModelsNode, "Model");
            addXmlAttribute(_nodesetDoc, nodesetModelNode, "ModelUri", _nodesetURL);
            if (_nodesetModelVersion != "")
            {
                addXmlAttribute(_nodesetDoc, nodesetModelNode, "Version", _nodesetModelVersion);
            }
            if (_nodesetURLBase != "")
            {
                XmlNode nodesetReqModelsNode = addXmlElement(_nodesetDoc, nodesetModelNode, "RequiredModels");
                XmlNode nodesetReqModelNode = addXmlElement(_nodesetDoc, nodesetReqModelsNode, "Model");
                addXmlAttribute(_nodesetDoc, nodesetReqModelNode, "ModelUri", _nodesetURLBase);
                if (_nodesetModelBaseVersion != "")
                {
                    addXmlAttribute(_nodesetDoc, nodesetReqModelNode, "Version", _nodesetModelBaseVersion);
                } 
            }

            addAliases();

            // binary types
            _binaryTypesDoc = new XmlDocument();
            _binaryTypesRootNode = addQualifiedXmlElement(_binaryTypesDoc, _binaryTypesDoc, "opc", "http://opcfoundation.org/BinarySchema/", "TypeDictionary");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:opc", "http://opcfoundation.org/BinarySchema/");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:ua", "http://opcfoundation.org/UA/");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:tns", _nodesetURL);
            if (_nodesetURLBase != "")
            {
                addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "xmlns:tnsbase", _nodesetURLBase);
            }
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "DefaultByteOrder", "LittleEndian");
            addXmlAttribute(_binaryTypesDoc, _binaryTypesRootNode, "TargetNamespace", _nodesetURL);

            addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, _binaryTypesRootNode, "opc", "http://opcfoundation.org/BinarySchema/", "Import", "Namespace", "http://opcfoundation.org/UA/", "Location", "Opc.Ua.BinarySchema.bsd");

            // xml types
            _xmlTypesDoc = new XmlDocument();
            _xmlTypesRootNode = addQualifiedXmlElement(_xmlTypesDoc, _xmlTypesDoc, "xs", "http://www.w3.org/2001/XMLSchema", "schema");
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:tns", _nodesetURL);
            if (_nodesetURLBase != "")
            {
                addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:tnsbase", _nodesetURLBase);
            }
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:xs", "http://www.w3.org/2001/XMLSchema");
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xmlns:ua", "http://opcfoundation.org/UA/2008/02/Types.xsd");
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "targetNamespace", _nodesetURL);
            addXmlAttribute(_xmlTypesDoc, _xmlTypesRootNode, "elementFormDefault", "qualified");
            addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "Import", "namespace", "http://opcfoundation.org/UA/2008/02/Types.xsd");
        }

        void importNodeset()
        {
            if (_nodesetImportFile == "")
            {
                return;
            }
             
            Console.WriteLine("Load Nodeset insert file: {0}", _nodesetImportFile);
            XmlDocument nodesetImport = new XmlDocument();
            try
            {
                nodesetImport.Load(_nodesetImportFile);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error loading file - {0}", e.Message);
                return;
            }
            foreach (XmlNode nodesetImportNode in nodesetImport.DocumentElement.ChildNodes)
            {
                XmlNode nodesetImportedNode = _nodesetDoc.ImportNode(nodesetImportNode, true);
                if (nodesetImportedNode.Name == "NamespaceUris")
                {
                    foreach (XmlNode nodesetImportedSubNode in nodesetImportedNode.ChildNodes)
                    {
                        XmlNode clone = nodesetImportedSubNode.Clone();
                        _nodesetNamespaceUrisNode.AppendChild(clone);
                    }
                }
                else if (nodesetImportedNode.Name == "Aliases")
                {
                    foreach (XmlNode nodesetImportedSubNode in nodesetImportedNode.ChildNodes)
                    {
                        bool nodesetSameAttr = false;
                        foreach (XmlNode nodesetAliasNode in _nodesetAliasesNode.ChildNodes)
                        {
                            if (nodesetImportedSubNode.Attributes["Alias"].Value == nodesetAliasNode.Attributes["Alias"].Value)
                            {
                                nodesetAliasNode.InnerText = nodesetImportedSubNode.InnerText;
                                nodesetSameAttr = true;
                                break;
                            }
                        }
                        if (!nodesetSameAttr)
                        {
                            XmlNode clone = nodesetImportedSubNode.Clone();
                            _nodesetAliasesNode.AppendChild(clone);
                        }
                    }
                }
                else
                {
                    _nodesetUANodeSetNode.AppendChild(nodesetImportedNode);
                }
            }
        }

        void addAliases()
        {
            _aliases = new string[,] {
                { "Boolean", "i=1", "opc:Boolean", "xs:boolean" },
                { "SByte", "i=2", "opc:SByte", "xs:byte" },
                { "Byte", "i=3", "opc:SByte", "xs:unsignedByte" },
                { "Int16", "i=4", "opc:Int16", "xs:short" },
                { "UInt16", "i=5", "opc:UInt16", "xs:unsignedShort" },
                { "Int32", "i=6", "opc:Int32", "xs:int" },
                { "UInt32", "i=7", "opc:UInt32", "xs:unsignedInt" },
                { "Int64", "i=8", "opc:Int64", "xs:long" },
                { "UInt64", "i=9", "opc:UInt64", "xs:unsignedLong" },
                { "Float", "i=10", "opc:Float", "xs:float" },
                { "Double", "i=11", "opc:Double", "xs:double" },
                { "String", "i=12", "opc:String", "xs:string" },
                { "ByteString", "i=15", "opc:ByteString", "xs:base64Binary" },          
                { "Structure", "i=22", "", "" },     
                { "BaseDataType", "i=24", "", "" },   
                { "Enumeration", "i=29", "", "" },          
                { "Organizes", "i=35", "", "" },                        
                { "HasModellingRule", "i=37", "", "" },
                { "HasEncoding", "i=38", "", "" },
                { "HasDescription", "i=39", "", "" },
                { "HasTypeDefinition", "i=40", "", "" },
                { "HasSubtype", "i=45", "", "" },
                { "HasProperty", "i=46", "", "" },
                { "HasComponent", "i=47", "", "" },
                { "PropertyType", "i=68", "", "" },               
                { "Mandatory", "i=78", "", "" },
                { "Optional", "i=80", "", "" },
                { "OptionalPlaceholder", "i=11508", "", "" },
                { "MandatoryPlaceholder", "i=11510", "", "" },
                { "DefaultVariableRefType", "i=47", "", "" },
                { "DefaultObjectRefType", "i=47", "", "" },
            };

            _nodesetAliasesNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "Aliases");
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                XmlNode nodesetAliasNode = addXmlElement(_nodesetDoc, _nodesetAliasesNode, "Alias", _aliases[i,1]);
                addXmlAttribute(_nodesetDoc, nodesetAliasNode, "Alias", _aliases[i, 0]);
            }
        }

        void loadEnNsdoc(string nsdFile)
        {
            XmlDocument nsdocEnDoc = new XmlDocument();

            Console.WriteLine("Load English NSDOC file: {0}-en.nsdoc", nsdFile);
            try
            {
                nsdocEnDoc.Load(nsdFile + "-en.nsdoc");
            }
            catch (Exception e)
            {
                Console.WriteLine("Error loading file - {0}", e.Message);
                return;
            }

            foreach (XmlNode nsdDoc in nsdocEnDoc.ChildNodes[1].ChildNodes)
            {
                if (nsdDoc.Name == "Doc")
                { 
                    string docVal = nsdDoc.InnerText;
                    docVal = System.Text.RegularExpressions.Regex.Replace(docVal, @"<(.|\n)*?>", "");
                    docVal = System.Text.RegularExpressions.Regex.Replace(docVal, @"&lt;","<", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    docVal = System.Text.RegularExpressions.Regex.Replace(docVal, @"&gt;",">", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    _nsdocEnMap.Add(nsdDoc.Attributes["id"].Value, docVal);
                }
            }
        }

        string getEnNsdoc(string id)
        {
            string ennsdoc = "";
            try
            {
                ennsdoc = _nsdocEnMap[id];
            }
            catch
            { }
            return ennsdoc;
        }

        void start(string[] args)
        {
            if (!parseCommandLineArgs(args))
            {
                Console.WriteLine("NSDoNodeset invalid command line");
                return;
            }

            try
            {
                if (_wordFileName != "")
                { 
                    _wordApp = new Application();
                    _wordDoc = _wordApp.Documents.Add();
                    _wordDoc.Paragraphs.SpaceAfter = 0;
                }
            }
            catch
            { }

            _nodesetNodeIdMap = new Dictionary<String, String>();
            _nsdocEnMap = new Dictionary<String, String>();
            _nsdObjectMap = new Dictionary<String, XmlNode>();
            _nodesetHasDT = false;

            // load _nodesetNodeIdMapBase
            try
            { 
                System.IO.StreamReader nodeIdMapFileBaseR = new System.IO.StreamReader(_nodeIdMapFileNameBase);
                string nodeIdMapFileBaseRLine;
                nodeIdMapFileBaseRLine = nodeIdMapFileBaseR.ReadLine();
                while ((nodeIdMapFileBaseRLine = nodeIdMapFileBaseR.ReadLine()) != null)
                {
                    string[] split = nodeIdMapFileBaseRLine.Split('\t');
                    _nodesetNodeIdMap[split[0]] = split[1];
                }
                nodeIdMapFileBaseR.Close();
            }
            catch
            { }
            
            // load _nodesetNodeIdMap
            try
            { 
                System.IO.StreamReader nodeIdMapFileR = new System.IO.StreamReader(_nodeIdMapFileName);
                string nodeIdMapFileRLine;
                if ((nodeIdMapFileRLine = nodeIdMapFileR.ReadLine()) != null)
                {
                     _nextNodeId = Int32.Parse(nodeIdMapFileRLine);
                }
                while ((nodeIdMapFileRLine = nodeIdMapFileR.ReadLine()) != null)
                {
                    string[] split = nodeIdMapFileRLine.Split('\t');
                    _nodesetNodeIdMap[split[0]] = split[1];
                }
                nodeIdMapFileR.Close();
            }
            catch
            { }

            initOutputXmlDocuments();
            importNodeset();

            // load nsd base documents
            foreach (string nsdFileBase in _nsdFilesBase)
            {
                XmlDocument nsdDocBase = new XmlDocument();

                Console.WriteLine("Load NSD base file: {0}.nsd", nsdFileBase);
                try
                {
                    nsdDocBase.Load(nsdFileBase + ".nsd");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error loading file - {0}", e.Message);
                    return;
                }

                loadEnNsdoc(nsdFileBase);

                foreach (XmlNode nsdBaseTopLevelNode in nsdDocBase.ChildNodes[1].ChildNodes)
                {
                    if ((nsdBaseTopLevelNode.Name == "ObjectType") || (nsdBaseTopLevelNode.Name == "AbstractObjectType") || (nsdBaseTopLevelNode.Name == "CDC") || (nsdBaseTopLevelNode.Name == "AbstractLNClass") || (nsdBaseTopLevelNode.Name == "LNClass"))
                    {
                        try
                        {
                            _nsdObjectMap.Add(nsdBaseTopLevelNode.Attributes["name"].Value, nsdBaseTopLevelNode);
                        }
                        catch
                        { }
                    }
                }
            }

            // load nsd documents
            foreach (string nsdFile in _nsdFiles)
            {
                XmlDocument nsdDoc = new XmlDocument();

                Console.WriteLine("Load NSD file: {0}.nsd", nsdFile);
                try
                {
                    nsdDoc.Load(nsdFile + ".nsd");
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error loading file - {0}", e.Message);
                    return;
                }

                loadEnNsdoc(nsdFile);

                foreach (XmlNode nsdTopLevelNode in nsdDoc.ChildNodes[1].ChildNodes)
                {
                    if (nsdTopLevelNode.Name == "Enumeration")
                    {
                        createNodesetEnumerationDataType(nsdTopLevelNode);
                    }
                    else if (nsdTopLevelNode.Name == "ConstructedAttribute")
                    {
                        createNodesetStructureDataType(nsdTopLevelNode);
                    }   
                    else if ((nsdTopLevelNode.Name == "ObjectType") || (nsdTopLevelNode.Name == "AbstractObjectType") || (nsdTopLevelNode.Name == "CDC") || (nsdTopLevelNode.Name == "AbstractLNClass") || (nsdTopLevelNode.Name == "LNClass"))
                    {
                        try
                        {
                            _nsdObjectMap.Add(nsdTopLevelNode.Attributes["name"].Value, nsdTopLevelNode);
                            createNodesetObjectType(nsdTopLevelNode);
                        }
                        catch
                        { }
                    }
                }
            }

            if (_nodesetHasDT)
            {
                createNodesetTypeDictionary();
            }
            // store _nodesetDoc
            Console.WriteLine("Save Nodeset file: {0}", _nodesetFile);
            _nodesetDoc.Save(_nodesetFile);

            // store _nodesetNodeIdMap
            System.IO.StreamWriter nodeIdMapFile = new System.IO.StreamWriter(_nodeIdMapFileName);
            string nodeIdStart = String.Format("ns={0}", _nsIdx);
            nodeIdMapFile.WriteLine("{0}", _nextNodeId);
            foreach (var pair in _nodesetNodeIdMap)
            {
                if (pair.Value.Substring(0, 4) == nodeIdStart)
                {
                    nodeIdMapFile.WriteLine("{0}\t{1}", pair.Key, pair.Value);
                }
            }
            nodeIdMapFile.Close();      
            
            // store word document
            if ((_wordApp != null) && (_wordDoc != null))
            {
                Console.WriteLine("Save Word file: {0}", _wordFileName);
                _wordApp.ActiveDocument.SaveAs(_wordFileName, WdSaveFormat.wdFormatDocumentDefault);
                _wordDoc.Close();

                _wordApp.Quit();
                 System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_wordApp);
            }
        }

        void createNodesetObjectType(XmlNode nsdObject)
        {
            string elementName = nsdObject.Name;
            string name = nsdObject.Attributes["name"].Value;
            string nodeId = getNodeId(String.Format("Object:{0}",  name));
            bool isAbstract = elementName.StartsWith("Abstract");

            XmlNode nodesetObjectTypeNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAObjectType", "NodeId", nodeId, "BrowseName", String.Format("{0}:{1}", _nsIdx, name));
            if (isAbstract)
            {
                addXmlAttribute(_nodesetDoc, nodesetObjectTypeNode, "IsAbstract", "true");
            }
            XmlNode nodesetDisplayNameNode = addXmlElement(_nodesetDoc, nodesetObjectTypeNode, "DisplayName", name);
            XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetObjectTypeNode, "References");

            string baseClassName = "";
            string baseClassNodeId = "";
            if (nsdObject.Attributes["base"] != null)
            {
                baseClassName = nsdObject.Attributes["base"].Value;
            }
            else
            {
                if (elementName == "CDC")
                {
                    baseClassName = "IEC61850DOBaseObjectType";
                }
                else if ((elementName == "AbstractLNClass") || (elementName == "LNClass"))
                {
                    baseClassName = "IEC61850LNodeBaseObjectType";
                }
                else 
                {
                    baseClassNodeId = "i=58";
                    baseClassName = "BaseObjectType";
                }
            }

            if (baseClassNodeId == "")
            {
                baseClassNodeId = getNodeId(String.Format("Object:{0}",  baseClassName));
            }
            XmlNode nodesetBackwardHasSubtypeNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", baseClassNodeId);
            addXmlAttribute(_nodesetDoc, nodesetBackwardHasSubtypeNode, "ReferenceType", "HasSubtype");
            addXmlAttribute(_nodesetDoc, nodesetBackwardHasSubtypeNode, "IsForward", "false");

            // documentation
            Paragraph parTable = null;
            if (_wordDoc != null)
            {
                Paragraph p2 = _wordDoc.Paragraphs.Add();
                p2.Range.Font.Name = "Arial";
                p2.Range.Font.Size = 10F;
                p2.Range.Text = name;
                p2.Range.InsertParagraphAfter();

                parTable = _wordDoc.Paragraphs.Add();
                _wordCurrentTable = _wordDoc.Tables.Add(parTable.Range, 5, 7);

                _wordCurrentTable.Range.Font.Name = "Arial";
                _wordCurrentTable.Range.Font.Size = 8F;
                _wordCurrentTable.Range.Font.Bold = 0;

                _wordCurrentTable.Columns[1].Width = 70;
                _wordCurrentTable.Columns[2].Width = 55;
                _wordCurrentTable.Columns[3].Width = 65;
                _wordCurrentTable.Columns[4].Width = 90;
                _wordCurrentTable.Columns[5].Width = 80;
                _wordCurrentTable.Columns[6].Width = 65;
                _wordCurrentTable.Columns[7].Width = 40;

                _wordCurrentTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                _wordCurrentTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            }

            // instance declaration objects
            addNodesetObjectTypeInstanceDeclarationObjects(nodesetReferencesNode, nsdObject, nodeId, name, false, null);

            if (parTable != null)
            {
                _wordCurrentTable.Rows[1].Range.Font.Bold = 1;
                _wordCurrentTable.Cell(1,1).Range.Text = "Attribute";
                _wordCurrentTable.Rows[1].Cells[2].Merge(_wordCurrentTable.Rows[1].Cells[7]);
                _wordCurrentTable.Cell(1,2).Range.Text = "Value";

                _wordCurrentTable.Cell(2,1).Range.Text = "BrowseName";
                _wordCurrentTable.Rows[2].Cells[2].Merge(_wordCurrentTable.Rows[2].Cells[7]);
                _wordCurrentTable.Cell(2,2).Range.Text = name;

                _wordCurrentTable.Cell(3,1).Range.Text = "IsAbstract";
                _wordCurrentTable.Rows[3].Cells[2].Merge(_wordCurrentTable.Rows[3].Cells[7]);
                _wordCurrentTable.Cell(3,2).Range.Text = isAbstract.ToString().ToLower();

                _wordCurrentTable.Rows[4].Range.Font.Bold = 1;
                _wordCurrentTable.Cell(4,1).Range.Text = "Reference";
                _wordCurrentTable.Cell(4,2).Range.Text = "NodeClass";
                _wordCurrentTable.Cell(4,3).Range.Text = "BrowseName";
                _wordCurrentTable.Cell(4,4).Range.Text = "DataType";
                _wordCurrentTable.Cell(4,5).Range.Text = "TypeDefinition";
                _wordCurrentTable.Cell(4,6).Range.Text = "ModellingRule";		 			
                _wordCurrentTable.Cell(4,7).Range.Text = "Access";		

                _wordCurrentTable.Rows[5].Cells[1].Merge(_wordCurrentTable.Rows[5].Cells[7]);
                _wordCurrentTable.Cell(5,1).Range.Text = "Subtype of " + baseClassName;

                _wordCurrentTable = null;
            }
        }

        void addNodesetObjectTypeInstanceDeclarationObjects(XmlNode nodesetNode, XmlNode nsdObject, string parentNodeId, string parentPath, bool isAggregation, List<String> membOfSub)
        {
            string baseClassTypeName = "";
            if (nsdObject.Attributes["base"] != null)
            {
                baseClassTypeName = nsdObject.Attributes["base"].Value;
            }

            if (isAggregation)
            { // add mandatory members base classes
                if (baseClassTypeName != "")
                {
                    List<String> membOfSubBase = new List<String>(membOfSub);
                    foreach (XmlNode nsdMember in nsdObject)
                    {
                        membOfSubBase.Add(nsdMember.Attributes["name"].Value);
                    }
                    addNodesetObjectTypeInstanceDeclarationObjects(nodesetNode, getNsdObject(baseClassTypeName), parentNodeId, parentPath, true, membOfSubBase);
                }
            }

            foreach (XmlNode nsdMember in nsdObject)
            {
                string name = nsdMember.Attributes["name"].Value;

                if (isAggregation)
                {
                    if (null != membOfSub.Find(s => s == name))
                        continue;
                }

                string modelingRule = "";
                bool isArray = false;

                if (nsdMember.Attributes["isArray"] != null)
                {
                    isArray = Boolean.Parse(nsdMember.Attributes["isArray"].Value);
                }
                if (nsdMember.Attributes["presCond"] != null)
                {
                    string presCond = nsdMember.Attributes["presCond"].Value;
                    if (presCond == "M")
                    {
                        modelingRule = "Mandatory";
                    }
                    else
                    {
                        modelingRule = "Optional";
                    }
                }

                if ((isAggregation) && (modelingRule == "Optional"))
                {
                    continue;
                }

                if (nsdMember.Name == "DataAttribute") 
                { // variable
                    if (name == "t")
                    {
                        continue; // ignore t
                    }

                    string nodeId = getNodeId(String.Format("{0}|Variable:{1}",  parentPath, name));
                    string dtForDoc = "";
                    string typeName = "";
                    string typeKind = "";
                    string dataTypeName;
                    string access = "R";
                    string refTypeToUse = "DefaultVariableRefType";
                    string fc = "None";
                        
                    if (nsdMember.Attributes["referenceType"] != null)
                    {
                        refTypeToUse = nsdMember.Attributes["referenceType"].Value;
                    }
                    if (nsdMember.Attributes["fc"] != null)
                    {
                        fc = nsdMember.Attributes["fc"].Value;
                    }
                    if (nsdMember.Attributes["type"] != null)
                    {
                        typeName = nsdMember.Attributes["type"].Value;
                    }
                    if (nsdMember.Attributes["typeKind"] != null)
                    {
                        typeKind = nsdMember.Attributes["typeKind"].Value;
                    }
                    dataTypeName = getDatatypeName(typeName, typeKind, true, ref dtForDoc);

                    XmlNode nodesetComponentReferenceNode = addXmlElement(_nodesetDoc, nodesetNode, "Reference", nodeId);
                    addXmlAttribute(_nodesetDoc, nodesetComponentReferenceNode, "ReferenceType", refTypeToUse);

                    XmlNode nodesetVariableNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable");
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "NodeId", nodeId);
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "BrowseName", String.Format("{0}:{1}", _nsIdx, name));
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "ParentNodeId", parentNodeId);
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "DataType", dataTypeName);

                    if (isArray)
                    {
                        addXmlAttribute(_nodesetDoc, nodesetVariableNode, "ValueRank", "1");
                        dtForDoc += "[]";
                    }

                    if ((fc == "SP") || (fc == "CF") || (fc == "SV") || (fc == "DC") || (fc == "None"))
                    {
                        addXmlAttribute(_nodesetDoc, nodesetVariableNode, "AccessLevel", "3");
                        addXmlAttribute(_nodesetDoc, nodesetVariableNode, "UserAccessLevel", "3");
                        access = "RW";
                    }

                    addXmlElement(_nodesetDoc, nodesetVariableNode, "DisplayName", name);

                    XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetVariableNode, "References");
                    XmlNode nodesetRefParentNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", parentNodeId);
                    addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "ReferenceType", refTypeToUse);
                    addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "IsForward", "false");

                    string variableType = "i=63"; // BaseVariableType
                    if (refTypeToUse == "HasProperty")
                    { 
                        variableType = "i=68";     // PropertyType
                    }
                    addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", variableType, "ReferenceType", "HasTypeDefinition");
                    XmlNode nodesetRefModelingRuleNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", modelingRule);
                    addXmlAttribute(_nodesetDoc, nodesetRefModelingRuleNode, "ReferenceType", "HasModellingRule");

                    // documentation
                    if ((_wordCurrentTable != null) && (!isAggregation))
                    {
                        Row row = _wordCurrentTable.Rows.Add();
                        row.Cells[1].Range.Text = refTypeToUse;
                        row.Cells[2].Range.Text = "Variable";
                        row.Cells[3].Range.Text = name;
                        row.Cells[4].Range.Text = dtForDoc;
                        if (variableType == "i=63")
                        { 
                            row.Cells[5].Range.Text = "BaseVariableType";
                        }
                        else if (variableType == "i=68")
                        { 
                            row.Cells[5].Range.Text = "PropertyType";
                        }
                        else
                        {
                            row.Cells[5].Range.Text = variableType;
                        }
                        row.Cells[6].Range.Text = modelingRule;
                        row.Cells[7].Range.Text = access;
                    }
                }
                else if ((nsdMember.Name == "SubDataObject") || (nsdMember.Name == "DataObject"))
                { // object
                    string nodeId = getNodeId(String.Format("{0}|Object:{1}", parentPath, name));
                    string typeName = nsdMember.Attributes["type"].Value;
                    string refTypeToUse = "DefaultObjectRefType";

                    if (nsdMember.Attributes["referenceType"] != null)
                    {
                        refTypeToUse = nsdMember.Attributes["referenceType"].Value;
                    }
                    if (nsdMember.Attributes["underlyingType"] != null)
                    {
                        string underlyingType = nsdMember.Attributes["underlyingType"].Value;
                        string underlyingTypeKind = "ENUMERATED";
                        if (nsdMember.Attributes["underlyingTypeKind"] != null)
                        {
                            underlyingTypeKind = nsdMember.Attributes["underlyingTypeKind"].Value;
                        }
                        addNodesetObjectTypeCDC(typeName, underlyingType, underlyingTypeKind);
                        typeName = String.Format("{0}{1}", typeName, underlyingType);
                    }
                    if (isArray)
                    {
                        modelingRule += "Placeholder";
                        name = "<" + name + ">";
                    }

                    XmlNode nodesetComponentReferenceNode = addXmlElement(_nodesetDoc, nodesetNode, "Reference", nodeId);
                    addXmlAttribute(_nodesetDoc, nodesetComponentReferenceNode, "ReferenceType", refTypeToUse);
                                            
                    XmlNode nodesetVariableNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAObject");
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "NodeId", nodeId);
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "BrowseName", String.Format("{0}:{1}", _nsIdx, name));
                    addXmlAttribute(_nodesetDoc, nodesetVariableNode, "ParentNodeId", parentNodeId);

                    addXmlElement(_nodesetDoc, nodesetVariableNode, "DisplayName", name);

                    XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetVariableNode, "References");
                    XmlNode nodesetRefParentNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", parentNodeId);
                    addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "ReferenceType", refTypeToUse);
                    addXmlAttribute(_nodesetDoc, nodesetRefParentNode, "IsForward", "false");

                    string typeNodeId =  getNodeId(String.Format("Object:{0}", typeName));
                    XmlNode nodesetRefTypeDefNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", typeNodeId);
                    addXmlAttribute(_nodesetDoc, nodesetRefTypeDefNode, "ReferenceType", "HasTypeDefinition");
                    XmlNode nodesetRefModelingRuleNode = addXmlElement(_nodesetDoc, nodesetReferencesNode, "Reference", modelingRule);
                    addXmlAttribute(_nodesetDoc, nodesetRefModelingRuleNode, "ReferenceType", "HasModellingRule");

                    XmlNode typeNsdObject = getNsdObject(typeName);
                    if ((parentNodeId != typeNodeId) && (typeNsdObject != null)) // prevent recursion and working on unknown types
                    {
                        List<String> membersOfSub = new List<String>();
                        addNodesetObjectTypeInstanceDeclarationObjects(nodesetReferencesNode, typeNsdObject, nodeId, String.Format("{0}|{1}", parentPath, name), true, membersOfSub);
                    }

                    // documentation
                    if ((_wordCurrentTable != null) && (!isAggregation))
                    {
                        Row row = _wordCurrentTable.Rows.Add();
                        row.Cells[1].Range.Text = refTypeToUse;
                        row.Cells[2].Range.Text = "Object";
                        row.Cells[3].Range.Text = name;
                        row.Cells[4].Range.Text = "";
                        row.Cells[5].Range.Text = typeName;
                        row.Cells[6].Range.Text = modelingRule;
                        row.Cells[7].Range.Text = "";
                    }
                }
                if (nsdMember.Name == "ServiceParameter")
                { // control method
                    if (name == "ctlVal")
                    {
                        string dtForDoc = "";
                        string typeName = "";
                        string typeKind = "";
                        string dataTypeName;

                        if (nsdMember.Attributes["type"] != null)
                        {
                            typeName = nsdMember.Attributes["type"].Value;
                        }
                        if (nsdMember.Attributes["typeKind"] != null)
                        {
                            typeKind = nsdMember.Attributes["typeKind"].Value;
                        }
                        dataTypeName = getDatatypeName(typeName, typeKind, false, ref dtForDoc);

                        string operateNodeId = getNodeId(String.Format("{0}|Method:Operate", parentPath));
                        XmlNode nodesetOperateMethodRef = addNodesetMethod(nodesetNode, "Operate", operateNodeId, parentNodeId, "Mandatory");
                        XmlNode nodesetOperateMethodIn = addNodesetMethodInputArguments(nodesetOperateMethodRef, operateNodeId);
                        addNodesetMethodArgument(nodesetOperateMethodIn, "ctlVal", dataTypeName);
                        addNodesetMethodArgument(nodesetOperateMethodIn, "test", getAliasDatatypeNodeId("Boolean"));
                        addNodesetMethodArgument(nodesetOperateMethodIn, "synchroCheck", getAliasDatatypeNodeId("Boolean"));
                        addNodesetMethodArgument(nodesetOperateMethodIn, "interlockCheck", getAliasDatatypeNodeId("Boolean"));
                        XmlNode nodesetOperateMethodOut = addNodesetMethodOutputArguments(nodesetOperateMethodRef, operateNodeId);
                        addNodesetMethodArgument(nodesetOperateMethodOut, "addCause", getNodeId("Enumeration:AddCauseKind"));

                        if ((baseClassTypeName != "ENC") || (isAggregation))
                        {
                            string cancelNodeId = getNodeId(String.Format("{0}|Method:Cancel", parentPath));
                            XmlNode nodesetCancelMethodRef = addNodesetMethod(nodesetNode, "Cancel", cancelNodeId, parentNodeId, "Mandatory");
                            XmlNode nodesetCancelMethodOut = addNodesetMethodOutputArguments(nodesetCancelMethodRef, cancelNodeId);
                            addNodesetMethodArgument(nodesetCancelMethodOut, "addCause", getNodeId("Enumeration:AddCauseKind"));

                            if (!isAggregation)
                            {
                                string selectNodeId = getNodeId(String.Format("{0}|Method:Select", parentPath));
                                XmlNode nodesetSelectMethodRef = addNodesetMethod(nodesetNode, "Select", selectNodeId, parentNodeId, "Optional");
                                XmlNode nodesetSelectMethodOut = addNodesetMethodOutputArguments(nodesetSelectMethodRef, selectNodeId);
                                addNodesetMethodArgument(nodesetSelectMethodOut, "addCause", getNodeId("Enumeration:AddCauseKind"));
                            }
                        }
                        if (!isAggregation)
                        {
                            string selectVNodeId = getNodeId(String.Format("{0}|Method:SelectWithVal", parentPath));
                            XmlNode nodesetSelectVMethodRef = addNodesetMethod(nodesetNode, "SelectWithVal", selectVNodeId, parentNodeId, "Optional");
                            XmlNode nodesetSelectVMethodIn = addNodesetMethodInputArguments(nodesetSelectVMethodRef, selectVNodeId);
                            addNodesetMethodArgument(nodesetSelectVMethodIn, "ctlVal", dataTypeName);
                            addNodesetMethodArgument(nodesetSelectVMethodIn, "operTm", getNodeId("Structure:Timestamp"));
                            addNodesetMethodArgument(nodesetSelectVMethodIn, "test", getAliasDatatypeNodeId("Boolean"));
                            addNodesetMethodArgument(nodesetSelectVMethodIn, "synchroCheck", getAliasDatatypeNodeId("Boolean"));
                            addNodesetMethodArgument(nodesetSelectVMethodIn, "interlockCheck", getAliasDatatypeNodeId("Boolean"));
                            XmlNode nodesetSelectVMethodOut = addNodesetMethodOutputArguments(nodesetSelectVMethodRef, selectVNodeId);
                            addNodesetMethodArgument(nodesetSelectVMethodOut, "addCause", getNodeId("Enumeration:AddCauseKind"));

                            string operateTNodeId = getNodeId(String.Format("{0}|Method:TimeActivatedOperate", parentPath));
                            XmlNode nodesetOperateTMethodRef = addNodesetMethod(nodesetNode, "TimeActivatedOperate", operateTNodeId, parentNodeId, "Optional");
                            XmlNode nodesetOperateTMethodIn = addNodesetMethodInputArguments(nodesetOperateTMethodRef, operateTNodeId);
                            addNodesetMethodArgument(nodesetOperateTMethodIn, "ctlVal", dataTypeName);
                            addNodesetMethodArgument(nodesetOperateTMethodIn, "operTm", getNodeId("Structure:Timestamp"));
                            addNodesetMethodArgument(nodesetOperateTMethodIn, "test", getAliasDatatypeNodeId("Boolean"));
                            addNodesetMethodArgument(nodesetOperateTMethodIn, "synchroCheck", getAliasDatatypeNodeId("Boolean"));
                            addNodesetMethodArgument(nodesetOperateTMethodIn, "interlockCheck", getAliasDatatypeNodeId("Boolean"));
                            XmlNode nodesetOperateTMethodOut = addNodesetMethodOutputArguments(nodesetOperateTMethodRef, operateTNodeId);
                            addNodesetMethodArgument(nodesetOperateTMethodOut, "addCause", getNodeId("Enumeration:AddCauseKind"));
                        }

                        // documentation
                        if ((_wordCurrentTable != null) && (!isAggregation))
                        {
                            Row row;
                            row = _wordCurrentTable.Rows.Add();
                            row.Cells[1].Range.Text = "HasMethod";
                            row.Cells[2].Range.Text = "Method";
                            row.Cells[3].Range.Text = "Operate";
                            row.Cells[4].Range.Text = "";
                            row.Cells[5].Range.Text = "OperateMethod";
                            row.Cells[6].Range.Text = "Mandatory";
                            row.Cells[7].Range.Text = "";
                            if (baseClassTypeName != "ENC")
                            {
                                row = _wordCurrentTable.Rows.Add();
                                row.Cells[1].Range.Text = "HasMethod";
                                row.Cells[2].Range.Text = "Method";
                                row.Cells[3].Range.Text = "Cancel";
                                row.Cells[4].Range.Text = "";
                                row.Cells[5].Range.Text = "CancelMethod";
                                row.Cells[6].Range.Text = "Mandatory";
                                row.Cells[7].Range.Text = "";
                                row = _wordCurrentTable.Rows.Add();
                                row.Cells[1].Range.Text = "HasMethod";
                                row.Cells[2].Range.Text = "Method";
                                row.Cells[3].Range.Text = "Select";
                                row.Cells[4].Range.Text = "";
                                row.Cells[5].Range.Text = "SelectMethod";
                                row.Cells[6].Range.Text = "Optional";
                                row.Cells[7].Range.Text = "";
                            }
                            row = _wordCurrentTable.Rows.Add();
                            row.Cells[1].Range.Text = "HasMethod";
                            row.Cells[2].Range.Text = "Method";
                            row.Cells[3].Range.Text = "SelectWithVal";
                            row.Cells[4].Range.Text = "";
                            row.Cells[5].Range.Text = "SelectWithValMethod";
                            row.Cells[6].Range.Text = "Optional";
                            row.Cells[7].Range.Text = "";
                            row = _wordCurrentTable.Rows.Add();
                            row.Cells[1].Range.Text = "HasMethod";
                            row.Cells[2].Range.Text = "Method";
                            row.Cells[3].Range.Text = "TimeActivatedOperate";
                            row.Cells[4].Range.Text = "";
                            row.Cells[5].Range.Text = "TimeActivatedOperateMethod";
                            row.Cells[6].Range.Text = "Optional";
                            row.Cells[7].Range.Text = "";
                        }
                    }
                }
            }
        }

        void addNodesetObjectTypeCDC(string baseCDC, string specializedCDC, string typeKind)
        {
            XmlNode node = null;
            try
            {
                node = _nsdObjectMap[String.Format("{0}{1}", baseCDC, specializedCDC)];
            }
            catch
            { }

            if (node == null)
            {
                string[] name = { "", "" };
                string[] fc = { "", "" };
                string[] presCond = { "", "" };
                string nameS = "";

                if (specializedCDC == "EnumDA")
                {
                    typeKind = "BASIC";
                }

                if (baseCDC == "ENC")
                {
                    name[0] = "stVal";
                    fc[0] = "SP";
                    presCond[0] = "M";
                    name[1] = "subVal";
                    fc[1] = "SV";
                    presCond[1] = "O";
                    nameS = "ctlVal";
                }
                else if (baseCDC == "ENG")
                {
                    name[0] = "setVal";
                    fc[0] = "SP";
                    presCond[0] = "M";
                }
                else if (baseCDC == "ENS")
                {
                    name[0] = "stVal";
                    fc[0] = "ST";
                    presCond[0] = "M";
                    name[1] = "subVal";
                    fc[1] = "SV";
                    presCond[1] = "O";
                }
                else if (baseCDC == "CTS")
                {
                    name[0] = "ctlVal";
                    fc[0] = "SR";
                    presCond[0] = "M";
                }

                string xml = String.Format("<CDC name=\"{0}{1}\"  base=\"{2}\">", baseCDC, specializedCDC, baseCDC);
                for (int i = 0; i < 2; i++)
                {
                    if (name[i] != "")
                    {  
                        xml += String.Format("<DataAttribute name=\"{0}\" fc=\"{1}\" type=\"{2}\" typeKind=\"{3}\"  presCond=\"{4}\"/>", name[i], fc[i], specializedCDC, typeKind, presCond[i]);
                    }
                }
                if (nameS != "")
                {     
                    xml += String.Format("<ServiceParameter name=\"{0}\" type=\"{1}\" typeKind=\"ENUMERATED\"/>", nameS, specializedCDC);
                }
                xml += "</CDC>";
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xml);
                
                _nsdObjectMap.Add(String.Format("{0}{1}", baseCDC, specializedCDC), doc.ChildNodes[0]);
                Table wordTableSave = _wordCurrentTable;
                createNodesetObjectType(doc.ChildNodes[0]);
                _wordCurrentTable = wordTableSave;
            } 
        }

        XmlNode addNodesetMethod(XmlNode nodesetNode, string name, string nodeId, string parentNodeId, string modelingRule)
        {
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetNode, "Reference", nodeId,  "ReferenceType", "HasComponent");
                                           
            XmlNode nodesetMethodNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAMethod");
            addXmlAttribute(_nodesetDoc, nodesetMethodNode, "NodeId", nodeId);
            addXmlAttribute(_nodesetDoc, nodesetMethodNode, "BrowseName", String.Format("{0}:{1}", _nsIdx, name));
            addXmlAttribute(_nodesetDoc, nodesetMethodNode, "ParentNodeId", parentNodeId);

            addXmlElement(_nodesetDoc, nodesetMethodNode, "DisplayName", name);

            XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetMethodNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesNode, "Reference", parentNodeId, "ReferenceType", "HasComponent", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", modelingRule, "ReferenceType", "HasModellingRule");
            return nodesetReferencesNode;
        }

        XmlNode addNodesetMethodInputArguments(XmlNode nodesetMethodRef, string nodeId)
        {
            string nodeIdInput =  getNodeId(String.Format("{0}-InputArguments", nodeId));

            addXmlElementAndOneAttribute(_nodesetDoc, nodesetMethodRef, "Reference", nodeIdInput, "ReferenceType", "HasProperty");

            XmlNode inputArgumentsNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable");
            addXmlAttribute(_nodesetDoc, inputArgumentsNode, "NodeId", nodeIdInput);
            addXmlAttribute(_nodesetDoc, inputArgumentsNode, "BrowseName", "InputArguments");
            addXmlAttribute(_nodesetDoc, inputArgumentsNode, "ParentNodeId", nodeId);
            addXmlAttribute(_nodesetDoc, inputArgumentsNode, "DataType", "i=296");
            addXmlAttribute(_nodesetDoc, inputArgumentsNode, "ValueRank", "-1");

            addXmlElement(_nodesetDoc, inputArgumentsNode, "DisplayName", "InputArguments");

            XmlNode nodesetInputReferencesNode = addXmlElement(_nodesetDoc, inputArgumentsNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetInputReferencesNode, "Reference", nodeId, "ReferenceType", "HasProperty", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetInputReferencesNode, "Reference", "Mandatory", "ReferenceType", "HasModellingRule");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetInputReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");

            XmlNode nodesetInputValueNode = addXmlElement(_nodesetDoc, inputArgumentsNode, "Value");
            return addXmlElementAndOneAttribute(_nodesetDoc, nodesetInputValueNode, "ListOfExtensionObject", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");
        }

        XmlNode addNodesetMethodOutputArguments(XmlNode nodesetMethodRef, string nodeId)
        {
            string nodeIdOutput =  getNodeId(String.Format("{0}-OutputArguments", nodeId));

            addXmlElementAndOneAttribute(_nodesetDoc, nodesetMethodRef, "Reference", nodeIdOutput, "ReferenceType", "HasProperty");

            XmlNode outputArgumentsNode = addXmlElement(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable");
            addXmlAttribute(_nodesetDoc, outputArgumentsNode, "NodeId", nodeIdOutput);
            addXmlAttribute(_nodesetDoc, outputArgumentsNode, "BrowseName", "OutputArguments");
            addXmlAttribute(_nodesetDoc, outputArgumentsNode, "ParentNodeId", nodeId);
            addXmlAttribute(_nodesetDoc, outputArgumentsNode, "DataType", "i=296");
            addXmlAttribute(_nodesetDoc, outputArgumentsNode, "ValueRank", "-1");

            addXmlElement(_nodesetDoc, outputArgumentsNode, "DisplayName", "OutputArguments");

            XmlNode nodesetOutputReferencesNode = addXmlElement(_nodesetDoc, outputArgumentsNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetOutputReferencesNode, "Reference", nodeId, "ReferenceType", "HasProperty", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetOutputReferencesNode, "Reference", "Mandatory", "ReferenceType", "HasModellingRule");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetOutputReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");

            XmlNode nodesetOutputValueNode = addXmlElement(_nodesetDoc, outputArgumentsNode, "Value");
            return addXmlElementAndOneAttribute(_nodesetDoc, nodesetOutputValueNode, "ListOfExtensionObject", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");
        }

        void addNodesetMethodArgument(XmlNode nodeEOL, string name, string datatypeNodeId)
        {
            XmlNode nodesetValueEONode = addXmlElement(_nodesetDoc, nodeEOL, "ExtensionObject");
            XmlNode nodesetValueTypeIdNode = addXmlElement(_nodesetDoc, nodesetValueEONode, "TypeId");
            XmlNode nodesetValueId = addXmlElement(_nodesetDoc, nodesetValueTypeIdNode, "Identifier", "i=297");
            XmlNode nodesetBodyNode = addXmlElement(_nodesetDoc, nodesetValueEONode, "Body");
            XmlNode nodesetArgNode = addXmlElement(_nodesetDoc, nodesetBodyNode, "Argument");
            addXmlElement(_nodesetDoc, nodesetArgNode, "Name", name);
            XmlNode nodesetArgDTNode = addXmlElement(_nodesetDoc, nodesetArgNode, "DataType");
            addXmlElement(_nodesetDoc, nodesetArgDTNode, "Identifier", datatypeNodeId);
            addXmlElement(_nodesetDoc, nodesetArgNode, "ValueRank", "-1");
            addXmlElement(_nodesetDoc, nodesetArgNode, "ArrayDimensions");
            XmlNode nodesetArgDesNode = addXmlElement(_nodesetDoc, nodesetArgNode, "Description");
            addXmlElement(_nodesetDoc, nodesetArgDesNode, "Locale");
            addXmlElement(_nodesetDoc, nodesetArgDesNode, "Text");
        }

        void createNodesetEnumerationDataType(XmlNode nsdEnumeration)
        {
            _nodesetHasDT = true;
            string enumName = nsdEnumeration.Attributes["name"].Value;
            string nodeId = getNodeId(String.Format("Enumeration:{0}",  enumName));
            string enumValuesNodeId = getNodeId(String.Format("Enumeration#EnumValues:{0}", enumName));

            // nodeset
            XmlNode nodesetDataTypeNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UADataType", "NodeId", nodeId, "BrowseName", String.Format("{0}:{1}", _nsIdx, enumName));
            addXmlElement(_nodesetDoc, nodesetDataTypeNode, "DisplayName", enumName);

            XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetDataTypeNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesNode, "Reference", "Enumeration", "ReferenceType", "HasSubtype", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesNode, "Reference", enumValuesNodeId, "ReferenceType", "HasProperty");
            XmlNode nodesetDefinitionNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetDataTypeNode, "Definition", "Name", enumName);

            XmlNode nodesetEnumValuesNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", enumValuesNodeId, "BrowseName", String.Format("{0}:EnumValues", _nsIdx));
            addXmlAttribute(_nodesetDoc, nodesetEnumValuesNode, "DataType", "i=7594");
            addXmlAttribute(_nodesetDoc, nodesetEnumValuesNode, "ValueRank", "1");
            addXmlElement(_nodesetDoc, nodesetEnumValuesNode, "DisplayName", "EnumValues");

            XmlNode nodesetReferencesEVNode = addXmlElement(_nodesetDoc, nodesetEnumValuesNode, "References");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesEVNode, "Reference", "PropertyType", "ReferenceType", "HasTypeDefinition");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetReferencesEVNode, "Reference", "Mandatory", "ReferenceType", "HasModellingRule");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesEVNode, "Reference", nodeId, "ReferenceType", "HasProperty", "IsForward", "false");
            XmlNode nodesetValueNode = addXmlElement(_nodesetDoc, nodesetEnumValuesNode, "Value");
            XmlNode nodesetListExONode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetValueNode, "ListOfExtensionObject", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            // binary types
            XmlNode binaryEnumTypeNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, _binaryTypesRootNode, "opc", "http://opcfoundation.org/BinarySchema/", "EnumeratedType", "Name", enumName, "LengthInBits", "32");

            // xml types
            XmlNode xmlEnumTypeNode = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "simpleType", "name", enumName);
            XmlNode xmlEnumTypeNodeRes = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, xmlEnumTypeNode, "xs", "http://www.w3.org/2001/XMLSchema", "restriction", "base", "xs:string");
            addQualifiedXmlElementAndTwoAttributes(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "element", "name", enumName, "type", String.Format("{0}", enumName));
            XmlNode xmlComplexTypeNode = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "complexType", "name", String.Format("ListOf{0}", enumName));
            XmlNode xmlComplexTypeNodeSeq = addQualifiedXmlElement(_xmlTypesDoc, xmlComplexTypeNode, "xs", "http://www.w3.org/2001/XMLSchema", "sequence");
            XmlNode xmlComplexTypeNodeEl = addQualifiedXmlElement(_xmlTypesDoc, xmlComplexTypeNodeSeq, "xs", "http://www.w3.org/2001/XMLSchema", "element");
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "name", enumName);
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "type", String.Format("{0}", enumName));
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "minOccurs", "0");
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl, "maxOccurs", "unbounded");
            XmlNode xmlComplexTypeNodeEl2 = addQualifiedXmlElement(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "element");
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl2, "name", String.Format("ListOf{0}", enumName));
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl2, "type", String.Format("ListOf{0}", enumName));
            addXmlAttribute(_xmlTypesDoc, xmlComplexTypeNodeEl2, "nillable", "true");

            // documentation
            Paragraph parTable = null;
            if (_wordDoc != null)
            {
                Paragraph p2 = _wordDoc.Paragraphs.Add();
                p2.Range.Font.Name = "Arial";
                p2.Range.Font.Size = 10F;
                p2.Range.Text = enumName;
                p2.Range.InsertParagraphAfter();

                parTable = _wordDoc.Paragraphs.Add();
                _wordCurrentTable = _wordDoc.Tables.Add(parTable.Range, 1, 2);

                _wordCurrentTable.Range.Font.Name = "Arial";
                _wordCurrentTable.Range.Font.Size = 8F;
                _wordCurrentTable.Range.Font.Bold = 0;

                _wordCurrentTable.Columns[1].Width = 100;
                _wordCurrentTable.Columns[2].Width = 100;

                _wordCurrentTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                _wordCurrentTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            }

            foreach (XmlNode nsdLiteral in nsdEnumeration)
            {
                _nodesetHasDT = true;
                string exoValue = nsdLiteral.Attributes["literalVal"].Value;
                string exoText = nsdLiteral.Attributes["name"].Value;;
                string exoDescription = "";

                if (nsdLiteral.Attributes["descID"] != null)
                {
                    exoDescription = getEnNsdoc(nsdLiteral.Attributes["descID"].Value);
                }

                // nodeset
                XmlNode nodesetFieldNode = addXmlElementAndTwoAttributes(_nodesetDoc, nodesetDefinitionNode, "Field", "Name", exoText, "Value", exoValue);
                if (exoDescription.Length > 0)
                {
                    addXmlElement(_nodesetDoc, nodesetFieldNode, "Description", exoDescription);
                }
                addEnumExtensionObject(_nodesetDoc, nodesetListExONode, exoValue, exoText, exoDescription);

                // binary types
                addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryEnumTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "EnumeratedValue", "Name", exoText, "Value", exoValue);

                // xml typess
                addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, xmlEnumTypeNodeRes, "xs", "http://www.w3.org/2001/XMLSchema", "enumeration", "value", String.Format("{0}_{1}", exoText, exoValue));

                // documentation
                if (_wordCurrentTable != null)
                {
                    Row row = _wordCurrentTable.Rows.Add();
                    row.Cells[1].Range.Text = exoText;
                    row.Cells[2].Range.Text = exoValue;
                }
            }
            if (_wordCurrentTable != null)
            {
                _wordCurrentTable.Rows[1].Range.Font.Bold = 1;
                _wordCurrentTable.Cell(1,1).Range.Text = "Name";
                _wordCurrentTable.Cell(1,2).Range.Text = "Value";
                _wordCurrentTable = null;
            }
        }

        void addEnumExtensionObject(XmlDocument doc, XmlNode node, string value, string text, string description)
        {
            XmlNode extensionObject = addXmlElement(_nodesetDoc, node, "ExtensionObject");
            XmlNode typeId = addXmlElement(_nodesetDoc, extensionObject, "TypeId");
            addXmlElement(_nodesetDoc, typeId, "Identifier", "i=7616");
            XmlNode body = addXmlElement(_nodesetDoc, extensionObject, "Body");
            XmlNode enumValueType = addXmlElement(_nodesetDoc, body, "EnumValueType");
            addXmlElement(_nodesetDoc, enumValueType, "Value", value);
            XmlNode displayName = addXmlElement(_nodesetDoc, enumValueType, "DisplayName");
            addXmlElement(_nodesetDoc, displayName, "Locale");
            addXmlElement(_nodesetDoc, displayName, "Text", text);
            XmlNode descriptionN = addXmlElement(_nodesetDoc, enumValueType, "Description");
            addXmlElement(_nodesetDoc, descriptionN, "Locale");
            addXmlElement(_nodesetDoc, descriptionN, "Text", description);
        }

        void createNodesetStructureDataType(XmlNode nsdStructure)
        {
            string structName = nsdStructure.Attributes["name"].Value;
            string nodeId = getNodeId(String.Format("Structure:{0}",  structName));

            // nodeset
            XmlNode nodesetDataTypeNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UADataType", "NodeId", nodeId, "BrowseName", String.Format("{0}:{1}", _nsIdx, structName));
            addXmlElement(_nodesetDoc, nodesetDataTypeNode, "DisplayName", structName);

            XmlNode nodesetReferencesNode = addXmlElement(_nodesetDoc, nodesetDataTypeNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetReferencesNode, "Reference", "Structure", "ReferenceType", "HasSubtype", "IsForward", "false");
            XmlNode nodesetDefinitionNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetDataTypeNode, "Definition", "Name", structName);

            XmlNode nodesetSchemaEncoding = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAObject", "NodeId", getNodeId(String.Format("{0}-BinaryEnconding", nodeId)), "BrowseName", "Default Binary");
            addXmlAttribute(_nodesetDoc, nodesetSchemaEncoding, "SymbolicName", "Default Binary");
            addXmlElement(_nodesetDoc, nodesetSchemaEncoding, "DisplayName", "Default Binary");
            XmlNode nodesetSchemaEncodingReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaEncoding, "References");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaEncodingReferencesNode, "Reference", "i=76", "ReferenceType", "HasTypeDefinition");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaEncodingReferencesNode, "Reference", nodeId, "ReferenceType", "HasEncoding", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaEncodingReferencesNode, "Reference", getNodeId(String.Format("{0}-BinaryDescription", nodeId)), "ReferenceType", "HasDescription");
 
            XmlNode nodesetSchemaXmlEncoding = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAObject", "NodeId", getNodeId(String.Format("{0}-XMLEnconding", nodeId)), "BrowseName", "Default XML");
            addXmlAttribute(_nodesetDoc, nodesetSchemaXmlEncoding, "SymbolicName", "Default XML");
            addXmlElement(_nodesetDoc, nodesetSchemaXmlEncoding, "DisplayName", "Default XML");
            XmlNode nodesetSchemaXmlEncodingReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaXmlEncoding, "References");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlEncodingReferencesNode, "Reference", "i=76", "ReferenceType", "HasTypeDefinition");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaXmlEncodingReferencesNode, "Reference", nodeId, "ReferenceType", "HasEncoding", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlEncodingReferencesNode, "Reference", getNodeId(String.Format("{0}-XMLDescription", nodeId)), "ReferenceType", "HasDescription");
                                             
            // nodeset binary schema types
            XmlNode nodesetSchemaDescription = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(String.Format("{0}-BinaryDescription", nodeId)), "BrowseName", String.Format("{0}:{1}", _nsIdx,structName));
            addXmlAttribute(_nodesetDoc, nodesetSchemaDescription, "ParentNodeId", getNodeId(_nodeIdTextBinarySchema));
            addXmlAttribute(_nodesetDoc, nodesetSchemaDescription, "DataType", "String");
            addXmlElement(_nodesetDoc, nodesetSchemaDescription, "DisplayName", structName);
            XmlNode nodesetSchemaDescriptionReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaDescription, "References");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaDescriptionReferencesNode, "Reference", "i=69", "ReferenceType", "HasTypeDefinition");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaDescriptionReferencesNode, "Reference", getNodeId(_nodeIdTextBinarySchema), "ReferenceType", "HasComponent", "IsForward", "false");
            XmlNode nodesetSchemaDescriptionValueNode = addXmlElement(_nodesetDoc, nodesetSchemaDescription, "Value");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaDescriptionValueNode, "String", structName, "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            // nodeset XML schema types
            XmlNode nodesetSchemaXmlDescription = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(String.Format("{0}-XMLDescription", nodeId)), "BrowseName", String.Format("{0}:{1}", _nsIdx,structName));
            addXmlAttribute(_nodesetDoc, nodesetSchemaXmlDescription, "ParentNodeId", getNodeId(_nodeIdTextXmlSchema));
            addXmlAttribute(_nodesetDoc, nodesetSchemaXmlDescription, "DataType", "String");
            addXmlElement(_nodesetDoc, nodesetSchemaXmlDescription, "DisplayName", structName);
            XmlNode nodesetSchemaXmlDescriptionReferencesNode = addXmlElement(_nodesetDoc, nodesetSchemaXmlDescription, "References");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlDescriptionReferencesNode, "Reference", "i=69", "ReferenceType", "HasTypeDefinition");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetSchemaXmlDescriptionReferencesNode, "Reference", getNodeId(_nodeIdTextXmlSchema), "ReferenceType", "HasComponent", "IsForward", "false");
            XmlNode nodesetSchemaXmlDescriptionValueNode = addXmlElement(_nodesetDoc, nodesetSchemaXmlDescription, "Value");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetSchemaXmlDescriptionValueNode, "String", String.Format("//xs:element[@name='{0}']", structName), "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            // binary types
            XmlNode binaryStructTypeNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, _binaryTypesRootNode, "opc", "http://opcfoundation.org/BinarySchema/", "StructuredType", "Name", structName, "BaseType", "ua:ExtensionObject");

            // XML types
            XmlNode xmlStructTypeNode = addQualifiedXmlElementAndOneAttribute(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "complexType", "name", structName);
            XmlNode xmlStructTypeSqNode = addQualifiedXmlElement(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "sequence");
            addQualifiedXmlElementAndTwoAttributes(_xmlTypesDoc, _xmlTypesRootNode, "xs", "http://www.w3.org/2001/XMLSchema", "element", "name", structName, "type", String.Format("tns:{0}", structName));

            // documentation
            Paragraph parTable = null;
            if (_wordDoc != null)
            {
                Paragraph p2 = _wordDoc.Paragraphs.Add();
                p2.Range.Font.Name = "Arial";
                p2.Range.Font.Size = 10F;
                p2.Range.Text = structName;
                p2.Range.InsertParagraphAfter();

                parTable = _wordDoc.Paragraphs.Add();
                _wordCurrentTable = _wordDoc.Tables.Add(parTable.Range, 1, 3);

                _wordCurrentTable.Range.Font.Name = "Arial";
                _wordCurrentTable.Range.Font.Size = 8F;
                _wordCurrentTable.Range.Font.Bold = 0;

                _wordCurrentTable.Columns[1].Width = 100;
                _wordCurrentTable.Columns[2].Width = 100;
                _wordCurrentTable.Columns[3].Width = 40;

                _wordCurrentTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                _wordCurrentTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            }

            // count the number of optional fields
            int numOptionalFields=0;
            int numFields=0;
            foreach (XmlNode nsdSubDataAttribute in nsdStructure)
            {
                numFields++;
                if (nsdSubDataAttribute.Attributes["presCond"].Value != "M")
                {
                    numOptionalFields++;
                }
            }

            // binary types - add switch bits
            for (int i = 0; i < numOptionalFields; i++)
            {
                addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryStructTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "Field", "Name", String.Format("Bit{0}", i), "TypeName", "opc:Bit");
            }
            int remainingBitsInByte = 8 - (numOptionalFields % 8);
            if (remainingBitsInByte < 8) 
            {
                XmlNode binaryRemainingBitsInByteNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryStructTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "Field", "Name", "Reserved", "TypeName", "opc:Bit");
                addXmlAttribute(_binaryTypesDoc, binaryRemainingBitsInByteNode, "Length", String.Format("{0}", remainingBitsInByte));
            }

            int curOptionalField = -1;
            foreach (XmlNode nsdSubDataAttribute in nsdStructure)
            {
                string fieldName = nsdSubDataAttribute.Attributes["name"].Value;
                string typeName = nsdSubDataAttribute.Attributes["type"].Value;
                string typeKind = "";
                string dataTypeNodeId = "?#?";
                string dataTypeBinary = "opc:?#?";
                string dataTypeXml = "xs:?#?";
                bool isOptional = (nsdSubDataAttribute.Attributes["presCond"].Value != "M");

                if (isOptional)
                {
                    curOptionalField++;
                }

                if (nsdSubDataAttribute.Attributes["typeKind"] != null)
                {
                    typeKind = nsdSubDataAttribute.Attributes["typeKind"].Value;
                }

                if ((typeKind == "") || (typeKind == "BASIC"))
                {
                    string alias = getAliasOfBasicDatatype(typeName);
                    if (alias.StartsWith("Enumeration:"))
                    {
                        dataTypeNodeId = getNodeId(alias);
                        typeKind = "ENUMERATED";
                    }
                    else if (alias.StartsWith("Structure:"))
                    {
                        dataTypeNodeId = getNodeId(alias);
                        typeKind = "CONSTRUCTED";
                    }
                    else
                    {
                        dataTypeNodeId = alias;
                    }
                }
                else if (typeKind == "ENUMERATED")
                {
                    dataTypeNodeId = getNodeId(String.Format("Enumeration:{0}", typeName));
                }
                else if (typeKind == "CONSTRUCTED")
                {
                    dataTypeNodeId = getNodeId(String.Format("Structure:{0}", typeName));
                }

                if (typeKind == "ENUMERATED")
                {
                    if (dataTypeNodeId.Substring(3,1) == _nsIdx.ToString())
                    {
                        dataTypeBinary = String.Format("tns:{0}", typeName);
                        dataTypeXml = String.Format("tns:{0}", typeName);
                    }
                    else
                    {
                        dataTypeBinary = String.Format("tnsbase:{0}", typeName);
                        dataTypeXml = String.Format("tnsbase:{0}", typeName);
                    }
                }
                else if (typeKind == "CONSTRUCTED")
                {
                    if (dataTypeNodeId.Substring(3,1) == _nsIdx.ToString())
                    {
                        dataTypeBinary = String.Format("tns:{0}", typeName);
                        dataTypeXml = String.Format("tns:{0}", typeName);
                    }
                    else
                    {
                        dataTypeBinary = String.Format("tnsbase:{0}", typeName);
                        dataTypeXml = String.Format("tnsbase:{0}", typeName);
                    }
                }
                else
                { 
                    dataTypeBinary = getBinaryDatatypeName(dataTypeNodeId);
                    dataTypeXml = getXmlDatatypeName(dataTypeNodeId);
                }

                // nodeset
                XmlNode nodesetFieldNode = addXmlElementAndTwoAttributes(_nodesetDoc, nodesetDefinitionNode, "Field", "Name", fieldName , "DataType", dataTypeNodeId);
                if (isOptional)
                {
                    addXmlAttribute(_nodesetDoc, nodesetFieldNode, "IsOptional", "true");
                }
                                            
                // binary types
                XmlNode binaryFieldNode = addQualifiedXmlElementAndTwoAttributes(_binaryTypesDoc, binaryStructTypeNode, "opc", "http://opcfoundation.org/BinarySchema/", "Field", "Name", fieldName, "TypeName", dataTypeBinary);
                if (isOptional)
                {
                    addXmlAttribute(_binaryTypesDoc, binaryFieldNode, "SwitchField", String.Format("Bit{0}", curOptionalField));
                    addXmlAttribute(_binaryTypesDoc, binaryFieldNode, "SwitchValue", "1");
                }

                // XML types
                XmlNode xmlFieldNode = addQualifiedXmlElementAndTwoAttributes(_xmlTypesDoc, xmlStructTypeSqNode, "xs", "http://www.w3.org/2001/XMLSchema", "element", "name", fieldName, "TypeName", dataTypeXml);
                if (isOptional)
                {
                    addXmlAttribute(_xmlTypesDoc, xmlFieldNode, "minOccurs", "0");
                    addXmlAttribute(_xmlTypesDoc, xmlFieldNode, "nillable", "true");
                }

                // documentation
                if (_wordCurrentTable != null)
                {
                    Row row = _wordCurrentTable.Rows.Add();
                    row.Cells[1].Range.Text = fieldName;
                    row.Cells[2].Range.Text = dataTypeBinary.Substring(dataTypeBinary.IndexOf(':') + 1);
                    if (isOptional)
                    {
                        row.Cells[3].Range.Text = "O";
                    }
                    else
                    {
                        row.Cells[3].Range.Text = "M";
                    }
                }
            }

            if (_wordCurrentTable != null)
            { 
                _wordCurrentTable.Rows[1].Range.Font.Bold = 1;
                _wordCurrentTable.Cell(1,1).Range.Text = "Element name";
                _wordCurrentTable.Cell(1,2).Range.Text = "DataType";
                _wordCurrentTable.Cell(1,3).Range.Text = "M/O";
                _wordCurrentTable = null;
            }
            
        }

        void createNodesetTypeDictionary()
        {
            _binaryTypesDoc.Save(_binaryTypesFileName);
            _xmlTypesDoc.Save(_xmlTypesFileName);

            string binarySchemaUriNodeId = getNodeId(String.Format("{0}_NamespaceUri", _nodeIdTextBinarySchema));
            XmlNode nodesetBinarySchemaNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(_nodeIdTextBinarySchema), "BrowseName", String.Format("{0}:Opc.Ua.{1}", _nsIdx, _nodesetTypeDictionaryName));
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaNode, "SymbolicName", "BinarySchema");
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaNode, "DataType", "ByteString");
            addXmlElement(_nodesetDoc, nodesetBinarySchemaNode, "DisplayName", String.Format("Opc.Ua.{0}", _nodesetTypeDictionaryName));

            XmlNode nodesetBinarySchemaReferencesNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetBinarySchemaReferencesNode, "Reference", "i=93","ReferenceType", "HasComponent", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaReferencesNode, "Reference", binarySchemaUriNodeId, "ReferenceType", "HasProperty");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaReferencesNode, "Reference", "i=72", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetBinarySchemaValueNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaNode, "Value");

            XmlNode nodesetBinarySchemaValueBSNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaValueNode, "ByteString", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            MemoryStream msBinary = new MemoryStream();
            XmlTextWriter xBinary = new XmlTextWriter(msBinary, new UTF8Encoding(false));
            xBinary.Formatting = Formatting.Indented;
            _binaryTypesDoc.Save(xBinary);
            xBinary.Close();
            nodesetBinarySchemaValueBSNode.InnerText = Convert.ToBase64String(msBinary.ToArray());

            XmlNode nodesetBinarySchemaUriNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", binarySchemaUriNodeId, "BrowseName",  String.Format("{0}:NamespaceUri", _nsIdx));
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaUriNode, "ParentNodeId", getNodeId(_nodeIdTextBinarySchema));
            addXmlAttribute(_nodesetDoc, nodesetBinarySchemaUriNode, "DataType", "String");
            addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "DisplayName", "NamespaceUri");
            addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "Description", "A URI that uniquely identifies the dictionary.");
            XmlNode nodesetBinarySchemaUriReferencesNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetBinarySchemaUriReferencesNode, "Reference", getNodeId(_nodeIdTextBinarySchema),"ReferenceType", "HasProperty", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaUriReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetBinarySchemaUriValueNode = addXmlElement(_nodesetDoc, nodesetBinarySchemaUriNode, "Value");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetBinarySchemaUriValueNode, "String", _nodesetURL, "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            string xmlSchemaUriNodeId = getNodeId(String.Format("{0}_NamespaceUri", _nodeIdTextXmlSchema));
            XmlNode nodesetxmlSchemaNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", getNodeId(_nodeIdTextXmlSchema), "BrowseName", String.Format("{0}:Opc.Ua.{1}", _nsIdx, _nodesetTypeDictionaryName));
            addXmlAttribute(_nodesetDoc, nodesetxmlSchemaNode, "SymbolicName", "XmlSchema");
            addXmlAttribute(_nodesetDoc, nodesetxmlSchemaNode, "DataType", "ByteString");
            addXmlElement(_nodesetDoc, nodesetxmlSchemaNode, "DisplayName", String.Format("Opc.Ua.{0}", _nodesetTypeDictionaryName));

            XmlNode nodesetxmlSchemaReferencesNode = addXmlElement(_nodesetDoc, nodesetxmlSchemaNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetxmlSchemaReferencesNode, "Reference", "i=92","ReferenceType", "HasComponent", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetxmlSchemaReferencesNode, "Reference", xmlSchemaUriNodeId, "ReferenceType", "HasProperty");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetxmlSchemaReferencesNode, "Reference", "i=72", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetxmlSchemaValueNode = addXmlElement(_nodesetDoc, nodesetxmlSchemaNode, "Value");

            XmlNode nodesetxmlSchemaValueBSNode = addXmlElementAndOneAttribute(_nodesetDoc, nodesetxmlSchemaValueNode, "ByteString", "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");

            MemoryStream msXml = new MemoryStream();
            XmlTextWriter xXml = new XmlTextWriter(msXml, new UTF8Encoding(false));
            xXml.Formatting = Formatting.Indented;
            _xmlTypesDoc.Save(xXml);
            xXml.Close();
            nodesetxmlSchemaValueBSNode.InnerText = Convert.ToBase64String(msXml.ToArray());

            XmlNode nodesetXmlSchemaUriNode = addXmlElementAndTwoAttributes(_nodesetDoc, _nodesetUANodeSetNode, "UAVariable", "NodeId", xmlSchemaUriNodeId, "BrowseName",  String.Format("{0}:NamespaceUri", _nsIdx));
            addXmlAttribute(_nodesetDoc, nodesetXmlSchemaUriNode, "ParentNodeId", getNodeId(_nodeIdTextXmlSchema));
            addXmlAttribute(_nodesetDoc, nodesetXmlSchemaUriNode, "DataType", "String");
            addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "DisplayName", "NamespaceUri");
            addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "Description", "A URI that uniquely identifies the dictionary.");
            XmlNode nodesetXmlSchemaUriReferencesNode = addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "References");
            addXmlElementAndTwoAttributes(_nodesetDoc, nodesetXmlSchemaUriReferencesNode, "Reference", getNodeId(_nodeIdTextXmlSchema),"ReferenceType", "HasProperty", "IsForward", "false");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetXmlSchemaUriReferencesNode, "Reference", "i=68", "ReferenceType", "HasTypeDefinition");
            XmlNode nodesetXmlSchemaUriValueNode = addXmlElement(_nodesetDoc, nodesetXmlSchemaUriNode, "Value");
            addXmlElementAndOneAttribute(_nodesetDoc, nodesetXmlSchemaUriValueNode, "String", _nodesetURL, "xmlns", "http://opcfoundation.org/UA/2008/02/Types.xsd");
        }

        string getNodeId(string strNodeId)
        {
            string nodeId = null;
            try
            {
                nodeId = _nodesetNodeIdMap[strNodeId];
            }
            catch
            { }

            if (nodeId == null)
            {
                nodeId = String.Format("ns={0};i={1}", _nsIdx, _nextNodeId);
                _nextNodeId++;
                _nodesetNodeIdMap.Add(strNodeId, nodeId);
            }

            return nodeId;
        }

        string getDatatypeNodeId(string typeName, string typeKind)
        {
            string dataTypeNodeId = "?!?";
            if ((typeKind == "") || (typeKind == "BASIC"))
            {
                string alias = getAliasOfBasicDatatype(typeName);
                if (alias.StartsWith("Enumeration:"))
                {
                    dataTypeNodeId = getNodeId(alias);
                    typeKind = "ENUMERATED";
                }
                else if (alias.StartsWith("Structure:"))
                {
                    dataTypeNodeId = getNodeId(alias);
                    typeKind = "CONSTRUCTED";
                }
                else
                {
                    dataTypeNodeId = alias;
                }
            }
            else if (typeKind == "ENUMERATED")
            {
                dataTypeNodeId = getNodeId(String.Format("Enumeration:{0}", typeName));
            }
            else if (typeKind == "CONSTRUCTED")
            {
                dataTypeNodeId = getNodeId(String.Format("Structure:{0}", typeName));
            }
            return dataTypeNodeId;
        }

        string getDatatypeName(string typeName, string typeKind, bool aliasAllowed, ref string readableName)
        {
            string dataTypeName = "?!?";
            if ((typeKind == "") || (typeKind == "BASIC"))
            {
                string alias = getAliasOfBasicDatatype(typeName);
                if (alias.StartsWith("Enumeration:"))
                {
                    dataTypeName = getNodeId(alias);
                    typeKind = "ENUMERATED";
                    readableName = alias.Substring("Enumeration:".Length);
                }
                else if (alias.StartsWith("Structure:"))
                {
                    dataTypeName = getNodeId(alias);
                    typeKind = "CONSTRUCTED";
                    readableName = alias.Substring("Structure:".Length);
                }
                else
                {
                    readableName = alias;
                    if (aliasAllowed)
                    {
                        dataTypeName = alias;
                    }
                    else
                    {
                        dataTypeName = getAliasDatatypeNodeId(alias);
                    }
                }
            }
            else if (typeKind == "ENUMERATED")
            {
                if (typeName != "")
                {
                    readableName = typeName;
                    dataTypeName = getNodeId(String.Format("Enumeration:{0}", typeName));
                }
                else
                {
                    readableName = "Enumeration";
                    if (aliasAllowed)
                    {
                        dataTypeName = readableName;
                    }
                    else
                    {
                        dataTypeName = getAliasDatatypeNodeId(readableName);
                    }
                }
            }
            else if (typeKind == "CONSTRUCTED")
            {
                readableName = typeName;
                dataTypeName = getNodeId(String.Format("Structure:{0}", typeName));
            }
            else if (typeKind == "undefined")
            {
                readableName = "BaseDataType";
                if (aliasAllowed)
                {
                    dataTypeName = readableName;
                }
                else
                {
                    dataTypeName = getAliasDatatypeNodeId(readableName);
                }
            }
            return dataTypeName;
        }

        XmlAttribute addXmlAttribute(XmlDocument doc, XmlNode node, string name, string value)
        {
            XmlAttribute attr = doc.CreateAttribute(name);
            attr.Value = value;
            node.Attributes.Append(attr);
            return attr;
        }

        XmlAttribute addXmlAttributeDeep(XmlDocument doc, XmlNode node, string name, string value)
        {
            XmlAttribute attr = doc.CreateAttribute(name);
            attr.Value = value;
            node.Attributes.Append(attr);

            foreach (XmlNode child in node.ChildNodes)
            {
                addXmlAttributeDeep(doc, child, name, value);
            }
            return attr;
        }

        XmlNode addXmlElement(XmlDocument doc, XmlNode parent, string name, string innerText)
        {
            XmlNode node = addXmlElement(doc, parent, name);
            node.InnerText = innerText;
            return node;
        }

        XmlNode addXmlElement(XmlDocument doc, XmlNode parent, string name)
        {
            XmlNode node = doc.CreateElement(name);
            parent.AppendChild(node);
            return node;
        }

        XmlNode addXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string elName, string attrName, string attrValue)
        {
            XmlNode node = addXmlElement(doc, parent, elName);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string elName, string innerText, string attrName, string attrValue)
        {
            XmlNode node = addXmlElement(doc, parent, elName, innerText);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string elName, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addXmlElement(doc, parent, elName);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        XmlNode addXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string elName, string innerText, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addXmlElement(doc, parent, elName, innerText);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        XmlNode addQualifiedXmlElement(XmlDocument doc, XmlNode parent, string prefix, string uri, string name, string innerText)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, name);
            node.InnerText = innerText;
            return node;
        }

        XmlNode addQualifiedXmlElement(XmlDocument doc, XmlNode parent, string prefix, string uri, string name)
        {
            XmlNode node = doc.CreateElement(prefix, name, uri);
            parent.AppendChild(node);
            return node;
        }

        XmlNode addQualifiedXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string attrName, string attrValue)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addQualifiedXmlElementAndOneAttribute(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string innerText, string attrName, string attrValue)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName, innerText);
            addXmlAttribute(doc, node, attrName, attrValue);
            return node;
        }

        XmlNode addQualifiedXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        XmlNode addQualifiedXmlElementAndTwoAttributes(XmlDocument doc, XmlNode parent, string prefix, string uri, string elName, string innerText, string attrName1, string attrValue1, string attrName2, string attrValue2)
        {
            XmlNode node = addQualifiedXmlElement(doc, parent, prefix, uri, elName, innerText);
            addXmlAttribute(doc, node, attrName1, attrValue1);
            addXmlAttribute(doc, node, attrName2, attrValue2);
            return node;
        }

        string getAliasOfBasicDatatype(string dt)
        {
            if (dt == "BOOLEAN") { return "Boolean"; }
            else if (dt == "INT8") { return "SByte"; }
            else if (dt == "INT16") { return "Int16"; }
            else if (dt == "INT32") { return "Int32"; }
            else if (dt == "INT64") { return "Int64"; }
            else if (dt == "INT8U") { return "Byte"; }
            else if (dt == "INT16U") { return "UInt16"; }
            else if (dt == "INT32U") { return "UInt32"; }
            else if (dt == "FLOAT32") { return "Float"; }
            else if (dt == "Octet64") { return "ByteString"; }
            else if (dt == "VisString64") { return "String"; }
            else if (dt == "VisString129") { return "String"; }
            else if (dt == "VisString255") { return "String"; }
            else if (dt == "Unicode255") { return "String"; }
            else if (dt == "PhyComAddr") { return "String"; }
            else if (dt == "ObjRef") { return "String"; }
            else if (dt == "EntryID") { return "ByteString"; }
            else if (dt == "Currency") { return "String"; }
            else if (dt == "Timestamp") { return "Structure:Timestamp"; }
            else if (dt == "Quality") { return "Structure:Quality"; }
            else if (dt == "EntryTime") { return "Structure:Timestamp"; }
            else if (dt == "TrgOps") { return "Structure:TrgOps"; }
            else if (dt == "OptFlds") { return "Structure:OptFlds"; }
            else if (dt == "SvOptFlds") { return "Structure:SvOptFlds"; }
            else if (dt == "Check") { return "Structure:Check"; }
            else if (dt == "Tcmd") { return "Enumeration:StepControlKind"; }
            else if (dt == "Dbpos") { return "Enumeration:DpStatusKind"; }
            else if (dt == "EnumDA") { return "Enumeration"; }
            return "?#?";
        }

        string getBinaryDatatypeName(string dt)
        {
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                if (_aliases[i, 0] == dt)
                {
                    return _aliases[i, 2]; 
                }
            }
            return "";
        }

        string getXmlDatatypeName(string dt)
        {
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                if (_aliases[i, 0] == dt)
                {
                    return _aliases[i, 3]; 
                }
            }
            return "";
        }

        string getAliasDatatypeNodeId(string dt)
        {
            for (int i = 0; i < (_aliases.Length) / 4; i++)
            {
                if (_aliases[i, 0] == dt)
                {
                    return _aliases[i, 1]; 
                }
            }
            return "";
        }
        
        XmlNode getNsdObject(string id)
        { 
            XmlNode nsd = null;
            try
            {
                nsd = _nsdObjectMap[id];
            }
            catch
            { }
            return nsd;
        }
    }
}
