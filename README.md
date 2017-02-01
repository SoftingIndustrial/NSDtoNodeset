# NSDtoNodeset
Tool to convert IEC 61850-7-7 NSD files to OPC UA nodeset files.

This tool is used to generate the OPC UA nodeset files for the IEC 61850 OPC UA companion specification.

## Build
NSDtoNodeset is written in C# using the .NET framework V4.5.2. <br>
Microsoft Visual Studio 2015 is used to build it. 

## Command line arguments
* /nsd < file name ><br>
  Name of the NSD file to convert.<br>
  <br>This argument could be used multiple times.
* /nsdBase < file name ><br>
  Name of the NSD file used for base type information.<br>
  <br>This argument could be used multiple times.
* /nodeset < file name ><br>
  Name of the generated nodeset file
* /nodesetUrl < URL string ><br>
  URL used for the generation of the nodeset
* /nodesetUrlBase < URL string ><br>
  Nodeset URL of the base types passed with /ndsBase
* /nodesetTypeDictionary < name ><br>
  Name of the type dictionary in the nodeset
  <br> [optional; default: "NSDtoNodeset"]
* /nodesetImport < file name ><br>
  Name of the nodeset file to import
  <br> [optional]
* /nodesetStartId < id ><br>
  Start integer node id for the conversion 
  <br> [optional; default: 0]
* /nodeIdMap < file name ><br>
  Name of the node id mapping file
  <br>[optional; default: "NodeIdMap.txt"]
* /nodeIdMapBase < file name ><br>
  Name of the base node id mapping file
  <br>[optional; default: "NodeIdMap.txt"]
* /binaryTypes < file name ><br>
  Name of the binary types schema file 
  <br> [optional, default: "BinaryTypes.xml"]
* /xmlTypes < file name ><br>
  Name of the XML types schema file 
  <br> [optional, default: "XMLTypes.xml"]
* /word < file name >  
  Name of the MS Word file to generate. An absolute path has to be used.
  <br>[optional]  

