[DocsMetaParser]
- Parsing Metadata of Documents.
- Supported extension: docx, pptx, xlsx
- OS: Windows 10

[Source]
- docProps/app.xml
- docProps/core.xml

[Usage]
- C:\DocsMetaParser.exe -f C:\filename.docx
Parse a file
  
- C:\DocsMetaParser.exe -d C:\DocumentFolder
   // Parse files in directory (not recursively)
  
- C:\DocsMetaParser.exe -r C:\
   // Parse files recursively in directory
  
- C:\DocsMetaParser.exe -csv -r C:\
   // Export results as csv file
