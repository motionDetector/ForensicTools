[DocsMetaParser]
- Parsing Metadata of Documents.
- Supported extension: docx, pptx, xlsx
- OS: Windows 10

[Source]
- docProps/app.xml
- docProps/core.xml

[Usage]
- Parse a file  
C:\DocsMetaParser.exe -f C:\filename.docx  

- Parse files in directory (not recursively)
C:\DocsMetaParser.exe -d C:\DocumentFolder
 
- Parse files recursively in directory
C:\DocsMetaParser.exe -r C:\
  
- Export results as csv file
C:\DocsMetaParser.exe -csv -r C:\
