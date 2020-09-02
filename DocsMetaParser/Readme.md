[DocsMetaParser]
- Parsing Metadata of Documents.
- Supported extension: pdf, docx, pptx, xlsx, doc, ppt, xls
- OS: Windows 10

[Source]
- docProps/app.xml (docx, pptx, xlsx)
- docProps/core.xml (docx, pptx, xlsx)
- a part of Summary Information stream (doc, ppt, xls)
- a part of Document Summary Information stream (doc, ppt, xls)
- xmpmeta (pdf)
- plain text (pdf)

[Usage]
- Parse a file (-f)
- Parse files in directory (-d)
- Parse files recursively in directory (-r)
- Save results as CSV file format (-csv)
#You need to import csv file as UTF-8 option in Excel program to see hangul text.

[Examples]
- C:\DocsMetaParser.exe -f C:\filename.docx 
- C:\DocsMetaParser.exe -d C:\DocumentFolder    
- C:\DocsMetaParser.exe -r C:\ 
- C:\DocsMetaParser.exe -csv -r C:\ 
