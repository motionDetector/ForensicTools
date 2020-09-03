package main

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/binary"
	"encoding/csv"
	"errors"
	"flag"
	"fmt"
	"io/ioutil"
	"log"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"syscall"
	"time"

	xj "github.com/basgys/goxml2json"
	"github.com/karrick/godirwalk"
	"golang.org/x/text/encoding/korean"
	"golang.org/x/text/encoding/unicode"
	"golang.org/x/text/transform"
)

func main() {
	file := flag.String("f", "", usageFile)
	dir := flag.String("d", "", usageDir)
	rcv := flag.String("r", "", usageRcv)
	csv := flag.Bool("csv", false, usageCsv)
	flag.Parse()
	st = time.Now()
	if flag.NFlag() == 0 {
		flag.Usage()
		return
	}
	if *file != "" {
		parseFile(*file)
	}
	if *dir != "" {
		parseDir(*dir)
	}
	if *rcv != "" {
		parseRcv(*rcv)
	}
	et = time.Now()
	printEndTime(et)
	dt = diffTime(st, et)
	fmt.Println("ElapsedTime:", dt)
	if *csv {
		createOxmlCsv(oxmlDocs)
		createDocCsv(docs)
		createPdfCsv(pdfDocs)
		fmt.Println("CSV file Extracted.")
	}
}
func parseFile(file string) {
	_, err := ioutil.ReadFile(file)
	checkFatalErr(err)
	if chk := checkOxmlExt(file); chk {
		if chk := oxmlHeaderCheck(file); chk {
			readOxmlFile(file)
		}
	}
	if chk := checkDocExt(file); chk {
		if chk := docHeaderCheck(file); chk {
			readDocFile(file)
		}
	}
	if chk := checkPdfExt(file); chk {
		if chk := pdfHeaderCheck(file); chk {
			readPdfFile(file)
		}
	}
	if !checkOxmlExt(file) && !checkPdfExt(file) && !checkDocExt(file) {
		log.Fatalln(errHelp)
	}
}
func parseDir(dir string) {
	files, err := ioutil.ReadDir(dir)
	checkFatalErr(err)
	printStartTime(st)
	fmt.Println("Checking extension and header of files...")
	for _, f := range files {
		path := (filepath.Join(dir, f.Name()))
		if chk := checkOxmlExt(path); chk {
			if chk := oxmlHeaderCheck(path); chk {
				readOxmlFile(path)
			}
		}
		if chk := checkDocExt(path); chk {
			if chk := docHeaderCheck(path); chk {
				readDocFile(path)
			}
		}
		if chk := checkPdfExt(path); chk {
			if chk := pdfHeaderCheck(path); chk {
				readPdfFile(path)
			}
		}
	}
	if (oxmlDocs == nil) && (pdfDocs == nil) && (docs == nil) {
		log.Fatalln(errHelp)
	}
}
func parseRcv(rcv string) {
	_, err := ioutil.ReadDir(rcv)
	checkFatalErr(err)
	printStartTime(st)
	fmt.Println("Checking extension and header of files...")
	files := dirWalk(rcv)
	for _, f := range files {
		if chk := checkOxmlExt(f); chk {
			if chk := oxmlHeaderCheck(f); chk {
				readOxmlFile(f)
			}
		}
		if chk := checkDocExt(f); chk {
			if chk := docHeaderCheck(f); chk {
				readDocFile(f)
			}
		}
		if chk := checkPdfExt(f); chk {
			if chk := pdfHeaderCheck(f); chk {
				readPdfFile(f)
			}
		}
	}
	if (oxmlDocs == nil) && (pdfDocs == nil) && (docs == nil) {
		log.Fatalln(errHelp)
	}
}
func readOxmlFile(file string) {
	var docMeta oxmlDocsMeta
	oxmlMeta := make(map[string]string)
	fi, err := os.Stat(file)
	checkErr(err)
	t := fi.Sys().(*syscall.Win32FileAttributeData)
	docMeta.fileSize = strconv.FormatInt(fi.Size(), 10)
	docMeta.fileExt = filepath.Ext(file)
	docMeta.filePath, _ = filepath.Split(file)
	docMeta.fileCreationTime = time.Unix(0, t.CreationTime.Nanoseconds()).String()
	docMeta.fileWriteTime = time.Unix(0, t.LastWriteTime.Nanoseconds()).String()
	docMeta.fileAccessTime = time.Unix(0, t.LastAccessTime.Nanoseconds()).String()
	docMeta.fileName = filepath.Base(file)
	docMeta.number = strconv.Itoa(cnt)
	cnt++
	r, err := zip.OpenReader(file)
	checkErr(err)
	defer r.Close()
	for _, f := range r.File {
		if f.Name == fileApp {
			fo, err := f.Open()
			checkErr(err)
			b := make([]byte, f.FileInfo().Size())
			rFo := bufio.NewReader(fo)
			_, err = rFo.Read(b)
			checkErr(err)
			fileAppMap := map[string][]byte{
				oxmlApp:                  []byte("\x3C\x41\x70\x70\x6C\x69\x63\x61\x74\x69\x6F\x6E\x3E"),
				oxmlAppVersion:           []byte("\x3C\x41\x70\x70\x56\x65\x72\x73\x69\x6F\x6E\x3E"),
				oxmlCompany:              []byte("\x3C\x43\x6F\x6D\x70\x61\x6E\x79\x3E"),
				oxmlTotalTime:            []byte("\x3C\x54\x6F\x74\x61\x6C\x54\x69\x6D\x65\x3E"),
				oxmlWords:                []byte("\x3C\x57\x6F\x72\x64\x73\x3E"),
				oxmlPages:                []byte("\x3C\x50\x61\x67\x65\x73\x3E"),
				oxmlCharacters:           []byte("\x3C\x43\x68\x61\x72\x61\x63\x74\x65\x72\x73\x3E"),
				oxmlLines:                []byte("\x3C\x4C\x69\x6E\x65\x73\x3E"),
				oxmlPrphs:                []byte("\x3C\x50\x61\x72\x61\x67\x72\x61\x70\x68\x73\x3E"),
				oxmlSlides:               []byte("\x3C\x53\x6C\x69\x64\x65\x73\x3E"),
				oxmlNotes:                []byte("\x3C\x4E\x6F\x74\x65\x73\x3E"),
				oxmlHiddenSlides:         []byte("\x3C\x48\x69\x64\x64\x65\x6E\x53\x6C\x69\x64\x65\x73\x3E"),
				oxmlMmClips:              []byte("\x3C\x4D\x4D\x43\x6C\x69\x70\x73\x3E"),
				oxmlTemplate:             []byte("\x3C\x54\x65\x6D\x70\x6C\x61\x74\x65\x3E"),
				oxmlPresentationFormat:   []byte("\x3C\x50\x72\x65\x73\x65\x6E\x74\x61\x74\x69\x6F\x6E\x46\x6F\x72\x6D\x61\x74\x3E"),
				oxmlLinksUpToDate:        []byte("\x3C\x4C\x69\x6E\x6B\x73\x55\x70\x54\x6F\x44\x61\x74\x65\x3E"),
				oxmlCharactersWithSpaces: []byte("\x3C\x43\x68\x61\x72\x61\x63\x74\x65\x72\x73\x57\x69\x74\x68\x53\x70\x61\x63\x65\x73\x3E"),
				oxmlSharedDoc:            []byte("\x3C\x53\x68\x61\x72\x65\x64\x44\x6F\x63\x3E"),
				oxmlHyperlinksChanged:    []byte("\x3C\x48\x79\x70\x65\x72\x6C\x69\x6E\x6B\x73\x43\x68\x61\x6E\x67\x65\x64\x3E"),
				oxmlDocSecurity:          []byte("\x3C\x44\x6F\x63\x53\x65\x63\x75\x72\x69\x74\x79\x3E"),
				oxmlScaleCrop:            []byte("\x3C\x53\x63\x61\x6C\x65\x43\x72\x6F\x70\x3E"),
			}
			for k, v := range fileAppMap {
				if chk := bytes.Contains(b, v); !chk {
					continue
				}
				temp := b
				idx := bytes.Index(b, v) + len(v)
				temp = temp[idx:]
				end := bytes.Index(temp, []byte("\x3C"))
				temp = temp[:end]
				oxmlMeta[k] = string(temp)
			}
			fo.Close()
		}
		if f.Name == fileCore {
			fo, err := f.Open()
			checkErr(err)
			b := make([]byte, f.FileInfo().Size())
			rFo := bufio.NewReader(fo)
			_, err = rFo.Read(b)
			checkErr(err)
			fileCoreMap := map[string][]byte{
				oxmlTitle:          []byte("\x3C\x64\x63\x3A\x74\x69\x74\x6C\x65\x3E"),                                                                                                                                         // "<dc:title>"
				oxmlCreator:        []byte("\x3C\x64\x63\x3A\x63\x72\x65\x61\x74\x6F\x72\x3E"),                                                                                                                                 // "<dc:creator>"
				oxmlLastModifiedBy: []byte("\x3C\x63\x70\x3A\x6C\x61\x73\x74\x4D\x6F\x64\x69\x66\x69\x65\x64\x42\x79\x3E"),                                                                                                     // "<cp:lastModifiedBy>"
				oxmlRevision:       []byte("\x3C\x63\x70\x3A\x72\x65\x76\x69\x73\x69\x6F\x6E\x3E"),                                                                                                                             // "<cp:revision>"
				oxmlLastPrinted:    []byte("\x3C\x63\x70\x3A\x6C\x61\x73\x74\x50\x72\x69\x6E\x74\x65\x64\x3E"),                                                                                                                 // "<cp:lastPrinted>"
				oxmlCreated:        []byte("\x3C\x64\x63\x74\x65\x72\x6D\x73\x3A\x63\x72\x65\x61\x74\x65\x64\x20\x78\x73\x69\x3A\x74\x79\x70\x65\x3D\x22\x64\x63\x74\x65\x72\x6D\x73\x3A\x57\x33\x43\x44\x54\x46\x22\x3E"),     // "<dcterms:created xsi:type=\"dcterms:W3CDTF\">"
				oxmlModified:       []byte("\x3C\x64\x63\x74\x65\x72\x6D\x73\x3A\x6D\x6F\x64\x69\x66\x69\x65\x64\x20\x78\x73\x69\x3A\x74\x79\x70\x65\x3D\x22\x64\x63\x74\x65\x72\x6D\x73\x3A\x57\x33\x43\x44\x54\x46\x22\x3E"), // "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">"
			}
			for k, v := range fileCoreMap {
				if chk := bytes.Contains(b, v); !chk {
					continue
				}
				temp := b
				idx := bytes.Index(b, v) + len(v)
				temp = temp[idx:]
				end := bytes.Index(temp, []byte("\x3C"))
				temp = temp[:end]
				oxmlMeta[k] = string(temp)
			}
			fo.Close()
		}
	}
	docMeta.application = oxmlMeta[oxmlApp]
	docMeta.appVersion = oxmlMeta[oxmlAppVersion]
	docMeta.company = oxmlMeta[oxmlCompany]
	docMeta.title = oxmlMeta[oxmlTitle]
	docMeta.creator = oxmlMeta[oxmlCreator]
	docMeta.lastModifiedBy = oxmlMeta[oxmlLastModifiedBy]
	docMeta.revision = oxmlMeta[oxmlRevision]
	docMeta.created = oxmlMeta[oxmlCreated]
	docMeta.modified = oxmlMeta[oxmlModified]
	docMeta.totalTime = oxmlMeta[oxmlTotalTime]
	docMeta.pages = oxmlMeta[oxmlPages]
	docMeta.words = oxmlMeta[oxmlWords]
	docMeta.characters = oxmlMeta[oxmlCharacters]
	docMeta.lines = oxmlMeta[oxmlLines]
	docMeta.paragraphs = oxmlMeta[oxmlPrphs]
	docMeta.template = oxmlMeta[oxmlTemplate]
	docMeta.charactersWithSpaces = oxmlMeta[oxmlCharactersWithSpaces]
	docMeta.lastPrinted = oxmlMeta[oxmlLastPrinted]
	docMeta.slides = oxmlMeta[oxmlSlides]
	docMeta.notes = oxmlMeta[oxmlNotes]
	docMeta.hiddenSlides = oxmlMeta[oxmlHiddenSlides]
	docMeta.presentationFormat = oxmlMeta[oxmlPresentationFormat]
	docMeta.mmClips = oxmlMeta[oxmlMmClips]
	docMeta.sharedDoc = oxmlMeta[oxmlSharedDoc]
	docMeta.hyperlinksChanged = oxmlMeta[oxmlHyperlinksChanged]
	docMeta.docSecurity = oxmlMeta[oxmlDocSecurity]
	docMeta.scaleCrop = oxmlMeta[oxmlScaleCrop]
	docMeta.linksUpToDate = oxmlMeta[oxmlLinksUpToDate]
	oxmlDocs = append(oxmlDocs, docMeta)
	printOxmlDoc(docMeta)
}
func readPdfFile(file string) {
	var docMeta pdfFileMeta
	var rdf []byte
	pdfMeta := make(map[string]string)
	fi, err := os.Stat(file)
	checkErr(err)
	t := fi.Sys().(*syscall.Win32FileAttributeData)
	docMeta.fileSize = strconv.FormatInt(fi.Size(), 10)
	docMeta.fileExt = filepath.Ext(file)
	docMeta.filePath, _ = filepath.Split(file)
	docMeta.fileCreationTime = time.Unix(0, t.CreationTime.Nanoseconds()).String()
	docMeta.fileWriteTime = time.Unix(0, t.LastWriteTime.Nanoseconds()).String()
	docMeta.fileAccessTime = time.Unix(0, t.LastAccessTime.Nanoseconds()).String()
	docMeta.fileName = filepath.Base(file)
	docMeta.number = strconv.Itoa(cnt)
	cnt++
	fo, _ := os.Open(file)
	defer fo.Close()
	fs, _ := fo.Stat()
	b := make([]byte, fs.Size())
	rFo := bufio.NewReader(fo)
	_, err = rFo.Read(b)
	checkErr(err)
	j := x2j(b)
	hb := make([]byte, 3)
	fo.Seek(5, 0)
	_, err = rFo.Read(hb)
	checkErr(err)
	docMeta.appVersion = string(hb)
	xmpRdfMap := map[string][]byte{
		pdfTitle:       []byte("\x22\x74\x69\x74\x6C\x65\x22\x3A"),
		pdfSubject:     []byte("\x22\x73\x75\x62\x6A\x65\x63\x74\x22\x3A"),
		pdfCreator:     []byte("\x22\x63\x72\x65\x61\x74\x6F\x72\x22\x3A"),
		pdfPublisher:   []byte("\x22\x70\x75\x62\x6C\x69\x73\x68\x65\x72\x22\x3A"),
		pdfDescription: []byte("\x22\x64\x65\x73\x63\x72\x69\x70\x74\x69\x6F\x6E\x22\x3A"),
		pdfRights:      []byte("\x22\x72\x69\x67\x68\x74\x73\x22\x3A"),
	}
	for k, v := range xmpRdfMap {
		if chk := bytes.Contains(j, v); !chk {
			continue
		}
		if (k == pdfTitle) || (k == pdfDescription) || (k == pdfRights) {
			rdf = []byte("\x22\x23\x63\x6F\x6E\x74\x65\x6E\x74\x22\x3A")
		}
		if (k == pdfCreator) || (k == pdfPublisher) || (k == pdfSubject) {
			rdf = []byte("\x22\x6C\x69\x22\x3A")
		}
		chk := bytes.Contains(j, rdf)
		item := bytes.Contains(j, v)
		if !(chk && item) {
			continue
		}
		temp := j
		idx := bytes.Index(j, v) + len(v) + 1
		temp = temp[idx:]
		idx = bytes.Index(temp, rdf) + len(rdf) + 1
		temp = temp[idx:]
		if temp[0] == byte(0x22) {
			if temp[1] != byte(0x22) {
				temp = temp[1:]
				end := bytes.Index(temp, []byte("\x22"))
				if end != -1 {
					temp = temp[:end]
					pdfMeta[k] = string(temp)
				}
			}
		}
		if temp[0] == byte(0x5B) {
			temp = temp[1:]
			end := bytes.Index(temp, []byte("\x5D"))
			if end != -1 {
				temp = temp[:end]
				pdfMeta[k] = string(temp)
			}
		}
	}
	xmpMap := map[string][]byte{
		pdfProducer:           []byte("\x22\x50\x72\x6F\x64\x75\x63\x65\x72\x22\x3A"),
		pdfCreatorTool:        []byte("\x22\x43\x72\x65\x61\x74\x6F\x72\x54\x6F\x6F\x6C\x22\x3A"),
		pdfCreatorTool2:       []byte("\x22\x2D\x43\x72\x65\x61\x74\x6F\x72\x54\x6F\x6F\x6C\x22\x3A"),
		pdfKeywords:           []byte("\x22\x4B\x65\x79\x77\x6F\x72\x64\x73\x22\x3A"),
		pdfCreateDate:         []byte("\x22\x43\x72\x65\x61\x74\x65\x44\x61\x74\x65\x22\x3A"),
		pdfCreateDate4:        []byte("\x22\x2D\x43\x72\x65\x61\x74\x65\x44\x61\x74\x65\x22\x3A"),
		pdfCreated:            []byte("\x22\x43\x72\x65\x61\x74\x65\x64\x22\x3A"),
		pdfModifyDate:         []byte("\x22\x4D\x6F\x64\x69\x66\x79\x44\x61\x74\x65\x22\x3A"),
		pdfModifyDate2:        []byte("\x22\x2D\x4D\x6F\x64\x69\x66\x79\x44\x61\x74\x65\x22\x3A"),
		pdfSourceModified:     []byte("\x22\x53\x6F\x75\x72\x63\x65\x4D\x6F\x64\x69\x66\x69\x65\x64\x22\x3A"),
		pdfLastSaved:          []byte("\x22\x4C\x61\x73\x74\x53\x61\x76\x65\x64\x22\x3A"),
		pdfMetadataDate:       []byte("\x22\x4D\x65\x74\x61\x64\x61\x74\x61\x44\x61\x74\x65\x22\x3A"),
		pdfMetadataDate2:      []byte("\x22\x2D\x4D\x65\x74\x61\x64\x61\x74\x61\x44\x61\x74\x65\x22\x3A"),
		pdfDocumentID:         []byte("\x22\x44\x6F\x63\x75\x6D\x65\x6E\x74\x49\x44\x22\x3A"),
		pdfInstanceID:         []byte("\x22\x49\x6E\x73\x74\x61\x6E\x63\x65\x49\x44\x22\x3A"),
		pdfTrapped:            []byte("\x22\x54\x72\x61\x70\x70\x65\x64\x22\x3A"),
		pdfAggregationType:    []byte("\x22\x61\x67\x67\x72\x65\x67\x61\x74\x69\x6F\x6E\x54\x79\x70\x65\x22\x3A"),
		pdfPublicationName:    []byte("\x22\x70\x75\x62\x6C\x69\x63\x61\x74\x69\x6F\x6E\x4E\x61\x6D\x65\x22\x3A"),
		pdfEdition:            []byte("\x22\x65\x64\x69\x74\x69\x6F\x6E\x22\x3A"),
		pdfCopyRight:          []byte("\x22\x63\x6F\x70\x79\x72\x69\x67\x68\x74\x22\x3A"),
		pdfISBN:               []byte("\x22\x69\x73\x62\x6E\x22\x3A"),
		pdfCoverDisplayDate:   []byte("\x22\x63\x6F\x76\x65\x72\x44\x69\x73\x70\x6C\x61\x79\x44\x61\x74\x65\x22\x3A"),
		pdfOriginalDocumentID: []byte("\x22\x6F\x72\x69\x67\x69\x6E\x61\x6C\x44\x6F\x63\x75\x6D\x65\x6E\x74\x49\x44\x22\x3A"),
	}
	for k, v := range xmpMap {
		if chk := bytes.Contains(j, v); !chk {
			continue
		}
		temp := j
		idx := bytes.Index(j, v) + len(v) + 1
		temp = temp[idx:]
		if temp[0] == byte(0x22) {
			if temp[1] != byte(0x22) {
				temp = temp[1:]
				end := bytes.Index(temp, []byte("\x22"))
				temp = temp[:end]
				pdfMeta[k] = string(temp)
			}
		}
		if temp[0] == byte(0x5B) {
			temp = temp[1:]
			end := bytes.Index(temp, []byte("\x5D"))
			temp = temp[:end]
			pdfMeta[k] = string(temp)
		}
	}
	xmpLastMap := map[string][]byte{
		pdfFormat:         []byte("\x22\x2D\x66\x6F\x72\x6D\x61\x74\x22\x3A"),
		pdfSoftwareAgent:  []byte("\x22\x2D\x73\x6F\x66\x74\x77\x61\x72\x65\x41\x67\x65\x6E\x74\x22\x3A"),
		pdfSoftwareAgent2: []byte("\x22\x73\x6F\x66\x74\x77\x61\x72\x65\x41\x67\x65\x6E\x74\x22\x3A"),
		pdfParameters:     []byte("\x22\x70\x61\x72\x61\x6D\x65\x74\x65\x72\x73\x22\x3A"),
		pdfResourcePath:   []byte("\x22\x66\x69\x6C\x65\x50\x61\x74\x68\x22\x3A"),
		pdfDocChangeCount: []byte("\x22\x44\x6F\x63\x43\x68\x61\x6E\x67\x65\x43\x6F\x75\x6E\x74\x22\x3A"),
	}
	for k, v := range xmpLastMap {
		if chk := bytes.Contains(j, v); !chk {
			continue
		}
		temp := j
		idx := bytes.LastIndex(j, v) + len(v) + 1
		temp = temp[idx:]
		if temp[0] == byte(0x22) {
			if temp[1] != byte(0x22) {
				temp = temp[1:]
				end := bytes.Index(temp, []byte("\x22"))
				temp = temp[:end]
				pdfMeta[k] = string(temp)
			}
		}
	}
	plainMap := map[string][]byte{
		pdfTitle2:      []byte("\x2F\x54\x69\x74\x6C\x65\x28"),
		pdfTitle3:      []byte("\x2F\x54\x69\x74\x6C\x65\x20\x28"),
		pdfAuthor:      []byte("\x2F\x41\x75\x74\x68\x6F\x72\x28"),
		pdfAuthor2:     []byte("\x2F\x41\x75\x74\x68\x6F\x72\x20\x28"),
		pdfCreateDate2: []byte("\x2F\x43\x72\x65\x61\x74\x69\x6F\x6E\x44\x61\x74\x65\x28"),
		pdfCreateDate3: []byte("\x2F\x43\x72\x65\x61\x74\x69\x6F\x6E\x44\x61\x74\x65\x20\x28"),
		pdfCreator2:    []byte("\x2F\x43\x72\x65\x61\x74\x6F\x72\x28"),
		pdfCreator3:    []byte("\x2F\x43\x72\x65\x61\x74\x6F\x72\x20\x28"),
		pdfKeywords2:   []byte("\x2F\x4B\x65\x79\x77\x6F\x72\x64\x73\x28"),
		pdfModDate:     []byte("\x2F\x4D\x6F\x64\x44\x61\x74\x65\x28"),
		pdfModDate2:    []byte("\x2F\x4D\x6F\x64\x44\x61\x74\x65\x20\x28"),
		pdfProducer2:   []byte("\x2F\x50\x72\x6F\x64\x75\x63\x65\x72\x28"),
		pdfProducer3:   []byte("\x2F\x50\x72\x6F\x64\x75\x63\x65\x72\x20\x28"),
		pdfLang:        []byte("\x2F\x4C\x61\x6E\x67\x28"),
		pdfEncrypt:     []byte("\x2F\x45\x6E\x63\x72\x79\x70\x74"),
		pdfSubject2:    []byte("\x2F\x53\x75\x62\x6A\x65\x63\x74\x28"),
		pdfPages:       []byte("\x2F\x43\x6F\x6E\x74\x65\x6E\x74\x73"),
	}
	for k, v := range plainMap {
		if k == pdfPages {
			cnt := 0
			temp := b
			val := bytes.Contains(b, v)
			for val {
				cnt++
				idx := bytes.Index(temp, v) + len(v)
				temp = temp[idx:]
				val = bytes.Contains(temp, v)
			}
			switch cnt {
			case 0:
				pdfMeta[k] = ""
			default:
				pdfMeta[k] = strconv.Itoa(cnt)
			}
		}
		if v := bytes.Contains(b, []byte("\x2F\x45\x6E\x63\x72\x79\x70\x74")); v {
			pdfMeta[pdfEncrypt] = "True"
			continue
		}
		if val := bytes.Contains(b, v); val && (k != pdfPages) {
			temp := b
			idx := bytes.LastIndex(b, v) + len(v)
			temp = temp[idx:]
			if temp[0] == 0x29 || temp[0] == 0x5C {
				continue
			}
			if mz := bytes.Contains(temp, []byte("\x4D\x6F\x7A\x69\x6C\x6C\x61")); mz {
				end := bytes.Index(temp, []byte("\x0A"))
				if end != -1 {
					temp = temp[:end]
					pdfMeta[k] = string(temp)
				}
			} else {
				end := bytes.Index(temp, []byte("\x29"))
				if end != -1 {
					temp = temp[:end]
					chk := temp[:3]
					if chk[0] == 0xFE && chk[1] == 0xFF && chk[2] == 0x00 {
						pdfMeta[k] = cleanStr2(temp)
					} else if chk[0] == 0xFE && chk[1] == 0xFF {
						pdfMeta[k] = cleanStr2(temp)
					} else {
						pdfMeta[k] = string(temp)
					}
				}
			}
		}
	}
	docMeta.title = pdfMeta[pdfTitle]
	docMeta.title2 = pdfMeta[pdfTitle2]
	docMeta.title3 = pdfMeta[pdfTitle3]
	docMeta.subject = pdfMeta[pdfSubject]
	docMeta.subject2 = pdfMeta[pdfSubject2]
	docMeta.pages = pdfMeta[pdfPages]
	docMeta.creator = pdfMeta[pdfCreator]
	docMeta.creator2 = pdfMeta[pdfCreator2]
	docMeta.creator3 = pdfMeta[pdfCreator3]
	docMeta.author = pdfMeta[pdfAuthor]
	docMeta.author2 = pdfMeta[pdfAuthor2]
	docMeta.description = pdfMeta[pdfDescription]
	docMeta.rights = pdfMeta[pdfRights]
	docMeta.producer = pdfMeta[pdfProducer]
	docMeta.producer2 = pdfMeta[pdfProducer2]
	docMeta.producer3 = pdfMeta[pdfProducer3]
	docMeta.creatorTool = pdfMeta[pdfCreatorTool]
	docMeta.creatorTool2 = pdfMeta[pdfCreatorTool2]
	docMeta.Keywords = pdfMeta[pdfKeywords]
	docMeta.Keywords2 = pdfMeta[pdfKeywords2]
	docMeta.created = pdfMeta[pdfCreated]
	docMeta.createDate = pdfMeta[pdfCreateDate]
	docMeta.createDate2 = pdfMeta[pdfCreateDate2]
	docMeta.createDate3 = pdfMeta[pdfCreateDate3]
	docMeta.createDate4 = pdfMeta[pdfCreateDate4]
	docMeta.modifyDate = pdfMeta[pdfModifyDate]
	docMeta.modifyDate2 = pdfMeta[pdfModifyDate2]
	docMeta.modDate = pdfMeta[pdfModDate]
	docMeta.modDate2 = pdfMeta[pdfModDate2]
	docMeta.sourceModified = pdfMeta[pdfSourceModified]
	docMeta.metaDataDate = pdfMeta[pdfMetadataDate]
	docMeta.metaDataDate2 = pdfMeta[pdfMetadataDate2]
	docMeta.lastSaved = pdfMeta[pdfLastSaved]
	docMeta.instanceID = pdfMeta[pdfInstanceID]
	docMeta.documentID = pdfMeta[pdfDocumentID]
	docMeta.format = pdfMeta[pdfFormat]
	docMeta.softwareAgent = pdfMeta[pdfSoftwareAgent]
	docMeta.softwareAgent2 = pdfMeta[pdfSoftwareAgent2]
	docMeta.parameters = pdfMeta[pdfParameters]
	docMeta.resourcePath = pdfMeta[pdfResourcePath]
	docMeta.docChangeCount = pdfMeta[pdfDocChangeCount]
	docMeta.publisher = pdfMeta[pdfPublisher]
	docMeta.aggregationType = pdfMeta[pdfAggregationType]
	docMeta.publicationName = pdfMeta[pdfPublicationName]
	docMeta.edition = pdfMeta[pdfEdition]
	docMeta.copyright = pdfMeta[pdfCopyRight]
	docMeta.coverDisplayDate = pdfMeta[pdfCoverDisplayDate]
	docMeta.originalDocumentID = pdfMeta[pdfOriginalDocumentID]
	docMeta.trapped = pdfMeta[pdfTrapped]
	docMeta.lang = pdfMeta[pdfLang]
	docMeta.encrypt = pdfMeta[pdfEncrypt]
	pdfDocs = append(pdfDocs, docMeta)
	printPdf(docMeta)
}
func readDocFile(file string) {
	var docMeta oxmlDocsMeta
	fi, err := os.Stat(file)
	checkErr(err)
	t := fi.Sys().(*syscall.Win32FileAttributeData)
	docMeta.fileSize = strconv.FormatInt(fi.Size(), 10)
	docMeta.fileExt = filepath.Ext(file)
	docMeta.filePath, _ = filepath.Split(file)
	docMeta.fileCreationTime = time.Unix(0, t.CreationTime.Nanoseconds()).String()
	docMeta.fileWriteTime = time.Unix(0, t.LastWriteTime.Nanoseconds()).String()
	docMeta.fileAccessTime = time.Unix(0, t.LastAccessTime.Nanoseconds()).String()
	docMeta.fileName = filepath.Base(file)
	docMeta.number = strconv.Itoa(cnt)
	cnt++
	fo, _ := os.Open(file)
	defer fo.Close()
	fs, _ := fo.Stat()
	b := make([]byte, fs.Size())
	rFo := bufio.NewReader(fo)
	_, err = rFo.Read(b)
	checkErr(err)
	var meta, smry, docSmry, f1, f2 []byte
	var idx int
	var num uint32
	var tt time.Time
	zero := time.Date(1601, 01, 01, 00, 00, 00, 00, time.UTC)
	idx = bytes.Index(b, []byte("\xE0\x85\x9F\xF2\xF9\x4F\x68\x10\xAB\x91\x08\x00\x2B\x27\xB3\xD9"))
	if idx != -1 && len(b) > idx+512 {
		smry = b[idx-28 : idx+700]
		f1 = smry[80:81]
		f2 = smry[96:97]
	} else {
		return
	}
	idx = bytes.Index(b, []byte("\x02\xD5\xCD\xD5\x9C\x2E\x1B\x10\x93\x97\x08\x00\x2B\x2C\xF9\xAE"))
	if idx != -1 && len(b) > idx+700 {
		docSmry = b[idx-28 : idx+700]
	}
	idx = bytes.Index(smry, []byte("\xFE\xFF\x00\x00"))
	if idx != -1 && len(smry) > idx {
		meta = smry[4:6]
		docMeta.winVersion = strconv.Itoa(int(meta[0])) + "." + strconv.Itoa(int(meta[1]))
	}
	idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
	if idx != -1 && len(smry) > idx {
		meta = smry[idx-4 : idx]
		_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
		docMeta.codePage = strconv.Itoa(int(num))
	}
	if (f1[0] == 0x04) && (f2[0] == 0x06) && (filepath.Ext(file) == ".doc") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.keywords = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.comments = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.pages = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]

		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.characters = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.lines = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x04) && (f2[0] == 0x07) && (filepath.Ext(file) == ".doc") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.keywords = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.pages = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.characters = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			meta = docSmry[idx+8 : idx+8+int(num)]
			docMeta.company = cleanStr(meta)
			docSmry = docSmry[idx+8+int(num):]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.lines = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x07) && (f2[0] == 0x09) && (filepath.Ext(file) == ".doc") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.pages = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.characters = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.lines = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x04) && (f2[0] == 0x06) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.keywords = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.comments = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x04) && (f2[0] == 0x08) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.keywords = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
	}
	if (f1[0] == 0x08) && (f2[0] == 0x0B) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}

		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x09) && (f2[0] == 0x0B) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
	}
	if (f1[0] == 0x12) && (f2[0] == 0x0C) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
	}
	if (f1[0] == 0x12) && (f2[0] == 0x0D) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
	}
	if (f1[0] == 0x0B) && (f2[0] == 0x0D) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x0C) && (f2[0] == 0x13) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x0D) && (f2[0] == 0x02) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x0D) && (f2[0] == 0x00) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x13) && (f2[0] == 0x1E) && (filepath.Ext(file) == ".xls") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.docSecurity = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
	}
	if (f1[0] == 0x04) && (f2[0] == 0x06) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.keywords = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.comments = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			meta = docSmry[idx+8 : idx+8+int(num)]
			docMeta.category = string(meta)
			docSmry = docSmry[idx+8+int(num):]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.manager = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x05) && (f2[0] == 0x07) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.keywords = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.comments = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x04) && (f2[0] == 0x08) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x04) && (f2[0] == 0x09) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.subject = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x06) && (f2[0] == 0x08) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.comments = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			meta = smry[idx+8 : idx+8+int(num)]
			docMeta.application = string(meta)
			smry = smry[idx+8+int(num):]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x07) && (f2[0] == 0x09) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.template = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x06) && (f2[0] == 0x09) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.comments = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.lastPrinted = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.manager = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.company = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x08) && (f2[0] == 0x12) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.creator = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.application = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.created = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	if (f1[0] == 0x09) && (f2[0] == 0x0D) && (filepath.Ext(file) == ".ppt") {
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.title = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.lastModifiedBy = cleanStr(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(smry) > idx+8+int(num) {
				meta = smry[idx+8 : idx+8+int(num)]
				docMeta.revision = string(meta)
				smry = smry[idx+8+int(num):]
			} else {
				smry = smry[idx+8:]
			}
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			df := docDiffTime(zero, tt)
			docMeta.totalTime = df
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x40\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+12 {
			meta = smry[idx+4 : idx+12]
			tt = timeStamp(int64(binary.LittleEndian.Uint64(meta)))
			docMeta.modified = tt.String()
			smry = smry[idx+12:]
		}
		idx = bytes.Index(smry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(smry) > idx+8 {
			meta = smry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.words = strconv.Itoa(int(num))
			smry = smry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x1E\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			if len(docSmry) > idx+8+int(num) {
				meta = docSmry[idx+8 : idx+8+int(num)]
				docMeta.presentationFormat = cleanStr(meta)
				docSmry = docSmry[idx+8+int(num):]
			} else {
				docSmry = docSmry[idx+8:]
			}
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.paragraphs = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.slides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.notes = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.hiddenSlides = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
		idx = bytes.Index(docSmry, []byte("\x03\x00\x00\x00"))
		if idx != -1 && len(docSmry) > idx+8 {
			meta = docSmry[idx+4 : idx+8]
			_ = binary.Read(bytes.NewReader(meta), binary.LittleEndian, &num)
			docMeta.mmClips = strconv.Itoa(int(num))
			docSmry = docSmry[idx+8:]
		}
	}
	docs = append(docs, docMeta)
	printDoc(docMeta)
}

func checkErr(e error) {
	if e != nil {
		fmt.Println(errHelp)
	}
}

func checkFatalErr(e error) {
	if e != nil {
		log.Fatalln(errHelp)
	}
}

func checkOxmlExt(file string) bool {
	str := filepath.Ext(file)
	if str == ".pptx" || str == ".docx" || str == ".xlsx" {
		return true
	}
	return false
}

func checkPdfExt(file string) bool {
	str := filepath.Ext(file)
	if str == ".pdf" {
		return true
	}
	return false
}

func checkDocExt(file string) bool {
	str := filepath.Ext(file)
	if str == ".ppt" || str == ".doc" || str == ".xls" {
		return true
	}
	return false
}

func oxmlHeaderCheck(file string) bool {
	f, err := os.Open(file)
	checkErr(err)
	b := make([]byte, 4)
	_, err = f.Read(b)
	checkErr(err)
	h := []byte("\x50\x4B\x03\x04")
	return bytes.Equal(b, h)
}

func pdfHeaderCheck(file string) bool {
	f, err := os.Open(file)
	checkErr(err)
	b := make([]byte, 4)
	_, err = f.Read(b)
	checkErr(err)
	h := []byte("\x25\x50\x44\x46")
	return bytes.Equal(b, h)
}

func docHeaderCheck(file string) bool {
	f, err := os.Open(file)
	checkErr(err)
	b := make([]byte, 4)
	_, err = f.Read(b)
	checkErr(err)
	h := []byte("\xD0\xCF\x11\xE0")
	return bytes.Equal(b, h)
}

func x2j(b []byte) []byte {
	var js []byte
	if v := bytes.Contains(b, []byte("\x3C\x3F\x78\x70\x61\x63\x6B\x65\x74\x20\x62\x65\x67\x69\x6E")); v {
		temp := b
		idx := bytes.LastIndex(temp, []byte("\x3C\x3F\x78\x70\x61\x63\x6B\x65\x74\x20\x62\x65\x67\x69\x6E"))
		temp = temp[idx:]
		idx = bytes.Index(temp, []byte("\x3C\x2F\x78\x3A\x78\x6D\x70\x6D\x65\x74\x61\x3E")) + 12
		temp = temp[:idx]
		x := strings.NewReader(string(temp))
		j, err := xj.Convert(x)
		checkErr(err)
		js = j.Bytes()
	}
	return js
}

func createOxmlCsv(docs []oxmlDocsMeta) {
	sfx := suffixTime(et)
	file, err := os.Create("DocsMeta(ooxml)" + "_" + sfx + ".csv")
	checkErr(err)
	w := csv.NewWriter(file)
	defer file.Close()
	h := oxmlDocHeaders()
	wErr := w.Write(h)
	checkErr(wErr)

	for n, meta := range docs {
		n++
		num := strconv.Itoa(n)
		metaSlice := []string{num, meta.fileName, meta.filePath, meta.fileCreationTime, meta.fileWriteTime, meta.fileAccessTime,
			meta.fileSize, meta.fileExt, meta.title, meta.creator, meta.lastModifiedBy, meta.revision, meta.lastPrinted, meta.created, meta.modified,
			meta.application, meta.appVersion, meta.company, meta.totalTime, meta.words, meta.pages, meta.characters, meta.lines, meta.paragraphs,
			meta.slides, meta.notes, meta.hiddenSlides, meta.mmClips, meta.template, meta.presentationFormat, meta.linksUpToDate, meta.charactersWithSpaces,
			meta.sharedDoc, meta.hyperlinksChanged, meta.docSecurity, meta.scaleCrop}
		wErr := w.Write(metaSlice)
		checkErr(wErr)
	}
	wErr = w.Write([]string{""})
	checkErr(wErr)
	clnSt := cleanTime(st).String()
	clnEt := cleanTime(et).String()
	stSlice := []string{"StartTime", clnSt}
	etSlice := []string{"EndtTime", clnEt}
	dtSlice := []string{"ElapsedTime", dt}
	wErr = w.Write(stSlice)
	checkErr(wErr)
	wErr = w.Write(etSlice)
	checkErr(wErr)
	wErr = w.Write(dtSlice)
	checkErr(wErr)
	w.Flush()
}

func createDocCsv(docs []oxmlDocsMeta) {
	sfx := suffixTime(et)
	file, err := os.Create("DocsMeta(ole)" + "_" + sfx + ".csv")
	checkErr(err)
	w := csv.NewWriter(file)
	defer file.Close()
	h := docHeaders()
	wErr := w.Write(h)
	checkErr(wErr)

	for n, meta := range docs {
		n++
		num := strconv.Itoa(n)
		metaSlice := []string{num, meta.fileName, meta.filePath, meta.fileCreationTime, meta.fileWriteTime, meta.fileAccessTime,
			meta.fileSize, meta.fileExt, meta.application, meta.title, meta.subject, meta.keywords, meta.comments, meta.template, meta.creator, meta.lastModifiedBy,
			meta.revision, meta.lastPrinted, meta.created, meta.modified,
			meta.winVersion, meta.company, meta.totalTime, meta.words, meta.pages, meta.characters, meta.lines, meta.paragraphs,
			meta.slides, meta.notes, meta.hiddenSlides, meta.mmClips, meta.linksUpToDate, meta.charactersWithSpaces,
			meta.sharedDoc, meta.hyperlinksChanged, meta.docSecurity, meta.presentationFormat, meta.manager, meta.slides, meta.notes, meta.hiddenSlides, meta.mmClips}
		wErr := w.Write(metaSlice)
		checkErr(wErr)
	}
	wErr = w.Write([]string{""})
	checkErr(wErr)
	clnSt := cleanTime(st).String()
	clnEt := cleanTime(et).String()
	stSlice := []string{"StartTime", clnSt}
	etSlice := []string{"EndtTime", clnEt}
	dtSlice := []string{"ElapsedTime", dt}
	wErr = w.Write(stSlice)
	checkErr(wErr)
	wErr = w.Write(etSlice)
	checkErr(wErr)
	wErr = w.Write(dtSlice)
	checkErr(wErr)
	w.Flush()
}

func createPdfCsv(docs []pdfFileMeta) {
	sfx := suffixTime(et)
	file, err := os.Create("PdfMeta" + "_" + sfx + ".csv")
	checkErr(err)
	w := csv.NewWriter(file)
	defer file.Close()
	h := pdfHeaders()
	wErr := w.Write(h)
	checkErr(wErr)

	for n, meta := range docs {
		n++
		num := strconv.Itoa(n)
		metaSlice := []string{num, meta.fileName, meta.filePath, meta.fileCreationTime, meta.fileWriteTime, meta.fileAccessTime,
			meta.fileSize, meta.fileExt, meta.appVersion, meta.title, meta.title2, meta.title3, meta.subject, meta.subject2, meta.pages,
			meta.creator, meta.creator2, meta.creator3, meta.author, meta.author2, meta.created, meta.createDate, meta.createDate2, meta.createDate3, meta.createDate4,
			meta.modifyDate, meta.modifyDate2, meta.sourceModified, meta.modDate, meta.modDate2, meta.lastSaved, meta.metaDataDate, meta.metaDataDate2, meta.description,
			meta.rights, meta.encrypt, meta.producer, meta.producer2, meta.producer3, meta.creatorTool, meta.creatorTool2, meta.softwareAgent, meta.softwareAgent2,
			meta.documentID, meta.instanceID, meta.originalDocumentID, meta.Keywords, meta.Keywords2, meta.trapped, meta.docChangeCount, meta.parameters, meta.format,
			meta.resourcePath, meta.publisher, meta.aggregationType, meta.publicationName, meta.edition, meta.copyright, meta.isbn, meta.coverDisplayDate, meta.lang}
		wErr := w.Write(metaSlice)
		checkErr(wErr)
	}
	wErr = w.Write([]string{""})
	checkErr(wErr)
	clnSt := cleanTime(st).String()
	clnEt := cleanTime(et).String()
	stSlice := []string{"StartTime", clnSt}
	etSlice := []string{"EndtTime", clnEt}
	dtSlice := []string{"ElapsedTime", dt}
	wErr = w.Write(stSlice)
	checkErr(wErr)
	wErr = w.Write(etSlice)
	checkErr(wErr)
	wErr = w.Write(dtSlice)
	checkErr(wErr)
	w.Flush()
}
func oxmlDocHeaders() []string {
	h := []string{
		"Number", "FileName", "FilePath", "FileCreated", "FileModified", "FileAccessed", "FileSize(bytes)", "FileExt",
		"Title", "Creator", "LastModifiedBy", "Revision", "LastPrinted", "Created", "Modified",
		"Application", "AppVersion", "Company", "TotalTime", "Words", "Pages", "Characters", "Lines", "Paragraphs", "Slides",
		"Notes", "HiddenSlides", "MMClips", "Template", "PresentationFormat", "LinksUpToDate", "CharactersWithSpaces",
		"SharedDoc", "HyperlinksChanged", "DocSecurity", "ScaleCrop"}
	return h
}
func docHeaders() []string {
	h := []string{
		"Number", "FileName", "FilePath", "FileCreated", "FileModified", "FileAccessed", "FileSize(bytes)", "FileExt",
		"Application", "Title", "Subject", "Keywords", "Comments", "Template", "Creator", "LastModifiedBy", "Revision", "LastPrinted", "Created", "LastSaved",
		"WinVersion", "Company", "TotalTime", "Words", "Pages", "Characters", "Lines", "Paragraphs", "Slides",
		"Notes", "HiddenSlides", "MMClips", "LinksUpToDate", "CharactersWithSpaces",
		"SharedDoc", "HyperlinksChanged", "DocSecurity", "PresentationTarget", "Manager", "Slides", "Notes", "HiddenSlides", "MMClips"}
	return h
}
func pdfHeaders() []string {
	h := []string{
		"Number", "FileName", "FilePath", "FileCreated", "FileModified", "FileAccessed", "FileSize(bytes)", "FileExt",
		"AppVersion", "Title", "Title2", "Title3", "Subject", "Subject2", "Pages", "Creator", "Creator2", "Creator3", "Author", "Author2",
		"Created", "CreationDate", "CreationDate2", "CreationDate3", "CreationDate4", "ModifyDate", "ModifyDate2", "SourceModified",
		"ModDate", "ModDate2", "LastSaved", "MetaDataDate", "MetaDataDate2", "Description", "Rights", "Encrypted", "Producer", "Producer2", "Producer3",
		"CreatorTool", "CreatorTool2", "SoftwareAgent", "SoftwareAgent2", "DocumentID", "InstanceID", "OriginalDocumentID", "Keywords", "Keywords2",
		"Trapped", "DocChangeCount", "Parameters", "Format", "ResourcePath", "Publisher", "AggregationType", "PublicationName", "Edition", "Copyright",
		"ISBN", "CoverDisplayDate", "Language"}
	return h
}
func dirWalk(dir string) []string {
	var file []string
	godirwalk.Walk(dir, &godirwalk.Options{
		Unsorted: true,
		Callback: func(osPathname string, de *godirwalk.Dirent) error {
			file = append(file, osPathname)
			return nil
		},
		ErrorCallback: func(osPathname string, err error) godirwalk.ErrorAction {
			return godirwalk.SkipNode
		},
	})
	return file
}
func cleanTime(t time.Time) time.Time {
	return time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), 0, time.Local)
}
func cleanStr2(val []byte) string {
	debyte, _, _ := transform.Bytes(unicode.UTF16(unicode.BigEndian, unicode.IgnoreBOM).NewDecoder(), val)
	deStr := string(debyte)
	return deStr
}
func cleanStr(meta []byte) string {
	var bufs bytes.Buffer
	wr := transform.NewWriter(&bufs, korean.EUCKR.NewDecoder())
	wr.Write(meta)
	wr.Close()
	return bufs.String()
}
func printEndTime(t time.Time) {
	fmt.Println()
	fmt.Println("--------------------------------------------------")
	fmt.Println("Parsing Done.")
	fmt.Println("End time:", cleanTime(t))
}
func printStartTime(t time.Time) {
	fmt.Println("DocsMetaParser v1.0")
	fmt.Println("Author: Bloody Mary")
	fmt.Println("Start time:", cleanTime(t))
	fmt.Println("--------------------------------------------------")
}
func diffTime(st, et time.Time) string {
	hs := et.Sub(st).Hours()
	hs, mf := math.Modf(hs)
	ms := mf * 60
	ms, sf := math.Modf(ms)
	ss := sf * 60
	hsStr := strconv.FormatFloat(hs, 'f', 0, 64)
	msStr := strconv.FormatFloat(ms, 'f', 0, 64)
	ssStr := strconv.FormatFloat(ss, 'f', 2, 64)
	return hsStr + " " + "hours" + " " + msStr + " " + "minutes" + " " + ssStr + " " + "seconds"
}
func docDiffTime(st, et time.Time) string {
	hs := et.Sub(st).Hours()
	hs, mf := math.Modf(hs)
	ms := mf * 60
	ms, sf := math.Modf(ms)
	ss := sf * 60
	hsStr := strconv.FormatFloat(hs, 'f', 0, 64)
	msStr := strconv.FormatFloat(ms, 'f', 0, 64)
	ssStr := strconv.FormatFloat(ss, 'f', 0, 64)
	return hsStr + "h " + msStr + "m " + ssStr + "s"
}
func timeStamp(v int64) time.Time {

	t := time.Date(1601, 1, 1, 0, 0, 0, 0, time.UTC)
	d := time.Duration(v)
	for i := 0; i < 100; i++ {
		t = t.Add(d)
	}
	t = time.Date(t.Year(), t.Month(), t.Day(), t.Hour(), t.Minute(), t.Second(), 0, time.UTC)
	return t
}
func suffixTime(t time.Time) string {
	s := time.Now().Format("20060102T150405")
	return s
}
func printOxmlDoc(docMeta oxmlDocsMeta) {
	fmt.Println()
	fmt.Println("Number:", docMeta.number)
	fmt.Println("FileName:", docMeta.fileName)
	fmt.Println("FilePath:", docMeta.filePath)
	fmt.Println("FileCreated:", docMeta.fileCreationTime)
	fmt.Println("FileModified:", docMeta.fileWriteTime)
	fmt.Println("FileAccessed:", docMeta.fileAccessTime)
	fmt.Println("FileSize:", docMeta.fileSize, "bytes")
	fmt.Println("FileExt:", docMeta.fileExt)
	fmt.Println("Application:", docMeta.application)
	fmt.Println("AppVersion:", docMeta.appVersion)
	fmt.Println("Company:", docMeta.company)
	fmt.Println("TotalTime:", docMeta.totalTime)
	fmt.Println("Words:", docMeta.words)
	fmt.Println("Pages:", docMeta.pages)
	fmt.Println("Characters:", docMeta.characters)
	fmt.Println("Lines:", docMeta.lines)
	fmt.Println("Paragraphs:", docMeta.paragraphs)
	fmt.Println("Slides:", docMeta.slides)
	fmt.Println("Notes:", docMeta.notes)
	fmt.Println("HiddenSlides:", docMeta.hiddenSlides)
	fmt.Println("MMClips:", docMeta.mmClips)
	fmt.Println("Template:", docMeta.template)
	fmt.Println("PresentationFormat:", docMeta.presentationFormat)
	fmt.Println("LinksUpToDate:", docMeta.linksUpToDate)
	fmt.Println("CharactersWithSpaces:", docMeta.charactersWithSpaces)
	fmt.Println("SharedDoc:", docMeta.sharedDoc)
	fmt.Println("HyperlinksChanged:", docMeta.hyperlinksChanged)
	fmt.Println("DocSecurity:", docMeta.docSecurity)
	fmt.Println("ScaleCrop:", docMeta.scaleCrop)
	fmt.Println("LastModifiedBy:", docMeta.lastModifiedBy)
	fmt.Println("Revision:", docMeta.revision)
	fmt.Println("LastPrinted:", docMeta.lastPrinted)
	fmt.Println("Title:", docMeta.title)
	fmt.Println("Creator:", docMeta.creator)
	fmt.Println("Created:", docMeta.created)
	fmt.Println("Modified:", docMeta.modified)
}
func printDoc(docMeta oxmlDocsMeta) {
	fmt.Println()
	fmt.Println("Number:", docMeta.number)
	fmt.Println("FileName:", docMeta.fileName)
	fmt.Println("FilePath:", docMeta.filePath)
	fmt.Println("FileCreated:", docMeta.fileCreationTime)
	fmt.Println("FileModified:", docMeta.fileWriteTime)
	fmt.Println("FileAccessed:", docMeta.fileAccessTime)
	fmt.Println("FileSize:", docMeta.fileSize, "bytes")
	fmt.Println("FileExt:", docMeta.fileExt)
	fmt.Println("Application:", docMeta.application)
	fmt.Println("CodePage:", docMeta.codePage)
	fmt.Println("WinVersion:", docMeta.winVersion)
	fmt.Println("Title:", docMeta.title)
	fmt.Println("Subject:", docMeta.subject)
	fmt.Println("Creator:", docMeta.creator)
	fmt.Println("Keywords:", docMeta.keywords)
	fmt.Println("Comments:", docMeta.comments)
	fmt.Println("Template:", docMeta.template)
	fmt.Println("LastSavedBy:", docMeta.lastModifiedBy)
	fmt.Println("Company:", docMeta.company)
	fmt.Println("Revision:", docMeta.revision)
	fmt.Println("TotalTime:", docMeta.totalTime)
	fmt.Println("LastPrinted:", docMeta.lastPrinted)
	fmt.Println("Created:", docMeta.created)
	fmt.Println("LastSaved:", docMeta.modified)
	fmt.Println("Pages:", docMeta.pages)
	fmt.Println("Words:", docMeta.words)
	fmt.Println("Characters:", docMeta.characters)
	fmt.Println("Lines:", docMeta.lines)
	fmt.Println("Paragraphs:", docMeta.paragraphs)
	fmt.Println("DocSecurity:", docMeta.docSecurity)
	fmt.Println("PresentationTarget:", docMeta.presentationFormat)
	fmt.Println("Manager:", docMeta.manager)
	fmt.Println("Slides:", docMeta.slides)
	fmt.Println("Notes:", docMeta.notes)
	fmt.Println("HiddenSlides:", docMeta.hiddenSlides)
	fmt.Println("MMClips:", docMeta.mmClips)
}
func printPdf(docMeta pdfFileMeta) {
	fmt.Println()
	fmt.Println("Number:", docMeta.number)
	fmt.Println("FileName:", docMeta.fileName)
	fmt.Println("FilePath:", docMeta.filePath)
	fmt.Println("FileCreated:", docMeta.fileCreationTime)
	fmt.Println("FileModified:", docMeta.fileWriteTime)
	fmt.Println("FileAccessed:", docMeta.fileAccessTime)
	fmt.Println("FileSize:", docMeta.fileSize, "bytes")
	fmt.Println("FileExt:", docMeta.fileExt)
	fmt.Println("AppVersion:", docMeta.appVersion)
	fmt.Println("Title:", docMeta.title)
	fmt.Println("Title2:", docMeta.title2)
	fmt.Println("Title3:", docMeta.title3)
	fmt.Println("Subject:", docMeta.subject)
	fmt.Println("Subject2:", docMeta.subject2)
	fmt.Println("Pages:", docMeta.pages)
	fmt.Println("Creator:", docMeta.creator)
	fmt.Println("Creator2:", docMeta.creator2)
	fmt.Println("Creator3:", docMeta.creator3)
	fmt.Println("Author:", docMeta.author)
	fmt.Println("Author2:", docMeta.author2)
	fmt.Println("Created:", docMeta.created)
	fmt.Println("CreationDate:", docMeta.createDate)
	fmt.Println("CreationDate2:", docMeta.createDate2)
	fmt.Println("CreationDate3:", docMeta.createDate3)
	fmt.Println("CreationDate4:", docMeta.createDate4)
	fmt.Println("ModifyDate:", docMeta.modifyDate)
	fmt.Println("ModifyDate2:", docMeta.modifyDate2)
	fmt.Println("SourceModified:", docMeta.sourceModified)
	fmt.Println("ModDate:", docMeta.modDate)
	fmt.Println("ModDate2:", docMeta.modDate2)
	fmt.Println("LastSaved:", docMeta.lastSaved)
	fmt.Println("MetaDataDate:", docMeta.metaDataDate)
	fmt.Println("MetaDataDate2:", docMeta.metaDataDate2)
	fmt.Println("Description:", docMeta.description)
	fmt.Println("Rights:", docMeta.rights)
	fmt.Println("Encrypted:", docMeta.encrypt)
	fmt.Println("Producer:", docMeta.producer)
	fmt.Println("Producer2:", docMeta.producer2)
	fmt.Println("Producer3:", docMeta.producer3)
	fmt.Println("CreatorTool:", docMeta.creatorTool)
	fmt.Println("CreatorTool2:", docMeta.creatorTool2)
	fmt.Println("SoftwareAgent:", docMeta.softwareAgent)
	fmt.Println("SoftwareAgent2:", docMeta.softwareAgent2)
	fmt.Println("DocumentID:", docMeta.documentID)
	fmt.Println("InstanceID:", docMeta.instanceID)
	fmt.Println("OriginalDocumentID:", docMeta.originalDocumentID)
	fmt.Println("Keywords:", docMeta.Keywords)
	fmt.Println("Keywords2:", docMeta.Keywords2)
	fmt.Println("Trapped:", docMeta.trapped)
	fmt.Println("DocChangeCount:", docMeta.docChangeCount)
	fmt.Println("Parameters:", docMeta.parameters)
	fmt.Println("Format:", docMeta.format)
	fmt.Println("ResourcePath:", docMeta.resourcePath)
	fmt.Println("Publisher:", docMeta.publisher)
	fmt.Println("AggregationType:", docMeta.aggregationType)
	fmt.Println("PublicationName:", docMeta.publicationName)
	fmt.Println("Edition:", docMeta.edition)
	fmt.Println("Copyright:", docMeta.copyright)
	fmt.Println("ISBN:", docMeta.isbn)
	fmt.Println("CoverDisplayDate:", docMeta.coverDisplayDate)
	fmt.Println("Language:", docMeta.lang)
}

var (
	oxmlApp                  string
	oxmlAppVersion           string
	oxmlCompany              string
	oxmlTotalTime            string
	oxmlWords                string
	oxmlPages                string
	oxmlCharacters           string
	oxmlLines                string
	oxmlPrphs                string
	oxmlSlides               string
	oxmlNotes                string
	oxmlHiddenSlides         string
	oxmlMmClips              string
	oxmlTemplate             string
	oxmlPresentationFormat   string
	oxmlLinksUpToDate        string
	oxmlCharactersWithSpaces string
	oxmlSharedDoc            string
	oxmlHyperlinksChanged    string
	oxmlDocSecurity          string
	oxmlScaleCrop            string
	oxmlTitle                string
	oxmlCreator              string
	oxmlLastModifiedBy       string
	oxmlRevision             string
	oxmlLastPrinted          string
	oxmlCreated              string
	oxmlModified             string
	fileApp                  = "docProps/app.xml"
	fileCore                 = "docProps/core.xml"
)
var (
	pdfApp                string
	pdfTitle              string
	pdfSubject            string
	pdfCreator            string
	pdfPublisher          string
	pdfDescription        string
	pdfRights             string
	pdfProducer           string
	pdfCreatorTool        string
	pdfKeywords           string
	pdfPages              string
	pdfCreateDate         string
	pdfModifyDate         string
	pdfMetadataDate       string
	pdfDocumentID         string
	pdfInstanceID         string
	pdfTrapped            string
	pdfAggregationType    string
	pdfPublicationName    string
	pdfEdition            string
	pdfCopyRight          string
	pdfISBN               string
	pdfCoverDisplayDate   string
	pdfOriginalDocumentID string
	pdfTitle2             string
	pdfTitle3             string
	pdfAuthor             string
	pdfAuthor2            string
	pdfCreateDate2        string
	pdfCreateDate3        string
	pdfCreator2           string
	pdfCreator3           string
	pdfKeywords2          string
	pdfModDate            string
	pdfModDate2           string
	pdfProducer2          string
	pdfProducer3          string
	pdfLang               string
	pdfEncrypt            string
	pdfSubject2           string
	pdfCreateDate4        string
	pdfModifyDate2        string
	pdfMetadataDate2      string
	pdfCreatorTool2       string
	pdfSourceModified     string
	pdfCreated            string
	pdfLastSaved          string
	pdfFormat             string
	pdfSoftwareAgent      string
	pdfSoftwareAgent2     string
	pdfParameters         string
	pdfResourcePath       string
	pdfDocChangeCount     string
)

var (
	errHelp  = errors.New("File or Folder Not Found. Use '-help' option")
	errFile  = errors.New("File Not Found. Use '-help' option")
	errDir   = errors.New("Folder Not Found. Use '-help' option")
	errExist = errors.New("No Supported Files Exist. Use '-help' option")
	cnt      = 1
	oxmlDocs []oxmlDocsMeta
	pdfDocs  []pdfFileMeta
	docs     []oxmlDocsMeta
	st, et   time.Time
	dt       string

	usageFile = "Parse a file\nSupported pptx, docx, xlsx files\nDocsMetaParser.exe -f C:\\filename.docx"
	usageDir  = "Parse files in directory\nSupported pptx, docx, xlsx files\nDocsMetaParser.exe -d C:\\Folder"
	usageRcv  = "Parse files recursively in directory\nSupported pptx, docx, xlsx files\nDocsMetaParser.exe -r C:\\"
	usageCsv  = "Save as CSV file format\nDocsMetaParser.exe -csv -r C:\\"
)

type oxmlDocsMeta struct {
	application          string
	appVersion           string
	winVersion           string
	company              string
	fileName             string
	fileCreationTime     string
	fileAccessTime       string
	fileWriteTime        string
	fileSize             string
	fileExt              string
	filePath             string
	title                string
	codePage             string
	subject              string
	creator              string
	keywords             string
	comments             string
	lastModifiedBy       string
	revision             string
	created              string
	modified             string
	totalTime            string
	pages                string
	words                string
	characters           string
	lines                string
	paragraphs           string
	template             string
	charactersWithSpaces string
	lastPrinted          string
	slides               string
	notes                string
	hiddenSlides         string
	presentationFormat   string
	mmClips              string
	sharedDoc            string
	hyperlinksChanged    string
	docSecurity          string
	scaleCrop            string
	linksUpToDate        string
	category             string
	manager              string
	number               string
}
type pdfFileMeta struct {
	application        string
	appVersion         string
	fileName           string
	fileCreationTime   string
	fileAccessTime     string
	fileWriteTime      string
	fileSize           string
	fileExt            string
	filePath           string
	title              string
	title2             string
	title3             string
	subject            string
	subject2           string
	creator            string
	creator2           string
	creator3           string
	author             string
	author2            string
	description        string
	rights             string
	producer           string
	producer2          string
	producer3          string
	creatorTool        string
	creatorTool2       string
	softwareAgent      string
	softwareAgent2     string
	Keywords           string
	Keywords2          string
	pages              string
	created            string
	createDate         string
	createDate2        string
	createDate3        string
	createDate4        string
	modifyDate         string
	modifyDate2        string
	sourceModified     string
	modDate            string
	modDate2           string
	lastSaved          string
	metaDataDate       string
	metaDataDate2      string
	instanceID         string
	documentID         string
	docChangeCount     string
	parameters         string
	format             string
	resourcePath       string
	publisher          string
	aggregationType    string
	publicationName    string
	edition            string
	copyright          string
	isbn               string
	coverDisplayDate   string
	lang               string
	originalDocumentID string
	trapped            string
	encrypt            string
	number             string
}
