(ns clj-xlstream.core
  (:use
   [clojure.java.io :only [input-stream]])
  (:import
   (org.xml.sax InputSource)
   (org.xml.sax.helpers XMLReaderFactory)
   (org.apache.poi POIXMLDocument)
   (org.apache.poi.openxml4j.opc OPCPackage)
   (org.apache.poi.poifs.filesystem POIFSFileSystem)
   (org.apache.poi.hssf.eventusermodel HSSFEventFactory
                                       HSSFRequest)
   (org.apache.poi.xssf.eventusermodel ReadOnlySharedStringsTable
                                       XSSFReader
                                       XSSFSheetXMLHandler
                                       XSSFSheetXMLHandler$SheetContentsHandler)))

(defn file-type
  [istream opt]
  (when (.markSupported istream)
    (cond 
     (POIFSFileSystem/hasPOIFSHeader istream) :hssf
     (POIXMLDocument/hasOOXMLHeader istream)  :xssf)))

(defmulti read-xls
  "Reads Excel file"
  file-type)

;; binary Excel files (.xls)
(defmethod read-xls :hssf
  [istream handler]
  (let [events (new HSSFEventFactory)
        request (new HSSFRequest)]
    (println "not implemented yet")))

;; OOXML Excel files (.xlsx)
(defmethod read-xls :xssf
  [istream handler]
  (let [package (OPCPackage/open istream)
        reader (new XSSFReader package)
        styles (.getStylesTable reader)
        strings (new ReadOnlySharedStringsTable package)
        sheethandler (new XSSFSheetXMLHandler styles strings handler true)]
    (doseq [sheet (iterator-seq (.getSheetsData reader))]
      (doto (XMLReaderFactory/createXMLReader)
        (.setContentHandler sheethandler)
        (.parse (new InputSource sheet))))))

(def myhandler 
  (proxy [XSSFSheetXMLHandler$SheetContentsHandler] []
    (cell
      [#^String cellReference #^String formattedValue]
      (println "#CellValue: " formattedValue))
    (endRow
      []
      (println "#EndRow"))
    (headerFooter
      [#^String text isHeader #^String tagName]
      (println text tagName))
    (startRow
      [rowNum]
      (println "#RowNum: " rowNum))))

(defn -main
  [fname]
  (with-open [istream (input-stream fname)]
    (read-xls istream myhandler)))
