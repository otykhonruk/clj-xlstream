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
                                       HSSFListener
                                       HSSFRequest)
   (org.apache.poi.hssf.record Record)
   (org.apache.poi.xssf.eventusermodel ReadOnlySharedStringsTable
                                       XSSFReader
                                       XSSFSheetXMLHandler
                                       XSSFSheetXMLHandler$SheetContentsHandler)))

(defn- file-type
  [istream opt]
  (when (.markSupported istream)
    (cond 
     (POIFSFileSystem/hasPOIFSHeader istream) :hssf
     (POIXMLDocument/hasOOXMLHeader istream)  :xssf)))


(defn- make-hssf-listener
  [handler]
  (reify 
    HSSFListener
    (processRecord
      [_ record]
      (println record))))


(defmulti read-xls
  "Reads Excel file"
  file-type)


;; binary Excel files (.xls)
(defmethod read-xls :hssf
  [istream handler]
  (let [events (new HSSFEventFactory)
        request (new HSSFRequest)]
    (.addListenerForAllRecords request (make-hssf-listener handler))
    (.processEvents events request istream)))


;; OOXML Excel files (.xlsx)
(defmethod read-xls :xssf
  [istream handler]
  (let [package (OPCPackage/open istream)
        reader (new XSSFReader package)
        styles (.getStylesTable reader)
        strings (new ReadOnlySharedStringsTable package)
        sheethandler (new XSSFSheetXMLHandler styles strings handler true)
        sheets (.getSheetsData reader)]
    (doseq [sheet (iterator-seq sheets)]
      (.startSheet handler (.getSheetName sheets))
      (doto (XMLReaderFactory/createXMLReader)
        (.setContentHandler sheethandler)
        (.parse (new InputSource sheet)))
      (.endSheet handler))))


(defprotocol SheetListener
  (startSheet [this name])
  (endSheet [this]))


(def myhandler
  (reify
    XSSFSheetXMLHandler$SheetContentsHandler
    (cell
      [_ cellReference formattedValue]
      (println "#CellValue: " formattedValue))
    (endRow
      [_]
      (println "#EndRow"))
    (headerFooter
      [_ text isHeader tagName]
      (println text tagName))
    (startRow
      [_ rowNum]
      (println "#RowNum: " rowNum))
    SheetListener
    (startSheet
      [_ name]
      (println "#Sheet: " name))
    (endSheet
      [_]
      (println "#EndSheet"))))


(defn -main
  [fname]
  (with-open [istream (input-stream fname)]
    (read-xls istream myhandler)))
