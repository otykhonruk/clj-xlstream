(ns clj-xlstream.core
  (:require 
   [clojure.java.jdbc :as jdbc])
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
      (println "#Cell" cellReference ": " formattedValue))
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


;; maps xls columns to db columns
(def col-mapping 
  {"A" "count"
   "B" "onpp" 
   "C" "author"
   "D" "title"
   "E" "std"
   "F" "price"
   "G" "publisher"
   "H" "translate"
   "I" "year"
   "J" "genre"
   "K" "series"
   "L" "format"
   "M" "isbn"
   "N" "pages"
   "O" "article"
   "P" "pricetax"
   "Q" "ean"
   "R" "category"
   "S" "header"})

(defn col-name [cellref]
  (keyword (col-mapping (str (first cellref)))))


(def ^:dynamic *row*)
(def ^:dynamic *rownum*)

(def db-loader
  (reify
    XSSFSheetXMLHandler$SheetContentsHandler
    (cell
      [_ cellref val]
      (set! *row* (assoc *row* (col-name cellref) val)))
    (startRow [_ rownum]
      (set! *rownum* rownum))
    (endRow
      [_]
      (when (pos? *rownum*)
        (do
          (jdbc/insert-record :prices *row* :transactions? false)
          ;; (println *row*)
          (set! *row* {}))))
    (headerFooter [_ text isHeader tagName])
    SheetListener
    (startSheet [_ name])
    (endSheet  [_])))


(defn load-from-file
  [fname dbspec]
  (with-open [istream (input-stream fname)]
    (binding [*row* {}
              *rownum* 0]
      (jdbc/with-connection dbspec
        (read-xls istream db-loader)))))

;; test db
;; 
;; (def test-db 
;;   {:subprotocol "sqlite"
;;    :subname "rs.db"
;;    :username ""
;;    :password ""})

;; test table
;; 
;; (defn create-prices
;;   [dbspec]
;;   (jdbc/with-connection dbspec
;;     (apply 
;;      jdbc/create-table
;;      :prices
;;      (for [c (vals col-mapping)] [(keyword c) "TEXT"]))))

(defn -main
  [fname]
  (load-from-file fname test-db))
