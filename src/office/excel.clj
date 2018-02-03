(ns office.excel
  (:require [ring.util.response :as r])
  (:import
   (java.awt Color)
   (java.io ByteArrayInputStream ByteArrayOutputStream)
   (org.apache.poi.ss.util CellRangeAddress)
   (org.apache.poi.common.usermodel HyperlinkType)
   (org.apache.poi.ss.usermodel CellStyle
                                CreationHelper
                                FillPatternType
                                HorizontalAlignment
                                IndexedColors)
   (org.apache.poi.hssf.usermodel HSSFCell)
   (org.apache.poi.xssf.usermodel XSSFWorkbook
                                  XSSFSheet
                                  XSSFFont
                                  XSSFColor
                                  XSSFHyperlink
                                  TextAlign
                                  XSSFRow)))

;; We want to ignore hiccup style classes and ids that might follow the tag
(defn td? [element]
  (cond
    (= "sum" (name element)) true
    (not (nil? (re-find #"^td" (name element)))) true
    :else false))

(defn th? [element]
  (if (nil? (re-find #"^th" (name element)))
    false
    true))

(defn tr? [element]
  (cond
    (not (keyword? element)) false
    (nil? (re-find #"^tr" (name element))) false
    :else true))

(defn table? [element]
  (cond
    (not (keyword? element)) false
    (= :spreadsheet element) true
    (nil? (re-find #"^table" (name element))) false
    :else true))
;;;;;;

;; get thead's tr
(defn get-thead-row [sexp]
  (loop [sexp sexp]
    (cond
      (empty? sexp) nil
      (map? (first sexp)) (recur (rest sexp))
      (= :wb (first sexp)) (recur (first (rest sexp)))
      (table? (first sexp)) (recur (rest sexp))
      (= :thead (first sexp)) sexp
      (= :thead (ffirst sexp)) (first sexp)
      (tr? (ffirst sexp)) (recur (first (rest sexp))))))

;;how many cells in the thead?
(defn column-count [sexp]
  (let [thead (get-thead-row sexp)]
    (loop [thead (rest thead) c 0]
      (cond
        (empty? thead) c
        (tr? (ffirst thead)) (recur (rest (first thead)) c)
        (map? (first thead)) (recur (rest thead) c)
        (th? (ffirst thead)) (recur (rest thead) (inc c))
        (td? (ffirst thead)) (recur (rest thead) (inc c))
        :else
        (recur (rest thead) c)))))

(defn set-cell-bg [cell style bg]
  (try
    (.setFillBackgroundColor style (.getIndex (IndexedColors/valueOf (.toUpperCase (name bg)))))
    (catch IllegalArgumentException e
      ;;not an indexed color, let's assume it's hex
      (.setFillForegroundColor style (new XSSFColor (Color/decode bg)))))
  (.setFillPattern style FillPatternType/SOLID_FOREGROUND)
  (.setCellStyle cell style))

(defn process-header-cell [wb row sexp num & bg]
  (let [cell (.createCell row num)
        font (.createFont wb)
        style (.createCellStyle wb)]
    (.setBold font true)
    (.setFont style font)
    (.setCellStyle cell style)
    (.setCellValue cell (second sexp))
    (if (not (nil? bg))
      (set-cell-bg cell style (first bg)))))

(defn process-cell-config [wb spreadsheet config cell row]
  (let [font (.createFont wb)
        style (.createCellStyle wb)]
    (cond
      (contains? config :colspan) (do
                                    ;; defaulting to bold & centered for now
                                    (.setBold font true)
                                    (.setFont style font)
                                    (.setAlignment style HorizontalAlignment/CENTER)
                                    (.setCellStyle cell style)
                                    (.addMergedRegion spreadsheet (new CellRangeAddress
                                                                       (.getRowNum row)
                                                                       (.getRowNum row)
                                                                       0
                                                                       (dec (Integer. (:colspan config))))))
      (contains? config :font-style) (do
                                       (cond
                                         (= "italic" (:font-style config)) (do
                                                                             (.setItalic font true)
                                                                             (.setFont style font)
                                                                             (.setCellStyle cell style))))
      (contains? config :font-weight) (do
                                        (cond
                                          (= "bold" (:font-weight config)) (do
                                                                             (.setBold font true)
                                                                             (.setFont style font)
                                                                             (.setCellStyle cell style)))))))

(defn process-sum [cell sexp]
  (let [formula (second sexp)]
    (.setCellType cell HSSFCell/CELL_TYPE_FORMULA)
    (.setCellFormula cell (str "SUM(" formula ")"))))

(defn process-cell [wb spreadsheet row sexp num & bg]
  (let [cell (.createCell row num)]
    (loop [sexp sexp]
      (cond
        (empty? sexp) spreadsheet
        (= :td (first sexp)) (recur (rest sexp))
        (= :sum (first sexp)) (do
                                (process-sum cell sexp)
                                (recur (rest sexp)))
        (map? (first sexp)) (do
                              (process-cell-config wb spreadsheet (first sexp) cell row)
                              (recur (rest sexp)))
        (vector? (first sexp)) (do
                                 (cond
                                   (= :a (ffirst sexp)) (let [url (:href (second (first sexp)))
                                                              text (nth (first sexp) 2)
                                                              create-helper (.getCreationHelper wb)
                                                              link (.createHyperlink create-helper HyperlinkType/URL)]
                                                          (.setCellValue cell text)
                                                          (.setAddress link url)
                                                          (.setHyperlink cell link)))
                                 (recur (rest sexp)))
        (string? (first sexp)) (do
                                 (if (nil? (re-find #"^[-+]?([0-9]*\.[0-9]+|[0-9]+)" (first sexp)))
                                   (.setCellValue cell (first sexp))
                                   (.setCellValue cell (Double. (first sexp))))
                                 (if (not (nil? bg))
                                   (let [style (.createCellStyle wb)]
                                     (set-cell-bg cell style (first bg))))
                                 (recur (rest sexp)))))))

(defn process-row-config [wb spreadsheet config cells row]
  (if (not (nil? (:background-color config))) (loop [cells cells num 0]
                                                (cond
                                                  (empty? cells) spreadsheet
                                                  (td? (ffirst cells)) (do (process-cell wb spreadsheet row (first cells) num (:background-color config))
                                                                           (recur (rest cells) (inc num)))
                                                  (th? (ffirst cells)) (do (process-header-cell wb row (first cells) num (:background-color config))
                                                                           (recur (rest cells) (inc num)))
                                                  :else
                                                  (throw (Exception. (str "Don't know what to do with " (first cells))))))))

(defn process-row [wb spreadsheet num sexp]
  (let [row (.createRow spreadsheet num)]
    (cond
      (map? (second sexp)) (process-row-config wb spreadsheet (second sexp) (rest (rest sexp)) row)
      (not (nil? (:background-color (second sexp))))
      (let [style (.createCellStyle wb)]
        (loop [cells (rest (rest sexp)) num 0]
          (cond
            (empty? cells) spreadsheet
            (td? (ffirst cells)) (do (process-cell wb spreadsheet row (first cells) num (:background-color (second sexp)))
                                     (recur (rest cells) (inc num)))
            (th? (ffirst cells)) (do (process-header-cell wb row (first cells) num (:background-color (second sexp)))
                                     (recur (rest cells) (inc num)))
            :else
            (throw (Exception. (str "Don't know what to do with " (first cells)))))) )
      :else
      (loop [cells (rest sexp) num 0]
        (cond
          (empty? cells) spreadsheet
          (td? (ffirst cells)) (do (process-cell wb spreadsheet row (first cells) num)
                                   (recur (rest cells) (inc num)))
          (th? (ffirst cells)) (do (process-header-cell wb row (first cells) num)
                                   (recur (rest cells) (inc num)))
          :else
          (throw (Exception. (str "Don't know what to do with " (first cells)))))))))

(defn process-spreadsheet [wb sexp]
  (if (not (map? (second sexp)))
    (throw (Exception. "Worksheet title is required.")))
  (let [spreadsheet (.createSheet wb (:title (second sexp)))]
    (loop [rows (rest (rest sexp))
           rowid 0]
      (cond
        (empty? rows) wb
        (tr? (ffirst rows)) (do
                              (process-row wb spreadsheet rowid (first rows))
                              (recur (rest rows) (inc rowid)))
        (= :thead (ffirst rows)) (do
                                   (process-row wb spreadsheet rowid (first (rest (first rows))))
                                   (recur (rest rows) (inc rowid)))
        (= :tbody (ffirst rows)) (do
                                   (recur (rest rows)
                                          (loop [body-rows (rest (first rows)) i rowid]
                                            (cond
                                              (empty? body-rows) i
                                              :else (do
                                                      (process-row wb spreadsheet i (first body-rows))
                                                      (recur (rest body-rows) (inc i)))))))
        (= :tfoot (ffirst rows)) (do
                                   (process-row wb spreadsheet rowid (first (rest (first rows))))
                                   (recur (rest rows) (inc rowid)))
        :else
        (throw (Exception. (str "Don't know what to do with " (ffirst rows))))))
    (loop [i 1]
      (cond
        (= i (inc (column-count sexp))) nil
        :else
        (do
          (.autoSizeColumn spreadsheet i)
          (recur (inc i)))))))

(defn excel [sexp]
  (let [wb (new XSSFWorkbook)]
    (loop [sexp sexp]
      (cond
        (empty? sexp) wb
        (= :wb (first sexp)) (recur (rest sexp))
        (or
         (table? (ffirst sexp))
         (= :spreadsheet (ffirst sexp)))(do (process-spreadsheet wb (first sexp))
                                            (recur (rest sexp)))
        :else
        (throw (Exception. (str "Syntax Error. Don't know what to do with " (first sexp))))))))

(defn excel-page [title excel-sexp]
  (let [stream (ByteArrayOutputStream.)]
    (.write (excel excel-sexp)
            stream)
    (-> (r/response (ByteArrayInputStream. (.toByteArray stream)))
        (r/header "Content-Type" "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;")
        (r/header "Content-Length" (.size stream))
        (r/header "Content-Disposition" (str "attachment; filename=" (clojure.string/replace title #"[ ]+" "_") ".xlsx")))))
