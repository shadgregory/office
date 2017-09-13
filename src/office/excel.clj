(ns office.excel
  (:require [ring.util.response :as r])
  (:import
   (java.awt Color)
   (java.io ByteArrayInputStream ByteArrayOutputStream)
   (org.apache.poi.ss.util CellRangeAddress)
   (org.apache.poi.ss.usermodel CellStyle
                                FillPatternType
                                IndexedColors)
   (org.apache.poi.xssf.usermodel XSSFWorkbook
                                  XSSFSheet
                                  XSSFFont
                                  XSSFColor
                                  TextAlign
                                  XSSFRow)))

(defn td? [element]
  (if (nil? (re-find #"^td" (name element)))
    false
    true))

(defn th? [element]
  (if (nil? (re-find #"^th" (name element)))
    false
    true))

(defn tr? [element]
  (if (nil? (re-find #"^tr" (name element)))
    false
    true))

(defn set-cell-bg [cell style bg]
  (try
    (.setFillBackgroundColor style (.getIndex (IndexedColors/valueOf (.toUpperCase (name bg)))))
    (catch IllegalArgumentException e
      ;;not an indexed color, let's assume it's hex
      (.setFillBackgroundColor style (new XSSFColor (Color/decode bg)))))
  (.setFillPattern style CellStyle/LEAST_DOTS)
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

(defn process-cell-config [wb spreadsheet config cells row]
  (if (not (nil? (:colspan config))) (let [cell (.createCell row 0)
                                           font (.createFont wb)
                                           style (.createCellStyle wb)]
                                       ;; defaulting to bold & centered for now
                                       (.setCellValue cell (first cells))
                                       (.setBold font true)
                                       (.setFont style font)
                                       (.setAlignment style CellStyle/ALIGN_CENTER)
                                       (.setCellStyle cell style)
                                       (.addMergedRegion spreadsheet (new CellRangeAddress
                                                                          (.getRowNum row)
                                                                          (.getRowNum row)
                                                                          0
                                                                          (dec (Integer. (:colspan config))))))))

(defn process-cell [wb spreadsheet row sexp num & bg]
  (let [cell (.createCell row num)]
    (cond
      (= "italic" (:font-style (second sexp)))
      (let [font (.createFont wb)
            style (.createCellStyle wb)]
        (.setItalic font true)
        (.setFont style font)
        (.setCellStyle cell style)
        (.setCellValue cell (nth sexp 2))
        (if (not (nil? bg))
          (set-cell-bg cell style (first bg))))
      (= "bold" (:font-weight (second sexp)))
      (let [font (.createFont wb)
            style (.createCellStyle wb)]
        (.setBold font true)
        (.setFont style font)
        (.setCellStyle cell style)
        (.setCellValue cell (nth sexp 2))
        (if (not (nil? bg))
          (set-cell-bg cell style (first bg))))
      (map? (second sexp)) (process-cell-config wb spreadsheet (second sexp) (rest (rest sexp)) row)
      :else
      (do
        (.setCellValue cell (str (second sexp)))
        (if (not (nil? bg))
          (let [style (.createCellStyle wb)]
            (set-cell-bg cell style (first bg))))))))

(defn process-row-config [wb spreadsheet config cells row row-num]
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
      (map? (second sexp)) (process-row-config wb spreadsheet (second sexp) (rest (rest sexp)) row num)
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
  (if (not (string? (second sexp)))
    (throw (Exception. "Worksheet title is required.")))
  (let [spreadsheet (.createSheet wb (second sexp))]
    (loop [rows (rest (rest sexp))
           rowid 0]
      (cond
        (empty? rows) wb
        (tr? (ffirst rows))(do
                             (process-row wb spreadsheet rowid (first rows))
                             (recur (rest rows) (inc rowid)))
        :else
        (throw (Exception. (str "Don't know what to do with " (ffirst rows))))))
    (loop [index (count (rest (rest sexp)))]
      (cond
        (zero? index) nil
        :else (do
                (.autoSizeColumn spreadsheet (short index))
                (recur (dec index)))))))

(defn excel [sexp]
  (let [wb (new XSSFWorkbook)]
    (loop [sexp sexp]
      (cond
        (empty? sexp) wb
        (= :wb (first sexp)) (recur (rest sexp))
        (= :spreadsheet (ffirst sexp)) (do (process-spreadsheet wb (first sexp))
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
