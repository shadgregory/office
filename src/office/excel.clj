(ns office.excel
  (:import
   (org.apache.poi.xssf.usermodel
            XSSFWorkbook
            XSSFSheet
            XSSFFont
            XSSFRow)))

(defn process-header-cell [wb row sexp num]
  (let [cell (.createCell row num)
        font (.createFont wb)
        style (.createCellStyle wb)]
    (.setBold font true)
    (.setFont style font)
    (.setCellStyle cell style)
    (.setCellValue cell (second sexp))))

(defn process-cell [wb row sexp num]
  (let [cell (.createCell row num)]
    (cond
      (= "italic" (:font-style (second sexp)))
      (let [font (.createFont wb)
            style (.createCellStyle wb)]
        (.setItalic font true)
        (.setFont style font)
        (.setCellStyle cell style)
        (.setCellValue cell (nth sexp 2)))
      (= "bold" (:font-weight (second sexp)))
      (let [font (.createFont wb)
            style (.createCellStyle wb)]
        (.setBold font true)
        (.setFont style font)
        (.setCellStyle cell style)
        (.setCellValue cell (nth sexp 2)))
      :else
      (.setCellValue cell (second sexp)))))

(defn process-row [wb spreadsheet num sexp]
  (let [row (.createRow spreadsheet num)]
    (cond
      (not (nil? (:background-color (second sexp))))
      (let [style (.getRowStyle row)]
        (.setFillBackgroundColor style )
        )
      :else
      (loop [cells (rest sexp) num 0]
        (cond
          (empty? cells) spreadsheet
          (= :cell (first (first cells))) (do (process-cell wb row (first cells) num)
                                              (recur (rest cells) (inc num)))
          (= :th (first (first cells))) (do (process-header-cell wb row (first cells) num)
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
        (= :row (first (first rows)))(do
                                       (process-row wb spreadsheet rowid (first rows))
                                       (recur (rest rows) (inc rowid)))
        :else
        (throw (Exception. (str "Don't know what to do with " (first (first rows)))))))
    (loop [index (count (rest (rest sexp)))]
      (cond
        (= index 0) nil
        :else (do
                (.autoSizeColumn spreadsheet (short index))
                (recur (dec index)))))))

(defn excel [sexp]
  (let [wb (new XSSFWorkbook)]
    (loop [sexp sexp]
      (cond
        (empty? sexp) wb
        (= :wb (first sexp)) (recur (rest sexp))
        (= :spreadsheet (first (first sexp))) (do (process-spreadsheet wb (first sexp))
                                                  (recur (rest sexp)))
        :else
        (throw (Exception. (str "Syntax Error. Don't know what to do with " (first sexp))))))))
