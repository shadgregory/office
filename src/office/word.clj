(ns office.word
  (:require [clojure.pprint])
  (:import (java.io File FileOutputStream FileInputStream)
           (org.apache.poi.util Units)
           (org.apache.poi.xwpf.usermodel XWPFDocument ParagraphAlignment Borders BreakType UnderlinePatterns))
  (:gen-class))

(defn align-code [code]
  (cond
    (= code "center") ParagraphAlignment/CENTER
    (= code "right") ParagraphAlignment/RIGHT
    (= code "left") ParagraphAlignment/LEFT
    (= code "both") ParagraphAlignment/BOTH))

(defn border-code [code]
  (cond
    (= code "double") Borders/DOUBLE
    (= code "single") Borders/SINGLE))

(defn font-config [config run]
  (if (not (nil? config))
    (when-not (nil? config)
      (if (contains? config :color) (.setColor run (get config :color)))
      (if (contains? config :text-position)
        (.setTextPosition run (get config :text-position)))
      (if (contains? config :font-family)
        (.setFontFamily run (get config :font-family)))
      (if (contains? config :font-size)
        (.setFontSize run (get config :font-size))))))

(defn paragraph [doc exp]
  (let [par (.createParagraph doc)
        run (.createRun par)]
    (loop [body (rest exp)
           run run]
      (let [current (first body)]
        (cond
          (nil? current) nil
          (vector? current)  (cond
                               (= :i (first current)) (do
                                                        (.setItalic run true)
                                                        (if (vector? (second current))
                                                          (recur (cons (first (subvec current 1)) (rest body)) run)
                                                          (recur (cons (subvec current 1) (rest body)) run)))
                               (= :u (first current)) (do
                                                        (.setUnderline run UnderlinePatterns/SINGLE)
                                                        (if (vector? (second current))
                                                          (recur (cons (first (subvec current 1)) (rest body)) run)
                                                          (recur (cons (subvec current 1) (rest body)) run)))
                               (= :strike (first current)) (do
                                                             (.setStrike run true)
                                                             (if (vector? (second current))
                                                               (recur (cons (first (subvec current 1)) (rest body)) run)
                                                               (recur (cons (subvec current 1) (rest body)) run)))
                               (= :b (first current)) (do
                                                        (.setBold run true)
                                                        (if (vector? (second current))
                                                          (recur (cons (first (subvec current 1)) (rest body)) run)
                                                          (recur (cons (subvec current 1) (rest body)) run)))
                               (= :run (first current)) (recur (cons (subvec current 1) (rest body)) (.createRun par))
                               (= :br (first current)) (do
                                                         (.addBreak run)
                                                         (recur (rest body) (.createRun par)))
                               (= :img (first current)) (cond
                                                          (string? (second current)) (let [text (second current)]
                                                                                       (.addBreak run)
                                                                                       (.addPicture run (new FileInputStream text)
                                                                                                    XWPFDocument/PICTURE_TYPE_JPEG
                                                                                                    text
                                                                                                    (Units/toEMU 200)
                                                                                                    (Units/toEMU 200))
                                                                                       (recur (cons (subvec current 2) (rest body)) run))
                                                          (map? (second current)) (let [config (second current)
                                                                                        file-name (nth current 2)]
                                                                                    (.addBreak run)
                                                                                    (.addPicture run (new FileInputStream file-name)
                                                                                                 XWPFDocument/PICTURE_TYPE_JPEG
                                                                                                 file-name
                                                                                                 (Units/toEMU (get config :width ))
                                                                                                 (Units/toEMU (get config :height)))
                                                                                    (recur (cons (subvec current 3) (rest body)) run)))
                               (map? (first current)) (do
                                                        (font-config (first current) run)
                                                        (recur (cons (subvec current 1) (rest body)) run))
                               (string? (first current)) (do
                                                           (.setText run (ffirst body))
                                                           (recur (rest body) run)))
          (map? (first body)) (let [config (first body)
                                    run (.createRun par)]
                                (if (contains? config :align) (.setAlignment par (align-code (get config :align))))
                                (if (contains? config :vert-align) (.setVerticalAlignment par (align-code (get config :vert-align))))
                                (if (contains? config :border-bottom) (.setBorderBottom par (border-code (get config :border-bottom))))
                                (if (contains? config :border-top) (.setBorderTop par (border-code (get config :border-top))))
                                (if (contains? config :border-right) (.setBorderRight par (border-code (get config :border-right))))
                                (if (contains? config :border-left) (.setBorderLeft par (border-code (get config :border-left))))
                                (if (contains? config :border-between) (.setBorderBetween par (border-code (get config :border-between))))
                                (recur (rest body) run))
          (string? (first body)) (do
                                   (.setText run (first body))
                                   (recur (rest body) (.createRun par))))))))

(defn word [& exps]
  (let [doc (new XWPFDocument)]
    (loop [exps exps]
      (cond
        (nil? exps) nil
        (empty? exps) nil
        (vector? (first exps)) (cond
                                 (= :p (ffirst exps)) (do
                                                        (paragraph doc (first exps))
                                                        (recur (rest exps)))
                                 (= :br (ffirst exps))(let [par (.createParagraph doc)
                                                            run (.createRun par)]
                                                        (.addBreak run)
                                                        (recur (rest exps))))
        :else
        (throw (Throwable. (str "Syntax Error! " (first exps))))))
    doc))

(defn -main
  [& args]
  (let [doc (word
             [:p {:border-bottom "double" :border-top "double" :align "center"} "This is " [:b "BOLD"] ", ya know! And this is" [:i " ITALIC"]]
             [:p [:b {:color "7CFC00" :font-size 30} "Green?"]]
             [:p "Calling for a " [:strike "STRIKE"]]
             [:p "No formatting."]
             [:p "The quick brown fox" [:run {:font-family "Courier"
                                              :font-size 20
                                              :color "0000FF"} " jumps"]]
             [:p {:align "center"} "This is center!"]
             [:p {:align "right"} "This is right!"]
             [:p {:align "both"} "This is both!"]
             [:p {:align "center" :border-top "single"} "Single Border Top"]
             [:br]
             [:p {:align "center" :border-top "double" :border-bottom "double" :border-left "double" :border-right "double"}
              "Double Border Top & Double Border Bottom & Double Border Left & Double Border Right"]
             [:p [:img {:height 236 :width 200} "img/MikeBensonPicture.jpg"]]
             [:p [:b [:i "Font Style One"]] [:br]
              [:run {:text-position 100} "Font Style Two"]]
             [:p [:u [:b "Bold and underlined"]]]
             )
        out (new FileOutputStream (new File "/tmp/test.docx"))]
    (.write doc out)
    (.close out)
    (println "/tmp/test.docx written successfully")))
