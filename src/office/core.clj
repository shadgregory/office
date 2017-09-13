(ns office.core
  (:require [office.excel :refer [excel]]
            [office.word  :refer [word]])
  (:import [java.io File FileOutputStream File]))

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
             [:p [:img {:height 236 :width 200} "img/john_mccarthy.jpg"]]
             [:p [:b [:i "Font Style One"]] [:br]
              [:run {:text-position 100} "Font Style Two"]]
             [:p [:u [:b "Bold and underlined"]]]
             )
        out (new FileOutputStream (new File "/tmp/test.docx"))]
    (.write doc out)
    (.close out)
    (println "/tmp/test.docx written successfully"))
  (let [out (new FileOutputStream (new File "/tmp/foo.xslx"))
        wb (excel
            [:wb
             [:spreadsheet " Employee Info "
              [:tr {:background-color "grey_25_percent"}
               [:th "EMP ID"]
               [:th "EMP NAME"]
               [:th "DESIGNATION"]]
              [:tr
               [:td "tp01"]
               [:td "Gopal"]
               [:td "Technical Manager"]]
              [:tr
               [:td "tp02"]
               [:td "Manisha"]
               [:td "Proof Reader"]]
              [:tr
               [:td "tp03"]
               [:td "Masthan"]
               [:td "Technical Writer"]]
              [:tr
               [:td "tp04"]
               [:td "Satish"]
               [:td "Technical Writer"]]
              [:tr
               [:td {:font-style "italic"} "tp05"]
               [:td {:font-weight "bold"} "Krishna"]
               [:td "Technical Writer"]]
              [:tr [:td {:colspan "3"} "This cell is three columns wide!"]]]])]
    (.write wb out)
    (.close out)
    (prn "/tmp/foo.xslx written!")))
