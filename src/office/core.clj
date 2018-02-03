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
  (let [out (new FileOutputStream (new File "/tmp/prez.xslx"))
        wb (excel
            [:wb
             [:table {:title "Test"}
              [:thead
               [:tr {:background-color "#8DBDD8"}
                [:td "President"]
                [:td "Born"]
                [:td "Died"]
                [:td "Wiki"]]]
              [:tbody 
               [:tr [:td "Abraham Lincoln"]
                [:td "1809"]
                [:td "1865"]
                [:td [:a {:href "https://en.wikipedia.org/wiki/Abraham_Lincoln"} "Bio"]]]
               [:tr
                [:td "Andrew Johnson"]
                [:td "1808"]
                [:td "1875"]
                [:td [:a {:href "https://en.wikipedia.org/wiki/Andrew_Johnson"} "Bio"]]]
               [:tr
                [:td "Ulysses S. Grant"]
                [:td "1822"]
                [:td "1885"]
                [:td
                 [:a {:href "https://en.wikipedia.org/wiki/Ulysses_S._Grant"} "Bio"]]]
               [:tr
                [:td "Rutherford B. Hayes"]
                [:td "1822"]
                [:td "1893"]
                [:td [:a {:href "https://en.wikipedia.org/wiki/Rutherford_B._Hayes"} "Bio"]]]]
              [:tfoot
               [:tr [:td {:colspan "4"} "Reconstruction Presidents."]]]]])]
    (.write wb out)
    (.close out))
  (let [out (new FileOutputStream (new File "/tmp/foo.xslx"))
        wb (excel
            [:wb
             [:spreadsheet {:title " Employee Info "}
              [:tr [:td {:colspan "4"} "Description Goes Here"]]
              [:thead
               [:tr {:background-color "#8DBDD8"}
                [:th "EMP ID"]
                [:th "EMP NAME"]
                [:th "DESIGNATION"]
                [:th "TIME EMPLOYED"]]]
              [:tbody
               [:tr
                [:td "tp01"]
                [:td "Gopal"]
                [:td "Technical Manager"]
                [:td "6"]]
               [:tr
                [:td "tp02"]
                [:td "Manisha"]
                [:td "Proof Reader"]
                [:td "7"]]
               [:tr
                [:td "tp03"]
                [:td "Masthan"]
                [:td "Technical Writer"]
                [:td "4"]]
               [:tr.ignore-this-class
                [:td "tp04"]
                [:td "Satish"]
                [:td "Technical Writer"]
                [:td "24"]]
               [:tr
                [:td {:font-style "italic"} "tp05"]
                [:td {:font-weight "bold"} "Krishna"]
                [:td "Technical Writer"]
                [:td "30"]]
               [:tr
                [:td [:a {:href "https://duckduckgo.com/"} "DuckDuckGo"]]
                [:td [:a {:href "https://www.google.com/"} "Google"]]
                [:td [:a {:href "https://www.bing.com/"} "Bing"]]
                [:sum "D3:D7"]]]
              [:tfoot
               [:tr [:td {:colspan "4"} "Footer Goes Here"]]]]])]
    (.write wb out)
    (.close out)
    (prn "/tmp/foo.xslx written!")))
