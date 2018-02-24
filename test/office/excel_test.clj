(ns office.excel-test
  (:import [java.io FileOutputStream File])
  (:require [office.excel :as excel]
            [clojure.test :refer [deftest is]]))

(deftest column-count
  (let [out (new FileOutputStream (new File "./count.xslx"))
        wb (excel/excel [:wb
                         [:table {:title "Column Count Test"}
                          [:tbody
                           [:tr
                            [:td "A"]
                            [:td "B"]]]]])]
    (.write wb out)
    (.close out)
    (is (= 2 (excel/max-column-count "./count.xslx"))))
  (let [out (new FileOutputStream (new File "./count.xslx"))
        wb (excel/excel [:wb
                         [:table {:title "Column Count Test"}
                          [:tbody
                           [:tr [:td "A"][:td "B"] [:td "C"]]
                           [:tr [:td "A"][:td {:colspan "2"} "B"]]]]])]
    (.write wb out)
    (.close out)
    (is (= 3 (excel/max-column-count "./count.xslx"))))
  (is (= 2 (excel/column-count [:wb
                                [:table {:title "Test"}
                                 [:thead
                                  [:tr [:th "A"] [:th "B"]]]
                                 [:tbody [:tr [:td "foo"] [:td "bar"]]]
                                 [:tfoot [:tr [:td "foobar"] [:td "barfoo"]]]]])))
  (is (= 3 (excel/column-count [:wb
                                [:spreadsheet {:title " Employee Info "}
                                 [:tr [:td {:colspan "3"} "Description Goes Here"]]
                                 [:thead
                                  [:tr {:background-color "grey_25_percent"}
                                   [:th "EMP ID"]
                                   [:th "EMP NAME"]
                                   [:th "DESIGNATION"]]]
                                 [:tbody
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
                                  [:tr.ignore-this-class
                                   [:td "tp04"]
                                   [:td "Satish"]
                                   [:td "Technical Writer"]]
                                  [:tr
                                   [:td {:font-style "italic"} "tp05"]
                                   [:td {:font-weight "bold"} "Krishna"]
                                   [:td "Technical Writer"]]
                                  [:tr
                                   [:td [:a {:href "https://duckduckgo.com/"} "DuckDuckGo"]]
                                   [:td [:a {:href "https://www.google.com/"} "Google"]]
                                   [:td [:a {:href "https://www.bing.com/"} "Bing"]]]]
                                 [:tfoot
                                  [:tr [:td {:colspan "3"} "Footer Goes Here"]]]]]))))

