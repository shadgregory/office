(ns office.core
  (:require [office.excel :refer [excel]])
  (:import [java.io File FileOutputStream File]))

(defn -main
  [& args]
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
               [:td "Technical Writer"]]]])]
    (.write wb out)
    (.close out)
    (prn "/tmp/foo.xslx written!")))
