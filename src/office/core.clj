(ns office.core
  (:require [office.excel :refer [excel]])
  (:import [java.io File FileOutputStream File]))

(defn -main
  [& args]
  (let [out (new FileOutputStream (new File "/tmp/foo.xslx"))
        wb (excel
            [:wb
             [:spreadsheet " Employee Info "
              [:row
               [:cell "EMP ID"]
               [:cell "EMP NAME"]
               [:cell "DESIGNATION"]]
              [:row
               [:cell "tp01"]
               [:cell "Gopal"]
               [:cell "Technical Manager"]]
              [:row
               [:cell "tp02"]
               [:cell "Manisha"]
               [:cell "Proof Reader"]]
              [:row
               [:cell "tp03"]
               [:cell "Masthan"]
               [:cell "Technical Writer"]]
              [:row
               [:cell "tp04"]
               [:cell "Satish"]
               [:cell "Technical Writer"]]
              [:row
               [:cell {:font-style "italic"} "tp05"]
               [:cell {:font-weight "bold"} "Krishna"]
               [:cell "Technical Writer"]]]])]
    (.write wb out)
    (.close out)
    (prn "/tmp/foo.xslx written!")))
