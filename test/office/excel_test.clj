(ns office.excel-test
  (:require [office.excel :as sut]
            [clojure.test :refer [deftest is]]))

(deftest column-count
  (is (= 2 (sut/column-count [:wb
                              [:table
                               [:thead
                                [:tr [:th "A"] [:th "B"]]]
                               [:tbody [:tr [:td "foo"] [:td "bar"]]]
                               [:tfoot [:tr [:td "foobar"] [:td "barfoo"]]]]]))))
