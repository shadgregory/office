# Office for Clojure
Office for Clojure provides a hiccup-like syntax for ms office docs.

```clojure
(excel
 [:wb
  [:table {:title "Test"}
   [:thead
    [:tr [:th "A"] [:th "B"]]]
   [:tbody [:tr [:td "foo"] [:td "bar"]]]
   [:tfoot [:tr [:td "foobar"] [:td "barfoo"]]]]])
```
![Screenshot](screenshot_excel.png)
