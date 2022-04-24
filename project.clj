(defproject office "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.11.1"]
                 [ring/ring-core "1.9.5"]
                 [org.apache.poi/poi "4.1.2"]
                 [org.apache.poi/poi-ooxml "4.1.2"]]
  :main ^:skip-aot office.core
  :plugins [[lein-ancient "1.0.0-RC3"]]
  :local-repo "local-m2"
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all}})
