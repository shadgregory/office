(defproject office "0.1.0-SNAPSHOT"
  :description "FIXME: write description"
  :url "http://example.com/FIXME"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.9.0"]
                 [ring/ring-core "1.6.2"]
                 [org.apache.poi/poi "3.17"]
                 [org.apache.poi/poi-ooxml "3.17"]]
  :main ^:skip-aot office.core
  :local-repo "local-m2"
  :target-path "target/%s"
  :profiles {:uberjar {:aot :all}})
