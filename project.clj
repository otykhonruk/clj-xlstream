(defproject clj-xlstream "0.1.0-SNAPSHOT"
  :description "Apache POI wrapper for streaming Excel spreadsheet content"
  :url "http://github.com/atihonruk/clj-xlstream"
  :license {:name "Eclipse Public License"
            :url "http://www.eclipse.org/legal/epl-v10.html"}
  :dependencies [[org.clojure/clojure "1.5.1"]
                 [org.clojure/java.jdbc "0.3.0-alpha4"]
                 [org.apache.poi/poi "3.9"]
                 [org.apache.poi/poi-ooxml "3.9"]]
  :repositories {"sonatype-oss-public" "https://oss.sonatype.org/content/groups/public/"}
  :main clj-xlstream.core)
