(ns csvxls.core)

(defn replaceSuffix
  "replace suffix of string"
  [target from to]
  (clojure.string/replace target
                          (re-pattern (str from "$"))
                          to))

(defn file-to-csv
  [filename]
  { :row_num 1 :col_num 2})

(defn csv2xls
  "replace csv to xls"
  [csvfile xls]
  (let [csv (file-to-csv csvfile)
        row_num (:row_num csv)
        col_num (:col_num csv)]
    (println row_num col_num)))

(defn -main
  [& args]
  (when (= (count args) 1)
    (let [csv (first args)
          xls (replaceSuffix csv ".csv" ".xls")]
      (csv2xls csv xls))))))
