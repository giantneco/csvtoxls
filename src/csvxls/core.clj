(ns csvxls.core)

(require '[clojure.data.csv :as csv]
         '[clojure.java.io :as io])

(import '(org.apache.poi.hssf.usermodel HSSFWorkbook)
        '(org.apache.poi.ss.usermodel Workbook)
        '(org.apache.poi.ss.usermodel Sheet)
        '(org.apache.poi.ss.usermodel Row)
        '(org.apache.poi.ss.usermodel Cell)
        '(java.io FileOutputStream)
        )

(defn readcsv
  "read csv"
  [filename]
  (with-open [in-file (io/reader filename)]
    (doall
     (csv/read-csv in-file))))

;; (defn row-col-csv
;;   "add row and column to csv"
;;   [csv]
;;   (map cons (iterate inc 0 )
;;    (map (partial map vector (iterate inc 0 )) csv)))

(defn replaceSuffix
  "replace suffix of string"
  [target from to]
  (clojure.string/replace target
                          (re-pattern (str from "$"))
                          to))

(defn do-dump-csv
  "dump csv data to xls"
  [helper sheet csv]
  (do
    (loop [row (.createRow sheet 0), row-num 0, col-num 0, rest csv, csv-column (first rest)]
      (if (empty? rest)
        ()
        (if (empty? csv-column)
          (recur (.createRow sheet (inc row-num)) (inc row-num) 0 (next rest) (first (next rest)))
          (do
            ;; (println (str row-num "," col-num " " (first csv-column)))
            (-> row (.createCell col-num) (.setCellValue (.createRichTextString helper (first csv-column))))
            (recur row row-num (inc col-num) rest (next csv-column))))))))

(defn csv2xls
  "replace csv to xls"
  [csvfile xls]
  (with-open [out (FileOutputStream. xls)]
    (let [wb (HSSFWorkbook.)
          helper (.getCreationHelper wb)
          sheet (.createSheet wb "sheet0")]
      (do-dump-csv helper sheet (readcsv csvfile))
      (.write wb out))))

(defn -main
  [& args]
  (when (= (count args) 1)
    (let [csv (first args)
          xls (replaceSuffix csv ".csv" ".xls")]
      (csv2xls csv xls))))
