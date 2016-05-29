package main

import (
	"strings"
	"github.com/tealeg/xlsx"
	"gopkg.in/yaml.v2"
	"io/ioutil"
	"path/filepath"
	"strconv"
	"os"
	"bufio"
)

type setting struct {
	// 各ステータスはyamlから書き込むため、大文字でないといけないみたい
	FileName 				string	`yaml:"file_name"`
	Sheet struct {
			SettingSheet	string	`yaml:"setting_sheet"`
			ResultSheet		string	`yaml:"result_sheet"`
		}
	Comments				string	`yaml:"comment"`
}

type sourceFile struct {
	filePath 	string
	fileName 	string
	columnNum 	int // intサイズ以上のソースは現実的でないので
}

var s setting
var xlFile *xlsx.File
var sourcePaths []string
var fileExtentions []string
var sourceFiles []sourceFile
var resultFiles []sourceFile

func main() {

	// 設定ファイル読込み
	getSettingData()

	// 集計ファイルチェック
	checkCalcDataFile()

	// 集計対象ファイル情報取得
	getSourceFilesInfo()

	// 集計対象ファイルパス取得
	getSourceFiles()

	// ステップ数集計
	calcSourceSteps()

	// 集計済みファイル取得
	getResultFiles()

	// 集計結果書込み
	writeCalcResult()

}

/* 設定ファイル読込み */
func getSettingData() {
	buf, err := ioutil.ReadFile("data.yaml")

	if err != nil {
		panic(err)
	}

	err = yaml.Unmarshal(buf, &s)

	// TODO: 設定ファイルのフォーマットチェック
}

/* 集計ファイルチェック */
func checkCalcDataFile() {
	fileName := s.FileName
	var err error
	xlFile, err = xlsx.OpenFile(fileName)

	if err != nil {
		panic(err)
	}

	// TODO: ファイルフォーマットチェック
}

/* 集計対象ファイル情報取得 */
func getSourceFilesInfo() {

	var sheet *xlsx.Sheet
	for _, sheet = range xlFile.Sheets {
		if sheet.Name == s.Sheet.SettingSheet {break}
	}

	rows := sheet.Rows

	// "B-4"から下の、入力されたファイルパスを取得
	rowNum := 3
	for {
		str := rows[rowNum].Cells[1].Value
		if len(str) == 0 {break}
		sourcePaths = append(sourcePaths, str)
		rowNum++
	}

	// "C-4"から下の、入力された拡張子を取得
	rowNum = 3
	for {
		str := rows[rowNum].Cells[2].Value
		if len(str) == 0 {break}
		fileExtentions = append(fileExtentions, str)
		rowNum++
	}
}

/* 集計対象ファイル取得 */
func getSourceFiles() {
	for _, path := range sourcePaths {
		searchSourcePath(path)
	}
}

/* ソースパス検索 */
func searchSourcePath(searchPath string) {
	files, err := ioutil.ReadDir(searchPath)

	// 集計対象ファイルのパスが存在しない場合など
	if err != nil {
		panic(err)
	}

	for _, fi := range files {
		fullPath := filepath.Join(searchPath, fi.Name())

		// 取得ファイルがディレクトリの場合、配下のファイルを検索する
		if fi.IsDir() {
			searchSourcePath(fullPath)
		} else {
			// TODO; ルートを指定して、相対パスで出力できるようにする？
			//rel, err := filepath.Rel(rootPath, fullPath)
			//if err != nil {
			//	panic(err)
			//}
			// 拡張子チェック
			if checkFileExtension(fullPath) {
				// フルパスからファイル名を取得する
				s := strings.Split(fullPath, "/")

				sf := sourceFile{searchPath + "/", s[len(s)-1], 0}
				sourceFiles = append(sourceFiles, sf)
			}
		}
	}
}

 /* 拡張子チェック */
func checkFileExtension(fullPath string) bool {
	for _, ex := range fileExtentions {
		if strings.HasSuffix(fullPath, ex) {
			return true
		}
	}
	return false
}

/* ステップ数集計 */
func calcSourceSteps() {
	// 除外コメント取得
	comments := strings.Split(s.Comments, ",")

	for i, sf := range sourceFiles {
		step := 0
		// ファイルを開く
		fi, err := os.Open(sf.filePath + sf.fileName)
		if err != nil {
			panic(err)
		}

		// メソッドがreturn時に閉じる
		defer fi.Close()

		scanner := bufio.NewScanner(fi)
		for scanner.Scan() {
			// ステップ換算する場合、ステップ数をインクリメントする
			if isStep(scanner.Text(), comments) {step++}
		}

		if serr := scanner.Err(); serr != nil {
			panic(serr)
		}

		// goのrangeを使用したfor文は、値渡しになるっぽい
		// sf.columnNum = stepは、NG
		sourceFiles[i].columnNum = step
	}
}

/* ステップ数集計 */
func isStep(str string, comments []string) bool {

	// 文字列からタブ文字と半角スペースを除外
	str = strings.Replace(str, "\t", "", -1)
	str = strings.Replace(str, " ", "", -1)

	// 文字列が存在しない場合、step数として換算しない
	if len(str) == 0 {return false}
	// 文字列がコメントの場合、step数として換算しない
	for _, c := range comments {
		if strings.HasPrefix(str, c) {
			return false
		}
	}

	return true
}

/* 集計済みファイル取得 */
func getResultFiles() {

	var sheet *xlsx.Sheet
	for _, sheet = range xlFile.Sheets {
		if sheet.Name == s.Sheet.ResultSheet {break}
	}

	rows := sheet.Rows

	// "C-3"と"D-3から下の、ファイルパスを取得
	rowNum := 2
	for {
		if sheet.MaxRow <= rowNum {break}
		sf := sourceFile{rows[rowNum].Cells[2].Value, rows[rowNum].Cells[3].Value, 0}
		resultFiles = append(resultFiles, sf)
		rowNum++
	}
}

/* 集計結果書込み */
func writeCalcResult() {

	var sheet *xlsx.Sheet
	for _, sheet = range xlFile.Sheets {
		if sheet.Name == s.Sheet.ResultSheet {break}
	}

	var cell *xlsx.Cell
	for i, sf := range sourceFiles {

		// 集計済みファイルの場合、ステップ数の上書き処理を行う
		if j := isResultFile(sf); j != -1 {
			row := sheet.Rows
			cell = row[j + 2].Cells[4]
			cell.Value = strconv.Itoa(sf.columnNum)

			// 新規追加ファイルの場合、列の新規追加を行う
		} else {
			row := sheet.AddRow()

			// "A"列に空セルを追加
			row.AddCell()
			// "B"列にNoを入力
			cell = row.AddCell()
			cell.Value = strconv.Itoa(i + 1)
			// "C"列にファイルパスを入力
			cell = row.AddCell()
			cell.Value = sf.filePath
			// "D"列にファイル名を入力
			cell = row.AddCell()
			cell.Value = sf.fileName
			// "E"列にステップ数を入力
			cell = row.AddCell()
			cell.Value = strconv.Itoa(sf.columnNum)
		}
	}
	xlFile.Save(s.FileName)

}

/*
 * 集計済みチェック
 * 集計済みの場合、resultFilesの添字を返却
 * 集計済みでない場合、-1を返却
 */
func isResultFile(sf sourceFile) int {
	for i, rf := range resultFiles {
		if (sf.filePath == rf.filePath) && (sf.fileName == rf.fileName) {
			return i
		}
	}
	return -1
}