package main

import (
	"io"
	"os"
	"os/exec"
	"strconv"
	"strings"

	"github.com/okonma-violet/confdecoder"
	"github.com/tealeg/xlsx"
)

var failedpath = "/FAILED/"
var convertedpath = "/converted/"

type cell struct {
	col       int
	row       int
	value_str string
	value_int int
}

func main() {
	args := os.Args[1:]
	if len(args) < 4 || args[0] == "--help" {
		println("format: [xlsx template file path] [new xlsx file path] [template config file path] [xlsx files path]")
		return
	}
	tconfig := new(struct {
		WorkingCoords []string
		ControlCoords []string
	})
	err := confdecoder.DecodeFile("1conf.txt", tconfig) //(args[2])
	if err != nil {
		println("Parsing template config file err: " + err.Error())
		return
	}

	// limsheet := make(map[int]*struct {
	// 	limrow_f []int
	// 	limcol_f []int
	// 	limrow_s []int
	// 	limcol_s []int
	// })

	limsheetcfg := make(map[int][][]int) // col,row,col,row
	for i := 0; i < len(tconfig.WorkingCoords); i++ {
		extc := strings.Split(tconfig.WorkingCoords[i], "-")
		sheetnum, err := strconv.Atoi(extc[0])
		if err != nil {
			println("Parsing template config file err: incorrect WorkingCords field: " + tconfig.WorkingCoords[i])
			return
		}
		// if _, ok := limsheet[sheetnum]; !ok {
		// 	limsheet[sheetnum] = &struct {
		// 		limrow_f []int
		// 		limcol_f []int
		// 		limrow_s []int
		// 		limcol_s []int
		// 	}{make([]int, 0,1), make([]int, 0,1), make([]int, 0,1), make([]int, 0,1)}
		// }
		if wc := strings.Split(extc[1], ":"); len(wc) == 2 {
			lmts := make([]int, 4)
			lmts[0], lmts[1], err = xlsx.GetCoordsFromCellIDString(wc[0])
			if err != nil {
				println("Parsing template config file err: incorrect WorkingCords field: " + tconfig.WorkingCoords[i])
				return
			}
			lmts[2], lmts[3], err = xlsx.GetCoordsFromCellIDString(wc[1])
			if err != nil {
				println("Parsing template config file err: incorrect WorkingCords field: " + tconfig.WorkingCoords[i])
				return
			}
			limsheetcfg[sheetnum] = append(limsheetcfg[sheetnum], lmts)
		} else {
			println("Parsing template config file err: incorrect WorkingCords field: " + tconfig.WorkingCoords[i])
			return
		}
	}
	controlcellscfg := make(map[int][][]int) // col,row,col,row
	for i := 0; i < len(tconfig.ControlCoords); i++ {
		extc := strings.Split(tconfig.ControlCoords[i], "-")
		sheetnum, err := strconv.Atoi(extc[0])
		if err != nil {
			println("Parsing template config file err: incorrect ControlCoords field: " + tconfig.ControlCoords[i])
			return
		}
		//sheetnum--
		if wc := strings.Split(extc[1], ":"); len(wc) == 2 {
			ctrl := make([]int, 4)
			ctrl[0], ctrl[1], err = xlsx.GetCoordsFromCellIDString(wc[0])
			if err != nil {
				println("Parsing template config file err: incorrect ControlCoords field: " + tconfig.ControlCoords[i])
				return
			}
			ctrl[2], ctrl[3], err = xlsx.GetCoordsFromCellIDString(wc[1])
			if err != nil {
				println("Parsing template config file err: incorrect ControlCoords field: " + tconfig.ControlCoords[i])
				return
			}
			controlcellscfg[sheetnum] = append(controlcellscfg[sheetnum], ctrl)
		} else {
			println("Parsing template config file err: incorrect ControlCoords field: " + tconfig.ControlCoords[i])
			return
		}
	}

	tf, err := xlsx.OpenFile("1.xlsx") //(args[0])
	if err != nil {
		println("Reading template file err: " + err.Error())
		return
	}

	controlcells := make(map[string][]cell)

	for sheetnum, ctrls := range controlcellscfg {
		if len(tf.Sheets) < sheetnum {
			println("Parsing template file err: incorrect WorkingCords field's sheet num: " + strconv.Itoa(sheetnum))
			return
		}
		sheet := tf.Sheets[sheetnum]
		cells := make([]cell, 0)
		for i := 0; i < len(ctrls); i++ {
			for cl := ctrls[i][0]; cl <= ctrls[i][2]; cl++ {
				if sheet.MaxCol <= cl {
					println("Parsing template file err: WorkingCords column overflow")
					return
				}
				for rw := ctrls[i][1]; rw <= ctrls[i][3]; rw++ {
					if sheet.MaxRow <= rw {
						println("Parsing template file err: WorkingCords row overflow")
						return
					}
					c, err := sheet.Cell(rw, cl)
					if err != nil {
						println("Parsing template file (sheet.Cell) err: " + err.Error())
						return
					}
					cells = append(cells, cell{col: cl, row: rw, value_str: c.String()})
				}
			}
		}
		controlcells[sheet.Name] = append(controlcells[sheet.Name], cells...)
	}

	outcells := make(map[string][]cell)

	for sheetnum, limcls := range limsheetcfg {
		if len(tf.Sheets) < sheetnum {
			println("Parsing template file err: incorrect ControlCoords field's sheet num: " + strconv.Itoa(sheetnum))
			return
		}
		sheet := tf.Sheets[sheetnum]
		cells := make([]cell, 0)
		for i := 0; i < len(limcls); i++ {
			for cl := limcls[i][0]; cl <= limcls[i][2]; cl++ {
				if sheet.MaxCol <= cl {
					println("Parsing template file err: ControlCoords column overflow")
					return
				}
				for rw := limcls[i][1]; rw <= limcls[i][3]; rw++ {
					if sheet.MaxRow <= rw {
						println("Parsing template file err: ControlCoords row overflow")
						return
					}
					_, err := sheet.Cell(rw, cl)
					if err != nil {
						println("Parsing template file (sheet.Cell) err: " + err.Error())
						return
					}
					cells = append(cells, cell{col: cl, row: rw})
				}
			}
		}
		outcells[sheet.Name] = append(outcells[sheet.Name], cells...)
	}

	files, err := os.ReadDir(args[3])
	if err != nil {
		println("ReadDir err: " + err.Error())
		return
	}
	convertedpath = args[3] + convertedpath
	if err := os.Mkdir(convertedpath, 0755); err != nil {
		println("Mkdir (convertedpath) err: " + err.Error())
		return
	}
	failedpath = args[3] + failedpath
	if err := os.Mkdir(failedpath, 0755); err != nil {
		println("Mkdir (failedpath) err: " + err.Error())
		return
	}
	curpath := args[3] + "/"
filesparsing:
	for i := 0; i < len(files); i++ {
		if strings.HasSuffix(strings.ToLower(files[i].Name()), ".xls") {
			if out, err := converttoxlsx(curpath, convertedpath); err != nil {
				println("Converttoxlsx err: " + err.Error() + "; output: " + out)
				return
			}
			continue
		}
		if strings.HasSuffix(strings.ToLower(files[i].Name()), ".xlsx") {
			curfile, err := xlsx.OpenFile(curpath + files[i].Name())
			if err != nil {
				println("Opening xlsx err: " + err.Error())
				return
			}
			for ctrlsheetname, ctrlcells := range controlcells {
				sht, ok := curfile.Sheet[ctrlsheetname]
				if !ok {
					println("File " + files[i].Name() + " does not pass control, no such sheet")
					goto failed
				}
				for k := 0; k < len(ctrlcells); k++ {
					crcl, err := sht.Cell(ctrlcells[k].row, ctrlcells[k].col)
					if err != nil {
						println("File " + files[i].Name() + " does not pass control, sheet.Cell() err: " + err.Error())
						goto failed
					}
					if crcl.String() != ctrlcells[k].value_str {
						println("File " + files[i].Name() + " does not pass control, not equal")
						goto failed
					}
				}
				continue
			failed:
				src, err := os.Open(curpath + files[i].Name())
				if err != nil {
					println("Opening file (failed) " + files[i].Name() + " err: " + err.Error())
					return
				}
				dst, err := os.Create(failedpath + files[i].Name())
				if err != nil {
					println("Creating file (failed) " + files[i].Name() + " err: " + err.Error())
					src.Close()
					return
				}
				if _, err = io.Copy(dst, src); err != nil {
					println("Copying file (failed) " + files[i].Name() + " err: " + err.Error())
					src.Close()
					dst.Close()
					os.Remove(failedpath + files[i].Name())
					return
				}
				src.Close()
				dst.Close()
				println("Failed file " + files[i].Name() + " copied to FAILED")
				continue
			}

			curoutcells := make(map[*cell]int)

			for sheetname, cls := range outcells {
				sht, ok := curfile.Sheet[sheetname]
				if !ok {
					println("File " + files[i].Name() + " does not pass control, no such sheet")
					goto failed2
				}
				for k := 0; k < len(cls); k++ {
					cl, err := sht.Cell(cls[k].row, cls[k].col)
					if err != nil {
						println("File " + files[i].Name() + " does not pass control, sheet.Cell() err: " + err.Error())
						goto failed2
					}
					val, err := cl.Int()
					if err != nil {
						println("File " + files[i].Name() + " does not pass control, cell.Int() err: " + err.Error())
						goto failed2
					}
					curoutcells[&cls[k]] = val
				}
			}
			////////////////////
		failed2:
			src, err := os.Open(curpath + files[i].Name())
			if err != nil {
				println("Opening file (failed) " + files[i].Name() + " err: " + err.Error())
				return
			}
			dst, err := os.Create(failedpath + files[i].Name())
			if err != nil {
				println("Creating file (failed) " + files[i].Name() + " err: " + err.Error())
				src.Close()
				return
			}
			if _, err = io.Copy(dst, src); err != nil {
				println("Copying file (failed) " + files[i].Name() + " err: " + err.Error())
				src.Close()
				dst.Close()
				os.Remove(failedpath + files[i].Name())
				return
			}
			src.Close()
			dst.Close()
			println("Failed file " + files[i].Name() + " copied to FAILED")
		}
	}
	if curpath == args[3]+"/" {
		curpath = convertedpath
		goto filesparsing
	}
}

func converttoxlsx(filepath string, outdir string) (string, error) {
	return run("soffice", []string{"--headless", "--convert-to", "xlsx", filepath, "--outdir", outdir})
	// //var b []byte
	// b, err := cmd.CombinedOutput()
	// out := string(b)
	// return out, err
}
func run(path string, args []string) (out string, err error) {

	cmd := exec.Command(path, args...)
	//logger.Println("------------- " + cmd.String())

	var b []byte
	b, err = cmd.CombinedOutput()
	out = string(b)

	return
}

// for k := len(args) - 1; k > 1; k -= 2 {
// 	xf, err := xlsx.OpenFile(args[k])
// 	if err != nil {
// 		println("Reading xlsx file err: " + err.Error())
// 		return
// 	}
// 	if len(xf.Sheets) == 0 {
// 		println("Reading xlsx file err: no sheets found in file") // яхз возможно ли это
// 		return
// 	}

// 	sht := xf.Sheets[0]
// 	for j := 0; j < len(tfd.Rows); j++ {
// 		if tfd.Rows[j].Key == "sheet" {
// 			shti, err := strconv.Atoi(tfd.Rows[j].Value)
// 			if err != nil {
// 				var ok bool
// 				if sht, ok = xf.Sheet[tfd.Rows[j].Value]; !ok {
// 					println("Xlsx template's config err: sheet specified neither by num nor by existing sheetname (" + args[k-1] + ")")
// 					return
// 				}
// 			} else {
// 				if len(xf.Sheets) < (shti + 1) {
// 					println("Xlsx templates  config'err: specified sheet num is bigger than num of sheets in readed xlsx: " + strconv.Itoa(shti) + " you want (zero based), " + strconv.Itoa(len(xf.Sheets)) + " xlsx has")
// 					return
// 				}
// 				sht = xf.Sheets[shti]
// 			}
// 			continue
// 		}

// 		if tfd.Rows[j].Key == "" || tfd.Rows[j].Value == "" {
// 			println("Xlsx template's config err: bad row - name or value is empty")
// 			return
// 		}
// 		icol, irow, err := xlsx.GetCoordsFromCellIDString(tfd.Rows[j].Value)
// 		if err != nil {
// 			println("Xlsx template's config err: getting coords from row value err: " + err.Error())
// 			return
// 		}
// 		tag := tagopen + tfd.Rows[j].Key + tagclose
// 		if irow >= sht.MaxRow {
// 			println("Xlsx template's config err: specified row num in coords is bigger than num of rows in sheet, " + strconv.Itoa(irow) + " you want (zero based), " + strconv.Itoa(sht.MaxRow) + " sheet has")
// 			return
// 		}
// 		if icol >= sht.MaxCol {
// 			println("Xlsx template's config err: specified column num in coords is bigger than num of columns in sheet, " + strconv.Itoa(icol) + " you want (zero based), " + strconv.Itoa(sht.MaxCol) + " sheet has")
// 			return
// 		}
// 		cell, err := sht.Cell(irow, icol)
// 		if err != nil {
// 			println("Xlsx template's config err: sheet.Cell() err: " + err.Error())
// 			return
// 		}
// 		if err = dx.Replace(tag, cell.String(), -1); err != nil {
// 			println("Replace err: replacing tag \"" + tag + "\" with \"" + cell.String() + "\" err: " + err.Error())
// 			return
// 		}
// 		println("Replace: replaced tag \"" + tag + "\" with \"" + cell.String() + "\"")
// 	}

// }

// if err = dx.WriteToFile(args[1]); err != nil {
// 	println("Writing result file err: " + err.Error())
// }
// println("Done! Result written to " + args[1])
