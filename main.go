package main

import (
	"errors"
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
	//value_int  int
	value_float float64
	tfcell      *xlsx.Cell
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
	err := confdecoder.DecodeFile(args[2], tconfig)
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

	tf, err := xlsx.OpenFile(args[0])
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
					c, err := sheet.Cell(rw, cl)
					if err != nil {
						println("Parsing template file (sheet.Cell) err: " + err.Error())
						return
					}
					if c.Formula() == "" {
						cells = append(cells, cell{col: cl, row: rw, tfcell: c})
					} else {
						c.Value = "" // !!! иначе все формулы в нуле будут стоять
					}
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
		if !errors.Is(err, os.ErrExist) {
			println("Mkdir (convertedpath) err: " + err.Error())
			return
		}
	}
	failedpath = args[3] + failedpath
	if err := os.Mkdir(failedpath, 0755); err != nil {
		if !errors.Is(err, os.ErrExist) {
			println("Mkdir (failedpath) err: " + err.Error())
			return
		}
	}
	var sucf, totalf int
	curpath := args[3] + "/"
filesparsing:
	for i := 0; i < len(files); i++ {
		if strings.HasSuffix(strings.ToLower(files[i].Name()), ".xls") {
			out, err := converttoxlsx(curpath+files[i].Name(), convertedpath)
			if err != nil {
				println("Converttoxlsx (file ", curpath+files[i].Name(), ") err: "+err.Error()+"; output: "+out)
				return
			}
			println("Converted " + curpath + files[i].Name() + " to " + convertedpath + files[i].Name() + "x")
			continue
		}
		if strings.HasSuffix(strings.ToLower(files[i].Name()), ".xlsx") {
			totalf++
			curfile, err := xlsx.OpenFile(curpath + files[i].Name())
			if err != nil {
				println("Opening xlsx err: " + err.Error())
				return
			}
			// CONTROL
			for ctrlsheetname, ctrlcells := range controlcells {
				sht, ok := curfile.Sheet[ctrlsheetname]
				if !ok {
					println("File " + files[i].Name() + " does not pass control, no such sheet")
					goto failedcontrol
				}
				for k := 0; k < len(ctrlcells); k++ {
					crcl, err := sht.Cell(ctrlcells[k].row, ctrlcells[k].col)
					if err != nil {
						println("File " + files[i].Name() + " does not pass control, sheet.Cell() err: " + err.Error())
						goto failedcontrol
					}
					if crcl.String() != ctrlcells[k].value_str {
						println("File " + files[i].Name() + " does not pass control, not equal")
						goto failedcontrol
					}
				}
				continue
			failedcontrol:
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
				println("File " + files[i].Name() + " failed, copied to FAILED")
				continue filesparsing
			}
			println("File " + curpath + files[i].Name() + " control done")

			// PARSING CELLS
			curoutcells := make(map[*cell]float64)
			for sheetname, cls := range outcells {
				sht, ok := curfile.Sheet[sheetname]
				if !ok {
					println("File " + files[i].Name() + " parse err: no such sheet")
					goto failedparsing
				}
				for k := 0; k < len(cls); k++ {
					cl, err := sht.Cell(cls[k].row, cls[k].col)
					origcoords := xlsx.GetCellIDStringFromCoords(cls[k].col, cls[k].row)
					if err != nil {
						println("File "+files[i].Name()+" parse err: sheet.Cell() err: "+err.Error(), ", coords: "+origcoords)
						goto failedparsing
					}
					if cl.Formula() != "" {
						println("File "+files[i].Name()+" parse err: formula found: "+cl.Formula(), ", coords: "+origcoords)
						goto failedparsing
					}
					valstr := cl.String()
					if valstr != "" {
						valflt, err := cl.Float()
						if err != nil {
							println("File "+files[i].Name()+" parse err: cl.Float() err: "+err.Error(), ", coords: "+origcoords)
							goto failedparsing
						}
						curoutcells[&cls[k]] = valflt
					}
				}
			}
			for cl, val := range curoutcells {
				cl.value_float += val
			}
			println("File " + curpath + files[i].Name() + " parsing cells done")
			sucf++
			continue

		failedparsing:
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
		files, err = os.ReadDir(curpath)
		if err != nil {
			println("ReadDir err: " + err.Error())
			return
		}
		goto filesparsing
	}
	println("Files total: " + strconv.Itoa(totalf) + ", successfully: " + strconv.Itoa(sucf) + ", failed: " + strconv.Itoa(totalf-sucf))
	// SET RESULT
	for _, cells := range outcells {
		for i := 0; i < len(cells); i++ {
			if cells[i].value_float == 0 {
				cells[i].tfcell.Value = ""
				continue
			}
			if isInt(cells[i].value_float) {
				cells[i].tfcell.SetInt(int(cells[i].value_float))
			} else {
				cells[i].tfcell.SetFloat(cells[i].value_float)
			}
		}
	}
	outname := args[1]
	if !strings.HasSuffix(outname, ".xlsx") {
		outname = outname + ".xlsx"
	}
	if err = tf.Save(outname); err != nil {
		println("Saving result to " + args[1] + " err: " + err.Error())
		return
	}
	println("Done! Result saved to " + outname)
}

func isInt(val float64) bool {
	return float64(int(val)) == val
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
