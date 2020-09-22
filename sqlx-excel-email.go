package main

import (
	//	"database/sql"
	"fmt"
	"log"

	"github.com/jmoiron/sqlx"

	"strconv"

	"strings"

	"github.com/360EntSecGroup-Skylar/excelize"

	_ "github.com/mattn/go-oci8"

	"time"

	"github.com/jordan-wright/email"

)

func do() {

	var rows [][]interface{}

	{
		db, err := sqlx.Open("oci8", global.config.Dsn)
		if err != nil {
			fmt.Println(err)
			return
		}
		//db.SetMaxIdleConns(2)
		defer db.Close()

		{
			_, err = db.Exec("alter session set current_schema=" + global.config.Schema)
			if err != nil {
				log.Panicln(err)
			}

			dbRows, err := db.Queryx(`select * from table`)
			if err != nil {
				log.Panicln(err)
			}
			defer dbRows.Close()
			/*			ct, _ := dbRows.ColumnTypes()
						for _, v := range ct {
							log.Print("\t",v.Name())
						}*/

			for dbRows.Next() {

				cols, err := dbRows.SliceScan()

				if err != nil {
					log.Panicln(err)
				}
				rows = append(rows, cols)
			}

		}

	}

	var saveFile string

	saveFile = global.folder + "\\output\\report_" + time.Now().Format("2006-01-02-1504") + ".xlsx"

	xlsx, err := excelize.OpenFile(global.folder + "\\template.xlsx")

	if err != nil {
		log.Panicln(err)
	}

	sheet := "Sheet1"
	begin := 2
	i := 0
	for _, v := range rows {
		for j, vv := range v {
			col, _ := excelize.ColumnNumberToName(j + 1)
			//log.Printf("type %T, %v", vv, vv)

			xlsx.SetCellValue(sheet, col+strconv.Itoa(begin+i), vv)

		}
		i++

	}

	xlsx.SaveAs(saveFile)

	if global.config.Recipients != "" && len(rows) != 0 {
		e := email.NewEmail()
		e.From = global.config.Sender
		e.To = strings.Split(global.config.Recipients, ";")
		e.Cc = []string{global.config.Sender}
		e.Subject = "Report"
		e.HTML = []byte("<div style='font-family:arial'>Please find report attached.</div>")
		e.AttachFile(saveFile)
		err = e.Send("smtpserver:2567", nil)
		if err != nil {
			log.Printf("%s\n", err)
		}
	}

}
