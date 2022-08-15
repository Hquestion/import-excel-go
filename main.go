package main

import (
	"database/sql"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"github.com/xuri/excelize/v2"
	"strconv"
	"strings"
	"time"
)

var sheetTypeMap = map[string]int64{
	"588":   1,
	"888":   2,
	"988":   3,
	"1088":  4,
	"1288":  5,
	"1388":  6,
	"1488":  7,
	"1688":  8,
	"1888":  9,
	"1988":  10,
	"1888新": 9,
}

type Ticket struct {
	Number     string    `json:"number"`
	Password   string    `json:"password"`
	Createtime time.Time `json:"createtime"`
	Type       int64     `json:"type"`
}

type Order struct {
	TicketId      string `json:"ticket_id"`
	Username      string `json:"username"`
	Phone         string `json:"phone"`
	Address       string `json:"address"`
	Createtime    string `json:"createtime"`
	ShunfengOrder string `json:"shunfeng_order"`
}

func main() {
	db, err := sql.Open("mysql", "root:123456@tcp(47.100.229.203)/pangxie")
	if err != nil {
		panic(err)
	}
	defer db.Close()
	// See "Important settings" section.
	db.SetConnMaxLifetime(time.Minute * 3)
	db.SetMaxOpenConns(10)
	db.SetMaxIdleConns(10)
	tickets := ReadExcelTicket()
	fmt.Printf("Total tockets : %v\n", len(tickets))
	_ = InsertTickets(db, tickets)
}

func ReadExcelTicket() []Ticket {
	f, err := excelize.OpenFile("ticket.xlsx")
	if err != nil {
		panic(err)
	}
	defer func() {
		if err = f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	sheetList := f.GetSheetList()
	var tickets = make([]Ticket, 0)
	for _, item := range sheetList {
		table, err := f.GetRows(item, excelize.Options{RawCellValue: true})
		if err != nil {
			panic(err)
		}
		for _, num := range table[1:] {
			if len(num) <= 2 {
				continue
			}
			if strings.Trim(num[0], "") == "" && strings.Trim(num[1], "") == "" {
				continue
			}
			var ticketItem Ticket
			ticketItem.Createtime = time.Now()
			//fmt.Printf("number 0 is : %v, length is %v\n", num[0], len(strings.Trim(num[0], "")))
			num0 := strings.Trim(num[0], "")
			if len(num0) > 10 || num0 == "" {
				ticketItem.Number = num[1]
				ticketItem.Password = num[2]
				if len(num) > 3 {
					t, err := strconv.ParseInt(num[3], 10, 64)
					if err != nil {
						ticketItem.Type = 0
					} else {
						ticketItem.Type = t
					}
				} else {
					ticketItem.Type = 0
				}
			} else {
				ticketItem.Number = num[0]
				ticketItem.Password = num[1]
				if len(num) > 2 {
					t, err := strconv.ParseInt(num[2], 10, 64)
					if err != nil {
						ticketItem.Type = 0
					} else {
						ticketItem.Type = t
					}
				} else {
					ticketItem.Type = 0
				}
			}
			ticketItem.Type = sheetTypeMap[item]
			tickets = append(tickets, ticketItem)
		}
	}
	return tickets
}

func InsertTickets(db *sql.DB, tickets []Ticket) map[string]int64 {
	stmt, err := db.Prepare("INSERT INTO ticket(number,password,createtime,type) VALUES (?,?,?,?)")
	if err != nil {
		panic(err)
	}
	count := 0
	var ticketNumIdMap = make(map[string]int64)
	for _, t := range tickets {
		res, err := stmt.Exec(t.Number, t.Password, t.Createtime, t.Type)
		if err != nil {
			fmt.Errorf("insert ticket [%v] error: %v\n", t.Number, err)
		}
		count++
		id, err := res.LastInsertId()
		if err != nil {
			fmt.Errorf("get insert id error: %v\n", err)
			continue
		}
		ticketNumIdMap[t.Number] = id
	}
	fmt.Printf("successfully inserted %v records\n", count)
	fmt.Printf("ticket number and id map: %v\n", ticketNumIdMap)
	return ticketNumIdMap
}
