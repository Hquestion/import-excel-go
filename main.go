package main

import (
	"database/sql"
	"fmt"
	_ "github.com/go-sql-driver/mysql"
	"github.com/xuri/excelize/v2"
	"os"
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

var orderTextMap = map[string]int64{
	"已发货": 2,
	"待发货": 1,
	"已结束": 3,
}

type Ticket struct {
	Number     string    `json:"number"`
	Password   string    `json:"password"`
	Createtime time.Time `json:"createtime"`
	Type       int64     `json:"type"`
}

type Order struct {
	TicketId      int64     `json:"ticket_id"`
	Username      string    `json:"username"`
	Phone         string    `json:"phone"`
	Address       string    `json:"address"`
	Createtime    time.Time `json:"createtime"`
	ShunfengOrder string    `json:"shunfeng_order"`
	Status        int64     `json:"status"`
}

func main() {
	mysqlUser := os.Getenv("MYSQL_USER")
	mysqlPwd := os.Getenv("MYSQL_PWD")
	mysqlServer := os.Getenv("MYSQL_SERVER")
	db, err := sql.Open("mysql", fmt.Sprintf("%v:%v@tcp(%v)/pangxie", mysqlUser, mysqlPwd, mysqlServer))
	if err != nil {
		panic(err)
	}
	defer db.Close()
	// See "Important settings" section.
	db.SetConnMaxLifetime(time.Minute * 3)
	db.SetMaxOpenConns(10)
	db.SetMaxIdleConns(10)
	tickets, usedTickets := ReadExcelTicket()
	fmt.Printf("Total tockets : %v\n", len(tickets))
	ticketIdMap := InsertTickets(db, tickets)
	orders := ReadExcelOrder(ticketIdMap)
	gOrders := GenerateUsedOrders(orders, usedTickets, ticketIdMap)
	orders = append(orders, gOrders...)
	fmt.Printf("used tickets: %v\n", len(usedTickets))
	InsertOrder(db, orders)

}

func ReadExcelTicket() ([]Ticket, []string) {
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
	var usedTickets = make([]string, 0)
	for _, item := range sheetList {
		table, err := f.GetRows(item, excelize.Options{RawCellValue: true})
		fmt.Printf("sheet %v length %v \n", item, len(table))
		if err != nil {
			panic(err)
		}
		for _, num := range table[1:] {
			if len(num) < 2 {
				fmt.Printf("ignored ticket: %v\n", num)
				continue
			}
			if strings.Trim(num[0], "") == "" && strings.Trim(num[1], "") == "" {
				fmt.Printf("ignored ticket 2: %v\n", num)
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
			if ticketItem.Type == 1 {
				usedTickets = append(usedTickets, ticketItem.Number)
			}
			ticketItem.Type = sheetTypeMap[item]
			tickets = append(tickets, ticketItem)
		}
	}
	return tickets, usedTickets
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
	fmt.Printf("successfully inserted %v tickets\n", count)
	return ticketNumIdMap
}

func ReadExcelOrder(ticketIdMap map[string]int64) (orders []Order) {
	f, err := excelize.OpenFile("order.xlsx")
	if err != nil {
		panic(err)
	}
	defer func() {
		if err = f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	sheetList := f.GetSheetList()
	for _, item := range sheetList {
		table, err := f.GetRows(item, excelize.Options{RawCellValue: true})
		if err != nil {
			panic(err)
		}
		for _, o := range table[1:] {
			var orderItem Order
			orderItem.Username = o[0]
			orderItem.Phone = o[1]
			orderItem.Address = o[2]
			orderItem.ShunfengOrder = o[3]
			orderItem.TicketId = ticketIdMap[o[4]]
			t := o[7]
			excelStartDate := time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
			createTime := time.Now()
			if t != "" {
				days, _ := strconv.ParseFloat(t, 64)
				createTime = excelStartDate.Add(time.Duration(days*86400) * time.Second)
			}
			orderItem.Createtime = createTime
			statusText := o[8]
			var status int64 = 3
			if statusText != "" {
				if s, ok := orderTextMap[statusText]; ok {
					status = s
				}
			}
			orderItem.Status = status
			orders = append(orders, orderItem)
		}
	}
	return orders
}

func InsertOrder(db *sql.DB, orders []Order) {
	stmt, err := db.Prepare("INSERT INTO order_record(ticket_id,username,phone,address,shunfeng_order,createtime,status) VALUES (?,?,?,?,?,?,?)")
	if err != nil {
		panic(err)
	}
	count := 0
	for _, t := range orders {
		_, err := stmt.Exec(t.TicketId, t.Username, t.Phone, t.Address, t.ShunfengOrder, t.Createtime, t.Status)
		if err != nil {
			fmt.Errorf("insert order [%v] error: %v\n", t.ShunfengOrder, err)
		}
		count++
		//id, err := res.LastInsertId()
		//if err != nil {
		//	fmt.Errorf("get insert id error: %v\n", err)
		//	continue
		//}
		//ticketNumIdMap[t.Number] = id
	}
	fmt.Printf("successfully inserted %v records\n", count)
	//fmt.Printf("ticket number and id map: %v\n", ticketNumIdMap)
	//return ticketNumIdMap
}

func GenerateUsedOrders(orders []Order, usedTickets []string, ticketIdMap map[string]int64) (gOrders []Order) {
	var ticketIds []int64
	for _, number := range usedTickets {
		ticketId := ticketIdMap[number]
		matched := false
		for _, order := range orders {
			if ticketId == order.TicketId {
				matched = true
			}
		}
		if matched == false {
			ticketIds = append(ticketIds, ticketId)
		}
	}
	fmt.Printf("used tickets with no order: %v\n", ticketIds)
	for _, id := range ticketIds {
		var orderItem Order
		orderItem.ShunfengOrder = ""
		orderItem.Status = 3
		orderItem.TicketId = id
		orderItem.Username = "unknown"
		orderItem.Phone = "unknown"
		orderItem.Address = "unknown"
		orderItem.Createtime = time.Now()
		gOrders = append(gOrders, orderItem)
	}
	return gOrders
}
