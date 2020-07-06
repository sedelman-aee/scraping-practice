package main

import (
    "fmt"
    "strconv"
    "strings"
    "github.com/360EntSecGroup-Skylar/excelize"
    "github.com/gocolly/colly"
)

func main() {

    c := colly.NewCollector()
    data := [][]string{}
    name := ""

    //Scrape committee name
    c.OnHTML("h1", func(e *colly.HTMLElement) {
    	name = strings.TrimSpace(e.Text)
    })

    //Scrape meeting materials
    c.OnHTML(".accordion", func(e *colly.HTMLElement) {
        
        meetingDate := e.ChildText(".date")

        e.ForEach(".meetingMaterial", func(_ int, elem *colly.HTMLElement) {
            date := elem.ChildText("span")
            docName := elem.ChildText("a[href]")
            link := "https://www.pjm.com" + elem.ChildAttr("a[href]","href")
            docType := elem.ChildText("i")

            entry := []string{docName, link, date, meetingDate, docType}
            data = append(data, entry)
        })

    })

    //Print when visiting
    c.OnRequest(func(r *colly.Request) {
        fmt.Println("Visiting", r.URL.String())
    })

    //Visit site
    c.Visit("https://www.pjm.com/committees-and-groups/committees/mrc")

    //Write to Excel file
    f := excelize.NewFile()

    f.SetCellValue("Sheet1", "A1", name+" Documents")
    headers := map[string]string{"A2": "Name", "B2": "Link", "C2": "Published On", "D2": "Meeting Date", "E2": "Type"}
    for k, v := range headers {
        f.SetCellValue("Sheet1", k, v)
    }

    for i := range data {
        row := strconv.Itoa(i+3)
        newRow := map[string]string{"A"+row: data[i][0], "B"+row: data[i][1], "C"+row: data[i][2], "D"+row: data[i][3], "E"+row: data[i][4]}
        for k, v := range newRow {
            f.SetCellValue("Sheet1", k, v)
        }
    }

    if err := f.SaveAs("pjm-mrc-docs.xlsx"); err != nil {
        println(err.Error())
    }
    
}
