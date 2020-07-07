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
    comName := ""

    //Scrape committee name
    c.OnHTML("h1", func(e *colly.HTMLElement) {
        comName = strings.TrimSpace(e.Text)
    })

    //Scrape meeting materials
    c.OnHTML(".accordion", func(e *colly.HTMLElement) {
        
        meetingDate := e.ChildText(".date")

        e.ForEach(".meetingMaterial", func(_ int, elem *colly.HTMLElement) {
            date := elem.ChildText("span")
            docName := elem.ChildText("a[href]")
            link := "https://www.pjm.com" + elem.ChildAttr("a[href]","href")
            docType := elem.ChildText("i")

            entry := []string{comName, docName, link, date, meetingDate, docType}
            data = append(data, entry)
        })

    })

    //Add other committees to the queue
    c.OnHTML("a[href]", func(e *colly.HTMLElement) {
        href := e.Attr("href")
        if strings.Contains(href, "committees-and-groups/committees") {
            c.Visit(e.Request.AbsoluteURL(href))
        }
        // } else if strings.Contains(href, "committees-and-groups/user-groups") {
        //     c.Visit(e.Request.AbsoluteURL(href))
        // } else if strings.Contains(href, "committees-and-groups/tech-change-forum") {
        //     c.Visit(e.Request.AbsoluteURL(href))
        // } else if strings.Contains(href, "committees-and-groups/subcommittees") {
        //     c.Visit(e.Request.AbsoluteURL(href))
        // } else if strings.Contains(href, "committees-and-groups/task-forces") {
        //     c.Visit(e.Request.AbsoluteURL(href))
        // } else if strings.Contains(href, "committees-and-groups/stakeholder-meetings") {
        //     c.Visit(e.Request.AbsoluteURL(href))
        // }
    })

    //Print when visiting
    c.OnRequest(func(r *colly.Request) {
        fmt.Println("Visiting", r.URL.String())
    })

    //Visit site
    c.Visit("https://www.pjm.com/committees-and-groups/committees")

    //Write to Excel file
    f := excelize.NewFile()

    headers := map[string]string{"A1": "Committee", "B1": "Name", "C1": "Link", "D1": "Published On", "E1": "Meeting Date", "F1": "Type"}
    for k, v := range headers {
        f.SetCellValue("Sheet1", k, v)
    }

    for i := range data {
        row := strconv.Itoa(i+2)
        newRow := map[string]string{"A"+row: data[i][0], "B"+row: data[i][1], "C"+row: data[i][2], "D"+row: data[i][3], "E"+row: data[i][4], "F"+row: data[i][5]}
        for k, v := range newRow {
            f.SetCellValue("Sheet1", k, v)
        }
    }

    if err := f.SaveAs("pjm-committee-docs.xlsx"); err != nil {
        println(err.Error())
    }
    
}

