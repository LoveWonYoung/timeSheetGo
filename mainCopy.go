package main

import (
	"bufio"
	"fmt"
	"os"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type yfTimeSheet struct {
	name       string
	_timesheet []everyDay
}

// 03-11-24

// isWeekday checks if the given time is a weekday (Monday to Friday).
func isWeekday(t time.Time) bool {
	return t.Weekday() != time.Saturday && t.Weekday() != time.Sunday
}

// workdaysCount calculates the number of workdays between two dates. 找到指定区间工作日
func workdaysCount(startDate, endDate time.Time) int {
	days := 0
	currentDate := startDate

	for currentDate.Before(endDate) || currentDate.Equal(endDate) {
		if isWeekday(currentDate) {
			days++
		}
		currentDate = currentDate.Add(24 * time.Hour) // Add one day
	}

	return days
}

// 读取excel
type everyDay struct {
	ProjectDay, ProjectName, ProjectCode string
	ProjectTimes                         float64
}

func init() {
	_ = createDirIfNotExists("result")

}
func createDirIfNotExists(dirname string) error {
	// 拼接当前目录和要创建的目录名
	currentDir, err := os.Getwd()
	if err != nil {
		return err
	}
	fullPath := currentDir + "/" + dirname

	// 使用MkdirAll创建目录，权限为0755
	err = os.MkdirAll(fullPath, 0755)
	if err != nil {
		// 如果目录已存在，MkdirAll会返回ErrExist，我们可以忽略这个错误
		if !os.IsExist(err) {
			return err
		}
	}
	return nil
}
func GetProjectCode(s string) string {
	ex := regexp.MustCompile("\\d{8}|[A-Z][A-Z]\\d{6}")

	ret := ex.FindAllString(strings.ToUpper(s), 1)
	if len(ret) != 0 {
		return ret[0]
	} else {
		return ""
	}
}

func sliceDeduplication(allCode []string) []string {
	keys := make(map[string]bool)
	var list []string
	for _, entry := range allCode {
		if _, found := keys[entry]; !found {
			keys[entry] = true
			list = append(list, entry)
		}
	}
	return list
}

/*
 */
func readXlsx(name string) []everyDay {
	f, _ := excelize.OpenFile(fmt.Sprintf("\\"+"path", name))
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()

	// f.GetCellValue("Daily Report", fmt.Sprintf("F%v", i+2))  获取Daily Report 的 F列，第i+2行数据的值
	/*
		F 项目名称
		E 日期
		I 工时
	*/

	var personDict []everyDay
	var projectCodeList []string
	for line := 0; ; line++ {
		sheetPName, _ := f.GetCellValue("Daily Report", fmt.Sprintf("F%v", line+2))
		sheetPDay, _ := f.GetCellValue("Daily Report", fmt.Sprintf("E%v", line+2))
		tmpPTimes, _ := f.GetCellValue("Daily Report", fmt.Sprintf("I%v", line+2))
		sheetPCode := GetProjectCode(sheetPName)
		sheetPTimes, _ := strconv.ParseFloat(tmpPTimes, 64) //转浮点

		// 判断以下数据为空直接break
		b, _ := f.GetCellValue("Daily Report", fmt.Sprintf("B%v", line+2))
		c, _ := f.GetCellValue("Daily Report", fmt.Sprintf("C%v", line+2))
		d, _ := f.GetCellValue("Daily Report", fmt.Sprintf("D%v", line+2))
		g, _ := f.GetCellValue("Daily Report", fmt.Sprintf("G%v", line+2))
		h, _ := f.GetCellValue("Daily Report", fmt.Sprintf("H%v", line+2))

		if sheetPName != "" && sheetPDay != "" && sheetPTimes != 0 && sheetPCode != "" {
			temp := everyDay{
				ProjectDay:   sheetPDay,
				ProjectName:  sheetPName,
				ProjectCode:  sheetPCode,
				ProjectTimes: sheetPTimes,
			}
			projectCodeList = append(projectCodeList, sheetPCode)
			personDict = append(personDict, temp)
		} else if b == "" && c == "" && d == "" && g == "" && h == "" && tmpPTimes == "" && sheetPDay == "" && sheetPName == "" {
			break
		} // 危险操作
	}
	return personDict
}
func (y yfTimeSheet) totalTimes() float64 {
	ret := 0.0
	for _, i := range y._timesheet {
		ret += i.ProjectTimes
	}
	return ret
}

// 返回去重项目号
func (y yfTimeSheet) getPersonProjectCode() []string {
	var allCode []string
	for _, code := range y._timesheet {
		allCode = append(allCode, code.ProjectCode)
	}
	return sliceDeduplication(allCode)
}

func (y yfTimeSheet) everyProjectTimes() {
	for _, i := range y.getPersonProjectCode() {
		var temp float64
		temp = 0
		for _, j := range y._timesheet {
			if j.ProjectCode == i {
				temp += j.ProjectTimes
			}
		}
		fmt.Printf("项目%v工时%v\n", i, temp)
	}
}

// 得到 上月21->此月20的总工时
func (y yfTimeSheet) signTimes() float64 {
	now := time.Now()         //获取当前时间
	const layout = "01-02-06" // excel 时间模板转换
	var tempTimes float64 = 0
	for _, i := range y._timesheet {
		stamp, _ := time.Parse(layout, i.ProjectDay)
		if (stamp.Year() == now.Year() && stamp.Month() == now.Month() && stamp.Day() <= 20) ||
			(stamp.Year() == now.Year() && stamp.Month() == now.Month()-1 && stamp.Day() >= 21) {
			tempTimes += i.ProjectTimes
		}
	}
	fmt.Printf("%v\n", tempTimes)
	return tempTimes
}
func (y yfTimeSheet) signTimesEveryProject() {
	f := excelize.NewFile()
	defer func() {
		if err := f.Close(); err != nil {
			fmt.Println(err)
		}
	}()
	// 创建一个工作表
	index, _ := f.NewSheet(fmt.Sprintf("%v", y.name))
	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "A1", "项目号")
	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "B1", "上月21->此月20的总工时")
	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "E1", "项目名称")
	now := time.Now()         //获取当前时间
	const layout = "01-02-06" // excel 时间模板转换
	var digtalTimes float64 = 0
	var line = 2

	for _, code := range y.getPersonProjectCode() {
		var codeTempTimes float64 = 0
		var tempName = ""
		for _, sign := range y._timesheet {
			parseTime, _ := time.Parse(layout, sign.ProjectDay)
			if (code == sign.ProjectCode && (parseTime.Year() == now.Year() && parseTime.Month() == now.Month() && parseTime.Day() <= 20)) ||
				(code == sign.ProjectCode && (parseTime.Year() == now.Year() && parseTime.Month() == now.Month()-1 && parseTime.Day() >= 21)) {
				codeTempTimes += sign.ProjectTimes
				digtalTimes += sign.ProjectTimes
				tempName = sign.ProjectName
			}
		}
		if codeTempTimes != 0 {
			//fmt.Printf("项目号:%v  工时:%v\n", code, codeTempTimes)
			_ = f.SetCellValue(fmt.Sprintf("%v", y.name), fmt.Sprintf("A%d", line), code)
			_ = f.SetCellValue(fmt.Sprintf("%v", y.name), fmt.Sprintf("B%d", line), codeTempTimes)
			_ = f.SetCellValue(fmt.Sprintf("%v", y.name), fmt.Sprintf("E%d", line), tempName)
			line += 1
		}
	}
	start := time.Date(time.Now().Year(), time.Now().Month()-1, 21, 0, 0, 0, 0, time.UTC)
	end := time.Date(time.Now().Year(), time.Now().Month(), 20, 0, 0, 0, 0, time.UTC)
	workdays := workdaysCount(start, end)

	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "D1", "应填工时")
	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "D2", digtalTimes)
	irate := (digtalTimes / (float64(workdays) * 8)) * 100
	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "C1", "工时占比")
	_ = f.SetCellValue(fmt.Sprintf("%v", y.name), "C2", fmt.Sprintf("%%%.2f", irate))

	f.SetActiveSheet(index)
	// 根据指定路径保存文件
	if err := f.SaveAs(fmt.Sprintf("./result/%v_总工时%v_应填工时%v_占比%%%.2f.xlsx", y.name, digtalTimes, workdays*8, irate)); err != nil {
		fmt.Println(err)
	}
	fmt.Printf("%v 总工时 %v 应填%v工时 占比%%%.2f\n", y.name, digtalTimes, workdays*8, irate)

}

// GetBetweenDates 获取两个日期之间的日期列表/*
func GetBetweenDates(sdate, edate string) []string {
	var d []string
	timeFormatTpl := "01-02-06"
	if len(timeFormatTpl) != len(sdate) {
		timeFormatTpl = timeFormatTpl[0:len(sdate)]
	}
	date, err := time.Parse(timeFormatTpl, sdate)
	if err != nil {
		// 时间解析，异常
		return d
	}
	date2, err := time.Parse(timeFormatTpl, edate)
	if err != nil {
		// 时间解析，异常
		return d
	}
	if date2.Before(date) {
		// 如果结束时间小于开始时间，异常
		return d
	}
	// 输出日期格式固定
	timeFormatTpl = "01-02-06"
	date2Str := date2.Format(timeFormatTpl)
	d = append(d, date.Format(timeFormatTpl))
	for {
		date = date.AddDate(0, 0, 1)
		dateStr := date.Format(timeFormatTpl)
		d = append(d, dateStr)
		if dateStr == date2Str {
			break
		}
	}
	return d
}

func (y yfTimeSheet) areaProjectTims(startTime, endTime, projectCode string) float64 {
	var areaList []string
	areaList = GetBetweenDates(startTime, endTime)
	var ret float64 = 0
	for _, day := range areaList {
		for _, yfTimeSheetDay := range y._timesheet {
			if yfTimeSheetDay.ProjectDay == day && yfTimeSheetDay.ProjectCode == projectCode {
				ret += yfTimeSheetDay.ProjectTimes
			}
		}
	}
	//fmt.Printf("%v", ret)
	return ret
}

// 构造函数
func newPerson(name string) yfTimeSheet {
	return yfTimeSheet{name: name, _timesheet: readXlsx(name)}
}
func main() {
	start := time.Now() // 获取当前时间

	name := newPerson("name")
	name.signTimesEveryProject()

	cost := time.Since(start) // 计算此时与start的时间差
	println()
	fmt.Printf("生成成功，耗时: %v \n", cost)

	//fmt.Printf("%v", GetBetweenDates("02-11-24", "02-15-24"))\
	reader := bufio.NewReader(os.Stdin)
	_, _ = reader.ReadByte()
}
