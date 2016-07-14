# RTCalendar
A simple DataPicker component.

## Quick start

### Using with workbook
```
Public Sub sample()
  Dim cal As RVCalendar
  Dim startRow As Integer
  Dim startCol As Integer
  Dim span As Integer
  Dim myYear As Integer
  Dim myMonth As Integer
  Dim myDay As Integer
  dim maxMonth as Integer
  Dim i As Integer

  Set cal = New RVCalendar
  cal.init(Date)

  startRow = 1
  startCol = 1
  span     = 10
  myYear   = 2016
  myMonth  = 1
  myDay    = 1
  maxMonth = 12

  Application.ScreenUpdating = False
  ActiveSheet.EnableCalculation = False
  Application.Calculation = xlCalculationManual

  For i = startRow To maxMonth
    myMonth = i
    cal.calDate = DateSerial(myYear, myMonth, myDay)
    Call cal.render(ActiveSheet, startRow, startCol)
    startRow = span * i
  Next

  Application.Calculation = xlCalculationAutomatic
  ActiveSheet.EnableCalculation = True
  Application.ScreenUpdating = True

  Set cal = Nothing
End Sub
```

### Using with Userform
```
Private Sub SomeTextBox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Load RVCalendarForm
  RVCalendarForm.form = Me
End Sub

Private Sub UserForm_Initialize()
  Me.Show vbModeless
End Sub
```

## Usage

### RVCalendar Class

#### color
```
' Saturday Color
saturdayBgColor = RGB(0, 125, 204)
saturdayColor   = RGB(255, 255, 255)

' Holiday Color
offdayBgColor = RGB(204, 0, 0)
offdayColor   = RGB(255, 255, 255)

' Workday Color
ondayBgColor = RGB(100, 100, 100)
ondayColor   = RGB(255, 255, 255)

' Header Color
fieldBgColor = RGB(50, 50, 50)
fieldColor   = RGB(255, 255, 255)

' today Color
highlightBgColor = RGB(204, 190, 0)
highlightColor   = RGB(255, 255, 255)
```

#### fontSize
```
fontSize = 10
```

#### outputType
`outputType` can be `book` or `form`.

When you set this value as `book`, a `Date` value will return in `ActiveCell.value` after the click event.

When you set this value as `form`, a `Date` value will return in `ActiveControl.value` after the click event.

The default value is `book`.

```
outputType = "book"
```

#### lang
Day of the week in `en` or `ja`.

```
cal.lang = "ja"
```

#### calDate
A value of Today.
```
cal.calDate = Date
```

#### holidayList
A list of holiday. Default is japanese holiday (2013-2032).

```
dim hList as Object
set hList = CreateObject("System.Collections.ArrayList")
hList.add CDate("2016/1/1")
holidayList = hList
set hList = nothing
```


#### Property

##### Let
```
Property Let calDate(ByVal val As Date)
Property Let calYear(ByVal val As Integer)
Property Let calMonth(ByVal val As Integer)
Property Let calDay(ByVal val As Integer)
Property Let saturdayBgColor(ByVal val As Long)
Property Let saturdayColor(ByVal val As Long)
Property Let ondayBgColor(ByVal val As Long)
Property Let ondayColor(ByVal val As Long)
Property Let offdayBgColor(ByVal val As Long)
Property Let offdayColor(ByVal val As Long)
Property Let fieldBgColor(ByVal val As Long)
Property Let highlightBgColor(ByVal val As Long)
Property Let highlightColor(ByVal val As Long)
Property Let fieldColor(ByVal val As Long)
Property Let fontSize(ByVal val As Integer)
Property Let lang(ByVal val As String)
Property Let holidayList(ByVal val As Object)
Property Let outputType(ByVal val As String)
Property Let selectedDate(ByVal val As Date)
```

##### Get
```
Property Get calDate() As Date
Property Get calYear() As Integer
Property Get calMonth() As Integer
Property Get calDay() As Integer
Property Get calWday() As Integer
Property Get prevYear() As Date
Property Get nextYear() As Date
Property Get prevMonth() As Date
Property Get nextMonth() As Date
Property Get prevDay() As Date
Property Get nextDay() As Date
Property Get saturdayBgColor() As Long
Property Get saturdayColor() As Long
Property Get ondayBgColor() As Long
Property Get ondayColor() As Long
Property Get offdayBgColor() As Long
Property Get offdayColor() As Long
Property Get fieldBgColor() As Long
Property Get fieldColor() As Long
Property Get highlightBgColor() As Long
Property Get highlightColor() As Long
Property Get fontSize() As Integer
Property Get lang() As String
Property Get holidayList() As Object
Property Get weekSize() As Integer
Property Get monthSize() As Integer
Property Get outputType() As String
Property Get selectedDate() As Date
```

#### Method
```
Public Sub init(ByVal t As Date)
Public Function isLeapYear(ByVal Y As Integer) As Boolean
Public Function numOfDays() As Integer
Public Function firstDayInMonth(ByVal dateVal As Date) As Date
Public Function lastDayInMonth(ByVal dateVal As Date) As Date
Public Function monthlyList() As Object
Public Function getCalWday(ByVal val As Date) As Integer
Public Function calHead() As Object
Public Sub render(ByVal sheet As Worksheet, ByVal firstRow As Long, ByVal firstColumn As Long, Optional ByVal hasHead As Boolean = True)
```


### RVDateItem Class

#### Property

##### Let
```
Property Let color(ByVal val As Long)
Property Let bgColor(ByVal val As Long)
Property Let fontSize(ByVal val As Integer)
Property Let value(ByVal val As Variant)
Property Let dateValue(ByVal val As Variant)
Property Let isHoliday(ByVal val As Boolean)
```

##### Get
```
Property Get color() As Long
Property Get bgColor() As Long
Property Get fontSize() As Integer
Property Get value() As Variant
Property Get dateValue() As Variant
Property Get isHoliday() As Boolean
```

#### Method
```
Public Sub init(ByVal iValue As Integer, ByVal iDateValue As Date, ByVal iIsHoliday As Boolean, ByVal iFontSize As Integer, ByVal iColor As Long, ByVal iBgColor As Long)
```


### RVCalendarForm

#### Property

##### Let
```
Property Let cal(ByVal val As RVCalendar)
Property Let form(ByVal val As MSForms.UserForm)
```

##### Get
```
Property Get cal() As RVCalendar
Property Get form() As MSForms.UserForm
Property Get formatDate() As String
```
