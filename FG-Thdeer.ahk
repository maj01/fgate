; برمجية نسخ التحاضير من إكسل إلى بوابة المستقبل من خلال محاكاة ضغطات لوحة المفاتيح
; بالاعتماد على autoHotKey
; يلزم توفر البرنامج على الجهاز
; autohotkey.com
; إعداد/ محمد الجارالله

Gui, Color, 4493a5
Gui, Font, s13, Segoe UI
Gui, Add, Text,w300 y15 Center cWhite, لبدء النسخ يجب أولا فتح الإكسل على ورقة التحضير ثم تصغيره ثم جعل صفحة بوابة المستقبل الخاصة بالتحضير مفتوحة في المقدمة وجعل مؤشر الكتابة على الحقل المسمى الوحدة ثم اضغط بدء النسخ
Gui, Add, Text,w300 Center cWhite, يلزم عدم تحريك الفأرة أو لوحة المفاتيح أثناء النسخ
Gui, Add, Button, default w300 gStart, ابدأ النسخ
Gui, Show,w320, تحضير بوابة المستقبل
return

Start:
ButtonOK:
Gui, Minimize

SetTitleMatchMode, 2
SetTitleMatchMode, Slow

if WinActive("التحضير") or WinActive("ahk_class" . ClassName)
{
		ClipSaved := ClipboardAll
		try	XL := ComObjActive("Excel.Application")
		catch{
			MsgBox , 0, ناسخ التحضير, لم يتم العثور على إكسل. يجب أن يكون إكسل مفتوح على الورقة المطلوب النسخ منها.
			Gui, Restore
			return
		}

		For cell in XL.Range["B4:B29"]{
				clipboard := ""
				clipboard := % cell.text
				;ClipWait
				if ((A_Index <= 2) or (A_Index >= 10 and A_Index <= 20)){
					if (A_Index >= 10 and A_Index <=17){
						Send {Tab}
						Sleep 200
					}
					Send ^a
					Sleep 400
					Send ^v
				}
				else if ((A_Index >= 3 and A_Index <= 9) or (A_Index >= 21 and A_Index<= 25)){
					if (cell.text="T"){
							Send {Space}
							Sleep 300
					}
				}
				else{
					Send {Tab 6}
					break
				}
				Send {Tab}
				Sleep 200
		}
		Sleep 200
		Send {Enter}
		Sleep 1000

		Clipboard := ClipSaved
		ClipSaved := ""
}
else
	MsgBox, يجب أن تكون صفحة التحضير ببوابة المستقبل مفتوحة وفي المقدمة

Gui, Restore
return

GuiClose:
ExitApp
