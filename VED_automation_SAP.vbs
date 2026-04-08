If Not IsObject(application) Then

   Set SapGuiAuto  = GetObject("SAPGUI")

   Set application = SapGuiAuto.GetScriptingEngine

End If

If Not IsObject(connection) Then

   Set connection = application.Children(0)

End If

If Not IsObject(session) Then

   Set session    = connection.Children(0)

End If

 

' ВКЛЮЧАЕМ РЕЖИМ ОЖИДАНИЯ (чтобы скрипт не вылетал, пока крутится индикатор)

session.TestToolMode = 1

 

' --- НАСТРОЙКИ ---

Dim zavody, mesyacy, god, path, zavod, mes, l_day, fname

zavody = Array("0001", "0002", "0003", "0004") ' Впишите сюда ВСЕ свои заводы через запятую

mesyacy = Array("01", "02", "03") ' Месяцы (накопительно)

god = "2026"

path = "C:\Users\ваш путь\files"

' -----------------

 

On Error Resume Next ' Пропускать ошибки, если данных по заводу нет

 

For Each zavod In zavody

    For Each mes In mesyacy

       

        ' 1. Логика последнего дня месяца для даты окончания

        Select Case mes

            Case "02": l_day = "28"

            Case "04", "06", "09", "11": l_day = "30"

            Case Else: l_day = "31"

        End Select

 

        ' 2. Вход в транзакцию

        session.findById("wnd[0]/tbar[0]/okcd").text = "/nJ3RFLVMOBVED"

        session.findById("wnd[0]").sendVKey 0

 

        ' 3. Заполнение полей (накопительный итог с 1 января)

        session.findById("wnd[0]/usr/ctxtPA_BUKRS").text = "" ' БЕ пустая

        session.findById("wnd[0]/usr/ctxtSO_WERKS-LOW").text = zavod

        session.findById("wnd[0]/usr/ctxtSO_BUDAT-LOW").text = "01.01." & god

        session.findById("wnd[0]/usr/ctxtSO_BUDAT-HIGH").text = l_day & "." & mes & "." & god

       

        ' Выбор вашего варианта просмотра

        session.findById("wnd[0]/usr/tabsTABSTRIP_TABS/tabpTAB20/ssub%_SUBSCREEN_TABS:J_3RMOBVED:0520/ctxtPA_VARNT").text = "вариант"

       

        ' ПРОСТАВЛЯЕМ ВСЕ 3 ГАЛОЧКИ

        session.findById("wnd[0]/usr/tabsTABSTRIP_TABS/tabpTAB20/ssub%_SUBSCREEN_TABS:J_3RMOBVED:0520/chkPA_ALLMA").selected = true ' Все материалы

        session.findById("wnd[0]/usr/tabsTABSTRIP_TABS/tabpTAB20/ssub%_SUBSCREEN_TABS:J_3RMOBVED:0520/chkPA_REVER").selected = true ' Сторнирования

        session.findById("wnd[0]/usr/tabsTABSTRIP_TABS/tabpTAB20/ssub%_SUBSCREEN_TABS:J_3RMOBVED:0520/chkPA_DOCS").selected = true  ' Документы

 

        ' 4. Запуск (F8)

        session.findById("wnd[0]").sendVKey 8

 

        ' 5. Экспорт (через цепочку меню из вашего пробного кода)

        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select

        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select

        session.findById("wnd[1]/tbar[0]/btn[0]").press

 

        ' 6. Сохранение файла с уникальным именем

        fname = "OSV_" & zavod & "_Jan_to_" & mes & ".xls"

       

        session.findById("wnd[1]/usr/ctxtDY_PATH").text = path

        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fname

        session.findById("wnd[1]/tbar[0]/btn[0]").press

        

        ' Если файл уже существует - подтверждаем замену (нажимаем "Да")

        If session.Children.Count > 1 Then

            session.findById("wnd[2]/tbar[0]/btn[0]").press

        End If

 

        ' Пауза 5 секунд, чтобы система "переварила" Excel

        WScript.Sleep 5000

    Next

Next

 

MsgBox "Готово! Все отчеты лежат в " & path