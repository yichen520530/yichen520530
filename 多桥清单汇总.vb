Module Module4
    Private brr(0 To 50000, 0 To 12), wb_HZ, sht_Hz, wb
    Private 钢筋所在列， wbNum, openWbNum As Int16
    Private 清单数量, 总数据数， i As Int32
    Private qLName As String
    Private qDarr(,), 未汇总细目号, 清单细目号, 清单细目号val, endrow, 错误提示, 部位1， 部位2， 部位3, 部位4， 部位5
    Private delRow As Excel.Range = Nothing
    Private maxCol, 合并单元格占位数, 图号所在列 As Int16
    Private B列最末行号, 单列合计值 As Double
    Private d = CreateObject("Scripting.Dictionary")
    'Dim Fso As Object, Folder As Object
    'Private d As New Dictionary(Of String, String)()
    Sub 多桥清单汇总()   '多数的桥梁汇总格式
        '在桥梁清单汇总表的基础上 增加计算过程函数 22.4.14
        '修改maxcol的方法，取所有单元格和的最大列，防止隐藏列汇总不了
        '修改find方法，在行或者列用完的情况下，find类不能竖向查找，会出现错误
        '22.8.3 增加汇总属于路基的400章的清单
        Dim outputarr(0 To 50000, 0 To 12)， qDarr(,), 未汇总细目号, 清单细目号, 清单细目号val, endrow, 错误提示, 部位1， 部位2， 部位3, 部位4， 部位5
        Dim delRow As Excel.Range = Nothing
        Dim maxCol, 合并单元格占位数, 图号所在列 As Int16
        Dim B列最末行号, 单列合计值 As Double
        Dim Fso As Object, Folder As Object
        Dim d = CreateObject("Scripting.Dictionary")
        'On Error GoTo line_Err
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlManual '关闭自动计算
        zlapp.ScreenUpdating = False '关闭屏幕自动刷新
        zlapp.DisplayAlerts = False '禁止显示提示和警告消息
        Fso = CreateObject("Scripting.FileSystemObject")
        Folder = Fso.GetFolder(CreateObject("Shell.Application").BrowseForFolder(0, "请选择文件夹", 0, "").Self.Path & "\")
        wb_HZ = zlapp.ActiveWorkbook.Name
        sht_Hz = zlapp.ActiveSheet.Name
        Call GetFiles(Folder, 需汇总工作薄, wbNum)
        With zlapp.ActiveWorkbook.Sheets(sht_Hz)
            .Select
            zlapp.Selection.ClearOutline
            endrow = .Cells.Find("*", .Cells(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole,
                                     SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row
            未汇总细目号 = .UsedRange.Find("未汇总细目号", LookAt:=Excel.XlLookAt.xlPart)
            If 未汇总细目号 IsNot Nothing Then
                qDarr = .Range("A4:J" & 未汇总细目号.Row - 1).value
            Else
                qDarr = .Range("A4:J" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).value
            End If
        End With
        清单数量 = 0 ： openWbNum = 0
        For wb_i = 0 To 需汇总工作薄.Length - 1
            Try
                wb = zlapp.Workbooks.Open(需汇总工作薄(wb_i).path, UpdateLinks:=0, ReadOnly:=True)    '打开   只读  
            Catch ex As Exception
                错误提示 = 需汇总工作薄(wb_i).name & " 工作薄打开错误，请修改格式后重新汇总 " & vbCrLf & 错误提示
            End Try
            zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlManual '关闭自动计算
            Try
                qLName = Right(zlapp.Sheets("桩基-扣除系梁高度").Range("a2").value, Len(zlapp.Sheets("桩基-扣除系梁高度").Range("a2").value) - 5)
            Catch ex As Exception
                qLName = wb.name
            End Try
            openWbNum += 1
            For Each sht In wb.Sheets
                If Not arrShtNo.Contains(sht.Name) Then
                    With sht
                        .AutoFilterMode = False '关闭筛选
                        清单细目号 = Nothing
                        部位1 = .UsedRange.Find("部位1", LookAt:=Excel.XlLookAt.xlWhole, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                        部位2 = .UsedRange.Find("部位2", LookAt:=Excel.XlLookAt.xlWhole, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                        部位3 = .UsedRange.Find("部位3", LookAt:=Excel.XlLookAt.xlWhole, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                        清单细目号 = .UsedRange.Find("清单细目号", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                        图号所在列 = 0
                        If 清单细目号 Is Nothing Then GoTo Nextsht
                        'maxCol = .Cells(清单清单细目号.Row, zlapp.Columns.Count).End(Excel.XlDirection.xlToLeft).Column   '最后一个清单号所在列，用来循环
                        图号所在列 = .UsedRange.Find("图号", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows).column
                        If 图号所在列 > 0 Then    '汇总到图号前一列
                            maxCol = 图号所在列 - 1
                        Else
                            maxCol = .usedrange.columns.count
                        End If
                        For j = 清单细目号.MergeArea.Columns.Count + 1 To maxCol
                            If TypeName(.Cells(清单细目号.Row - 1, j).value) = "String" Then
                                MsgBox(wb.name & "中的工作表： " & sht.name & " 第" & 清单细目号.Row - 1 & "行，第" & j & "列汇总错误")              '汇总不是数值，弹出错误提示，结束程序
                                zlapp.DisplayAlerts = True '显示提示和警告消息
                                zlapp.ScreenUpdating = True '显示屏幕自动刷新
                                Exit Sub
                            End If
                            清单细目号val = .Cells(清单细目号.Row, j).value
                            If .Cells(清单细目号.Row - 1, j) IsNot Nothing And .Cells(清单细目号.Row - 1, j).value <> 0 Then
                                If 清单细目号val IsNot Nothing Then
                                    d(清单细目号val) += .Cells(清单细目号.Row - 1, j).value
                                    单列合计值 = 0
                                    If .Cells(3, j).MergeArea.Cells(1, 1).value = .Cells(5, j).MergeArea.Cells(1, 1).value Then   '路基使用
                                        部位4 = .Cells(3, j).MergeArea.Cells(1, 1).value
                                        部位5 = ""
                                    Else
                                        部位4 = .Cells(3, j).MergeArea.Cells(1, 1).Value
                                        部位5 = .Cells(5, j).MergeArea.Cells(1, 1).Value
                                    End If
                                    For i = 6 To 清单细目号.Row - 2
                                        合并单元格占位数 = .Cells(i, j).MergeArea.Count
                                        Try
                                            If .Cells(i, j).value <> 0 Then
                                            End If
                                        Catch e As Exception
                                            .Cells(i, j).value = 0
                                            错误提示 = wb.name & " " & sht.name & " 第： " & i & " 行，第： " & j & "列" & vbCrLf & 错误提示
                                        End Try
                                        If .Cells(i, j).value Then
                                            If 部位1 Is Nothing Then
                                                brr(清单数量, 8) = .Cells(i, j).MergeArea.Cells(1, 1).value   '工程量
                                                brr(清单数量, 2) = qLName
                                                brr(清单数量, 3) = .Cells(i, 1).MergeArea.Cells(1, 1).value
                                                brr(清单数量, 4) = .Cells(i, 2).MergeArea.Cells(1, 1).value
                                                brr(清单数量, 5) = .Cells(i, 3).MergeArea.Cells(1, 1).value
                                                brr(清单数量, 9) = Replace(.Cells(i, 图号所在列).MergeArea.Cells(1, 1).value, Chr(10), "")
                                                brr(清单数量, 1) = 清单细目号val
                                                If Left(sht.name, 2) <> "桩基" And sht.name <> "支座" And Left(清单细目号val, 3) <> "403" _
                                                And Left(清单细目号val, 5).ToString <> "411-5" Then        '计算过程
                                                    If 是否输出计算式 And .cells(i, j).MergeArea.Cells(1, 1).HasFormula Then   '  And 清单细目号 IsNot Nothing 
                                                        brr(清单数量, 10) = LoopFormulaValue(.cells(i, j).MergeArea.Cells(1, 1).FormulaLocal, sht)
                                                    End If
                                                End If
                                            Else      '路基格式使用
                                                If 部位1.MergeArea.Count > 1 Then
                                                    brr(清单数量, 2) = .Cells(i, 部位1.Column).Value & "--" & .Cells(i, 部位1.Column + 2).Value
                                                Else
                                                    brr(清单数量, 2) = Replace(.Cells(i, 部位1.Column).Value, " ", "")
                                                End If
                                                If 部位3 IsNot Nothing Then
                                                    brr(清单数量, 2) = brr(清单数量, 2) & "--" & .Cells(i, 部位3.Column).Value
                                                End If
                                                If 部位2 IsNot Nothing Then
                                                    brr(清单数量, 4) = .Cells(i, 部位2.Column).Value
                                                End If
                                                brr(清单数量, 8) = .Cells(i, j).MergeArea.Cells(1, 1).value   '工程量
                                                brr(清单数量, 3) = sht.Name
                                                brr(清单数量, 5) = 部位4
                                                brr(清单数量, 6) = 部位5
                                                brr(清单数量, 9) = .Cells(i, 图号所在列).MergeArea.Cells(1, 1).Value
                                                brr(清单数量, 1) = 清单细目号val
                                            End If
                                            单列合计值 += brr(清单数量, 8)
                                            清单数量 += 1
                                        End If
                                        i = i + 合并单元格占位数 - 1
                                    Next i
                                    If Math.Abs(单列合计值 - .Cells(清单细目号.Row - 1, j).value) > 0.01 Then
                                        错误提示 = wb.name & " " & sht.name & " 第： " & j & "列汇总数据不一致，请自行检查" & vbCrLf & 错误提示
                                    End If
                                End If
                            End If

                        Next j
                    End With
                End If
Nextsht:
            Next sht
            wb.Close(False) '关闭所打开的工作簿
        Next
        '以下为合并数组
        With zlapp.Sheets(sht_Hz)
            总数据数 = 0
            For i = 1 To UBound(qDarr)
                If qDarr(i, 1) IsNot Nothing Then
                    outputarr(总数据数， 0) = qDarr(i， 1)
                    outputarr(总数据数， 1) = qDarr(i， 2)
                    outputarr(总数据数， 6) = qDarr(i， 7)
                    outputarr(总数据数， 7) = d(qDarr(i， 1))
                    d.Remove(qDarr(i， 1))                         '删除对应字典的键位
                    For 临时数据1 = 0 To 清单数量 - 1
                        If qDarr(i, 1).ToString = brr(临时数据1, 1).ToString Then
                            总数据数 += 1
                            outputarr(总数据数, 1) = brr(临时数据1, 2)
                            outputarr(总数据数, 2) = brr(临时数据1, 3)
                            outputarr(总数据数, 3) = brr(临时数据1, 4)
                            outputarr(总数据数, 4) = brr(临时数据1, 5)
                            outputarr(总数据数, 5) = brr(临时数据1, 6)
                            outputarr(总数据数, 7) = brr(临时数据1, 8)
                            outputarr(总数据数, 8) = brr(临时数据1, 9)
                            'outputarr(总数据数, 9) = brr(临时数据1, 10)
                        End If
                    Next
                    总数据数 += 1
                End If
            Next
            If .Cells(4, 1).value.ToString IsNot Nothing Then
                .Range("A4:J" & endrow + 1).ClearContents
                'zlapp.Selection.ClearOutline  '取消分组，
            End If
            With zlapp.ActiveSheet.Outline        '设置分组格式
                .AutomaticStyles = False
                .SummaryRow = Microsoft.Office.Core.XlConstants.xlAbove
                .SummaryColumn = Microsoft.Office.Core.XlConstants.xlLeft
            End With
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).value = outputarr
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Font.name = "宋体"
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Font.Size = 10
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlCenter
            .cells(4, 2).Resize(总数据数 + 1, 1).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlLeft
            .cells(4, 8).Resize(总数据数 + 1, 1).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlLeft
            zlapp.Goto(Reference:= .Cells(3, 1), Scroll:=True)
            For i = 4 To .Range("A65535").End(Excel.XlDirection.xlUp).Row
                'If StrComp(.Cells(i, 1).value, .Cells(i + 1, 1).value, 1) = 0 Then
                If .Cells(i, 1).VALUE Is Nothing Then
                    .Cells(i, 1).Rows.Group
                End If
            Next i
            endrow = .Cells.Find("*", .Cells(1, 1), Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                                     Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
            .Rows(UBound(outputarr) + 6 & ":" & endrow + 1).DELETE
            If d.count > 0 Then  '多的细目号

                B列最末行号 = (.Cells(zlapp.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row) + 1
                .Cells(B列最末行号, 2) = "以下为未汇总细目号"
                .Cells(B列最末行号 + 1, 1).Resize(d.count, 1).value = zlapp.WorksheetFunction.Transpose(d.Keys)
                .Cells(B列最末行号 + 1, 8).Resize(d.count, 1).value = zlapp.WorksheetFunction.Transpose(d.items)
                qDarr = .Range("A" & B列最末行号 + 1 & ":L" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).value
                .Range("A" & B列最末行号 + 1 & ":L" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).ClearContents
                总数据数 = 0
                ReDim outputarr(0 To UBound(outputarr), 0 To UBound(outputarr, 2))
                For i = 1 To UBound(qDarr)
                    outputarr(总数据数， 0) = qDarr(i， 1)
                    outputarr(总数据数， 1) = qDarr(i， 2)
                    outputarr(总数据数， 6) = qDarr(i， 7)
                    outputarr(总数据数， 7) = d(qDarr(i， 1))
                    d.Remove(qDarr(i, 1))
                    For 临时数据1 = 0 To 清单数量 - 1
                        If qDarr(i, 1).ToString = brr(临时数据1, 1).ToString Then
                            总数据数 += 1
                            outputarr(总数据数, 1) = brr(临时数据1, 2)
                            outputarr(总数据数, 2) = brr(临时数据1, 3)
                            outputarr(总数据数, 3) = brr(临时数据1, 4)
                            outputarr(总数据数, 4) = brr(临时数据1, 5)
                            outputarr(总数据数, 5) = brr(临时数据1, 6)
                            outputarr(总数据数, 7) = brr(临时数据1, 8)
                            outputarr(总数据数, 8) = brr(临时数据1, 9)
                        End If
                    Next
                    总数据数 += 1
                Next
                .Range("A" & B列最末行号 + 1).Resize(总数据数, UBound(outputarr, 2)).value = outputarr
                zlapp.Goto(Reference:= .Cells(B列最末行号 - 10, 1), Scroll:=True)
                MsgBox("存在未汇总细目号，请检查后重新汇总")
            End If
            '.Range("A4:Z" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).Interior.Color = 16777215
            .Calculate
        End With
        Erase outputarr
line_Exit:

        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        'Erase brr : Erase 需汇总工作薄 : r = 0
        If openWbNum <> 需汇总工作薄.Length Then
            MsgBox("汇总完成" & Chr(13) & "汇总文件夹里文件数：" & 需汇总工作薄.Length & Chr(13) & "已汇总文件数：" & openWbNum & Chr(13) & "请务必检查是否汇总完成", , "错误提示")
        End If
        If 需汇总工作薄.Length > 20 Then
            MsgBox("汇总完成" & Chr(13) & "汇总文件夹里文件数：" & 需汇总工作薄.Length & Chr(13) & "已汇总文件数：" & openWbNum & Chr(13),, "错误提示")
        End If
        If 错误提示 IsNot Nothing Then
            MsgBox(错误提示 & "数据有问题，请核实，如果不为0，请重新汇总 ")
        End If
        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlAutomatic  '工作薄开启自动重算
        Exit Sub
line_Err:
        MsgBox(Err.Number & ":" & Err.Description & Chr(13) & "请仔细检查数据。如无法处理，请联系管理员", , "错误提示")
        'Resume line_Exit
        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlAutomatic  '工作薄开启自动重算
    End Sub
    Sub 本桥清单汇总_标准()   '单做桥梁汇总格式
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlManual '关闭自动计算
        zlapp.ScreenUpdating = False '关闭屏幕自动刷新
        zlapp.DisplayAlerts = False '禁止显示提示和警告消息
        Call 收集桥梁数据()
        Call 输出数据格式_标准()
    End Sub
    Sub 本桥清单汇总_弈睿泽高速()   '单做桥梁汇总格式
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlManual '关闭自动计算
        zlapp.ScreenUpdating = False '关闭屏幕自动刷新
        zlapp.DisplayAlerts = False '禁止显示提示和警告消息
        Call 收集桥梁数据()
        Call 输出数据格式_弈睿泽高速()
    End Sub
    Sub 收集桥梁数据()
        wb_HZ = zlapp.ActiveWorkbook.Name
        sht_Hz = zlapp.ActiveSheet.Name
        'Call GetFiles(Folder, 需汇总工作薄, wbNum)
        With zlapp.ActiveWorkbook.Sheets(sht_Hz)
            .Select
            zlapp.Selection.ClearOutline
            endrow = .Cells.Find("*", .Cells(1, 1), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole,
                                     SearchOrder:=Excel.XlSearchOrder.xlByRows, SearchDirection:=Excel.XlSearchDirection.xlPrevious).Row
            未汇总细目号 = .UsedRange.Find("未汇总细目号", LookAt:=Excel.XlLookAt.xlPart)
            If 未汇总细目号 IsNot Nothing Then
                qDarr = .Range("A4:M" & 未汇总细目号.Row - 1).value
            Else
                qDarr = .Range("A4:M" & .Cells(zlapp.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row).value
            End If
        End With
        清单数量 = 0 ： openWbNum = 0
        Try
            qLName = Right(zlapp.Sheets("基础数据").Range("a2").value, Len(zlapp.Sheets("基础数据").Range("a2").value) - 5)
        Catch ex As Exception
            qLName = wb_HZ
        End Try
        'openWbNum += 1
        For Each sht In zlapp.ActiveWorkbook.Sheets
            If Not arrShtNo.Contains(sht.Name) Then
                With sht
                    .AutoFilterMode = False '关闭筛选
                    清单细目号 = Nothing
                    部位1 = .UsedRange.Find("部位1", LookAt:=Excel.XlLookAt.xlWhole, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                    部位2 = .UsedRange.Find("部位2", LookAt:=Excel.XlLookAt.xlWhole, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                    部位3 = .UsedRange.Find("部位3", LookAt:=Excel.XlLookAt.xlWhole, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                    清单细目号 = .UsedRange.Find("清单细目号", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows)  '查找部分，先查找列
                    图号所在列 = 0
                    If 清单细目号 Is Nothing Then GoTo Nextsht
                    'maxCol = .Cells(清单清单细目号.Row, zlapp.Columns.Count).End(Excel.XlDirection.xlToLeft).Column   '最后一个清单号所在列，用来循环
                    图号所在列 = .UsedRange.Find("图号", LookAt:=Excel.XlLookAt.xlPart, SearchOrder:=Excel.XlSearchOrder.xlByRows).column
                    If 图号所在列 > 0 Then    '汇总到图号前一列
                        maxCol = 图号所在列 - 1
                    Else
                        maxCol = .usedrange.columns.count
                    End If
                    For j = 清单细目号.MergeArea.Columns.Count + 1 To maxCol
                        If TypeName(.Cells(清单细目号.Row - 1, j).value) = "String" Then
                            MsgBox(wb_HZ & "中的工作表： " & sht.name & " 第" & 清单细目号.Row - 1 & "行，第" & j & "列汇总错误")              '汇总不是数值，弹出错误提示，结束程序
                            zlapp.DisplayAlerts = True '显示提示和警告消息
                            zlapp.ScreenUpdating = True '显示屏幕自动刷新
                            Exit Sub
                        End If
                        清单细目号val = .Cells(清单细目号.Row, j).value
                        If .Cells(清单细目号.Row - 1, j) IsNot Nothing And .Cells(清单细目号.Row - 1, j).value <> 0 Then
                            If 清单细目号val IsNot Nothing Then
                                d(清单细目号val) += .Cells(清单细目号.Row - 1, j).value
                                'If d.ContainsKey(清单细目号val.ToString) Then
                                '    d(清单细目号val.ToString) += .Cells(清单细目号.Row - 1, j).value
                                'Else
                                '    d(清单细目号val.ToString) = .Cells(清单细目号.Row - 1, j).value
                                'End If
                                单列合计值 = 0
                                If .Cells(3, j).MergeArea.Cells(1, 1).value = .Cells(5, j).MergeArea.Cells(1, 1).value Then   '路基使用
                                    部位4 = .Cells(3, j).MergeArea.Cells(1, 1).value
                                    部位5 = ""
                                Else
                                    部位4 = .Cells(3, j).MergeArea.Cells(1, 1).Value
                                    部位5 = .Cells(5, j).MergeArea.Cells(1, 1).Value
                                End If
                                For i = 6 To 清单细目号.Row - 2
                                    合并单元格占位数 = .Cells(i, j).MergeArea.Count
                                    Try
                                        If .Cells(i, j).value <> 0 Then
                                        End If
                                    Catch e As Exception
                                        .Cells(i, j).value = 0
                                        错误提示 = wb_HZ & " " & sht.name & " 第： " & i & " 行，第： " & j & "列" & vbCrLf & 错误提示
                                    End Try
                                    If .Cells(i, j).value Then
                                        If 部位1 Is Nothing Then
                                            brr(清单数量, 8) = .Cells(i, j).MergeArea.Cells(1, 1).value   '工程量
                                            brr(清单数量, 2) = qLName  '桥梁名称
                                            brr(清单数量, 3) = .Cells(i, 1).MergeArea.Cells(1, 1).value
                                            brr(清单数量, 4) = .Cells(i, 2).MergeArea.Cells(1, 1).value
                                            brr(清单数量, 5) = .Cells(i, 3).MergeArea.Cells(1, 1).value
                                            brr(清单数量, 9) = Replace(.Cells(i, 图号所在列).MergeArea.Cells(1, 1).value, Chr(10), "")
                                            brr(清单数量, 1) = 清单细目号val
                                            brr(清单数量, 6) = sht.name   '工作表名称
                                            If Left(sht.name, 2) <> "桩基" And sht.name <> "支座" And Left(清单细目号val, 3) <> "403" _
                                            And Left(清单细目号val, 5).ToString <> "411-5" Then        '计算过程
                                                If 是否输出计算式 And .cells(i, j).MergeArea.Cells(1, 1).HasFormula Then   '  And 清单细目号 IsNot Nothing 
                                                    brr(清单数量, 10) = LoopFormulaValue(.cells(i, j).MergeArea.Cells(1, 1).FormulaLocal, sht)
                                                End If
                                            End If
                                        Else      '路基格式使用
                                            If 部位1.MergeArea.Count > 1 Then
                                                brr(清单数量, 2) = .Cells(i, 部位1.Column).Value & "--" & .Cells(i, 部位1.Column + 2).Value
                                            Else
                                                brr(清单数量, 2) = Replace(.Cells(i, 部位1.Column).Value, " ", "")
                                            End If
                                            If 部位3 IsNot Nothing Then
                                                brr(清单数量, 2) = brr(清单数量, 2) & "--" & .Cells(i, 部位3.Column).Value
                                            End If
                                            If 部位2 IsNot Nothing Then
                                                brr(清单数量, 4) = .Cells(i, 部位2.Column).Value
                                            End If
                                            brr(清单数量, 8) = .Cells(i, j).MergeArea.Cells(1, 1).value   '工程量
                                            brr(清单数量, 3) = sht.Name
                                            brr(清单数量, 5) = 部位4
                                            brr(清单数量, 6) = 部位5
                                            brr(清单数量, 9) = .Cells(i, 图号所在列).MergeArea.Cells(1, 1).Value
                                            brr(清单数量, 1) = 清单细目号val
                                        End If
                                        单列合计值 += brr(清单数量, 8)
                                        清单数量 += 1
                                    End If
                                    i = i + 合并单元格占位数 - 1
                                Next i
                                If Math.Abs(单列合计值 - .Cells(清单细目号.Row - 1, j).value) > 0.01 Then
                                    错误提示 = wb_HZ & " " & sht.name & " 第： " & j & "列汇总数据不一致，请自行检查" & vbCrLf & 错误提示
                                End If
                            End If
                        End If

                    Next j
                End With
            End If
Nextsht:
        Next sht
        'wb.Close(False) '关闭所打开的工作簿
        'Next
        '以下为合并数组
    End Sub
    Sub 输出数据格式_标准()
        Dim outputarr(0 To 50000, 0 To 14)
        With zlapp.Sheets(sht_Hz)
            总数据数 = 0
            For i = 1 To UBound(qDarr)
                If qDarr(i, 1) IsNot Nothing Then
                    outputarr(总数据数， 0) = qDarr(i， 1)
                    outputarr(总数据数， 1) = qDarr(i， 2)
                    outputarr(总数据数， 6) = qDarr(i， 7)
                    outputarr(总数据数， 7) = d(qDarr(i， 1))
                    d.Remove(qDarr(i， 1))                         '删除对应字典的键位
                    For 临时数据1 = 0 To 清单数量 - 1
                        If qDarr(i, 1).ToString = brr(临时数据1, 1).ToString Then
                            总数据数 += 1
                            outputarr(总数据数, 1) = brr(临时数据1, 2)
                            outputarr(总数据数, 2) = brr(临时数据1, 3)
                            outputarr(总数据数, 3) = brr(临时数据1, 4)
                            outputarr(总数据数, 4) = brr(临时数据1, 5)
                            outputarr(总数据数, 5) = brr(临时数据1, 6)
                            outputarr(总数据数, 7) = brr(临时数据1, 8)
                            outputarr(总数据数, 8) = brr(临时数据1, 9)
                            'outputarr(总数据数, 9) = brr(临时数据1, 10)
                        End If
                    Next
                    总数据数 += 1
                End If
            Next
            If .Cells(4, 1).value IsNot Nothing Then
                .Range("A4:M" & endrow + 1).ClearContents
                'zlapp.Selection.ClearOutline  '取消分组，
            End If
            With zlapp.ActiveSheet.Outline        '设置分组格式
                .AutomaticStyles = False
                .SummaryRow = Microsoft.Office.Core.XlConstants.xlAbove
                .SummaryColumn = Microsoft.Office.Core.XlConstants.xlLeft
            End With
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).value = outputarr
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Font.name = "宋体"
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Font.Size = 10
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlCenter
            .cells(4, 2).Resize(总数据数 + 1, 1).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlLeft
            .cells(4, 8).Resize(总数据数 + 1, 1).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlLeft
            zlapp.Goto(Reference:= .Cells(3, 1), Scroll:=True)
            For i = 4 To .Range("A65535").End(Excel.XlDirection.xlUp).Row
                'If StrComp(.Cells(i, 1).value, .Cells(i + 1, 1).value, 1) = 0 Then
                If .Cells(i, 1).VALUE Is Nothing Then
                    .Cells(i, 1).Rows.Group
                End If
            Next i
            endrow = .Cells.Find("*", .Cells(1, 1), Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                                     Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
            .Rows(UBound(outputarr) + 6 & ":" & endrow + 1).DELETE
            If d.count > 0 Then  '多的细目号

                B列最末行号 = (.Cells(zlapp.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row) + 1
                .Cells(B列最末行号, 2) = "以下为未汇总细目号"
                .Cells(B列最末行号 + 1, 1).Resize(d.count, 1).value = zlapp.WorksheetFunction.Transpose(d.Keys)
                Dim arr_d() As String = d.Values.ToArray()
                '.Cells(B列最末行号 + 1, 8).Resize(d.Count, 1).value = zlapp.WorksheetFunction.Transpose(d.items)
                .Cells(B列最末行号 + 1, 8).Resize(d.Count, 1).value = zlapp.WorksheetFunction.Transpose(arr_d)
                qDarr = .Range("A" & B列最末行号 + 1 & ":L" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).value
                .Range("A" & B列最末行号 + 1 & ":L" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).ClearContents
                总数据数 = 0
                ReDim outputarr(0 To UBound(outputarr), 0 To UBound(outputarr, 2))
                For i = 1 To UBound(qDarr)
                    outputarr(总数据数， 0) = qDarr(i， 1)
                    outputarr(总数据数， 1) = qDarr(i， 2)
                    outputarr(总数据数， 6) = qDarr(i， 7)
                    outputarr(总数据数， 7) = d(qDarr(i， 1))
                    d.Remove(qDarr(i, 1))
                    For 临时数据1 = 0 To 清单数量 - 1
                        If qDarr(i, 1).ToString = brr(临时数据1, 1).ToString Then
                            总数据数 += 1
                            outputarr(总数据数, 1) = brr(临时数据1, 2)
                            outputarr(总数据数, 2) = brr(临时数据1, 3)
                            outputarr(总数据数, 3) = brr(临时数据1, 4)
                            outputarr(总数据数, 4) = brr(临时数据1, 5)
                            outputarr(总数据数, 5) = brr(临时数据1, 6)
                            outputarr(总数据数, 7) = brr(临时数据1, 8)
                            outputarr(总数据数, 8) = brr(临时数据1, 9)
                        End If
                    Next
                    总数据数 += 1
                Next
                .Range("A" & B列最末行号 + 1).Resize(总数据数, UBound(outputarr, 2)).value = outputarr
                zlapp.Goto(Reference:= .Cells(B列最末行号 - 10, 1), Scroll:=True)
                MsgBox("存在未汇总细目号，请检查后重新汇总")
            End If
            '.Range("A4:Z" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).Interior.Color = 16777215
            .Calculate
        End With
        Erase outputarr
line_Exit:

        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        'Erase brr : Erase 需汇总工作薄 : r = 0
        'If openWbNum <> 需汇总工作薄.Length Then
        '    MsgBox("汇总完成" & Chr(13) & "汇总文件夹里文件数：" & 需汇总工作薄.Length & Chr(13) & "已汇总文件数：" & openWbNum & Chr(13) & "请务必检查是否汇总完成", , "错误提示")
        'End If
        'If 需汇总工作薄.Length > 20 Then
        '    MsgBox("汇总完成" & Chr(13) & "汇总文件夹里文件数：" & 需汇总工作薄.Length & Chr(13) & "已汇总文件数：" & openWbNum & Chr(13),, "错误提示")
        'End If
        If 错误提示 IsNot Nothing Then
            MsgBox(错误提示 & "数据有问题，请核实，如果不为0，请重新汇总 ")
        End If
        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlAutomatic  '工作薄开启自动重算
        Exit Sub
line_Err:
        MsgBox(Err.Number & ":" & Err.Description & Chr(13) & "请仔细检查数据。如无法处理，请联系管理员", , "错误提示")
        'Resume line_Exit
        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlAutomatic  '工作薄开启自动重算
    End Sub
    Sub 输出数据格式_弈睿泽高速()
        Dim outputarr(0 To 50000, 0 To 14)
        With zlapp.Sheets(sht_Hz)
            总数据数 = 0
            For i = 1 To UBound(qDarr)
                If qDarr(i, 2) IsNot Nothing Then
                    outputarr(总数据数， 0) = qDarr(i， 1)
                    outputarr(总数据数， 1) = qDarr(i， 2)
                    outputarr(总数据数， 2) = qDarr(i， 3)         '单位
                    outputarr(总数据数， 3) = qDarr(i， 4)         '单位
                    If d.Exists（qDarr(i， 2)） Then
                        d.Remove(qDarr(i， 2))                         '删除对应字典的键位，先删除，下面这句才不会出错
                    End If
                    outputarr(总数据数， 13) = qDarr(i， 13)         '清单工程量
                    outputarr(总数据数， 15) = "=M" & 总数据数 + 4 & "-N" & 总数据数 + 4
                    'Dim aa As String = qDarr(i， 2).ToString
                    'If d.ContainsKey(aa) Then
                    '    outputarr(总数据数， 12) = d(qDarr(i， 2))             '汇总的工程量
                    '    d.Remove(outputarr(总数据数， 1))
                    'End If
                    'If d.exists(qDarr(i, 2)) Then

                    'End If '删除对应字典的键位
                    For 临时数据1 = 0 To 清单数量 - 1
                        If qDarr(i, 2).ToString = brr(临时数据1, 1).ToString Then
                            总数据数 += 1
                            outputarr(总数据数, 4) = brr(临时数据1, 2)         '桥名
                            outputarr(总数据数, 5) = brr(临时数据1, 6)         '部位名称
                            outputarr(总数据数, 6) = brr(临时数据1, 3)         '左右幅
                            outputarr(总数据数, 7) = brr(临时数据1, 4)         '名称
                            outputarr(总数据数, 8) = brr(临时数据1, 5)         '适用范围
                            outputarr(总数据数, 9) = brr(临时数据1, 7)         '
                            outputarr(总数据数, 12) = brr(临时数据1, 8)         '工程量
                            outputarr(总数据数, 10) = brr(临时数据1, 9)         '图号
                            outputarr(总数据数, 11) = brr(临时数据1, 10)        '计算式
                        End If
                    Next
                    总数据数 += 1
                End If
            Next
            If .Cells(4, 2).value IsNot Nothing Then
                .Range("A4:M" & endrow + 1).ClearContents
                'zlapp.Selection.ClearOutline  '取消分组，
            End If
            With zlapp.ActiveSheet.Outline        '设置分组格式
                .AutomaticStyles = False
                .SummaryRow = Microsoft.Office.Core.XlConstants.xlAbove
                .SummaryColumn = Microsoft.Office.Core.XlConstants.xlLeft
            End With
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).value = outputarr
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Font.name = "宋体"
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Font.Size = 10
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).Borders.LineStyle = Excel.XlLineStyle.xlContinuous
            .cells(4, 1).Resize(总数据数 + 1, UBound(outputarr, 2)).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlCenter
            .cells(4, 2).Resize(总数据数 + 1, 1).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlLeft
            .cells(4, 8).Resize(总数据数 + 1, 1).HorizontalAlignment = Microsoft.Office.Core.XlConstants.xlLeft
            zlapp.Goto(Reference:= .Cells(3, 1), Scroll:=True)
            For i = 4 To .Range("B65535").End(Excel.XlDirection.xlUp).Row
                'If StrComp(.Cells(i, 1).value, .Cells(i + 1, 1).value, 1) = 0 Then
                If .Cells(i, 2).VALUE Is Nothing Then
                    .Cells(i, 2).Rows.Group
                End If
            Next i
            endrow = .Cells.Find("*", .Cells(1, 1), Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole,
                                     Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
            .Rows(UBound(outputarr) + 6 & ":" & endrow + 1).DELETE
            If d.count > 0 Then  '多的细目号
                B列最末行号 = (.Cells(zlapp.Rows.Count, "B").End(Excel.XlDirection.xlUp).Row) + 1
                .Cells(B列最末行号, 2) = "以下为未汇总细目号"
                .Cells(B列最末行号 + 1, 1).Resize(d.Count, 1).value = zlapp.WorksheetFunction.Transpose(d.Keys)
                Dim arr_d() As String = d.Values.ToArray()
                '.Cells(B列最末行号 + 1, 8).Resize(d.Count, 1).value = zlapp.WorksheetFunction.Transpose(d.items)
                .Cells(B列最末行号 + 1, 8).Resize(d.Count, 1).value = zlapp.WorksheetFunction.Transpose(arr_d)
                qDarr = .Range("A" & B列最末行号 + 1 & ":M" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).value
                .Range("A" & B列最末行号 + 1 & ":M" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).ClearContents
                总数据数 = 0
                ReDim outputarr(0 To UBound(outputarr), 0 To UBound(outputarr, 2))
                For i = 1 To UBound(qDarr)
                    outputarr(总数据数， 0) = qDarr(i， 1)
                    outputarr(总数据数， 1) = qDarr(i， 2)
                    outputarr(总数据数， 2) = qDarr(i， 3)
                    outputarr(总数据数， 3) = qDarr(i， 4)
                    outputarr(总数据数， 13) = qDarr(i， 13)         '清单工程量
                    'outputarr(总数据数， 13) = "=L" & 总数据数 + 4 & "-M" & 总数据数 + 4
                    If d.ContainsKey(qDarr(i， 2)) Then
                        outputarr(总数据数， 12) = d(qDarr(i， 2))           '汇总的工程量
                        d.Remove(qDarr(i， 2))
                    End If

                    For 临时数据1 = 0 To 清单数量 - 1
                        If qDarr(i, 2).ToString = brr(临时数据1, 1).ToString Then
                            总数据数 += 1
                            outputarr(总数据数, 4) = brr(临时数据1, 2)         '桥名
                            outputarr(总数据数, 5) = brr(临时数据1, 6)         '部位名称
                            outputarr(总数据数, 6) = brr(临时数据1, 3)         '左右幅
                            outputarr(总数据数, 7) = brr(临时数据1, 4)         '名称
                            outputarr(总数据数, 8) = brr(临时数据1, 5)         '适用范围
                            outputarr(总数据数, 9) = brr(临时数据1, 7)         '
                            outputarr(总数据数, 12) = brr(临时数据1, 8)         '工程量
                            outputarr(总数据数, 10) = brr(临时数据1, 9)         '图号
                            outputarr(总数据数, 11) = brr(临时数据1, 10)        '计算式
                        End If
                    Next
                    总数据数 += 1
                Next
                .Range("B" & B列最末行号 + 1).Resize(总数据数, UBound(outputarr, 2)).value = outputarr
                zlapp.Goto(Reference:= .Cells(B列最末行号 - 10, 1), Scroll:=True)
                MsgBox("存在未汇总细目号，请检查后重新汇总")
            End If
            '.Range("A4:Z" & .Cells(zlapp.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row).Interior.Color = 16777215
            .Calculate
        End With
        Erase outputarr
line_Exit:

        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        'Erase brr : Erase 需汇总工作薄 : r = 0
        'If openWbNum <> 需汇总工作薄.Length Then
        '    MsgBox("汇总完成" & Chr(13) & "汇总文件夹里文件数：" & 需汇总工作薄.Length & Chr(13) & "已汇总文件数：" & openWbNum & Chr(13) & "请务必检查是否汇总完成", , "错误提示")
        'End If
        'If 需汇总工作薄.Length > 20 Then
        '    MsgBox("汇总完成" & Chr(13) & "汇总文件夹里文件数：" & 需汇总工作薄.Length & Chr(13) & "已汇总文件数：" & openWbNum & Chr(13),, "错误提示")
        'End If
        If 错误提示 IsNot Nothing Then
            MsgBox(错误提示 & "数据有问题，请核实，如果不为0，请重新汇总 ")
        End If
        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlAutomatic  '工作薄开启自动重算
        Exit Sub
line_Err:
        MsgBox(Err.Number & ":" & Err.Description & Chr(13) & "请仔细检查数据。如无法处理，请联系管理员", , "错误提示")
        'Resume line_Exit
        zlapp.ActiveSheet.Outline.ShowLevels(RowLevels:=1)
        zlapp.DisplayAlerts = True '显示提示和警告消息
        zlapp.ScreenUpdating = True '显示屏幕自动刷新
        zlapp.Calculation = Microsoft.Office.Interop.Excel.Constants.xlAutomatic  '工作薄开启自动重算
    End Sub
End Module
