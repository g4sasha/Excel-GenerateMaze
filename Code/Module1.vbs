' Размеры лабиринта
Dim mazeWidth As Integer
Dim mazeHeight As Integer

' Позиция лабиринта (от левого и верхнего угла)
Const mazePositionX = 2
Const mazePositionY = 10

' Карты лабиринта (далее всё превращается в двумерные массивы)
Dim freeCells As Variant ' Карта свободных ячеек (при значениях False в элементах - стоят стены)
Dim visited As Variant ' Карта посещённых ячеек для алгоритма генерации лабиринта
Dim robotVisited As Variant ' Карта посещённых мест роботом

Dim pathFound As Boolean ' Переменная, отвечающая за состояние найденного пути
Dim isGenerated As Boolean ' Переменная учитывает состояние лабиринта

Dim maxSteps As Integer ' Максимальное количество шагов, которое потребовалось для прохождения
Dim steps As Integer ' Основное количество шагов
' При нажании на главную кнопку
Sub NewMaze()
    
    Cells.Interior.ColorIndex = xlNone ' Очистка абсолютно всех ячеек от цвета
    
    ' Создаём пустую карту стен и накладываем её на лабиринт
    ReDim NewMaze(mazeWidth, mazeHeight) As Boolean
    freeCells = NewMaze
    
    ' Устанавливаем обычные размеры лабиринта
    mazeWidth = 22
    mazeHeight = 32
    
    ' Удаляем информацию о пройденных шагах
    maxSteps = 0
    steps = 0
    Range("M3") = ""
    Range("M4") = ""
    
    ' Создаём пустую карту посещённых ячеек и накладываем её на лабиринт
    ReDim newVisitMap(mazeWidth, mazeHeight) As Boolean
    visited = newVisitMap
    
    ' Создаём пустую карту посещённых мест робота и накладываем её на лабиринт
    ReDim newRobotMap(mazeWidth, mazeHeight) As Boolean
    robotVisited = newRobotMap
    
    ' Сразу устанавливаем внешние стены как посещённые
    Dim i As Integer
    Dim j As Integer
    For i = 0 To mazeWidth
        For j = 0 To mazeHeight
            If i = 0 Or j = 0 Or i = mazeWidth Or j = mazeHeight Then ' Если одна из границ
                visited(i, j) = True
            End If
        Next
    Next
    
    pathFound = False ' Ставим состояние лабиринта как "не пройденный"
    
    Call PaveTheWay(1, mazeHeight - 1) ' Начинаем алгоритм генерации
    Call Draw ' Отрисовка всего лабиринта
    Call PostProcess ' Ставим старт и финиш
    
End Sub
' Процедура вызываеться по кнопке "Пройти лабиринт" и предназначена для поиска пути роботом
Sub RobotPath()

    ' Обнуляем шаги, чтобы они не росли с каждым новым прохождением
    steps = 0
    maxSteps = 0

    ' Предотвращение генерации лабиринта с пустым массивом стен
    If Not isGenerated Then
        MsgBox "Сначала сгенерируйте лабиринт"
    Else
        ' Сразу устанавливаем внешние стены как посещённые
        Dim i As Integer
        Dim j As Integer
        For i = 0 To mazeWidth
            For j = 0 To mazeHeight
                If i = 0 Or j = 0 Or i = mazeWidth Or j = mazeHeight Then ' Если одна из границ
                    robotVisited(i, j) = True
                End If
            Next
        Next
        
        ' Запускаем поиск
        Call FindPath(1, mazeHeight - 1)
    End If
    
    Range("M3") = maxSteps
    Range("M4") = steps
    
End Sub
' Алгоритм поиска пути в лабиринте
Sub FindPath(x As Integer, y As Integer)
    
    ' Указываем визуально текущий шаг
    robotVisited(x, y) = True
    Cells(y + mazePositionY, x + mazePositionX).Interior.Color = RGB(255, 150, 150)
    
    ' Прибавляем шаги
    steps = steps + 1
    maxSteps = maxSteps + 1
    
    ' Хранит все направления (варианты)
    Dim directionUsed(3) As Integer
    directionUsed(0) = 1
    directionUsed(1) = 2
    directionUsed(2) = 3
    directionUsed(3) = 4
    
    ' Пока есть куда идти
    While (directionUsed(0) <> -1 Or directionUsed(1) <> -1 Or directionUsed(2) <> -1 Or directionUsed(3) <> -1) And Not pathFound
    
        ' Выбираем случайное направление
        Dim direction As Integer
        direction = RandomFromArray(directionUsed())
        
        ' Если выбрали идти вверх
        If direction = 1 Then
                ' Если наверх возможно пройти
                If Not robotVisited(x, y - 1) And freeCells(x, y - 1) Then
                    Call FindPath(x, y - 1) ' Продолжаем путь сверху
                Else
                    directionUsed(0) = -1 ' Блокируем возможность идти вверх и ищем другое направление
                End If
        End If
        
        ' Если выбрали идти вправо
        If direction = 2 Then
                ' Если вправо возможно пройти
                If Not robotVisited(x + 1, y) And freeCells(x + 1, y) Then
                    Call FindPath(x + 1, y) ' Продолжаем путь справа
                Else
                    directionUsed(1) = -1 ' Блокируем возможность идти вправо и ищем другое направление
                End If
        End If
        
        ' Если выбрали идти вниз
        If direction = 3 Then
                ' Если вниз возможно пройти
                If Not robotVisited(x, y + 1) And freeCells(x, y + 1) Then
                    Call FindPath(x, y + 1) ' Продолжаем путь вниз
                Else
                    directionUsed(2) = -1 ' Блокируем возможность идти вниз и ищем другое направление
                End If
        End If
        
        ' Если выбрали идти влево
        If direction = 4 Then
                ' Если влево возможно пройти
                If Not robotVisited(x - 1, y) And freeCells(x - 1, y) Then
                    Call FindPath(x - 1, y) ' Продолжаем путь слева
                Else
                    directionUsed(3) = -1 ' Блокируем возможность идти влево и ищем другое направление
                End If
        End If
        
        ' Если мы дошли до финиша - прекращаем искать
        If x = mazeWidth - 1 And y = 1 Then
            pathFound = True
        End If
        
    Wend
    
    If Not pathFound Then
        ' Благодаря этой строчке мы не оставляем тупики в пройденном пути
        Cells(y + mazePositionY, x + mazePositionX).Interior.ColorIndex = xlNone
        steps = steps - 1
    End If
    
End Sub
' Процедура будет генерировать лабиринт до тех пор, пока все ячейки не окажутся посещёнными
Sub PaveTheWay(x As Integer, y As Integer)
    
    freeCells(x, y) = True ' Удаляем стену в текущей ячейке
    visited(x, y) = True ' Устанавливаем эту ячейку как посещённую
    
    ' Хранит все направления (варианты)
    Dim directionUsed(3) As Integer
    directionUsed(0) = 1
    directionUsed(1) = 2
    directionUsed(2) = 3
    directionUsed(3) = 4
    
    ' Пока есть куда идти
    While directionUsed(0) <> -1 Or directionUsed(1) <> -1 Or directionUsed(2) <> -1 Or directionUsed(3) <> -1
    
        ' Выбираем случайное направление
        Dim direction As Integer
        direction = RandomFromArray(directionUsed())
        
        ' Если выбрали идти вверх
        If direction = 1 Then
            If y <> 1 Then
                ' Если наверх возможно пройти
                If Not visited(x, y - 2) Then
                    visited(x, y - 1) = True
                    freeCells(x, y - 1) = True
                    Call PaveTheWay(x, y - 2) ' Продолжаем путь сверху
                Else
                    directionUsed(0) = -1 ' Блокируем возможность идти вверх и ищем другое направление
                End If
            Else
                directionUsed(0) = -1
            End If
        End If
        
        ' Если выбрали идти вправо
        If direction = 2 Then
            If x <> mazeWidth - 1 Then
                ' Если вправо возможно пройти
                If Not visited(x + 2, y) Then
                    visited(x + 1, y) = True
                    freeCells(x + 1, y) = True
                    Call PaveTheWay(x + 2, y) ' Продолжаем путь справа
                Else
                    directionUsed(1) = -1 ' Блокируем возможность идти вправо и ищем другое направление
                End If
            Else
                directionUsed(1) = -1
            End If
        End If
        
        ' Если выбрали идти вниз
        If direction = 3 Then
            If y <> mazeHeight - 1 Then
                ' Если возможно пойти вниз
                If Not visited(x, y + 2) Then
                    visited(x, y + 1) = True
                    freeCells(x, y + 1) = True
                    Call PaveTheWay(x, y + 2) ' Продолжаем путь снизу
                Else
                    directionUsed(2) = -1 ' Блокируем возможность идти вниз и ищем другое направление
                End If
            Else
                directionUsed(2) = -1
            End If
        End If
        
        ' Если выбрали идти влево
        If direction = 4 Then
            If x <> 1 Then
                ' Если возможно пойти влево
                If Not visited(x - 2, y) Then
                    visited(x - 1, y) = True
                    freeCells(x - 1, y) = True
                    Call PaveTheWay(x - 2, y) ' Продолжаем путь слева
                Else
                    directionUsed(3) = -1 ' Блокируем возможность идти влево и ищем другое направление
                End If
            Else
                directionUsed(3) = -1
            End If
        End If
        
    Wend
    
End Sub
' Визуально закрашиваем ячейки при значениях False для freeCells
Sub Draw()
    
    ' Пробегаемся по каждой ячейке лабиринта
    Dim stepX As Integer
    Dim stepY As Integer
    For stepX = 0 To mazeWidth ' Строки
        For stepY = 0 To mazeHeight ' Столбцы
            If freeCells(stepX, stepY) = True Then ' Если пусто
                Cells(stepY + mazePositionY, stepX + mazePositionX).Interior.ColorIndex = xlNone ' Убирает цвет ячейке
            Else
                Cells(stepY + mazePositionY, stepX + mazePositionX).Interior.Color = RGB(0, 0, 0) ' Закрашивает ячейку чёрным цветом
            End If
        Next
    Next
    
End Sub
' Добавляет старт и финиш в виде закрашеных ячеек
Sub PostProcess()
    Cells(mazeHeight + mazePositionY, 1 + mazePositionX).Interior.Color = RGB(0, 255, 0)
    Cells(1 + mazePositionY, mazeWidth + mazePositionX).Interior.Color = RGB(255, 0, 0)
    isGenerated = True
End Sub
' Функция выбирает случайный элемент из массива, который не равен -1
Function RandomFromArray(arr() As Integer) As Integer
    
    Dim validElements() As Integer
    Dim i As Integer
    Dim count As Integer
    For i = LBound(arr) To UBound(arr)
        If arr(i) <> -1 Then
            ReDim Preserve validElements(count)
            validElements(count) = arr(i)
            count = count + 1
        End If
    Next i
    If count = 0 Then
        RandomFromArray = -2
        Exit Function
    End If
    RandomFromArray = validElements(Int(Rnd() * count))
    
End Function
' По кнопке можно установить свои стены. Этим занимается этот массив
Sub SetWallsFromColor()

    Call DeleteWay

    ' Создаём пустую карту стен и накладываем её на лабиринт
    ReDim NewMaze(mazeWidth, mazeHeight) As Boolean
    freeCells = NewMaze
    
    ' Создаём пустую карту посещённых мест робота и накладываем её на лабиринт
    ReDim newRobotMap(mazeWidth, mazeHeight) As Boolean
    robotVisited = newRobotMap

    ' Пробегаемся по каждой ячейке лабиринта
    Dim stepX As Integer
    Dim stepY As Integer
    For stepX = 0 To mazeWidth ' Строки
        For stepY = 0 To mazeHeight ' Столбцы
            If Cells(stepY + mazePositionY, stepX + mazePositionX).DisplayFormat.Interior.Color = 0 Then ' Если пусто
                freeCells(stepX, stepY) = False
            Else
                freeCells(stepX, stepY) = True
            End If
        Next
    Next
    
    ' Завершающие процедуры
    pathFound = False
    Call PostProcess

End Sub
' Сбросить всё
Sub DefaultMaze()
    Call StartProgram
End Sub
' Начальные процедуры для улучшения работы с документом
Sub StartProgram()
    mazeWidth = 23
    mazeHeight = 33
    ThisWorkbook.Sheets("Лабиринт по умолчанию").Activate
    ThisWorkbook.Sheets("Лабиринт").Activate
    Cells.Interior.ColorIndex = xlNone ' Очистка абсолютно всех ячеек от цвета
    Dim i As Integer
    Dim j As Integer
    For i = 0 To mazeWidth
        For j = 0 To mazeHeight
            Cells(j + mazePositionY, i + mazePositionX).Interior.Color = ThisWorkbook.Sheets("Лабиринт по умолчанию").Cells(j + 1, i + 1).DisplayFormat.Interior.Color
        Next
    Next
    Range("M3") = ""
    Range("M4") = ""
    Call SetWallsFromColor
End Sub
' Удаляет все ячейки красного цвета
Sub DeleteWay()
    Dim i As Integer
    Dim j As Integer
    For i = 0 To mazeWidth
        For j = 0 To mazeHeight
            If Cells(j + mazePositionY, i + mazePositionX).Interior.Color = RGB(255, 150, 150) Then
                Cells(j + mazePositionY, i + mazePositionX).Interior.Color = xlNone
            End If
        Next
    Next
End Sub