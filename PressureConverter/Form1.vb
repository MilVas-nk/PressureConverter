Public Class Form1

#Region "Переменные класса"
    Private decimalSeparator As Char
    Private isUpdating As Boolean = False ' Флаг для предотвращения циклических обновлений

    ' Словарь для хранения коэффициентов преобразования (убран ReadOnly)
    Private conversionMatrix As Dictionary(Of Integer, Double())

    ' Объявляем массивы контролов на уровне класса
    Private Labels() As Label
    Private TextBoxes() As TextBox
#End Region

#Region "Инициализация"
    Public Sub New()
        InitializeComponent()

        ' Получаем десятичный разделитель один раз
        decimalSeparator = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator(0)

        ' Инициализируем матрицу коэффициентов преобразования
        InitializeConversionMatrix()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Привязываем общие обработчики событий для всех текстбоксов давления
        AttachCommonEventHandlers()
        LinkLabel1.TabStop = False

        ' Инициализируем массивы с нужными контролами
        Labels = {Label13, Label14, Label15, Label16, Label17, Label18, Label19, Label20, Label21, Label22, Label23}
        TextBoxes = {TextBox13, TextBox14, TextBox15, TextBox16, TextBox17, TextBox18, TextBox19, TextBox20, TextBox21, TextBox22, TextBox23}
    End Sub

    Private Sub InitializeConversionMatrix()
        conversionMatrix = New Dictionary(Of Integer, Double()) From {
            {1, New Double() {1, 0.001, 0.000001, 0.000010197162129779282, 0.01, 0.00001, 0.10197162129779283, 0.0075006375541921064, 0.000009869, 0.0001450377}}, ' Па
            {2, New Double() {1000, 1, 0.001, 0.010197162129779282, 10, 0.01, 101.97162129779282, 7.50062, 0.009869, 0.1450377}}, ' кПа
            {3, New Double() {1000000, 1000, 1, 10.197162129779283, 10000, 10, 101971.62129779282, 7500.62, 9.869, 145.0377}}, ' МПа
            {4, New Double() {98066.5, 98.0665, 0.0980665, 1, 980.665, 0.980665, 10000, 735.55970124, 0.96791655, 14.22334251}}, ' кгс/см2
            {5, New Double() {100, 0.1, 0.0001, 0.0010197162129779282, 1, 0.001, 10.197162129779283, 0.750062, 0.0009869, 0.01450377}}, ' мбар
            {6, New Double() {100000, 100, 0.1, 1.0197162129779282, 1000, 1, 10197.162129779283, 750.062, 0.9869, 14.50377}}, ' бар
            {7, New Double() {9.80665, 0.00980665, 0.00000980665, 0.00010197162129779283, 0.0980665, 0.0000980665, 1, 0.073555970124, 0.000096791655, 0.001422334251}}, ' мм вод.ст.
            {8, New Double() {133.322, 0.133322, 0.000133322, 0.00136, 1.33322, 0.00133322, 13.5951, 1, 0.001316, 0.019325}}, ' мм рт.ст.
            {9, New Double() {101325, 101.325, 0.101325, 1.0332274478, 1013.249997, 1.013249997, 10332.274478, 759.999869, 1, 14.695949}}, ' атм
            {10, New Double() {6894.757185, 6.894757185, 0.006894757185, 0.070307, 68.947572, 0.068947572, 703.06958, 51.714925, 0.068046, 1}} ' psi
        }
    End Sub

    Private Sub AttachCommonEventHandlers()
        ' Массив всех текстбоксов для конвертации давления
        Dim pressureTextBoxes() As TextBox = {
            TextBox1, TextBox2, TextBox3, TextBox4, TextBox5,
            TextBox6, TextBox7, TextBox8, TextBox9, TextBox10
        }

        ' Привязываем общие обработчики
        For Each tb In pressureTextBoxes
            AddHandler tb.MouseClick, AddressOf PressureTextBox_MouseClick
            AddHandler tb.KeyPress, AddressOf NumericTextBox_KeyPress
            AddHandler tb.KeyUp, AddressOf PressureTextBox_KeyUp
        Next

        ' Для текстбоксов с отрицательными значениями
        For Each tb In {TextBox11, TextBox12, TextBox24}
            AddHandler tb.KeyPress, AddressOf NumericWithNegativeTextBox_KeyPress
        Next
    End Sub
#End Region

#Region "Общие обработчики событий"
    Private Sub PressureTextBox_MouseClick(sender As Object, e As MouseEventArgs)
        DirectCast(sender, TextBox).SelectAll()
    End Sub

    Private Sub NumericTextBox_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Not Char.IsDigit(e.KeyChar) AndAlso
           e.KeyChar <> decimalSeparator AndAlso
           e.KeyChar <> vbBack Then
            e.Handled = True
        End If
    End Sub

    Private Sub NumericWithNegativeTextBox_KeyPress(sender As Object, e As KeyPressEventArgs)
        Dim tb = DirectCast(sender, TextBox)

        If Not Char.IsDigit(e.KeyChar) AndAlso
           e.KeyChar <> decimalSeparator AndAlso
           e.KeyChar <> vbBack AndAlso
           Not (e.KeyChar = "-"c AndAlso tb.SelectionStart = 0 AndAlso Not tb.Text.Contains("-")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub PressureTextBox_KeyUp(sender As Object, e As KeyEventArgs)
        If isUpdating Then Return ' Предотвращаем циклические обновления

        Dim sourceTextBox = DirectCast(sender, TextBox)
        Dim sourceIndex = GetTextBoxIndex(sourceTextBox)

        If sourceIndex = 0 Then Return

        Dim value As Double
        If Double.TryParse(sourceTextBox.Text, value) Then
            UpdateAllPressureValues(sourceIndex, value)
        End If
    End Sub
#End Region

#Region "Методы конвертации давления"
    Private Function GetTextBoxIndex(tb As TextBox) As Integer
        Select Case tb.Name
            Case "TextBox1" : Return 1
            Case "TextBox2" : Return 2
            Case "TextBox3" : Return 3
            Case "TextBox4" : Return 4
            Case "TextBox5" : Return 5
            Case "TextBox6" : Return 6
            Case "TextBox7" : Return 7
            Case "TextBox8" : Return 8
            Case "TextBox9" : Return 9
            Case "TextBox10" : Return 10
            Case Else : Return 0
        End Select
    End Function

    Private Sub UpdateAllPressureValues(sourceIndex As Integer, value As Double)
        isUpdating = True

        Try
            If conversionMatrix IsNot Nothing AndAlso conversionMatrix.ContainsKey(sourceIndex) Then
                Dim factors = conversionMatrix(sourceIndex)
                Dim textBoxes() As TextBox = {
                    TextBox1, TextBox2, TextBox3, TextBox4, TextBox5,
                    TextBox6, TextBox7, TextBox8, TextBox9, TextBox10
                }

                For i = 0 To 9
                    If i <> sourceIndex - 1 Then
                        textBoxes(i).Text = FormatValue(value * factors(i))
                    End If
                Next
            End If
        Finally
            isUpdating = False
        End Try
    End Sub

    Private Function FormatValue(value As Double) As String
        Return Math.Round(value, 8).ToString("0.########")
    End Function
#End Region

#Region "Расчет точек и тока"
    Private Sub TextBox11_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyUp
        Dim value As Double
        If Double.TryParse(TextBox11.Text, value) Then
            РасчетТочек()
            РасчетТочекТока()
        End If
    End Sub

    Private Sub TextBox12_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyUp
        Dim value As Double
        If Double.TryParse(TextBox12.Text, value) Then
            РасчетТочек()
            РасчетТочекТока()
        End If
    End Sub

    Private Sub TextBox24_keyup(sender As Object, e As KeyEventArgs) Handles TextBox24.KeyUp
        Dim value As Double
        If Double.TryParse(TextBox24.Text, value) Then
            РасчетТочек()
            РасчетТочекТока()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        РасчетТока()
        РасчетТочекТока()
    End Sub
    Private Sub РасчетТочек()
        Dim minValue, maxValue As Double

        If Not Double.TryParse(TextBox11.Text, minValue) OrElse
           Not Double.TryParse(TextBox12.Text, maxValue) Then
            Return
        End If

        Dim range As Double = maxValue - minValue

        ' 5 точек (0%, 25%, 50%, 75%, 100%)
        TextBox13.Text = (minValue + range * 0).ToString()
        TextBox14.Text = (minValue + range * 0.25).ToString()
        TextBox15.Text = (minValue + range * 0.5).ToString()
        TextBox16.Text = (minValue + range * 0.75).ToString()
        TextBox17.Text = (minValue + range * 1).ToString()

        ' 6 точек (0%, 20%, 40%, 60%, 80%, 100%)
        TextBox18.Text = (minValue + range * 0).ToString()
        TextBox19.Text = (minValue + range * 0.2).ToString()
        TextBox20.Text = (minValue + range * 0.4).ToString()
        TextBox21.Text = (minValue + range * 0.6).ToString()
        TextBox22.Text = (minValue + range * 0.8).ToString()
        TextBox23.Text = (minValue + range * 1).ToString()

        РасчетТока()
    End Sub

    Private Sub РасчетТока()
        Dim minValue, maxValue, currentValue As Double

        If Double.TryParse(TextBox11.Text, minValue) AndAlso
           Double.TryParse(TextBox12.Text, maxValue) AndAlso
           Double.TryParse(TextBox24.Text, currentValue) Then

            Dim range As Double = maxValue - minValue

            If Math.Abs(range) > 0.0000001 Then ' Проверка деления на ноль с учетом точности
                If CheckBox1.Checked Then
                    Dim procentValue = (currentValue - minValue) * 100 / range
                    Dim currentMA = Math.Sqrt(procentValue) * 10 * 16 / 100 + 4
                    TextBox25.Text = String.Format("{0:0.###} мA", Math.Round(currentMA, 3))
                Else
                    Dim normalizedValue = (currentValue - minValue) / range
                    Dim currentMA = normalizedValue * 16 + 4
                    TextBox25.Text = String.Format("{0:0.###} мA", Math.Round(currentMA, 3))
                End If
            Else
                TextBox25.Text = "4.000 мA"
            End If
        Else
            TextBox25.Text = "0 мA"
        End If
    End Sub

    Private Sub РасчетТочекТока()
        Dim minValue, maxValue As Double

        If Double.TryParse(TextBox11.Text, minValue) AndAlso
           Double.TryParse(TextBox12.Text, maxValue) Then

            Dim range As Double = maxValue - minValue

            If Math.Abs(range) > 0.0000001 Then ' Проверка деления на ноль с учетом точности
                If CheckBox1.Checked Then
                    For i As Integer = 0 To 10
                        Dim normalizedValue = Math.Sqrt((TextBoxes(i).Text - minValue) * 100 / range) * 10 * 16 / 100 + 4
                        Labels(i).Text = String.Format("{0:0.000} мА", Math.Round(normalizedValue, 3))
                    Next i
                Else
                    For i As Integer = 0 To 10
                        Dim normalizedValue = (TextBoxes(i).Text - minValue) / range * 16 + 4
                        Labels(i).Text = String.Format("{0:0.000} мА", Math.Round(normalizedValue, 3))
                    Next i
                End If
            End If
        End If

    End Sub
#End Region

#Region "Прочие обработчики"
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Process.Start("http://www.milvas.ru")
        Catch ex As Exception
            MessageBox.Show("Не удалось открыть ссылку: " & ex.Message, "Ошибка",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class