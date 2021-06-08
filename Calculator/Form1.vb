Option Explicit On
'Option Strict On
Public Class Form1

  Dim calc As Calculation
  Dim numberOfDigits As Integer

  Protected Overrides Function ProcessCmdKey(ByRef msg As System.Windows.Forms.Message,
                                             ByVal keyData As System.Windows.Forms.Keys) As Boolean

    If keyData = Keys.Enter Then
      'Do something here when Enter is pressed
      Evaluate()
      Return True 'Stop the form from processing the key any further
    End If
    Return MyBase.ProcessCmdKey(msg, keyData)

  End Function

  Private Sub Button_Click(sender As Object, e As EventArgs) Handles Button0.Click,
    Button1.Click, Button2.Click, Button3.Click, Button4.Click, Button5.Click,
    Button6.Click, Button7.Click, Button8.Click, Button9.Click, ButtonDot.Click,
    ButtonClear.Click, ButtonClearError.Click,
    ButtonAdd.Click, ButtonSubtract.Click, ButtonMultiply.Click, ButtonDivide.Click,
    ButtonSign.Click, ButtonShift.Click, ButtonRoot.Click, ButtonRatio.Click,
    ButtonPercent.Click, ButtonMemorySub.Click, ButtonMemorySave.Click, ButtonMemoryRead.Click,
    ButtonMemoryClear.Click, ButtonMemoryAdd.Click, ButtonEqual.Click, ButtonSqrt.Click

    ActionManagement(sender.Name)

  End Sub

  Private Sub Form_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
    ActionManagement(e.KeyCode.ToString())
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    calc = New Calculation
    numberOfDigits = 12
  End Sub

  Sub ActionManagement(source As String)

    Select Case source
      Case "Button0", "Button1", "Button2", "Button3", "Button4",
           "Button5", "Button6", "Button7", "Button8", "Button9",
           "D0", "NumPad0", "D1", "NumPad1", "D2", "NumPad2",
           "D3", "NumPad3", "D4", "NumPad4", "D5", "NumPad5",
           "D6", "NumPad6", "D7", "NumPad7", "D8", "NumPad8",
           "D9", "NumPad9"
        AddDigit(source)
      Case "ButtonClear", "ButtonClearError", "Delete"
        ClearTable()
      Case "ButtonDot", "Decimal"
        AddDot()
      Case "ButtonEqual", "Enter"
        Evaluate()
      Case "ButtonAdd", "ButtonSubtract", "ButtonMultiply", "ButtonDivide",
           "Divide", "Multiply", "Subtract", "Add"
        SaveOperation(source)
      Case "ButtonRatio", "ButtonSqrt"
        SaveOperation(source)
        EvaluateUnary()
      Case "ButtonSign", "OemMinus"
        ChangeSign()
      Case "ButtonShift", "Back"
        Shift()
      Case "ButtonPercent"
        CalculatePercent()
      Case "ButtonMemoryAdd", "ButtonMemorySub", "ButtonMemoryClear", "ButtonMemoryRead", "ButtonMemorySave"
        ActionsInMemory(source)
      Case Else
        MsgBox("No action defined for the button or the key")
    End Select

  End Sub

  Sub ClearTable()

    calc.Clear()
    Label1.Text = "0"
    Label2.Text = ""

  End Sub

  Sub Evaluate()

    calc.operand2 = Label1.Text
    calc.MathOp()
    Label1.Text = calc.result
    Label2.Text = ""

  End Sub

  Sub CalculatePercent()

    If calc.operand1 = "" OrElse calc.operand1 = "0" Then
      Label1.Text = "0"
    Else
      Dim res As Decimal = (CDec(calc.operand1) / 100) * CDec(Label1.Text)
      Label1.Text = $"{res:##########0.##########}"
    End If

  End Sub

  Sub ActionsInMemory(source As String)

    source = Normalize(source)
    calc.SetMemory(source, Label1.Text, LabelMemory.Text)

  End Sub

  Sub EvaluateUnary()

    calc.MathOp()
    Label1.Text = calc.result

  End Sub

  Sub ChangeSign()

    If Label1.Text = "" OrElse Label1.Text = "0" Then
      Return
    End If
    If Label1.Text.IndexOf("-") = -1 Then
      Label1.Text = "-" & Label1.Text
    Else
      Label1.Text = Replace(Label1.Text, "-", "")
    End If

  End Sub

  Sub SaveOperation(name As String)

    name = Normalize(name)
    calc.operand1 = Label1.Text
    calc.SetCurrentop(name)
    Label2.Text = calc.postfix

  End Sub

  Sub AddDot()

    If Label1.Text.IndexOf(",") > -1 OrElse Label1.Text.Length = numberOfDigits Then
      Return
    Else
      Label1.Text &= ","
      calc.newNumber = False
    End If

  End Sub

  Shared Function Normalize(source As String) As String

    source = Replace(source, "Button", "")
    source = Replace(source, "Memory", "")
    If source.IndexOf("Divide") = -1 Then
      source = Replace(source, "D", "")
    End If
    source = Replace(source, "NumPad", "")
    Return source

  End Function

  Sub AddDigit(digit As String)

    digit = Normalize(digit)
    If Label1.Text.Length = 0 And digit = "0" Then
      Return
    End If
    If Label1.Text = "0" OrElse calc.newNumber Then
      Label1.Text = digit
      calc.newNumber = False
    ElseIf Label1.Text.Length = numberOfDigits Then
      Return
    Else
      Label1.Text &= digit
    End If

  End Sub

  Sub Shift()

    If Label1.Text.Length <= 1 OrElse Label1.Text = "0" Then
      Label1.Text = "0"
      Return
    End If
    Label1.Text = Label1.Text.Substring(0, Label1.Text.Length - 1)
    If Label1.Text.Length > 0 AndAlso Label1.Text.Substring(Label1.Text.Length - 1, 1) = "," Then
      Label1.Text = Label1.Text.Substring(0, Label1.Text.Length - 1)
    End If

  End Sub

End Class

Public Class Calculation

  Enum Operation As Byte
    add
    subtract
    multiply
    divide
    ratio
    sqrt
  End Enum

  Dim currentop As Nullable(Of Operation)
  Public newNumber As Boolean
  Public postfix As String
  Public operand1 As String
  Public operand2 As String
  Public result As String
  Public memory As String

  Sub New()

    Clear()
    memory = "0"

  End Sub

  Public Sub SetCurrentop(value As String)

    newNumber = True
    Select Case value
      Case "Add"
        currentop = Operation.add
        postfix = $"{operand1} +"
      Case "Subtract"
        currentop = Operation.subtract
        postfix = $"{operand1} -"
      Case "Multiply"
        currentop = Operation.multiply
        postfix = $"{operand1} *"
      Case "Divide"
        currentop = Operation.divide
        postfix = $"{operand1} /"
      Case "Ratio"
        currentop = Operation.ratio
        postfix = $"reciproc({operand1})"
      Case "Sqrt"
        currentop = Operation.sqrt
        postfix = $"sqrt({operand1})"
      Case Else
        MsgBox("Operation not defined")
        newNumber = False
    End Select

  End Sub

  Public Sub MathOp()

    Dim res As Decimal
    Select Case currentop
      Case Operation.add
        res = CDec(operand1) + CDec(operand2)
      Case Operation.subtract
        res = CDec(operand1) - CDec(operand2)
      Case Operation.multiply
        res = CDec(operand1) * CDec(operand2)
      Case Operation.divide
        If CDec(operand2) = 0 Then
          res = 0
        Else
          res = CDec(operand1) / CDec(operand2)
        End If
      Case Operation.ratio
        If CDec(operand1) = 0 Then
          res = 0
        Else
          res = 1 / CDec(operand1)
        End If
      Case Operation.sqrt
        res = Math.Sqrt(CDec(operand1))
    End Select
    result = $"{res:##########0.##########}"
    operand1 = result
    operand2 = "0"
    currentop = Nothing
    newNumber = True

  End Sub

  Public Sub Clear()

    operand1 = "0"
    operand2 = "0"
    currentop = Nothing
    newNumber = True
    postfix = ""
    result = ""

  End Sub

  Public Sub SetMemory(value As String, ByRef label1 As String, ByRef labelMemory As String)

    Dim res As Decimal
    newNumber = True
    Select Case value
      Case "Add"
        res = CDec(memory) + CDec(label1)
        memory = $"{res:##########0.##########}"
        labelMemory = "M"
      Case "Sub"
        res = CDec(memory) - CDec(label1)
        memory = $"{res:##########0.##########}"
        labelMemory = "M"
      Case "Save"
        memory = label1
        labelMemory = "M"
      Case "Read"
        label1 = memory
      Case "Clear"
        memory = "0"
        labelMemory = ""
      Case Else
        MsgBox("Memory operations are undefined")
        newNumber = False
    End Select

  End Sub

End Class
