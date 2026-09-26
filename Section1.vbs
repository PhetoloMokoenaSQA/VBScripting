
Class Account

    Dim strAccountNumber
    Dim strAccountHolder
    Dim dblBalance
   
    public Function SetAccountHolder(strName)
        strAccountHolder = strName
    End Function

    public Function SetAccountNumber(strAccNo)
        strAccountNumber = strAccNo
    End Function

    public Function Withdrawal()
        Dim amount
        Dim strInput
        Do
            strInput = InputBox("Please enter the amount you would like to withdraw:")
            If IsNumeric(strInput) Then
                amount = CDbl(strInput)
                If amount > 0 Then
                    If amount <= dblBalance Then
                        dblBalance = dblBalance - amount
                        MsgBox "Withdrawal successful!" & vbNewLine & _
                               "Amount withdrawn: R" & amount & vbNewLine & _
                               "New balance: R" & dblBalance
                        Exit Do
                    Else
                        MsgBox "Insufficient funds." & vbNewLine & _
                               "Available balance: R" & dblBalance
                    End If
                Else
                    MsgBox "Please enter an amount greater than zero."
                End If
            Else
                MsgBox "Please enter a valid numeric amount."
            End If
        Loop

    End Function

    public Function Deposit()
        Dim amount
        Dim strInput
        Do
            strInput = InputBox( _
                "Please enter the amount you would like to deposit:")
            If IsNumeric(strInput) Then
                amount = CDbl(strInput)
                If amount > 0 Then
                    dblBalance = dblBalance + amount
                    MsgBox "Deposit successful!" & vbNewLine & _
                           "Amount deposited: R" & amount & vbNewLine & _
                           "New balance: R" & dblBalance
                    Exit Do
                Else
                    MsgBox "Please enter an amount greater than zero."
                End If
            Else
                MsgBox "Please enter a valid numeric amount."
            End If
        Loop

    End Function

    Function Transfer()
        Dim accountNumber
        Dim amount
        Dim strInput

        Do
            accountNumber = Trim(InputBox( _
                "Please enter the account number you want to transfer money to:"))
            If Len(accountNumber) >= 7 And Len(accountNumber) <= 10 Then
                strInput = InputBox( _
                    "Please enter the amount you would like to transfer:")
                If IsNumeric(strInput) Then
                    amount = CDbl(strInput)
                    If amount > 0 Then
                        If amount <= dblBalance Then
                            dblBalance = dblBalance - amount
                            MsgBox "Electronic transfer successful!" & vbNewLine & _
                                   "Amount transferred: R" & amount & vbNewLine & _
                                   "To account: " & accountNumber & vbNewLine & _
                                   "New balance: R" & dblBalance
                            Exit Do
                        Else
                            MsgBox "Insufficient funds." & vbNewLine & _
                                   "Available balance: R" & dblBalance
                        End If
                    Else
                        MsgBox "Please enter an amount greater than zero."
                    End If
                Else
                    MsgBox "Please enter a valid numeric amount."
                End If
            Else
                MsgBox "Please enter a valid account number." & vbNewLine & _
                       "Account number must contain between 7 and 10 digits."
            End If
        Loop

    End Function

    Function CheckBalance()
        CheckBalance = dblBalance
    End Function

    Function GetAccountDetails()
        GetAccountDetails = _
            "Account Holder: " & strAccountHolder & vbNewLine & _
            "Account Number: " & strAccountNumber & vbNewLine & _
            "Available Balance: R" & dblBalance

    End Function

    Function SetBalance(balance)
        dblBalance = balance
    End Function 

End Class



''''''''''''''''''''''' MAIN SCRIPT''''''''''''''''''''''


Dim objAcc
Dim strName
Dim strAccNo
Dim choice 




dblBalance = 1000.00
Set objAcc = New Account
objAcc.SetBalance dblBalance 



Do
    strName = Trim(InputBox( _
        "Please enter your name:"))
    If strName = "" Then
        MsgBox "Please enter your name."
    End If
Loop Until strName <> ""


Do
    strAccNo = Trim(InputBox( _
        "Please enter your account number:" & vbNewLine & _
        "Account number must contain between 7 and 10 digits."))
    If Len(strAccNo) >= 7 And Len(strAccNo) <= 10 Then
        If IsNumeric(strAccNo) Then
            Exit Do
        Else
            MsgBox "Account number must contain numbers only."
        End If
    Else
        MsgBox "Account number must contain between 7 and 10 digits."
    End If
Loop


objAcc.SetAccountHolder strName
objAcc.SetAccountNumber strAccNo


Do
    choice = InputBox( _
        "Welcome, " & strName & "!" & vbNewLine & vbNewLine & _
        "Account Number: " & strAccNo & vbNewLine & _
        "Current Balance: R" & objAcc.CheckBalance() & vbNewLine & vbNewLine & _
        "Please select an option:" & vbNewLine & vbNewLine & _
        "1. Withdraw" & vbNewLine & _
        "2. Deposit" & vbNewLine & _
        "3. Electronic Transfer" & vbNewLine & _
        "4. Check Balance" & vbNewLine & _
        "5. Account Details" & vbNewLine & _
        "6. Exit")
    Select Case choice

        Case "1"
            objAcc.Withdrawal
        Case "2"
            objAcc.Deposit
        Case "3"
            objAcc.Transfer
        Case "4"
            MsgBox "Current balance: R" & _
                   objAcc.CheckBalance()
        Case "5"
            MsgBox objAcc.GetAccountDetails()
        Case "6"
            MsgBox "Thank you, " & strName & "." & vbNewLine & _
                   "Account Number: " & strAccNo & vbNewLine & _
                   "Final Balance: R" & objAcc.CheckBalance() & vbNewLine & vbNewLine & _
                   "Goodbye!"
            Exit Do
        Case Else
            MsgBox "Invalid option." & vbNewLine & _
                   "Please select an option from 1 to 6."
    End Select

Loop