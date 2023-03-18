Attribute VB_Name = "Erros"
Sub logErro(msgErr As String)
    Dim endCell As String
    
    endCell = "H" & (plan_testes.Range("n_err") + 2)

    plan_testes.Range(endCell) = msgErr
End Sub
