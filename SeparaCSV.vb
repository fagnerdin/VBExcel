Sub TXT()
Dim arquivo As String
arquivo = "C:\ARQUIVO.csv"

Open arquivo For Input As #1

Dim linha
linha = 0

Dim sl, ls, sp
sl = 2
ls = 2
sp = 2

Do Until EOF(1)
    Line Input #1, linefromfile
    
    lineitens = Split(linefromfile, ";")
    Dim vItem
    
    Dim coluna
    coluna = 0
    
    If linha = 0 Then
        For Each vLinha In lineitens
            vItem = Replace(lineitens(coluna), """", "")
            Cells(linha + 1, coluna + 1) = vItem
            coluna = coluna + 1
        Next vLinha
    Else
    
        Select Case Left(lineitens(8), 2)
            Case 35
                Sheets("GRUPO1").Select
                For Each vLinha In lineitens
                    vItem = Replace(lineitens(coluna), """", "")
                    Cells(sp, coluna + 1).Value = vItem
                    coluna = coluna + 1
                Next vLinha
                sp = sp + 1
                
            Case 41, 42, 43
                Sheets("GRUPO2").Select
                For Each vLinha In lineitens
                    vItem = Replace(lineitens(coluna), """", "")
                    Cells(sl, coluna + 1).Value = vItem
                    coluna = coluna + 1
                Next vLinha
                sl = sl + 1

            Case Else
                Sheets("GRUPO3").Select
                For Each vLinha In lineitens
                    vItem = Replace(lineitens(coluna), """", "")
                    Cells(ls, coluna + 1).Value = vItem
                    coluna = coluna + 1
                Next vLinha
                ls = ls + 1
                
        End Select
    End If
    
    linha = linha + 1
       
Loop

Close #1

End Sub
