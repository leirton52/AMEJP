Attribute VB_Name = "testes"
Option Explicit

Sub teste()
'On Error Resume Next
    'teste 1 - verificando se os dados estão sendo passados corretamente para o objeto, saída na tabela(testes) celulas c1 1 c2
    Dim dNode1 As Variant
    Dim dNode2(1 To 1, 1 To 11) As Variant
    Dim nodes As New Collection
    Dim cntI As Integer 'contador para for
    Dim cntJ As Integer 'contador para for
    Dim tempBool As Boolean 'serve para guardar temporariamente um estado boleano
    
    
    dNode1 = Planilha1.Range("b4:L4")
    dNode2(1, 1) = 150
    dNode2(1, 2) = 200
    dNode2(1, 3) = 3
    dNode2(1, 4) = 4
    dNode2(1, 5) = 5
    dNode2(1, 6) = "fixa"
    dNode2(1, 7) = "livre"
    dNode2(1, 8) = "mola"
    dNode2(1, 9) = ""
    dNode2(1, 10) = ""
    dNode2(1, 11) = 30
    
    Dim node1 As New cls_Node
    Dim node2 As New cls_Node
    
    node1.constructor (dNode1)
    plan_testes.Range("c1") = argNodeIguais(node1, dNode1)
    node2.constructor (dNode2)
    plan_testes.Range("c2") = argNodeIguais(node2, dNode2)
    
    nodes.Add node1
    nodes.Add node2

    'teste2:
    Dim dLine1 As Variant
    Dim dLine2(1 To 1, 1 To 13) As Variant
    
    dLine1 = Planilha1.Range("O4:AA4")
     
    dLine2(1, 1) = 1
    dLine2(1, 2) = 2
    dLine2(1, 3) = 10
    dLine2(1, 4) = 20
    dLine2(1, 5) = 30
    dLine2(1, 6) = 40
    dLine2(1, 7) = "retangular"
    dLine2(1, 8) = 30
    dLine2(1, 9) = 10
    dLine2(1, 10) = "concreto"
    dLine2(1, 11) = 2400
    dLine2(1, 12) = 0.02
    dLine2(1, 13) = 0.00001
    
    Dim line1 As New cls_Line
    Dim line2 As New cls_Line
    
    line1.constructor dLine1, nodes
    plan_testes.Range("c3") = argLineIguais(line1, dLine1)
    line2.constructor dLine2, nodes
    plan_testes.Range("c4") = argLineIguais(line2, dLine2)
    
    Dim area2 As Double
    Dim inercia2 As Double
    Dim length2 As Double
    Dim cos2 As Double
    Dim sen2 As Double
    Dim mAuxGlobalToLocal2 As Variant
    Dim mAuxLocalToGlobal2
    Dim MRlocal
    Dim MRglobal
    Dim MElocal
    Dim MEglobal
    
    area2 = 300
    inercia2 = 2500
    length2 = Sqr((nodes.Item(2).posX - nodes.Item(1).posX) ^ 2 + (nodes.Item(2).posY - nodes.Item(1).posY) ^ 2)
    cos2 = (nodes.Item(2).posX - nodes.Item(1).posX) / length2
    sen2 = (nodes.Item(2).posY - nodes.Item(1).posY) / length2
    mAuxGlobalToLocal2 = line2.mAuxGlobalToLocal
    mAuxLocalToGlobal2 = line2.mAuxLocalToGlobal
    MRlocal = line2.mRigidezLocal
    MRglobal = line2.mRigidezGlobal
    MElocal = line2.mEngasteLocal
    MEglobal = line2.mEngasteGlobal
    
    'testando se a area esta sendo calculada corretamente
    If line2.area = area2 Then
        plan_testes.Range("c5") = "OK"
    Else
        plan_testes.Range("c5") = "momento de area nao esta sendo calculado corretamente"
    End If
    'testando se o momento de inercia esta sendo calculada corretamente
    If line2.inercia = inercia2 Then
        plan_testes.Range("c6") = "OK"
    Else
        plan_testes.Range("c6") = "momento de inercia nao esta sendo calculado corretamente"
    End If
    'testando se o momento de inercia esta sendo calculada corretamente
    If line2.length = length2 Then
        plan_testes.Range("c7") = "OK"
    Else
        plan_testes.Range("c7") = "comprimento da linha nao esta sendo calculado corretamente"
    End If
    'testando a matriz auxiliar que transforma coordenadados globais para locais
    If mAuxGlobalToLocal2(0, 0) <> cos2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(0, 1) <> sen2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(1, 0) <> -sen2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(1, 1) <> cos2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(3, 3) <> cos2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(3, 4) <> sen2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(4, 3) <> -sen2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    ElseIf mAuxGlobalToLocal2(4, 4) <> cos2 Then
        plan_testes.Range("c8") = "argumentos de mAuxGlobalToLocal calculados errados"
    Else
        plan_testes.Range("c8") = "OK"
    End If
    'testando a matriz auxiliar que transforma coordenadados locais para Globais
    If mAuxLocalToGlobal2(0, 0) <> cos2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(0, 1) <> -sen2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(1, 0) <> sen2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(1, 1) <> cos2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(3, 3) <> cos2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(3, 4) <> -sen2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(4, 3) <> sen2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    ElseIf mAuxLocalToGlobal2(4, 4) <> cos2 Then
        plan_testes.Range("c9") = "argumentos de mAuxLocalToGlobal calculados errados"
    Else
        plan_testes.Range("c9") = "OK"
    End If
    
    'testando a matriz de rigidez local
    Dim mrigidezLocal2(0 To 5, 0 To 5) As Double
    
    'atribuindo valores a linha 1 da matriz
    mrigidezLocal2(0, 0) = 6439.87577519939
    mrigidezLocal2(0, 1) = 0
    mrigidezLocal2(0, 2) = 0
    mrigidezLocal2(0, 3) = -6439.87577519939
    mrigidezLocal2(0, 4) = 0
    mrigidezLocal2(0, 5) = 0
    
    'atribuindo valores a linha 2 da matriz
    mrigidezLocal2(1, 0) = 0
    mrigidezLocal2(1, 1) = 51.5190062015952
    mrigidezLocal2(1, 2) = 2880
    mrigidezLocal2(1, 3) = 0
    mrigidezLocal2(1, 4) = -51.5190062015952
    mrigidezLocal2(1, 5) = 2880
    
    'atribuindo valores a linha 3 da matriz
    mrigidezLocal2(2, 0) = 0
    mrigidezLocal2(2, 1) = 2880
    mrigidezLocal2(2, 2) = 214662.52583998
    mrigidezLocal2(2, 3) = 0
    mrigidezLocal2(2, 4) = -2880
    mrigidezLocal2(2, 5) = 107331.26291999
    
    'atribuindo valores a linha 4 da matriz
    mrigidezLocal2(3, 0) = -6439.87577519939
    mrigidezLocal2(3, 1) = 0
    mrigidezLocal2(3, 2) = 0
    mrigidezLocal2(3, 3) = 6439.87577519939
    mrigidezLocal2(3, 4) = 0
    mrigidezLocal2(3, 5) = 0
    
    'atribuindo valores a linha 5 da matriz
    mrigidezLocal2(4, 0) = 0
    mrigidezLocal2(4, 1) = -51.5190062015952
    mrigidezLocal2(4, 2) = -2880
    mrigidezLocal2(4, 3) = 0
    mrigidezLocal2(4, 4) = 51.5190062015952
    mrigidezLocal2(4, 5) = -2880
    
    'atribuindo valores a linha 6 da matriz
    mrigidezLocal2(5, 0) = 0
    mrigidezLocal2(5, 1) = 2880
    mrigidezLocal2(5, 2) = 107331.26291999
    mrigidezLocal2(5, 3) = 0
    mrigidezLocal2(5, 4) = -2880
    mrigidezLocal2(5, 5) = 214662.52583998
    
    tempBool = True
    For cntI = 0 To 5
        For cntJ = 0 To 5
            If Round(mrigidezLocal2(cntI, cntJ), 4) = Round(MRlocal(cntI, cntJ), 4) And tempBool Then
                plan_testes.Range("c10") = "OK"
            Else
                plan_testes.Range("c10") = "mRigidez local calculada errada, (" & cntI & " ," & cntJ & ")"
                tempBool = False
                Exit For
            End If
        Next
        
        If Not tempBool Then
            Exit For
        End If
    Next
    
    'Testando a matriz de rigidez em coordenadas globais de uma barra
    Dim mRigidezGlobal2(0 To 5, 0 To 5) As Double
    
    mRigidezGlobal2(0, 0) = 2880
    mRigidezGlobal2(0, 1) = 5760
    mRigidezGlobal2(0, 2) = 0
    mRigidezGlobal2(0, 3) = -2880
    mRigidezGlobal2(0, 4) = -5760
    mRigidezGlobal2(0, 5) = 0
    mRigidezGlobal2(1, 0) = -46.08
    mRigidezGlobal2(1, 1) = 23.04
    mRigidezGlobal2(1, 2) = 2880
    mRigidezGlobal2(1, 3) = 46.08
    mRigidezGlobal2(1, 4) = -23.04
    mRigidezGlobal2(1, 5) = 2880
    mRigidezGlobal2(2, 0) = -2575.95031007975
    mRigidezGlobal2(2, 1) = 1287.97515503987
    mRigidezGlobal2(2, 2) = 214662.525839979
    mRigidezGlobal2(2, 3) = 2575.95031007975
    mRigidezGlobal2(2, 4) = -1287.97515503987
    mRigidezGlobal2(2, 5) = 107331.262919989
    mRigidezGlobal2(3, 0) = -2880
    mRigidezGlobal2(3, 1) = -5760
    mRigidezGlobal2(3, 2) = 0
    mRigidezGlobal2(3, 3) = 2880
    mRigidezGlobal2(3, 4) = 5760
    mRigidezGlobal2(3, 5) = 0
    mRigidezGlobal2(4, 0) = 46.08
    mRigidezGlobal2(4, 1) = -23.04
    mRigidezGlobal2(4, 2) = -2880
    mRigidezGlobal2(4, 3) = -46.08
    mRigidezGlobal2(4, 4) = 23.04
    mRigidezGlobal2(4, 5) = -2880
    mRigidezGlobal2(5, 0) = -2575.95031007975
    mRigidezGlobal2(5, 1) = 1287.97515503987
    mRigidezGlobal2(5, 2) = 107331.262919989
    mRigidezGlobal2(5, 3) = 2575.95031007975
    mRigidezGlobal2(5, 4) = -1287.97515503987
    mRigidezGlobal2(5, 5) = 214662.525839979

    tempBool = True
    For cntI = 0 To 5
        For cntJ = 0 To 5
            If Round(mRigidezGlobal2(cntI, cntJ), 4) = Round(MRglobal(cntI, cntJ), 4) And tempBool Then
                plan_testes.Range("c11") = "OK"
            Else
                plan_testes.Range("c11") = "mRigidez global calculada errada, (" & cntI & " ," & cntJ & ")"
                tempBool = False
                Exit For
            End If
        Next
        
        If Not tempBool Then
            Exit For
        End If
    Next
    
    'Testando a matriz de reações de engastamento perfeiro em coordenadas locais
    Dim mEngasteLocal2(0 To 5)
    
    mEngasteLocal2(0) = -745.355992499929
    mEngasteLocal2(1) = -1844.75608143732
    mEngasteLocal2(2) = -35416.6666666666
    mEngasteLocal2(3) = -931.694990624912
    mEngasteLocal2(4) = -2068.3628791873
    mEngasteLocal2(5) = 37500
    
    For cntI = 0 To 5
        If Round(mEngasteLocal2(cntI), 4) = Round(MElocal(cntI), 4) Then
            plan_testes.Range("c12") = "OK"
        Else
            plan_testes.Range("c12") = "mEngasteLocal calculada errada, (" & cntI & ")"
            Exit For
        End If
    Next
    
    'Testando a matriz de reações de engastamento perfeiro em coordenadas locais
    Dim mEngasteGlobal2(0 To 5)
    
    mEngasteGlobal2(0) = 1316.66666666666
    mEngasteGlobal2(1) = -1491.66666666666
    mEngasteGlobal2(2) = -35416.6666666666
    mEngasteGlobal2(3) = 1433.33333333333
    mEngasteGlobal2(4) = -1758.33333333333
    mEngasteGlobal2(5) = 37500

    For cntI = 0 To 5
        If Round(mEngasteGlobal2(cntI), 4) = Round(MEglobal(cntI), 4) Then
            plan_testes.Range("c13") = "OK"
        Else
            plan_testes.Range("c13") = "mEngasteGlobal calculada errada, (" & cntI & ")"
            Exit For
        End If
    Next
        
End Sub


Private Function argNodeIguais(node1 As cls_Node, dNode As Variant)
    Dim msgErr As String
    
    If (node1.posX <> dNode(1, 1)) Then
        msgErr = "construção do objeto node errada, posX inconsistente"
    ElseIf (node1.posY <> dNode(1, 2)) Then
        msgErr = "construção do objeto node errada, posY inconsistente"
    ElseIf (node1.Fx <> dNode(1, 3)) Then
        msgErr = "construção do objeto node errada, Fx inconsistente"
    ElseIf (node1.Fy <> dNode(1, 4)) Then
        msgErr = "construção do objeto node errada, Fy inconsistente"
    ElseIf (node1.Mz <> dNode(1, 5)) Then
        msgErr = "construção do objeto node errada, Mz inconsistente"
    ElseIf (node1.restrictionX <> dNode(1, 6)) Then
        msgErr = "construção do objeto node errada, apoioX inconsistente"
    ElseIf (node1.restrictionY <> dNode(1, 7)) Then
        msgErr = "construção do objeto node errada, apoioY inconsistente"
    ElseIf (node1.restrictionZ <> dNode(1, 8)) Then
        msgErr = "construção do objeto node errada, apoioZ inconsistente"
    ElseIf (node1.apoioElasticoX <> dNode(1, 9)) Then
        msgErr = "construção do objeto node errada, Kx inconsistente"
    ElseIf (node1.apoioElasticoY <> dNode(1, 10)) Then
        msgErr = "construção do objeto node errada, Ky inconsistente"
    ElseIf (node1.apoioElasticoZ <> dNode(1, 11)) Then
        msgErr = "construção do objeto node errada, Kz inconsistente"
    Else
        msgErr = "OK"
    End If
    
    argNodeIguais = msgErr
End Function
    
Private Function argLineIguais(line1 As cls_Line, dLine As Variant)
    Dim msgErr As String
    
    If (line1.nodeI <> dLine(1, 1)) Then
        msgErr = "construção do objeto line errada, nodeI inconsistente"
    ElseIf (line1.nodeF <> dLine(1, 2)) Then
        msgErr = "construção do objeto line errada, nodeF inconsistente"
    ElseIf (line1.cargaXI <> dLine(1, 3)) Then
        msgErr = "construção do objeto line errada, cargaXI inconsistente"
    ElseIf (line1.cargaXF <> dLine(1, 4)) Then
        msgErr = "construção do objeto line errada, cargaXF inconsistente"
    ElseIf (line1.cargaYI <> dLine(1, 5)) Then
        msgErr = "construção do objeto line errada, cargaYI inconsistente"
    ElseIf (line1.cargaYF <> dLine(1, 6)) Then
        msgErr = "construção do objeto line errada, cargaYF inconsistente"
    ElseIf (line1.forma <> dLine(1, 7)) Then
        msgErr = "construção do objeto line errada, forma inconsistente"
    ElseIf (line1.base <> dLine(1, 8)) Then
        msgErr = "construção do objeto line errada, base inconsistente"
    ElseIf (line1.altura <> dLine(1, 9)) Then
        msgErr = "construção do objeto line errada, altura inconsistente"
    ElseIf (line1.tipo <> dLine(1, 10)) Then
        msgErr = "construção do objeto line errada, tipo inconsistente"
    ElseIf (line1.modElasticidade <> dLine(1, 11)) Then
        msgErr = "construção do objeto line errada, modElasticidade inconsistente"
    ElseIf (line1.coefPoison <> dLine(1, 12)) Then
        msgErr = "construção do objeto line errada, coefPoison inconsistente"
    ElseIf (line1.coefTermico <> dLine(1, 13)) Then
        msgErr = "construção do objeto line errada, coefTermico inconsistente"
    Else
        msgErr = "OK"
    End If
    
    argLineIguais = msgErr
End Function










