VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Autor: Raimundo Leirton
'Data inicio: 12/02/2023
'Ultima aualização: 24/02/2023
'Descrissão: Classe das Linhas da estrutura que as vezes será denomido como barra

Option Explicit

'Declaração das variáveis
'Posição da Linha
Public nodes As Collection
Public nodeI As Integer
Public nodeF As Integer

'Cargas distibuídas
Public cargaXI As Double
Public cargaXF As Double
Public cargaYI As Double
Public cargaYF As Double

'Informações a geometria
Public forma As String
Public base As Double
Public altura As Double

'Informações do material
Public tipo As String
Public modElasticidade As Double
Public coefPoison As Double
Public coefTermico As Double

'função que inicializa os valores de uma linha
Public Sub constructor(dadosLine, nodesOrigin As Collection)
    Dim Line As Variant
    Line = dadosLine

    'testando se todos os parâmetros para se iniciar um ponto foram dados
    If WorksheetFunction.CountA(Line) < 13 Then
        logErro ("parametros insuficitente")
        Exit Sub
    End If
    
    nodeI = dadosLine(1, 1)
    nodeF = dadosLine(1, 2)
    
    cargaXI = dadosLine(1, 3)
    cargaXF = dadosLine(1, 4)
    cargaYI = dadosLine(1, 5)
    cargaYF = dadosLine(1, 6)
    
    forma = dadosLine(1, 7)
    base = dadosLine(1, 8)
    altura = dadosLine(1, 9)
    
    tipo = dadosLine(1, 10)
    modElasticidade = dadosLine(1, 11)
    coefPoison = dadosLine(1, 12)
    coefTermico = dadosLine(1, 13)
    
    Set nodes = nodesOrigin
End Sub

'calcula a área da secção transversal
Property Get area() As Double
    If LCase(forma) = LCase("retangular") Then
        area = base * altura
    ElseIf LCase(forma) = LCase("circular") Then
        area = WorksheetFunction.Pi() * base ^ 2 / 4
    Else
        logErro ("forma nao cadastrada para calculo da area")
        area = 0
    End If
End Property

Property Get inercia() As Double
    If LCase(forma) = LCase("retangular") Then
        inercia = base * altura ^ 3 / 12
    ElseIf LCase(forma) = LCase("circular") Then
        inercia = WorksheetFunction.Pi() * base ^ 4 / 64
    Else
        logErro ("forma nao cadastrada para calculo da inercia")
        inercia = 0
    End If
End Property

Property Get length() As Double
    'cacula o tamanha da linha
    Dim Xi As Double
    Dim Xf As Double
    Dim Yi As Double
    Dim Yf As Double
    
    Xi = nodes.Item(nodeI).posX
    Xf = nodes.Item(nodeF).posX
    
    Yi = nodes.Item(nodeI).posY
    Yf = nodes.Item(nodeF).posY
    
    length = Sqr((Xf - Xi) ^ 2 + (Yf - Yi) ^ 2)
End Property

Property Get mAuxGlobalToLocal()
    'matriz auxiliar para tanformação das coordenadas de global para local
    Dim Xi As Double
    Dim Xf As Double
    Dim Yi As Double
    Dim Yf As Double
    Dim cos As Double
    Dim sen As Double
    Dim L As Double
    
    Dim matriz(0 To 5, 0 To 5) As Double
    
    Xi = nodes.Item(nodeI).posX
    Xf = nodes.Item(nodeF).posX
    Yi = nodes.Item(nodeI).posY
    Yf = nodes.Item(nodeF).posY
    
    L = length
    
    cos = (Xf - Xi) / L
    sen = (Yf - Yi) / L
    
    matriz(0, 0) = cos: matriz(0, 1) = sen: matriz(0, 2) = 0: matriz(0, 3) = 0: matriz(0, 4) = 0: matriz(0, 5) = 0
    matriz(1, 0) = -sen: matriz(1, 1) = cos: matriz(1, 2) = 0: matriz(1, 3) = 0: matriz(1, 4) = 0: matriz(1, 5) = 0
    matriz(2, 0) = 0: matriz(2, 1) = 0: matriz(2, 2) = 1: matriz(2, 3) = 0: matriz(2, 4) = 0: matriz(2, 5) = 0
    matriz(3, 0) = 0: matriz(3, 1) = 0: matriz(3, 2) = 0: matriz(3, 3) = cos: matriz(3, 4) = sen: matriz(3, 5) = 0
    matriz(4, 0) = 0: matriz(4, 1) = 0: matriz(4, 2) = 0: matriz(4, 3) = -sen: matriz(4, 4) = cos: matriz(4, 5) = 0
    matriz(5, 0) = 0: matriz(5, 1) = 0: matriz(5, 2) = 0: matriz(5, 3) = 0: matriz(5, 4) = 0: matriz(5, 5) = 1
    
    mAuxGlobalToLocal = matriz
End Property

Property Get mAuxLocalToGlobal()
    'matriz auxiliar para tanformação das coordenadas de Local para global
    Dim Xi As Double
    Dim Xf As Double
    Dim Yi As Double
    Dim Yf As Double
    Dim cos As Double
    Dim sen As Double
    Dim L As Double
    
    Dim matriz(0 To 5, 0 To 5) As Double
    
    Xi = nodes.Item(nodeI).posX
    Xf = nodes.Item(nodeF).posX
    Yi = nodes.Item(nodeI).posY
    Yf = nodes.Item(nodeF).posY
    
    L = length
    
    cos = (Xf - Xi) / L
    sen = (Yf - Yi) / L
    
    matriz(0, 0) = cos: matriz(0, 1) = -sen: matriz(0, 2) = 0: matriz(0, 3) = 0: matriz(0, 4) = 0: matriz(0, 5) = 0
    matriz(1, 0) = sen: matriz(1, 1) = cos: matriz(1, 2) = 0: matriz(1, 3) = 0: matriz(1, 4) = 0: matriz(1, 5) = 0
    matriz(2, 0) = 0: matriz(2, 1) = 0: matriz(2, 2) = 1: matriz(2, 3) = 0: matriz(2, 4) = 0: matriz(2, 5) = 0
    matriz(3, 0) = 0: matriz(3, 1) = 0: matriz(3, 2) = 0: matriz(3, 3) = cos: matriz(3, 4) = -sen: matriz(3, 5) = 0
    matriz(4, 0) = 0: matriz(4, 1) = 0: matriz(4, 2) = 0: matriz(4, 3) = sen: matriz(4, 4) = cos: matriz(4, 5) = 0
    matriz(5, 0) = 0: matriz(5, 1) = 0: matriz(5, 2) = 0: matriz(5, 3) = 0: matriz(5, 4) = 0: matriz(5, 5) = 1
    
    mAuxLocalToGlobal = matriz
End Property


Property Get mRigidezLocal()
    'matriz de rigidez nas coordenadas locais
    Dim L As Double 'comrpiemto da barra
    Dim a As Double 'area da barra
    Dim ine As Double ' Inercia da barra
    
    Dim matriz(0 To 5, 0 To 5) As Double
    
    L = length
    a = area
    ine = inercia
    
    'calculando os valores da matriz de rigidez local
    'linha 0
    matriz(0, 0) = modElasticidade * a / L
    matriz(0, 1) = 0
    matriz(0, 2) = 0
    matriz(0, 3) = -modElasticidade * a / L
    matriz(0, 4) = 0
    matriz(0, 5) = 0
      
    'Linha 1
    matriz(1, 0) = 0
    matriz(1, 1) = 12 * modElasticidade * ine / (L ^ 3)
    matriz(1, 2) = 6 * modElasticidade * ine / (L ^ 2)
    matriz(1, 3) = 0
    matriz(1, 4) = -12 * modElasticidade * ine / (L ^ 3)
    matriz(1, 5) = 6 * modElasticidade * ine / (L ^ 2)
    
    'Linha 2
    matriz(2, 0) = 0
    matriz(2, 1) = 6 * modElasticidade * ine / (L ^ 2)
    matriz(2, 2) = 4 * modElasticidade * ine / L
    matriz(2, 3) = 0
    matriz(2, 4) = -6 * modElasticidade * ine / (L ^ 2)
    matriz(2, 5) = 2 * modElasticidade * ine / L
    
    'Linha 3
    matriz(3, 0) = -modElasticidade * a / L
    matriz(3, 1) = 0
    matriz(3, 2) = 0
    matriz(3, 3) = modElasticidade * a / L
    matriz(3, 4) = 0
    matriz(3, 5) = 0
    
    'Linha 4
    matriz(4, 0) = 0
    matriz(4, 1) = -12 * modElasticidade * ine / (L ^ 3)
    matriz(4, 2) = -6 * modElasticidade * ine / (L ^ 2)
    matriz(4, 3) = 0
    matriz(4, 4) = 12 * modElasticidade * ine / (L ^ 3)
    matriz(4, 5) = -6 * modElasticidade * ine / (L ^ 2)
    
    'Linha 5
    matriz(5, 0) = 0
    matriz(5, 1) = 6 * modElasticidade * ine / (L ^ 2)
    matriz(5, 2) = 2 * modElasticidade * ine / L
    matriz(5, 3) = 0
    matriz(5, 4) = -6 * modElasticidade * ine / (L ^ 2)
    matriz(5, 5) = 4 * modElasticidade * ine / L
    
    mRigidezLocal = matriz
End Property

Property Get mRigidezGlobal()
    'matriz matriz de rigidez nas coordenadas globais
    Dim MRlocal 'matriz de rigidez local
    Dim Mauxiliar ' matriz auxiliar que transforma coordenadas locais para globais
    Dim matriz(0 To 5, 0 To 5) As Double
    Dim cntI As Integer 'variável para interação em for
    Dim cntJ As Integer 'variável para interação em for
    Dim cntK As Integer 'variável para interação em for
    
    MRlocal = mRigidezLocal
    Mauxiliar = mAuxLocalToGlobal
    
    For cntI = 0 To 5
        For cntJ = 0 To 5
            For cntK = 0 To 5
                matriz(cntI, cntJ) = matriz(cntI, cntJ) + Mauxiliar(cntJ, cntK) * MRlocal(cntI, cntK)
            Next
        Next
    Next
    
    mRigidezGlobal = matriz
End Property

Property Get mEngasteLocal()
    'calcula as reações de engatamento perfeito na barra
    
    Dim L As Double
    Dim matriz(0 To 5) As Double
    
    L = length
    
    'Encontrando as reaçoes de engastamento perfeito axiais: f1' e f4'
    matriz(0) = -L * (2 * cargaXI + cargaXF) / 6
    matriz(3) = -L * (2 * cargaXF + cargaXI) / 6
    
    'Encontrando as reaçoes de engatamento perfeito transversais f2 e f5
    matriz(1) = -L * (7 * cargaYI + 3 * cargaYF) / 20
    matriz(4) = -L * (3 * cargaYI + 7 * cargaYF) / 20
    
    'Encontrando as reaçoes de engatamento perfeito de momento f3 e f6
    matriz(2) = -(L ^ 2) * (3 * cargaYI + 2 * cargaYF) / 60
    matriz(5) = (L ^ 2) * (2 * cargaYI + 3 * cargaYF) / 60
    
    mEngasteLocal = matriz
End Property

Property Get mEngasteGlobal()
    'calcula as reações de engatamento perfeito na barra em coodenadas globais
    
    Dim MElocal 'matriz de engatamento perfeito local
    Dim Mauxiliar ' matriz auxiliar que transforma coordenadas locais para globais
    Dim matriz(0 To 5) As Double
    Dim cntI As Integer 'variável para interação em for
    Dim cntJ As Integer 'variável para interação em for
    
    MElocal = mEngasteLocal
    Mauxiliar = mAuxLocalToGlobal
    
    For cntI = 0 To 5
        For cntJ = 0 To 5
            matriz(cntI) = matriz(cntI) + Mauxiliar(cntI, cntJ) * MElocal(cntJ)
        Next
    Next
    
    mEngasteGlobal = matriz
End Property






















