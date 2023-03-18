VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Autor: Raimundo Leirton
'Data inicio: 12/02/2023
'Ultima aualização: 12/02/2023
'Descrissão: Classe dos nós da estrutura que as vezes será denomido como ponto ou Node

Option Explicit

'Declaração das variáveis
'Posição do ponto
Public posX As Double
Public posY As Double

'Cargas pontuais
Public Fx As Double
Public Fy As Double
Public Mz As Double

'Informações dos apoios
'apoio terá três valores possíveis: livre, fixa, mola
Public restrictionX As String
Public restrictionY As String
Public restrictionZ As String
'Valores atribuídos quando o apoio for do tipo mola
Public apoioElasticoX As String
Public apoioElasticoY As String
Public apoioElasticoZ As String

'função que inicializa os valores de um nó
Public Sub constructor(dadosNode)
    Dim Node As Variant
    Node = dadosNode

    'testando se todos os parâmetros para se iniciar um ponto foram dados
    If WorksheetFunction.CountA(Node) < 11 Then
        logErro ("parametros insuficitente para iniciar um node")
        Exit Sub
    End If
    
    posX = dadosNode(1, 1)
    posY = dadosNode(1, 2)
    
    Fx = dadosNode(1, 3)
    Fy = dadosNode(1, 4)
    Mz = dadosNode(1, 5)
    
    restrictionX = dadosNode(1, 6)
    restrictionY = dadosNode(1, 7)
    restrictionZ = dadosNode(1, 8)
    
    If restrictionX = "mola" Then
        apoioElasticoX = dadosNode(1, 9)
    End If
    If restrictionY = "mola" Then
        apoioElasticoY = dadosNode(1, 10)
    End If
    If restrictionZ = "mola" Then
        apoioElasticoZ = dadosNode(1, 11)
    End If
End Sub


















