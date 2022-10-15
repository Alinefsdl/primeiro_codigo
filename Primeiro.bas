Attribute VB_Name = "Módulo1"
Sub primeiro()
'o comando DIM(dimension) é utilizado para declarar variavel
'a variavel nome foi tipada como String (texto)
Dim nome As String

'o comando Imputbox abre uma caixa de entrada de
'assim o usuario digita o nome e aloca na
'variavel nome
nome = InputBox("digite o seu nome")
'o comando range permite selecionar uma celula na planilha de comandos
'do excel. Assim selecionamos a celula AL e adicionamos usando uma variavel nome
Range("A1").Value = nome
End Sub
