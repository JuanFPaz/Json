# Segundo Video - MsgBox - If Else End If - Abriendo ventanas

- URL:[Video 2](https://www.youtube.com/watch?v=nthHIpCXKUw)

## MsgBox

Vamos a crear un Mensaje con un `Command` (boton de VB6).

### Nomenglaturas

Las nomenglaturas son el estandar que se utiliza en VB6 para los elementos:

|abrevitarua | control |
------------------------
chk checkbox
cmd command button
drv drive list box
frm form
hsb horizontal scroll bar
lbl label
lst list
opt option button
shp shape
tmr timer
cbo combo y drop-list box
dir dir list box
fil file list box
fra frame
img image
lin line
mnu menu
pct pictureBox
txt text edit box
vsb vertical scroll bar

Por ejemplo, cuando cfreamos un boton en nuestro formulario, le asignamos el nombre `cmdMensaje`, para tener referncia de este objeto como un `Command`, por su contraparte, el Caption es el nombre con el que va aaparecer en el formulario:

![image.PNG](./image.PNG)

- Minuto 2:04 MsgBox estructura

- Doc Microsoft : [https://learn.microsoft.com/msgbox-function](https://learn.microsoft.com/es-es/office/vba/language/reference/user-interface-help/msgbox-function)

Estructura de un MsgBox:

```vb6
MsgBox "Mensaje Principal", <botones, botonxdefecto, iconos, propiedades del msg box>, "Titulo" (opcional), <helper y context (opcionales)>

'Para obtener el boton presiona,envolvemos en parenteis y guardamos en una variable
'
Dim Response
Response = MsgBox(Msg, 1 + 32, Title)

```

Nuestro Codigo debe verse asi:

```vb
Private Sub cmdMensaje_Click()
Dim Msg, Title, Response, MyString

Msg = "Estamos probando el Msg, queres continuar?"
Title = "Titulo del MsgBox"

Response = MsgBox(Msg, vbYesNo Or vbQuestion, Title)

If Response = 6 Then
MyString = "Si, queres continuar :)"
lblMensaje.Caption = MyString
cmdAbrirVentana.Visible = True

ElseIf Response = 7 Then
MyString = "Que pena, no queres aprender VB6 :("
lblMensaje.Caption = MyString
End If

End Sub
```

- Minuto 9:18 usa el `Me.` paracambiar el caption del formulario

- Minuto 10:49 Creamos un nuevo Button y un nuevo formulario en el proyecto.
Hacemos click derecho en Proyecto1, agregar, formulario ,estandar.

Le cambiamos el nombre por una nomenglura : `fnmLogin` (?)

Creamos un nuevo `Command button`, le agregamos un nombre `cmdAbrirVentana` y le establecemos su propiedad visible a False.

- Quede minuto 16:15