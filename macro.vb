Option Explicit

Sub EnviarMensajesWhatsApp()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim phone As String
    Dim msg As String
    Dim url As String

    ' Establecer la hoja de trabajo
    Set ws = ThisWorkbook.Sheets("Hoja1") ' Cambia "Hoja1" por el nombre de tu hoja

    ' Encontrar la última fila con datos en la columna A (Número de Teléfono)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Recorrer cada fila y enviar el mensaje
    For i = 2 To lastRow ' Asumiendo que la primera fila es el encabezado
        phone = ws.Cells(i, 2).Value ' Número de Teléfono en la columna B
        msg = ws.Cells(i, 3).Value ' Mensaje en la columna C

        ' Crear la URL de WhatsApp
        url = "https://web.whatsapp.com/send?phone=" & phone & "&text=" & Application.EncodeURL(msg)

        ' Abrir la URL en el navegador
        ThisWorkbook.FollowHyperlink url

        ' Esperar unos segundos para permitir que la página se cargue
        Application.Wait (Now + TimeValue("00:00:10"))

        ' enviar el mensaje presionando Enter
        Application.SendKeys "~", True

        Application.Wait (Now + TimeValue("00:00:1"))

        ' Cerrar la pestaña del navegador
        Application.SendKeys "^w", True

        ' Marcar como enviado en la columna D
        ws.Cells(i, 4) = "1"

    Next i

    MsgBox "Mensajes enviados con éxito.", vbInformation
End Sub
