Private WithEvents InboxItems As Outlook.Items
Private OutlookApp As Outlook.Application ' Declarar una variable a nivel de módulo para la aplicación de Outlook

Private Sub Application_Startup()
    Set OutlookApp = Outlook.Application
    Dim Namespace As Outlook.Namespace
    Set Namespace = OutlookApp.GetNamespace("MAPI")
    Set InboxItems = Namespace.GetDefaultFolder(olFolderInbox).Items
End Sub

Private Sub InboxItems_ItemAdd(ByVal Item As Object)
    Dim Msg As Outlook.MailItem
    Dim Correo As Outlook.MailItem
    
    If TypeOf Item Is Outlook.MailItem Then
        Set Msg = Item
        
        ' Verificar el asunto del correo
        If Msg.Subject = "carmijo@favoritafc.com" Then
            ' Ejecutar script de Python
            Dim PythonPath As String
            Dim ScriptPath As String
            Dim Cmd As String
            Dim Variable As String
            
            ' Ruta del intérprete de Python
            PythonPath = "D:\anaconda\python.exe"
            
            ' Ruta del script de Python
            ScriptPath = "C:\xampp\htdocs\py\script.py"
            
            ' Variable a pasar al script de Python
            Variable = Msg.Body
            
            ' Comando para ejecutar el script de Python con la variable como argumento
            Cmd = PythonPath & " " & ScriptPath & " """ & Variable & """"
            
            ' Ejecutar el comando
            Shell Cmd, vbNormalFocus
            
            ' Agrega aquí las acciones adicionales que deseas realizar después de ejecutar el script de Python
        ElseIf Msg.Subject = "actualizarpy3" Then
            ' Crea un nuevo correo
             Set Correo = OutlookApp.CreateItem(olMailItem)
             With Correo
                ' Establece la dirección de correo "De" como el buzón compartido
                .SentOnBehalfOfName = "asistenciafavorita@favoritafc.com"
                ' Establece los demás campos del correo
                .To = "pesanchez@favoritafc.com"
                .CC = "asistenciafavorita@favoritafc.com; gtomala@favoritafc.com; cavila@favoritafc.com; Operador-Sistemas@favoritafc.com; beugenio@favoritafc.com; dtahurtare@favoritafc.com; jsilva@favoritafc.com; karteaga@favoritafc.com; fsalazar@favoritafc.com; ricvera@favoritafc.com ;tquimi@favoritafc.com; dlopez@favoritafc.com; dtalindao@favoritafc.com; lcavero@favoritafc.com; ivacacela@favoritafc.com; mjumbo@favoritafc.com ;gereyes@favoritafc.com"
                .Subject = "Actualizar Py3 - " & Format(Date, "dd/MM/yyyy")
                .Body = Msg.Body
                ' Envía el correo
                 .Send
             End With
            
            ' Libera los objetos utilizados
             Set Correo = Nothing
            
        ElseIf Msg.Subject = "OS" Or Msg.Subject = "OC" Then
            ' Ejecutar script de Python
    Dim PythonPath1 As String
    Dim ScriptPath1 As String
    Dim Cmd1 As String
    Dim numero As String
    Dim orden As String
    
    ' Ruta del intérprete de Python
    PythonPath1 = "D:\anaconda\python.exe"
    
    ' Ruta del script de Python
    ScriptPath1 = "C:\xampp\htdocs\py\orden.py"
    
    ' Variable a pasar al script de Python
    numero = Msg.Body
    orden = Msg.Subject  ' Reemplaza "valor_de_orden" con el valor real de la variable orden
    
    ' Comando para ejecutar el script de Python con las variables como argumentos
    Cmd1 = PythonPath1 & " " & ScriptPath1 & " """ & numero & """ """ & orden & """"
    
    ' Ejecutar el comando
    Shell Cmd1, vbNormalFocus
        End If
    End If
End Sub
