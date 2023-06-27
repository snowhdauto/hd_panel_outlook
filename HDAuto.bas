Attribute VB_Name = "HDAuto"






Option Explicit
    Public G_myFolder As String
    Public G_myColor As String
    Public G_myFullName As String
    Public G_msgRequestRegistered As String, G_mailFileName As String
    Public G_msgResetPassword As String, G_msgNewPasswordMsg As String, G_msgNewPasswordHeader As String
    Public G_instuctionDescruption()

'Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszName As String, ByVal dwFlags As Long) As Long

Public Function DeclareVariables()
    Dim dump As Integer
    'G_myColor = "Green Category"
    'G_myColor = "Mishina Yuliya"
    G_myFullName = GetCurrentUser()
    G_myFolder = G_myFullName
    G_myColor = G_myFullName
    
    

    
End Function
Function SendReplyMsg(ReplyMSG As Object, replyBody As String, replyInput As String, replyAll As Boolean, insertSubject As Boolean)
    'replyMsg - объект, на который отвечаем
    'replyBody - текст ответа на сообщение
    'replyInput - текст замены для  %INPUTVALUE%
    'replyAll - режим ответа (TRUE - ответ всем, FALSE - только автору)
    'insertSubject - вставить %INPUTVALUE% в Subject письма
    
    Dim myReply As Object
    replyBody = Replace(replyBody, "INPUTVALUE", replyInput)
    
    If (replyAll = True) Then
        Set myReply = ReplyMSG.replyAll
    Else
        Set myReply = ReplyMSG.Reply
    End If
    
    If (insertSubject = True) Then
        myReply.Subject = replyInput & " | " & myReply.Subject
    End If
    
    myReply.BodyFormat = olFormatHTML
    myReply.htmlBody = replyBody & vbCrLf & myReply.htmlBody
    myReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
    myReply.Display
End Function

Function SendNewMsg(newBody As String, newInput As String, newSubject, newTo, newCC)

    'newBody - пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅпїЅ
    'newInput - пїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅпїЅпїЅпїЅ пїЅпїЅпїЅ  %INPUTVALUE%
    Dim objMsg As MailItem
    
    newBody = Replace(newBody, "INPUTVALUE", newInput)
    newSubject = Replace(newSubject, "%INPUTVALUE%", newInput)
    
    Set objMsg = Application.CreateItem(olMailItem)
    
    With objMsg
     .To = newTo
     .CC = newCC
     .Subject = newSubject
     .htmlBody = newBody
     .SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
     .BodyFormat = olFormatHTML
     .Display
     .Categories = G_myColor
    End With
End Function
Function SendReplyOnlyOneMsg(ReplyMSG As Object, replyBody As String, replyInput As String, replyAll As Boolean, insertSubject As Boolean)
    'replyMsg - объект, на который отвечаем
    'replyBody - тест ответа на сообщение
    'replyInput - текст замены для  %INPUTVALUE%
    'replyAll - режим ответа (TRUE - ответ всем, FALSE - только автору)
    'insertSubject - вставить %INPUTVALUE% в Subject письма
    
    Dim myReply As Object
    replyBody = Replace(replyBody, "INPUTVALUE", replyInput)
    
    If (replyAll = False) Then
        Set myReply = ReplyMSG.replyAll
    Else
        Set myReply = ReplyMSG.Reply
    End If
        
    If (insertSubject = True) Then
        myReply.Subject = replyInput & " | " & myReply.Subject
    End If
    
    myReply.BodyFormat = olFormatHTML
    myReply.htmlBody = replyBody & vbCrLf & myReply.htmlBody
    myReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
    myReply.Display
End Function
Function GeneratePassword() As String
    'Простой генератор паролей

    Dim charString As String
    Dim numsString As String, functionResult As String
    Dim len1 As Integer, rndChar As String, rndNum As String
    Dim i As Integer, functionResul As String
    
    charString = "wertypasdfhkzxcvnm"
    numsString = "123456789"
    
    len1 = Len(charString)
    rndChar = Mid(charString, CInt(Int((len1 * Rnd()) + 1)), 1)
    len1 = Len(numsString)
    rndNum = Mid(numsString, CInt(Int((len1 * Rnd()) + 1)), 1)
    
    For i = 1 To 4
        functionResult = functionResult & rndChar & rndNum
    Next i
    GeneratePassword = functionResult
    Exit Function
        
    
End Function
Function GetInstructionArray()
    'Функция считывает все документы из Tools\Instructions  в массив для генерации формы
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim obj As Object, numberOfInstructions As Integer, counter As Integer
    
    
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myDestFolder = myInbox.Folders("Tools").Folders("Instructions")


    
    
    
    For Each obj In myDestFolder.Items
        numberOfInstructions = numberOfInstructions + 1
    Next obj
    ReDim G_instuctionDescruption(1 To numberOfInstructions, 1 To 2)
    
    For Each obj In myDestFolder.Items
        counter = counter + 1
        G_instuctionDescruption(counter, 1) = obj.Subject
        G_instuctionDescruption(counter, 2) = obj.htmlBody
    Next obj
    
End Function


Function GetInfWaitingArray()
    'Функция считывает все документы из Tools\Information\Test_Waiting  в массив для генерации формы
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim obj As Object, numberOfInstructions As Integer, counter As Integer
    
    
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myDestFolder = myInbox.Folders("Tools").Folders("Information\Test_Waiting")

    
    
    
    
    For Each obj In myDestFolder.Items
        numberOfInstructions = numberOfInstructions + 1
    Next obj
    ReDim G_instuctionDescruption(1 To numberOfInstructions, 1 To 2)
    
    For Each obj In myDestFolder.Items
        counter = counter + 1
        G_instuctionDescruption(counter, 1) = obj.Subject
        G_instuctionDescruption(counter, 2) = obj.htmlBody
    Next obj
    
End Function
Function GetConfWaitingArray()
    'Функция считывает все документы из Tools\Confirmation Waiting  в массив для генерации формы
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim obj As Object, numberOfInstructions As Integer, counter As Integer
    
    
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myDestFolder = myInbox.Folders("Tools").Folders("Confirmation Waiting")

    
    
    
    
    For Each obj In myDestFolder.Items
        numberOfInstructions = numberOfInstructions + 1
    Next obj
    ReDim G_instuctionDescruption(1 To numberOfInstructions, 1 To 2)
    
    For Each obj In myDestFolder.Items
        counter = counter + 1
        G_instuctionDescruption(counter, 1) = obj.Subject
        G_instuctionDescruption(counter, 2) = obj.htmlBody
    Next obj
    
End Function

Function GetTextFromTemplate(searchSubject As String) As String
    'Функция получения текста письма из шаблона (папка Tools\Email)
    
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim obj As Object
    
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myDestFolder = myInbox.Folders("Tools").Folders("Email")

    
    
    
    For Each obj In myDestFolder.Items
        If (obj.Subject = searchSubject) Then
            GetTextFromTemplate = obj.htmlBody
        End If
    Next obj
    
    
End Function
Sub HD_RequestRegistred()
    'Запрос зарегистрирован, переносим его в папку TT opend, вешаем цветовую категорию
    'Если в получателях были не только мы, то создаёт ответное письмо по шаблону
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object
    Dim inputValue As String
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Requests").Folders("TT opened")

    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
         inputValue = InputBox("Введите номер запроса", "Номер запроса", "INC")
         If (msg.To = "Regional HelpDesk" Or msg.To = "Help Desk Russia") And msg.CC = "" Then
           'Письмо только нам, ничего не делаем
         Else
           
           dump = SendReplyMsg(msg, GetTextFromTemplate("G_msgRequestRegistered"), inputValue, True, True)
         End If
                  
         msg.Categories = G_myColor
         msg.Subject = inputValue & " | " & msg.Subject
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_Phishing()
    'Запрос зарегистрирован, переносим его в папку TT opend, вешаем цветовую категорию
    'Отвечаем сотруднику шаблоном для фишинговых писем
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object
    Dim inputValue As String
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Requests").Folders("TT opened")

    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
        Set msg = obj
         inputValue = InputBox("Введите номер запроса", "Номер запроса", "IR")
         
         dump = SendReplyOnlyOneMsg(msg, GetTextFromTemplate("G_msgPhishing_Simulation"), inputValue, True, True)
                                    
         msg.Categories = G_myColor
         msg.Subject = inputValue & " | " & msg.Subject
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToMyFolder()
    'Перемешение письма в мою папку и маркирование цветом
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Personal Folders").Folders(G_myFolder)
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = True
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub


Sub HD_MoveToWithout_TT()
    'Письмо не регистрируем, переносим в папку Without TT и помечаем цветом
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Without TT")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub


Sub HD_SelectPersonalFolder()
    'Перейти в свою папку
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myDestFolder = myInbox.Folders("personal folders").Folders(G_myFolder)
    Set Application.ActiveExplorer.CurrentFolder = myDestFolder
        
    
End Sub


Sub HD_SelectSharedInbox()
    ' Переход в общие входяшие
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox)
    Set Application.ActiveExplorer.CurrentFolder = myInbox
        
    
End Sub

    
Sub HD_SendNewPassword()
    ' Выслать новый пароль
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object
    Dim inputValue As String, inputValue2 As String
    
    
    
    inputValue = InputBox("Введите пароль", "Введите пароль", GeneratePassword())
    inputValue2 = InputBox("Введите номер запроса", "Номер запроса", "IR")
    dump = SendNewMsg(GetTextFromTemplate("G_msgResetPassword"), inputValue, inputValue2 & " |Ваш новый пароль", "", "")
    
    
End Sub
Sub HD_SendInstruction()
    Dim dump As Integer
    dump = DeclareVariables()
    dump = GetInstructionArray()
    SendInstruction.Show
End Sub
Sub HD_SendInfWaiting()
    Dim dump As Integer
    dump = DeclareVariables()
    dump = GetInfWaitingArray()
    SendInfWaiting.Show
End Sub
Sub HD_SendConfWaiting()
    Dim dump As Integer
    dump = DeclareVariables()
    dump = GetConfWaitingArray()
    SendConfWaiting.Show
End Sub


Function GetCurrentUser() As String
    Dim myolApp As Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Set myolApp = CreateObject("Outlook.Application")
    Set myNamespace = myolApp.GetNamespace("MAPI")
    GetCurrentUser = myNamespace.CurrentUser
End Function



Sub HD_MoveToTTStatus()
    'Переложить в папку TT Status\Info
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Requests").Folders("TT status \ info")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub


Sub HD_MarkMyColor()
    'Отметить своим цветом
    Dim dump As Integer
    dump = DeclareVariables()
    Dim msg  As Object, obj As Object
     
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
         msg.Categories = G_myColor
         
       End If
     Next obj
    
End Sub
Sub HD_UnMarkMyColor()
    'Убрать мой цвет
    Dim dump As Integer
    dump = DeclareVariables()
    Dim msg  As Object, obj As Object
     
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
         msg.Categories = " "
         
       End If
     Next obj
    
End Sub


Sub HD_MoveToSimCard()
    'Переложить в папку Sim Card
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Sim Card")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub
Sub HD_MoveToGray()
    'Переложить в папку Серые
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("ODR").Folders("Серые")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub
Sub HD_MoveToNotWorking()
    'Переложить в папку Не работает/работает
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("ODR").Folders("Не работает/работает")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub
Sub HD_MoveToMaintenance()
    'Переложить в папку Технические работы
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("ODR").Folders("Технические работы")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToNetwork()
    'Переложить в папку Network
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("network")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToJira()
    'Переложить в папку Jira
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("JIRA")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToMTSComm()
    'Переложить в папку MTS Communicator
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("MTS Communicator")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub
Sub HD_MoveToServiceDesk()
    'Переложить в папку IDM \ Service Desk
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("IDM \ Service Desk")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub



Sub HD_MoveToOther()
    'Переложить в папку Mail router
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("Mail router")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToThanks()
    'Переложить в папку Thanks
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Thanks")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_Сomplaint()
    'Жалоба
    '
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object
    Dim myReply As Object
    

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Complaint").Folders("ToDo")
    
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
             '---
            
            Set myReply = obj.Forward
    
    
            myReply.BodyFormat = olFormatHTML
            myReply.htmlBody = "Поступила жалоба, оригинал письма находится в папке Complaint\ToDo" & vbCrLf & myReply.htmlBody
            myReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            myReply.To = "RU.BSS.IT.Supervisors@multonpartners.com"
            myReply.Send
             '---
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = True
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub
Sub HD_Mobile_Scan()
    'Переложить в папку Mobile Scan
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("Mobile Scan")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToSBIS()
 'Переложить в папку Тензор (СБИС)
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("Тензор (СБИС)")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToTT_opened()
'Переложить в папку TT Opened
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Requests").Folders("TT Opened")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
     
End Sub
Sub HD_MoveToAirwatch()
 'Переложить в папку Airwatch
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Technical Notifications").Folders("Airwatch")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
     
End Sub

Sub HD_MoveToTest()
    'Переложить в папку Test
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Test")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub HD_MoveToNSR()
    'Переложить в папку Night shift
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Night shift")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

'''''''''''''''''''''''''''''''Функции для Quick Answers''''''''''''''''''''''''''''''''''''

Sub move_TTopened()
    
    'Отправка выбранного письма в папку "TT opened" и добавление номера тикета к телу письма
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object
    Dim inputValue As String
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Requests").Folders("TT opened")

    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
        Set msg = obj
         inputValue = InputBox("Введите номер запроса", "Номер запроса", "INC")
         
         
                                    
         msg.Categories = G_myColor
         msg.Subject = inputValue & " | " & msg.Subject
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub move_TTstatus()
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object

    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Requests").Folders("TT status \ info")
    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
         Set msg = obj
                  
         msg.Categories = G_myColor
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
         
       End If
     Next obj
    
End Sub

Sub CreateMail_Windows_Password_Reply()
Dim Pass As String, i As Byte, x As Byte
Dim arr(1 To 10) As Integer, a As Integer, b As Integer, c As Integer
Randomize Timer
 
For i = 1 To 10: arr(i) = i: Next i
For i = 1 To 50
a = Int((Rnd * 10) + 1): b = Int((Rnd * 10) + 1)
c = arr(b): arr(b) = arr(a): arr(a) = c
Next i
 
For i = 1 To 10
Select Case arr(i)
Case 1, 10
     x = (Rnd * 9) + 48
Case 2, 9
     x = (Rnd * 25) + 97
Case 3, 8
     x = (Rnd * 25) + 65
Case 4, 7
     x = (Rnd * 25) + 97
Case 5, 6
     x = (Rnd * 25) + 65
End Select
 
Pass = Pass & Chr(x)
Next


Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> <br>По вашему обращению, направляем временный пароль от входа в Windows - " & "!" & Pass & "!" & "<br>При смене пароля, пожалуйста следуйте правилам для нового пароля:</br><ul><li>Минимум 12 символов</li><li>Пароль может состоять как из букв, так из цифр</li><li>Система не позволит ввести слишком простые или уязвимые общеиспользуемые пароли, а также, пароли, содержащие ваши личные данные (например имя/фамилию), корпоративные названия и бренды.</li></ul>"




    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = htmlBody & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub

Public Sub CreateMail_Windows_Password()

Dim Pass As String, i As Byte, x As Byte
Dim arr(1 To 10) As Integer, a As Integer, b As Integer, c As Integer
Randomize Timer
 
For i = 1 To 10: arr(i) = i: Next i
For i = 1 To 50
a = Int((Rnd * 10) + 1): b = Int((Rnd * 10) + 1)
c = arr(b): arr(b) = arr(a): arr(a) = c
Next i
 
For i = 1 To 10
Select Case arr(i)
Case 1, 10
     x = (Rnd * 9) + 48
Case 2, 9
     x = (Rnd * 25) + 97
Case 3, 8
     x = (Rnd * 25) + 65
Case 4, 7
     x = (Rnd * 25) + 97
Case 5, 6
     x = (Rnd * 25) + 65
End Select
 
Pass = Pass & Chr(x)
Next


Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> <br>По вашему обращению, направляем временный пароль от входа в Windows - " & "!" & Pass & "!" & "<br>При смене пароля, пожалуйста следуйте правилам для нового пароля:</br><ul><li>Минимум 12 символов</li><li>Пароль может состоять как из букв, так из цифр</li><li>Система не позволит ввести слишком простые или уязвимые общеиспользуемые пароли, а также, пароли, содержащие ваши личные данные (например имя/фамилию), корпоративные названия и бренды.</li></ul>" & Signature
With MyEmail
    
.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
.Subject = "Временный пароль от входа в Windows "
.htmlBody = htmlBody

End With
MyEmail.Display
End Sub


Public Sub CreateMail_SAP_Dialog_Password()

Dim Pass As String, i As Byte, x As Byte
Dim arr(1 To 10) As Integer, a As Integer, b As Integer, c As Integer
Randomize Timer
 
For i = 1 To 10: arr(i) = i: Next i
For i = 1 To 50
a = Int((Rnd * 10) + 1): b = Int((Rnd * 10) + 1)
c = arr(b): arr(b) = arr(a): arr(a) = c
Next i
 
For i = 1 To 10
Select Case arr(i)
Case 1, 10
     x = (Rnd * 9) + 48
Case 2, 9
     x = (Rnd * 25) + 65
Case 3, 8
     x = (Rnd * 25) + 65
Case 4, 7
     x = (Rnd * 25) + 97
Case 5, 6
     x = (Rnd * 25) + 97
End Select
 
Pass = Pass & Chr(x)
Next


Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> <br>По вашему обращению, направляем временный пароль от входа в SAP\NEXT - " & "!" & Pass & "!" & "<br>При смене пароля, пожалуйста следуйте правилам для нового пароля:</br><ul><li>Минимум 8 символов</li><li>Пароль должен состоять из: букв, цифр, одна заглавная буква, один спецсимвол</li><li>Система не позволит ввести слишком простые или уязвимые общеиспользуемые пароли, а также, пароли, содержащие ваши личные данные (например имя/фамилию), корпоративные названия и бренды.</li></ul>" & Signature
With MyEmail
    
.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
.Subject = "Временный пароль от входа в SAP\NEXT"
.htmlBody = htmlBody

End With
MyEmail.Display
End Sub

Sub CreateMail_SAP_Dialog_Password_Reply()
    Dim Pass As String, i As Byte, x As Byte
Dim arr(1 To 10) As Integer, a As Integer, b As Integer, c As Integer
Randomize Timer
 
For i = 1 To 10: arr(i) = i: Next i
For i = 1 To 50
a = Int((Rnd * 10) + 1): b = Int((Rnd * 10) + 1)
c = arr(b): arr(b) = arr(a): arr(a) = c
Next i
 
For i = 1 To 10
Select Case arr(i)
Case 1, 10
     x = (Rnd * 9) + 48
Case 2, 9
     x = (Rnd * 25) + 65
Case 3, 8
     x = (Rnd * 25) + 65
Case 4, 7
     x = (Rnd * 25) + 97
Case 5, 6
     x = (Rnd * 25) + 97
End Select
 
Pass = Pass & Chr(x)
Next


Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> <br>По вашему обращению, направляем временный пароль от входа в SAP\NEXT - " & "!" & Pass & "!" & "<br>При смене пароля, пожалуйста следуйте правилам для нового пароля:</br><ul><li>Минимум 8 символов</li><li>Пароль должен состоять из: букв, цифр, одна заглавная буква, один спецсимвол</li><li>Система не позволит ввести слишком простые или уязвимые общеиспользуемые пароли, а также, пароли, содержащие ваши личные данные (например имя/фамилию), корпоративные названия и бренды.</li></ul>"
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = htmlBody & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub

Sub CreateMail_SAP_Password_Reply()
Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If

Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Доброго времени суток!</br> <br>Ваш  пароль для входа в систему SAP/NEXT - Wave2rus/"
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = htmlBody & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub


Public Sub CreateMail_SAP_Password()
Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If

Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Доброго времени суток!</br> <br>Ваш  пароль для входа в систему SAP/NEXT - Wave2rus/" & Signature
With MyEmail
    
.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
.Subject = "Пароль для входа в SAP/NEXT"
.htmlBody = htmlBody

End With
MyEmail.Display
End Sub

Public Sub CreateMail_MyID_instruction()




Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> <br>Для решения вашего вопроса необходимо выполнить заявку в системе <a href='https://ois.eur.cchbc.com/dashboard.aspx'>MyID</a><br>Ниже приведены инструкции, которые помогут в работе с MyID</br><ul><li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B%20%D0%B8%20%D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID%20-%20%D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5%20%D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0%20%D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BE%D0%B2%20%D0%B2%20SAP.aspx?OR=Teams-HL&CT=1643147401521&sourceId=&params=%7B%22AppName%22%3A%22Teams-Desktop%22%2C%22AppVersion%22%3A%2227%2F21110108720%22%7D'>Создание запроса на доступ пользователя в SAP</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5 %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BE%D0%B2 %D0%B2 %D0%BD%D0%B5SAP %D1%81%D0%B8%D1%81%D1%82%D0%B5%D0%BC%D1%8B.aspx'>Создание запроса на доступ пользователя в не-SAP системы</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%9F%D1%80%D0%BE%D1%81%D0%BC%D0%BE%D1%82%D1%80 %D0%B8%D0%BD%D1%84%D0%BE%D1%80%D0%BC%D0%B0%D1%86%D0%B8%D0%B8 %D0%BE %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0%D1%85, %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%B0%D1%85 %D0%B8 c%D0%BE%D1%82%D1%80%D1%83%D0%B4%D0%BD%D0%B8%D0%BA%D0%B0%D1%85.aspx'>Просмотр информации о запросах, доступах и сотрудниках</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%91%D0%BB%D0%BE%D0%BA%D0%B8%D1%80%D0%BE%D0%B2%D0%BA%D0%B0 %D0%B8 %D1%80%D0%B0%D0%B7%D0%B1%D0%BB%D0%BE%D0%BA%D0%B8%D1%80%D0%BE%D0%B2%D0%BA%D0%B0 %D1%83%D1%87%D0%B5%D1%82%D0%BD%D0%BE%D0%B9 %D0%B7%D0%B0%D0%BF%D0%B8%D1%81%D0%B8 %D1%81%D0%BE%D1%82%D1%80%D1%83%D0%B4%D0%BD%D0%B8%D0%BA%D0%B0.aspx'>Блокировка и разблокировка учетной записи сотрудника</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%9E%D0%B4%D0%BE%D0%B1%D1%80%D0%B5%D0%BD%D0%B8%D0%B5 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BE%D0%B2 %D1%81%D0%BE%D1%82%D1%80%D1%83%D0%B4%D0%BD%D0%B8%D0%BA%D0%B0.aspx'>Одобрение доступов сотрудника</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A0%D0%B5%D0%B3%D0%B8%D1%81%D1%82%D1%80%D0%B0%D1%86%D0%B8%D1%8F %D0%B2%D0%BD%D0%B5%D1%88%D0%BD%D0%B5%D0%B3%D0%BE %D0%BF%D0%BE%D0%BB%D1%8C%D0%B7%D0%BE%D0%B2%D0%B0%D1%82%D0%B5%D0%BB%D1%8F.aspx'>Регистрация внешнего пользователя</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%9E%D0%B3%D1%80%D0%B0%D0%BD%D0%B8%D1%87%D0%B5%D0%BD%D0%B8%D0%B5 %D0%B8 %D0%BF%D1%80%D0%BE%D0%B4%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5 %D1%81%D1%80%D0%BE%D0%BA%D0%B0 %D0%B4%D0%B5%D0%B9%D1%81%D1%82%D0%B2%D0%B8%D1%8F %D1%83%D1%87%D0%B5%D0%BD%D0%BE%D0%B9 %D0%B7%D0%B0%D0%BF%D0%B8%D1%81%D0%B8 %D0%B2%D0%BD%D0%B5%D1%88%D0%BD%D0%B5%D0%B3%D0%BE %D0%BF%D0%BE%D0%BB%D1%8C%D0%B7%D0%BE%D0%B2%D0%B0%D1%82%D0%B5%D0%BB%D1%8F.aspx'>Ограничение и продление срока действия ученой записи внешнего пользователя</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%A3%D0%B4%D0%B0%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5 (%D0%BE%D0%B3%D1%80%D0%B0%D0%BD%D0%B8%D1%87%D0%B5%D0%BD%D0%B8%D0%B5) %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%B0.aspx'>Удаление (ограничение) доступа</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5 %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0 %D0%BD%D0%B0 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF %D0%BA %D0%BF%D0%B0%D0%BF%D0%BA%D0%B0%D0%BC %D0%BD%D0%B0 %D1%81%D0%B5%D1%82%D0%B5%D0%B2%D1%8B%D1%85 %D0%B4%D0%B8%D1%81%D0%BA%D0%B0%D1%85.aspx'>Создание запроса на доступ к папкам на сетевых дисках </a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5 %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0 %D0%BD%D0%B0 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF %D0%B2 %D1%81%D0%B8%D1%81%D1%82%D0%B5%D0%BC%D1%83 CMS (Contract Management System).aspx'>Создание запроса на доступ в систему CMS (Contract Management System) </a></li></ul>" & Signature





With MyEmail
    
.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
.Subject = "Инструкция по работе с MyID"
.htmlBody = htmlBody

End With
MyEmail.Display
End Sub



Public Sub CreateMail_Intune_instruction()




Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> <br>Для решения вашего вопроса необходимо выполнить все шаги по инструкции ниже <a href='https://multonpartners.sharepoint.com/sites/spaces-BSS-RUHD/KB/MDM_MAM_Intune/iPad/iPad_registration(Intune).pdf'>Intune</a><br>" & Signature





With MyEmail
    
.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
.Subject = "Инструкция по работе с Intune"
.htmlBody = htmlBody

End With
MyEmail.Display
End Sub


Public Sub CreateMail_MyID_instruction_reply()
Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = "<br>Здравствуйте!</br> <br>Для решения вашего вопроса необходимо выполнить заявку в системе <a href='https://ois.eur.cchbc.com/dashboard.aspx'>MyID</a><br>Ниже приведены инструкции, которые помогут в работе с MyID</br><ul><li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B%20%D0%B8%20%D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID%20-%20%D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5%20%D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0%20%D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BE%D0%B2%20%D0%B2%20SAP.aspx?OR=Teams-HL&CT=1643147401521&sourceId=&params=%7B%22AppName%22%3A%22Teams-Desktop%22%2C%22AppVersion%22%3A%2227%2F21110108720%22%7D'>Создание запроса на доступ пользователя в SAP</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5 %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BE%D0%B2 %D0%B2 %D0%BD%D0%B5SAP %D1%81%D0%B8%D1%81%D1%82%D0%B5%D0%BC%D1%8B.aspx'>Создание запроса на доступ пользователя в не-SAP системы</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%9F%D1%80%D0%BE%D1%81%D0%BC%D0%BE%D1%82%D1%80 %D0%B8%D0%BD%D1%84%D0%BE%D1%80%D0%BC%D0%B0%D1%86%D0%B8%D0%B8 %D0%BE %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0%D1%85, %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%B0%D1%85 %D0%B8 c%D0%BE%D1%82%D1%80%D1%83%D0%B4%D0%BD%D0%B8%D0%BA%D0%B0%D1%85.aspx'>Просмотр информации о запросах, доступах и сотрудниках</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%91%D0%BB%D0%BE%D0%BA%D0%B8%D1%80%D0%BE%D0%B2%D0%BA%D0%B0 %D0%B8 %D1%80%D0%B0%D0%B7%D0%B1%D0%BB%D0%BE%D0%BA%D0%B8%D1%80%D0%BE%D0%B2%D0%BA%D0%B0 %D1%83%D1%87%D0%B5%D1%82%D0%BD%D0%BE%D0%B9 %D0%B7%D0%B0%D0%BF%D0%B8%D1%81%D0%B8 %D1%81%D0%BE%D1%82%D1%80%D1%83%D0%B4%D0%BD%D0%B8%D0%BA%D0%B0.aspx'>Блокировка и разблокировка учетной записи сотрудника</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%9E%D0%B4%D0%BE%D0%B1%D1%80%D0%B5%D0%BD%D0%B8%D0%B5 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%BE%D0%B2 %D1%81%D0%BE%D1%82%D1%80%D1%83%D0%B4%D0%BD%D0%B8%D0%BA%D0%B0.aspx'>Одобрение доступов сотрудника</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A0%D0%B5%D0%B3%D0%B8%D1%81%D1%82%D1%80%D0%B0%D1%86%D0%B8%D1%8F %D0%B2%D0%BD%D0%B5%D1%88%D0%BD%D0%B5%D0%B3%D0%BE %D0%BF%D0%BE%D0%BB%D1%8C%D0%B7%D0%BE%D0%B2%D0%B0%D1%82%D0%B5%D0%BB%D1%8F.aspx'>Регистрация внешнего пользователя</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%9E%D0%B3%D1%80%D0%B0%D0%BD%D0%B8%D1%87%D0%B5%D0%BD%D0%B8%D0%B5 %D0%B8 %D0%BF%D1%80%D0%BE%D0%B4%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5 %D1%81%D1%80%D0%BE%D0%BA%D0%B0 %D0%B4%D0%B5%D0%B9%D1%81%D1%82%D0%B2%D0%B8%D1%8F %D1%83%D1%87%D0%B5%D0%BD%D0%BE%D0%B9 %D0%B7%D0%B0%D0%BF%D0%B8%D1%81%D0%B8 %D0%B2%D0%BD%D0%B5%D1%88%D0%BD%D0%B5%D0%B3%D0%BE %D0%BF%D0%BE%D0%BB%D1%8C%D0%B7%D0%BE%D0%B2%D0%B0%D1%82%D0%B5%D0%BB%D1%8F.aspx'>Ограничение и продление срока действия ученой записи внешнего пользователя</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/%D0%94%D0%BE%D1%81%D1%82%D1%83%D0%BF%D1%8B %D0%B8 %D0%BA%D0%B2%D0%BE%D1%82%D1%8B/MyID - %D0%A3%D0%B4%D0%B0%D0%BB%D0%B5%D0%BD%D0%B8%D0%B5 (%D0%BE%D0%B3%D1%80%D0%B0%D0%BD%D0%B8%D1%87%D0%B5%D0%BD%D0%B8%D0%B5) %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF%D0%B0.aspx'>Удаление (ограничение) доступа</a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5 %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0 %D0%BD%D0%B0 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF %D0%BA %D0%BF%D0%B0%D0%BF%D0%BA%D0%B0%D0%BC %D0%BD%D0%B0 %D1%81%D0%B5%D1%82%D0%B5%D0%B2%D1%8B%D1%85 %D0%B4%D0%B8%D1%81%D0%BA%D0%B0%D1%85.aspx'>Создание запроса на доступ к папкам на сетевых дисках </a></li>" & _
"<li><a href='https://multonpartners.sharepoint.com/sites/spaces-RU-RUBSSHLP/Wiki/MyID - %D0%A1%D0%BE%D0%B7%D0%B4%D0%B0%D0%BD%D0%B8%D0%B5 %D0%B7%D0%B0%D0%BF%D1%80%D0%BE%D1%81%D0%B0 %D0%BD%D0%B0 %D0%B4%D0%BE%D1%81%D1%82%D1%83%D0%BF %D0%B2 %D1%81%D0%B8%D1%81%D1%82%D0%B5%D0%BC%D1%83 CMS (Contract Management System).aspx'>Создание запроса на доступ в систему CMS (Contract Management System) </a></li></ul>" & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem

    
End Sub
Sub CreateMail_Intune_instruction_reply()

    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = "<br>Здравствуйте!</br> <br>Для решения вашего вопроса необходимо выполнить все шаги по инструкции ниже <a href='https://multonpartners.sharepoint.com/sites/spaces-BSS-RUHD/KB/MDM_MAM_Intune/iPad/iPad_registration(Intune).pdf'>Intune</a><br>" & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub
Sub Mass_problem_red()

    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = "<br>Здравствуйте!</br> На данный момент сервис не доступен. При восстановлении работы сервиса, будет направлено дополнительное информационное сообщение.<br>" & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub
Sub Mass_problem_green()

    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = "<br>Здравствуйте!</br> <br>На данный момент работоспособность сервиса  восстановлена." & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub

Sub Mass_problem_yellow()

    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = "<br>Здравствуйте!</br> <br>На данный момент наблюдается снижение работоспособности уровня сервиса. При восстановлении работы сервиса, будет направлено дополнительное информационное сообщение." & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub
Sub Reply()




    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.htmlBody = "<br>Доброго времени суток!</br> <br>Ваш новый пароль для входа в систему SAP/NEXT - " & Pass & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem
    
End Sub

Function GetBoiler(ByVal sFile As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function


Public Sub CreateMail_SST_instruction()




Dim SigString As String
Dim Signature As String
SigString = Environ("appdata") & _
                "\Microsoft\Signatures\HD.htm"

    If Dir(SigString) <> "" Then
        Signature = GetBoiler(SigString)
    Else
        Signature = ""
    End If
    
    
    


Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br><font color='red'>Self Service Tool</font></br><br>Удобный и быстрый инструмент, который помогает найти ответы на технические, наиболее часто возникающие вопросы и упростить пользователю решение его проблем, доступен на компьютерах всем сотрудникам Компании.</br>" & _
"<br><font color='red'>Для запуска инструмента SST</font></br><br>Используйте ярлык Self Service Tool в панели уведомлений:</br>" & _
"<br><a><img src='\\runnvfd0\Photo\SST_foto\1.png'></a></br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\2.png'></a><br>" & _
"<br><font color='red'>SST включает в себя несколько опций, таких как:</font></br><br><strong>Первая помощь компьютеру</strong> - собирает информацию о сети при проблемах с доступом в интернет,обновляет групповые политики, очищает временные файлы.</br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\3.png'></a><br>" & _
"<br><font color='red'>Для сбора информации о сети </font></br>выбираете опцию <strong>Проблемы с доступом в интернет</strong>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\4.png'></a><br>" & _
"<br>Нажимаете на кнопку <strong>Снять логи</strong></br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\5.png'></a><br>" & _
"<br>Система генерирует технические файлы формата .txt, которые находятся на диске <strong>C:\Temp</strong> </br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\6.png'></a><br>" & _
"<br>Все файлы необходимо скопировать и направить на почту <a href='mailto:helpdesk.regional@multonpartners.com'>helpdesk.regional@multonpartners.com</a> для дальнейшего анализа проблемы с доступом в интернет на Вашем ПК.</br>" & _
"<br><font color='red'>Для обновления групповых политик</font></br> выбираете опцию <strong>Обновить групповые политики</strong>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\7.png'></a><br>" & _
"<br>Нажимаете на кнопку <strong>Обновить групповые политики.</strong></br> Это необходимо для синхронизации вашего ПК с последними обновлениями групповой политики компании</strong><br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\8.png'></a><br></br>" & _
"<br><font color='red'>Для чистки временных файлов </font></br> выбираете опцию <strong>Чистка временных файлов</strong><br>Опция очищает компьютер от файлов, которые занимают лишнее место и влияют на быстродействие компьютера в программах. После очистки временных файлов быстродействие работы компьютера увеличивается.</br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\9.png'></a></br><br>Нажимаете на кнопку <strong>Очистить временные файлы</strong></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\10.png'></a></br><br><strong>Первая помощь</strong> браузеру - очищает кэш и куки браузера. </br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\11.png'></a></br><br>В зависимости от того браузера (основные браузера - Google Chrome, Edge), который вы используете, нажимаете на кнопку <strong>Почистить Cookies/Cache</strong></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\12.png'></a></br><br><strong>Информация о компьютере</strong> - показывает информацию о количестве подключенных сетевых дисках и принтеров, сколько занято места временными файлами и сколько занято и свободно памяти на данный момент.</br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\13.png'></a></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\14.png'></a></br><br><strong>Отправить сообщение через Service Portal, написать боту Velera, написать письмо - способ регистрации проблемы.</br></strong><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\15.png'></a></br><br>Выбираете более удобный способ и нажимаете на кнопку, программа автоматически переведет вас в соответствующий канал связи с нами.</br>" & _
"<br>Почитать <strong>Советы и подсказки</strong> - в этом разделе вы найдете полезную информацию.</br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\16.png'></a></br><br>При нажатии на эту кнопку вас перенаправит на web-страницу Советы и подсказки.</br><br><strong>Запросить доступ через MyID - запрос доступа к различным внутренним системам компании</strong></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\17.png'></a><br><br>Если у вас пропал доступ к внутренним системам компании. Нажимаете на кнопку <strong>Запросить доступ через MyID</strong>, вас автоматически перенаправить на сайт, где вы сможете самостоятельно запросить доступ.</br>" & _
"<br><font color='red'>Запрос дополнительной информации для специалиста </font></br><br>Когда с вами связался специалист и ему необходимо от вас информация - какой vpn клиент используете, ваш rf/ru, модель ПК. Вы не знаете, что ответить. Открываете Self Service Tool и на главном экране программы есть вся необходимая информация.</br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\18.png'></a><br>" & Signature




With MyEmail
    
.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
.Subject = "Инструкция по работе с SST"
.htmlBody = htmlBody

End With
MyEmail.Display
End Sub


Public Sub CreateMail_SST_instruction_reply()
Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = "<br><font color='red'>Self Service Tool</font></br><br>Удобный и быстрый инструмент, который помогает найти ответы на технические, наиболее часто возникающие вопросы и упростить пользователю решение его проблем, доступен на компьютерах всем сотрудникам Компании.</br>" & _
"<br><font color='red'>Для запуска инструмента SST</font></br><br>Используйте ярлык Self Service Tool в панели уведомлений:</br>" & _
"<br><a><img src='\\runnvfd0\Photo\SST_foto\1.png'></a></br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\2.png'></a><br>" & _
"<br><font color='red'>SST включает в себя несколько опций, таких как:</font></br><br><strong>Первая помощь компьютеру</strong> - собирает информацию о сети при проблемах с доступом в интернет,обновляет групповые политики, очищает временные файлы.</br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\3.png'></a><br>" & _
"<br><font color='red'>Для сбора информации о сети </font></br>выбираете опцию <strong>Проблемы с доступом в интернет</strong>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\4.png'></a><br>" & _
"<br>Нажимаете на кнопку <strong>Снять логи</strong></br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\5.png'></a><br>" & _
"<br>Система генерирует технические файлы формата .txt, которые находятся на диске <strong>C:\Temp</strong> </br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\6.png'></a><br>" & _
"<br>Все файлы необходимо скопировать и направить на почту <a href='mailto:helpdesk.regional@multonpartners.com'>helpdesk.regional@multonpartners.com</a> для дальнейшего анализа проблемы с доступом в интернет на Вашем ПК.</br>" & _
"<br><font color='red'>Для обновления групповых политик</font></br> выбираете опцию <strong>Обновить групповые политики</strong>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\7.png'></a><br>" & _
"<br>Нажимаете на кнопку <strong>Обновить групповые политики.</strong></br> Это необходимо для синхронизации вашего ПК с последними обновлениями групповой политики компании</strong><br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\8.png'></a><br></br>" & _
"<br><font color='red'>Для чистки временных файлов </font></br> выбираете опцию <strong>Чистка временных файлов</strong><br>Опция очищает компьютер от файлов, которые занимают лишнее место и влияют на быстродействие компьютера в программах. После очистки временных файлов быстродействие работы компьютера увеличивается.</br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\9.png'></a></br><br>Нажимаете на кнопку <strong>Очистить временные файлы</strong></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\10.png'></a></br><br><strong>Первая помощь</strong> браузеру - очищает кэш и куки браузера. </br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\11.png'></a></br><br>В зависимости от того браузера (основные браузера - Google Chrome, Edge), который вы используете, нажимаете на кнопку <strong>Почистить Cookies/Cache</strong></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\12.png'></a></br><br><strong>Информация о компьютере</strong> - показывает информацию о количестве подключенных сетевых дисках и принтеров, сколько занято места временными файлами и сколько занято и свободно памяти на данный момент.</br>" & _
"<br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\13.png'></a></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\14.png'></a></br><br><strong>Отправить сообщение через Service Portal, написать боту Velera, написать письмо - способ регистрации проблемы.</br></strong><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\15.png'></a></br><br>Выбираете более удобный способ и нажимаете на кнопку, программа автоматически переведет вас в соответствующий канал связи с нами.</br>" & _
"<br>Почитать <strong>Советы и подсказки</strong> - в этом разделе вы найдете полезную информацию.</br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\16.png'></a></br><br>При нажатии на эту кнопку вас перенаправит на web-страницу Советы и подсказки.</br><br><strong>Запросить доступ через MyID - запрос доступа к различным внутренним системам компании</strong></br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\17.png'></a><br><br>Если у вас пропал доступ к внутренним системам компании. Нажимаете на кнопку <strong>Запросить доступ через MyID</strong>, вас автоматически перенаправить на сайт, где вы сможете самостоятельно запросить доступ.</br>" & _
"<br><font color='red'>Запрос дополнительной информации для специалиста </font></br><br>Когда с вами связался специалист и ему необходимо от вас информация - какой vpn клиент используете, ваш rf/ru, модель ПК. Вы не знаете, что ответить. Открываете Self Service Tool и на главном экране программы есть вся необходимая информация.</br><br><a><img width='500' height='300' src='\\runnvfd0\Photo\SST_foto\18.png'></a><br>" & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        Call move_TTstatus
    Next olItem

    
End Sub

Public Sub HD_lite_advisor()
'Отправка выбранного письма в папку "TT opened" и добавление номера тикета к телу письма
    Dim dump As Integer
    dump = DeclareVariables()
    Dim myolApp As New Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Dim myInbox As Outlook.MAPIFolder
    Dim myMailBox As Outlook.Recipient
    Dim myDestFolder As Outlook.MAPIFolder
    Dim myItems As Outlook.Items
    Dim myItem  As Object, msg  As Object, obj As Object
    Dim inputValue As String
    
    
    Dim objCategory As Category
    Set myNamespace = myolApp.GetNamespace("MAPI")
    Set myMailBox = myNamespace.CreateRecipient("helpdesk.regional@multonpartners.com")
    Set myInbox = myNamespace.GetSharedDefaultFolder(myMailBox, olFolderInbox).Parent
    Set myItems = myInbox.Items
    Set myDestFolder = myInbox.Folders("Ext _users")

    
    
    
    For Each obj In ActiveExplorer.Selection
       If TypeName(obj) = "MailItem" Then
        Set msg = obj
         inputValue = InputBox("Введите номер запроса", "Номер запроса", "INC")
         
         
                                    
         msg.Categories = G_myColor
         msg.Subject = inputValue & " | " & msg.Subject
         
         Dim MyEmail As MailItem
Dim htmlBody As String


Set MyEmail = Application.CreateItem(olMailItem)


htmlBody = "<br>Здравствуйте!</br> Ваше письмо принято в работу - " & inputValue




    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem
    For Each olItem In Application.ActiveExplorer.Selection
    Set olReply = olItem.Reply
            olReply.SentOnBehalfOfName = "helpdesk.regional@multonpartners.com"
            olReply.htmlBody = htmlBody & vbCrLf & olReply.htmlBody
        olReply.Display

        'olReply.Send
        
    Next olItem
    
    
    
    
    
        
         
         On Error Resume Next
           msg.Move myDestFolder
           msg.UnRead = False
         On Error GoTo -1
   
       End If
       
     Next obj
End Sub

