Imports Microsoft.Exchange.WebServices.Data
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim count As Integer = 0
            Dim service As New ExchangeService(ExchangeVersion.Exchange2010_SP2)
            'set username, password
            service.Credentials = New WebCredentials("xxxxx", "xxxxx")
            'set email server
            service.AutodiscoverUrl("xxxxx")
            Dim view As ItemView = New ItemView(10)
            Dim findResults As FindItemsResults(Of Item) = service.FindItems(WellKnownFolderName.Inbox, New ItemView(10))
            If findResults IsNot Nothing AndAlso findResults.Items IsNot Nothing AndAlso findResults.Items.Count > 0 Then

                For Each item As Item In findResults.Items
                    Dim message As EmailMessage = EmailMessage.Bind(service, item.Id, New PropertySet(BasePropertySet.IdOnly, ItemSchema.Attachments, ItemSchema.HasAttachments))
                    Dim attachmentCount As Integer = 1

                    For Each attachment As Attachment In message.Attachments

                        If TypeOf attachment Is FileAttachment Then
                            Dim fileAttachment As FileAttachment = TryCast(attachment, FileAttachment)
                            'set path of saving item and file name
                            fileAttachment.Load("D:\Attachments\" & item.DateTimeSent.ToString("yyyyMMddHHmm") & "-" & attachmentCount & "-" & fileAttachment.Name)
                            Console.WriteLine("Attachment name: " & fileAttachment.Name)
                            count += 1
                            attachmentCount += 1
                            Console.WriteLine(attachmentCount)
                        End If
                    Next

                    Console.WriteLine(item.Subject)
                Next
            Else
                Console.WriteLine("no items")
            End If
            MessageBox.Show("Successfully saved " & count & " attachments to folder: D:\Attachments\ ")
        Catch ex As Exception
            MessageBox.Show(String.Format("Error: {0}", ex.Message))
        End Try
    End Sub
End Class
