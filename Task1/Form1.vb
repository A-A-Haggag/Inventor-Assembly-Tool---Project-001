Imports Inventor
Imports System.Runtime.InteropServices

Public Class Form1
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim inventorApp As Inventor.Application

        Try
            inventorApp = Marshal.GetActiveObject("Inventor.Application")
            MessageBox.Show("Connected with Inventor succesfully")

        Catch ex As Exception
            MessageBox.Show("Cannot connect to Inventor")
        End Try



        Dim oAssemDoc As AssemblyDocument
        oAssemDoc = inventorApp.ActiveDocument

        Dim n = numberOfOcc(oAssemDoc.ComponentDefinition)

        MsgBox(n & " occurences")
        MsgBox(oAssemDoc.AllReferencedDocuments.Count & " kind of parts")


        If ComboBox1.Text = ("Names") Then
            ListBox1.Items.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(1).Name)
            ListBox1.Items.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(2).Name)
        ElseIf ComboBox1.Text = ("File Paths") Then
            ListBox1.Items.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.FullDocumentName)
            ListBox1.Items.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.FullDocumentName)
        ElseIf ComboBox1.Text = ("Materials") Then
            ListBox1.Items.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.ActiveMaterial.DisplayName)
            ListBox1.Items.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.ActiveMaterial.DisplayName)
        ElseIf ComboBox1.Text = ("Document Type") Then
            If oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}" Then
                ListBox1.Items.Add("part")
            ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                ListBox1.Items.Add("sheet metal")
            ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Then
                ListBox1.Items.Add("Assembley")
            End If
            If oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}" Then
                ListBox1.Items.Add("part")
            ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                ListBox1.Items.Add("sheet metal")
            ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Then
                ListBox1.Items.Add("Assembley")
            End If
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim inventorApp As Inventor.Application

        Try
            inventorApp = Marshal.GetActiveObject("Inventor.Application")
            MessageBox.Show("Connected with Inventor succesfully")

        Catch ex As Exception
            MessageBox.Show("Cannot connect to Inventor")
        End Try

        Dim oAssemDoc As AssemblyDocument
        oAssemDoc = inventorApp.ActiveDocument

        Dim n = numberOfOcc(oAssemDoc.ComponentDefinition)

        MsgBox(n & " occurences")
        MsgBox(oAssemDoc.AllReferencedDocuments.Count & " kind of parts")

        TreeView1.Nodes.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(1).Name)
        TreeView1.Nodes(0).Nodes.Add(New TreeNode(oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.FullDocumentName))
        TreeView1.Nodes(0).Nodes.Add(New TreeNode(oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.ActiveMaterial.DisplayName))
        If oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}" Then
            TreeView1.Nodes(0).Nodes.Add(New TreeNode("part"))
        ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
            TreeView1.Nodes(0).Nodes.Add(New TreeNode("sheet metal"))
        ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(1).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Then
            TreeView1.Nodes(0).Nodes.Add(New TreeNode("Assembley"))
        End If

        TreeView1.Nodes.Add(oAssemDoc.ComponentDefinition.Occurrences.Item(2).Name)
        TreeView1.Nodes(1).Nodes.Add(New TreeNode(oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.FullDocumentName))
        TreeView1.Nodes(1).Nodes.Add(New TreeNode(oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.ActiveMaterial.DisplayName))
        If oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{4D29B490-49B2-11D0-93C3-7E0706000000}" Then
            TreeView1.Nodes(1).Nodes.Add(New TreeNode("part"))
        ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
            TreeView1.Nodes(1).Nodes.Add(New TreeNode("sheet metal"))
        ElseIf oAssemDoc.ComponentDefinition.Occurrences.Item(2).ReferencedDocumentDescriptor.ReferencedDocument.SubType = "{E60F81E1-49B3-11D0-93C3-7E0706000000}" Then
            TreeView1.Nodes(1).Nodes.Add(New TreeNode("Assembley"))
        End If

    End Sub
    Public Function numberOfOcc(def As ComponentDefinition) As Integer
        Dim n As Integer = def.Occurrences.Count

        For Each occ As ComponentOccurrence In def.Occurrences
            If (occ.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                n = n + numberOfOcc(occ.Definition)
            End If
        Next

        Return n
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim inventorApp As Inventor.Application
        Try
            inventorApp = Marshal.GetActiveObject("Inventor.Application")
            ''MessageBox.Show("Connected with Inventor succesfully")

        Catch ex As Exception
            MessageBox.Show("Cannot connect to Inventor")
        End Try

        Dim oAssemDoc As AssemblyDocument
        oAssemDoc = inventorApp.ActiveDocument


        Dim selpart As ComponentOccurrence
        Dim part As PartDocument
        selpart = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select Part.")
        part = selpart.ReferencedDocumentDescriptor.ReferencedDocument


        MsgBox("part name:" & selpart._DisplayName)

        Dim invSummaryInfo As PropertySet
        invSummaryInfo = part.PropertySets.Item("Inventor Summary Information")
        Dim Author As [Property] = invSummaryInfo.Item("Author")
        MsgBox("Part Author: " & Author.Value)

        Dim invDesignInfo As PropertySet
        invDesignInfo = part.PropertySets.Item("Design Tracking Properties")
        Dim invPartNumberProperty As [Property]
        invPartNumberProperty = invDesignInfo.Item("Part Number")
        MsgBox("part no: " & invPartNumberProperty.Value)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim inventorApp As Inventor.Application
        Try
            inventorApp = Marshal.GetActiveObject("Inventor.Application")
            ''MessageBox.Show("Connected with Inventor succesfully")

        Catch ex As Exception
            MessageBox.Show("Cannot connect to Inventor")
        End Try

        Dim oAssemDoc As AssemblyDocument
        oAssemDoc = inventorApp.ActiveDocument


        Dim selpart As ComponentOccurrence
        Dim part As String
        selpart = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select Part.")
        part = selpart.ReferencedDocumentDescriptor.FullDocumentName


        'get place of\ in path
        'Dim oSource As String = part
        ' Dim oTarget As String = "\"
        ' MsgBox("InStr(oSource, oTarget) = " & InStr(oSource, oTarget))

        MsgBox(part)
        MsgBox(FilePath(part))
        MsgBox(BaseFilename(part))
        MsgBox(FileExtension(part))


    End Sub

    ' Return the path of the input filename.
    Public Function FilePath(ByVal fullFilename As String) As String
        ' Extract the path by getting everything up to and including the last backslash "\".
        FilePath = Microsoft.VisualBasic.Strings.Left(fullFilename, InStrRev(fullFilename, "\") - 1)
    End Function

    ' Return the base name of the input filename, without the path or the extension.
    Public Function BaseFilename(ByVal fullFilename As String) As String
        ' Extract the filename by getting everything to the right of the last backslash.
        Dim temp As String
        temp = Microsoft.VisualBasic.Strings.Right(fullFilename, (Len(fullFilename) - InStrRev(fullFilename, "\")))

        ' Get the base filename by getting everything to the left of the last period ".".
        BaseFilename = Microsoft.VisualBasic.Strings.Left(temp, InStrRev(temp, ".") - 1)
    End Function

    ' Return the extension of the input filename.
    Public Function FileExtension(ByVal fullFilename As String) As String
        ' Extract the filename by getting everything to the right of the last backslash.
        Dim temp As String
        temp = Microsoft.VisualBasic.Strings.Right(fullFilename, (Len(fullFilename) - InStrRev(fullFilename, "\")))

        ' Get the base filename by getting everything to the right of the last period ".".
        FileExtension = Microsoft.VisualBasic.Strings.Right(temp, Len(temp) - InStrRev(temp, ".") + 1)
    End Function

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form2.Show()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        System.Diagnostics.Process.Start("https://enerfacprojects.com/")
    End Sub
End Class