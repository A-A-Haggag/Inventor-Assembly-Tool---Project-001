Imports System.Runtime.InteropServices
Imports Inventor

Public Class Form2

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
        Dim part As PartDocument
        selpart = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select Part.")
        part = selpart.ReferencedDocumentDescriptor.ReferencedDocument

        Dim customPropSet As PropertySet
        customPropSet = part.PropertySets.Item("User Defined Properties")

        '''' Get the property named "testz".
        'Dim customProp As [Property]
        'customProp = customPropSet.Item("testz")
        ' Display the value of the iProperty.
        'MsgBox("Sample1 = " & customProp.Value)

        If ComboBox1.Text = ("Text") Then
            Dim textValue As String
            textValue = TextBox5.Text
            Call customPropSet.Add(textValue, TextBox4.Text)

        ElseIf ComboBox1.Text = ("Number") Then
            ' Create a new number property.
            Dim numberValue As Double
            numberValue = TextBox5.Text
            Call customPropSet.Add(numberValue, TextBox4.Text)
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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

        Dim invSummaryInfo As PropertySet
        invSummaryInfo = part.PropertySets.Item("Inventor Summary Information")
        Dim userName As [Property] = invSummaryInfo.Item("Author")
        userName.Value = TextBox2.Text
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
        Dim oParameters As Parameters
        selpart = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select Part.")
        part = selpart.ReferencedDocumentDescriptor.ReferencedDocument

        oParameters = part.ComponentDefinition.Parameters

        ' Get the parameter named "Length".
        Dim oLengthParam As Parameter
        oLengthParam = oParameters.Item("d0")
        ' Change the equation of the parameter.
        oLengthParam.Expression = TextBox3.Text + ComboBox2.Text
        ' Update the document.
        inventorApp.ActiveDocument.Update()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
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
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
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
        Dim oParameters As Parameters
        selpart = inventorApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Select Part.")
        part = selpart.ReferencedDocumentDescriptor.ReferencedDocument

        oParameters = part.ComponentDefinition.Parameters

        ' Get the parameter named "Length".
        Dim oaddParam As Parameter
        Dim oUserParams As UserParameters
        oUserParams = oParameters.UserParameters
        If ComboBox3.Text = "mm" Then
            oaddParam = oUserParams.AddByExpression(TextBox7.Text, TextBox6.Text, UnitsTypeEnum.kMillimeterLengthUnits)
        ElseIf ComboBox3.Text = "inch" Then
            oaddParam = oUserParams.AddByExpression(TextBox7.Text, TextBox6.Text, UnitsTypeEnum.kInchLengthUnits)
        ElseIf ComboBox3.Text = "default" Then
            oaddParam = oUserParams.AddByValue(TextBox7.Text, TextBox6.Text, UnitsTypeEnum.kDefaultDisplayLengthUnits)
        End If
        inventorApp.ActiveDocument.Update()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        System.Diagnostics.Process.Start("https://enerfacprojects.com/")
    End Sub

End Class