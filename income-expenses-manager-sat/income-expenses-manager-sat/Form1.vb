Imports System.Xml

Public Class Form1
    Dim xmlDoc As New XmlDocument()
    Dim xmlFilePath As String = "IncomeExpenseRecords.xml"
    Dim intDisplayedRecords As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If System.IO.File.Exists(xmlFilePath) Then
            xmlDoc.Load(xmlFilePath)
            PopulateComboBox()
            LoadRecords()
        Else
            Dim xmlDeclaration As XmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
            xmlDoc.AppendChild(xmlDeclaration)
            Dim rootNode As XmlNode = xmlDoc.CreateElement("Records")
            xmlDoc.AppendChild(rootNode)
            xmlDoc.Save(xmlFilePath)
        End If
    End Sub

    Private Sub btnAddCategory_Click(sender As Object, e As EventArgs) Handles btnAddCategory.Click
        Dim strName As String = InputBox("Enter the name of the category:", "Add Category", "Category")

        Dim categoryNode As XmlNode = xmlDoc.CreateElement("Category")

        Dim nameNode As XmlNode = xmlDoc.CreateElement("Name")
        nameNode.InnerText = strName
        categoryNode.AppendChild(nameNode)

        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
        rootNode.AppendChild(categoryNode)

        xmlDoc.Save(xmlFilePath)

        PopulateComboBox()
        LoadRecords()

        MessageBox.Show("Category added successfully!", "Add Category")
    End Sub

    Private Sub btnEditCategory_Click(sender As Object, e As EventArgs) Handles btnEditCategory.Click
        If cbxCategory.SelectedIndex = -1 Or cbxCategory.SelectedIndex = 0 Then
            MessageBox.Show("Please select a category!", "Edit Category")
            Return
        End If

        Dim strName As String = InputBox("Enter the name of the category:", "Edit Category", "Category")

        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/Category")

        For Each categoryNode As XmlNode In categoryNodes
            Dim nameNode As XmlNode = categoryNode.SelectSingleNode("Name")

            If nameNode IsNot Nothing And nameNode.InnerText = cbxCategory.SelectedItem Then
                categoryNode.RemoveChild(nameNode)
                Dim newNameNode As XmlNode = xmlDoc.CreateElement("Name")
                newNameNode.InnerText = strName
                categoryNode.AppendChild(newNameNode)

                Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

                For Each recordNode As XmlNode In recordNodes
                    Dim selectedCategoryNode As XmlNode = recordNode.SelectSingleNode("Category")

                    If selectedCategoryNode IsNot Nothing And selectedCategoryNode.InnerText = cbxCategory.SelectedItem Then
                        recordNode.RemoveChild(selectedCategoryNode)
                        Dim newCategoryNode As XmlNode = xmlDoc.CreateElement("Category")
                        newCategoryNode.InnerText = strName
                        recordNode.AppendChild(newCategoryNode)
                    End If
                Next

                xmlDoc.Save(xmlFilePath)

                PopulateComboBox()
                LoadRecords()

                MessageBox.Show("Category edited successfully!", "Edit Category")
            End If
        Next
    End Sub

    Private Sub btnDeleteCategory_Click(sender As Object, e As EventArgs) Handles btnDeleteCategory.Click
        If cbxCategory.SelectedIndex = -1 Or cbxCategory.SelectedIndex = 0 Then
            MessageBox.Show("Please select a category!", "Delete Category")
            Return
        End If

        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/Category")
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")

        For Each recordNode As XmlNode In recordNodes
            Dim selectedCategoryNode As XmlNode = recordNode.SelectSingleNode("Category")

            If selectedCategoryNode IsNot Nothing And selectedCategoryNode.InnerText = cbxCategory.SelectedItem Then
                rootNode.RemoveChild(recordNode)

                For Each categoryNode As XmlNode In categoryNodes
                    Dim nameNode As XmlNode = categoryNode.SelectSingleNode("Name")

                    If nameNode IsNot Nothing And nameNode.InnerText = cbxCategory.SelectedItem Then
                        rootNode.RemoveChild(categoryNode)
                    End If
                Next

                xmlDoc.Save(xmlFilePath)

                PopulateComboBox()
                LoadRecords()

                MessageBox.Show("Category deleted successfully!", "Delete Category")
            End If
        Next
    End Sub

    Private Sub btnAddIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnAddIncomeExpense.Click
        If cbxCategory.SelectedIndex = -1 Or cbxCategory.SelectedIndex = 0 Then
            MessageBox.Show("Please select a category!", "Add Income/Expense")
            Return
        End If

        Dim intMaxId As Integer = -1
        Dim intNewId As Integer

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        For Each selectedNode As XmlNode In recordNodes
            Dim selectedIdNode As XmlNode = selectedNode.SelectSingleNode("Id")

            If selectedIdNode IsNot Nothing And Val(selectedIdNode.InnerText) > intMaxId Then
                intMaxId = Val(selectedNode.SelectSingleNode("Id").InnerText)
            End If
        Next

        intNewId = intMaxId + 1

        Dim recordNode As XmlNode = xmlDoc.CreateElement("Record")

        Dim idNode As XmlNode = xmlDoc.CreateElement("Id")
        idNode.InnerText = intNewId
        recordNode.AppendChild(idNode)

        Dim categoryNode As XmlNode = xmlDoc.CreateElement("Category")
        categoryNode.InnerText = cbxCategory.SelectedItem
        recordNode.AppendChild(categoryNode)

        Dim amountNode As XmlNode = xmlDoc.CreateElement("Amount")
        amountNode.InnerText = Val(txtIncomeExpense.Text)
        recordNode.AppendChild(amountNode)

        Dim createdAtNode As XmlNode = xmlDoc.CreateElement("CreatedAt")
        createdAtNode.InnerText = Date.Now()
        recordNode.AppendChild(createdAtNode)

        Dim updatedAtNode As XmlNode = xmlDoc.CreateElement("UpdatedAt")
        updatedAtNode.InnerText = Date.Now()
        recordNode.AppendChild(updatedAtNode)

        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
        rootNode.AppendChild(recordNode)

        xmlDoc.Save(xmlFilePath)

        LoadRecords()

        MessageBox.Show("Income/Expense added successfully!", "Add Income/Expense")
    End Sub

    Private Sub btnEditIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnEditIncomeExpense.Click
        If txtIncomeExpense.Text = "" Then
            MessageBox.Show("Please enter an amount!", "Edit Income/Expense")
            Return
        End If
        If txtSelectedRecord.Text = "" Then
            MessageBox.Show("Please select a record!", "Edit Income/Expense")
            Return
        End If

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        For Each recordNode As XmlNode In recordNodes
            Dim idNode As XmlNode = recordNode.SelectSingleNode("Id")
            Dim amountNode As XmlNode = recordNode.SelectSingleNode("Amount")
            Dim updatedAtNode As XmlNode = recordNode.SelectSingleNode("UpdatedAt")

            If idNode IsNot Nothing And idNode.InnerText = Val(txtSelectedRecord.Text) Then
                If amountNode IsNot Nothing And updatedAtNode IsNot Nothing Then
                    recordNode.RemoveChild(recordNode.SelectSingleNode("Amount"))
                    Dim newAmountNode As XmlNode = xmlDoc.CreateElement("Amount")
                    newAmountNode.InnerText = txtIncomeExpense.Text
                    recordNode.AppendChild(newAmountNode)

                    recordNode.RemoveChild(recordNode.SelectSingleNode("UpdatedAt"))
                    Dim newUpdatedAtNode As XmlNode = xmlDoc.CreateElement("UpdatedAt")
                    newUpdatedAtNode.InnerText = Date.Now()
                    recordNode.AppendChild(newUpdatedAtNode)

                    xmlDoc.Save(xmlFilePath)

                    LoadRecords()

                    MessageBox.Show("Income/Expense edited successfully!", "Edit Income/Expense")
                End If
            End If
        Next
    End Sub

    Private Sub btnDeleteIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnDeleteIncomeExpense.Click
        If txtSelectedRecord.Text = "" Then
            MessageBox.Show("Please select a record!", "Edit Income/Expense")
            Return
        End If

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")

        For Each recordNode As XmlNode In recordNodes
            Dim idNode As XmlNode = recordNode.SelectSingleNode("Id")

            If idNode IsNot Nothing And idNode.InnerText = Val(txtSelectedRecord.Text) Then
                rootNode.RemoveChild(recordNode)

                xmlDoc.Save(xmlFilePath)

                LoadRecords()

                MessageBox.Show("Income/Expense deleted successfully!", "Delete Income/Expense")
            End If
        Next
    End Sub

    Private Sub btnViewRecords_Click(sender As Object, e As EventArgs) Handles btnViewRecords.Click
        LoadRecords()
    End Sub

    Private Sub btnViewReport_Click(sender As Object, e As EventArgs) Handles btnViewReport.Click
        LoadFinancialReport()
    End Sub

    Private Sub btnSortRecords_Click(sender As Object, e As EventArgs) Handles btnSortRecords.Click
        LoadRecords()

        Dim intNbrsArray(intDisplayedRecords - 1) As Integer
        Dim intCurrentIndex As Integer = 0

        Dim strSortOption As String = InputBox("Enter the sort option for the records (Id or Amount):", "Sort Records", "Amount")
        Dim strSortOrder As String = InputBox("Sort the records in descending order (Y or N):", "Sort Records", "N")

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        For Each recordNode As XmlNode In recordNodes
            Dim selectedNode As XmlNode = recordNode.SelectSingleNode(strSortOption)
            intNbrsArray(intCurrentIndex) = (Val(selectedNode.InnerText))
            intCurrentIndex = intCurrentIndex + 1
        Next

        Dim intLowerBound As Integer = LBound(intNbrsArray)
        Dim intUpperBound As Integer = UBound(intNbrsArray)

        Dim i As Integer = 0
        Dim j As Integer = 0

        Dim intMinValueIndex As Integer
        Dim intSwapValueIndex As Integer
        Dim intSwapValue As Integer

        For i = intLowerBound To intUpperBound
            intMinValueIndex = i
            intSwapValue = 0
            intSwapValueIndex = 0

            For j = i + 1 To intUpperBound
                If intNbrsArray(j) < intNbrsArray(intMinValueIndex) Then
                    If intSwapValue = 0 Then
                        intSwapValue = intNbrsArray(j)
                        intSwapValueIndex = j
                    Else
                        If intNbrsArray(j) < intSwapValue Then
                            intSwapValue = intNbrsArray(j)
                            intSwapValueIndex = j
                        End If
                    End If
                End If
            Next

            If intSwapValue <> 0 Then
                intNbrsArray(intSwapValueIndex) = intNbrsArray(intMinValueIndex)
                intNbrsArray(intMinValueIndex) = intSwapValue
            End If
        Next

        If strSortOrder = "Y" Then
            Array.Reverse(intNbrsArray)
        End If

        lstDisplay.Items.Clear()

        For c = 0 To intNbrsArray.Length - 1
            For Each recordNode As XmlNode In recordNodes
                If recordNode.SelectSingleNode(strSortOption).InnerText = intNbrsArray(c) Then
                    lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
                End If
            Next
        Next
    End Sub

    Private Sub btnSearchRecords_Click(sender As Object, e As EventArgs) Handles btnSearchRecords.Click
        LoadRecords()

        Dim intNbrsArray(intDisplayedRecords - 1) As Integer
        Dim intCurrentIndex As Integer = 0

        Dim strSearchRequestType As String = InputBox("Enter the search request type (Id or Amount):", "Search Records", "Amount")
        Dim strSearchRequestItem As String = InputBox("Enter the search request item:", "Search Records", "100")

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        For Each recordNode As XmlNode In recordNodes
            Dim selectedNode As XmlNode = recordNode.SelectSingleNode(strSearchRequestType)
            intNbrsArray(intCurrentIndex) = (Val(selectedNode.InnerText))
            intCurrentIndex = intCurrentIndex + 1
        Next

        Dim blnItemFound As Boolean = False
        Dim blnEndOfArray As Boolean = False
        Dim intSubscript As Integer = 0
        Dim intMaxSubscript As Integer = 0

        intMaxSubscript = intNbrsArray.Length - 1

        While blnItemFound = False And blnEndOfArray = False
            If strSearchRequestItem = intNbrsArray(intSubscript) Then
                blnItemFound = True
            ElseIf intSubscript = intMaxSubscript Then
                blnEndOfArray = True
            Else
                intSubscript = intSubscript + 1
            End If
        End While

        If blnItemFound = True Then
            lstDisplay.Items.Clear()

            For Each recordNode As XmlNode In recordNodes
                If recordNode.SelectSingleNode(strSearchRequestType).InnerText = strSearchRequestItem Then
                    lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
                End If
            Next
        Else
            MessageBox.Show("Item not found!", "Search Records")
        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        cbxCategory.SelectedIndex = 0
        txtIncomeExpense.Clear()
        txtSelectedRecord.Clear()
        lstDisplay.Items.Clear()
        cbxCategory.Focus()
    End Sub

    Private Sub LoadRecords()
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        intDisplayedRecords = 0

        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Category: " & cbxCategory.SelectedItem)

        For Each recordNode As XmlNode In recordNodes
            If cbxCategory.SelectedItem = "All" Or recordNode.SelectSingleNode("Category").InnerText = cbxCategory.SelectedItem Then
                lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
                intDisplayedRecords = intDisplayedRecords + 1
            End If
        Next
    End Sub

    Private Sub LoadFinancialReport()
        Dim intTotalIncome As Integer
        Dim intTotalExpenses As Integer
        Dim strTimePeriod As String = InputBox("Enter the time period for the financial report (e.g. 2024):", "Financial Report", "2024")

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        lstDisplay.Items.Clear()

        For Each recordNode As XmlNode In recordNodes
            Dim updatedAtNode As XmlNode = recordNode.SelectSingleNode("UpdatedAt")

            If updatedAtNode IsNot Nothing And updatedAtNode.InnerText.Contains(strTimePeriod) Then
                If recordNode.SelectSingleNode("Amount").InnerText.StartsWith("-") Then
                    intTotalExpenses = intTotalExpenses + Val(recordNode.SelectSingleNode("Amount").InnerText)
                Else
                    intTotalIncome = intTotalIncome + Val(recordNode.SelectSingleNode("Amount").InnerText)
                End If
            End If
        Next

        lstDisplay.Items.Add("Total Income: " & intTotalIncome)
        lstDisplay.Items.Add("Total Expenses: " & intTotalExpenses)
        lstDisplay.Items.Add("Net Income: " & (intTotalIncome + intTotalExpenses))
        lstDisplay.Items.Add("Time Period: " & strTimePeriod)
    End Sub

    Private Sub PopulateComboBox()
        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/Category")

        cbxCategory.Items.Clear()
        cbxCategory.Items.Add("All")

        For Each categoryNode As XmlNode In categoryNodes
            cbxCategory.Items.Add(categoryNode.SelectSingleNode("Name").InnerText)
        Next

        cbxCategory.SelectedIndex = 0
    End Sub
End Class
