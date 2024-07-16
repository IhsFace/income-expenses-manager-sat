Imports System.Xml

Public Class Form1
    Dim xmlDoc As New XmlDocument()
    Dim xmlFilePath As String = "IncomeExpenseRecords.xml"

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

    Private Sub btnViewRecords_Click(sender As Object, e As EventArgs) Handles btnViewRecords.Click
        LoadRecords()
    End Sub

    Private Sub btnViewReport_Click(sender As Object, e As EventArgs) Handles btnViewReport.Click
        LoadFinancialReport()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        cbxCategory.SelectedIndex = 0
        txtIncomeExpense.Clear()
        txtSelectedRecord.Clear()
        lstDisplay.Items.Clear()
        cbxCategory.Focus()
    End Sub

    Private Sub LoadRecords()
        Dim strCategory As String = cbxCategory.SelectedItem
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Category: " & strCategory)

        For Each recordNode As XmlNode In recordNodes
            If strCategory = "All" Or recordNode.SelectSingleNode("Category").InnerText = strCategory Then
                lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
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
