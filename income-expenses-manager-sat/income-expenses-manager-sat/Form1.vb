' PROGRAM NAME: Income Expenses Manager
' PURPOSE: To manage income and expense records
' AUTHOR: Ihsaan Ifzal
' DATE CREATED: 17/06/2024
Imports System.Xml

Public Class Form1
    ' Declare the XmlDocument object and the XML file path
    Dim xmlDoc As New XmlDocument()
    Dim xmlFilePath As String = "IncomeExpenseRecords.xml"
    ' Declare the integer variable to store the number of displayed records in the list box
    ' This variable is used across multiple subroutines so it is declared at the class level
    Dim intDisplayedRecords As Integer = 0
    ' Declare the boolean variable to store whether the user is editing a record
    Dim blnEditing As Boolean = False

    ' The Form1_Load event handler is executed when the form is loaded
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Check if the XML file exists
        If System.IO.File.Exists(xmlFilePath) Then
            ' Load the XML file
            xmlDoc.Load(xmlFilePath)
        Else
            ' Create a new XML file with the root node
            Dim xmlDeclaration As XmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
            xmlDoc.AppendChild(xmlDeclaration)
            Dim rootNode As XmlNode = xmlDoc.CreateElement("records")
            xmlDoc.AppendChild(rootNode)

            ' Create the budgetLimit node and the limit node
            Dim budgetLimitNode As XmlNode = xmlDoc.CreateElement("budgetLimit")
            Dim limitNode As XmlNode = xmlDoc.CreateElement("limit")

            ' Set the inner text of the limit node to 0 and append it to the budgetLimit node, then append the budgetLimit node to the root node
            limitNode.InnerText = "0"
            budgetLimitNode.AppendChild(limitNode)
            rootNode.AppendChild(budgetLimitNode)

            ' Save the XML file
            xmlDoc.Save(xmlFilePath)
        End If

        'Populate the combo box and display the financial report
        PopulateComboBox()
        LoadFinancialReport("all")
    End Sub

    ' The btnAddCategory_Click event handler is executed when the Add Category button is clicked
    Private Sub btnAddCategory_Click(sender As Object, e As EventArgs) Handles btnAddCategory.Click
        ' Get the name entered by the user
        Dim strName As String = txtName.Text

        ' Check if the user has entered a name
        If String.IsNullOrWhiteSpace(strName) Then
            ' Display a message box if the user has not entered a name
            MessageBox.Show("Please enter the name of the category!", "Add Category")
            Return
        End If

        ' Create the category node and the name node
        Dim categoryNode As XmlNode = xmlDoc.CreateElement("category")
        Dim nameNode As XmlNode = xmlDoc.CreateElement("name")

        ' Set the inner text of the name node to the name entered by the user and append it to the category node
        nameNode.InnerText = strName
        categoryNode.AppendChild(nameNode)

        ' Get the root node and append the category node to it
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("records")
        rootNode.AppendChild(categoryNode)

        xmlDoc.Save(xmlFilePath)

        ' Populate the combo box and display the financial report
        PopulateComboBox()
        LoadFinancialReport("all")

        ' Display a message box to confirm that the category has been added
        MessageBox.Show("Category added successfully!", "Add Category")
    End Sub

    ' The btnEditCategory_Click event handler is executed when the Edit Category button is clicked
    Private Sub btnEditCategory_Click(sender As Object, e As EventArgs) Handles btnEditCategory.Click
        ' Check if the user has selected a category
        If cbxCategory.SelectedIndex <= 0 Then
            ' Display a message box if the user has not selected a category
            MessageBox.Show("Please select a category!", "Edit Category")
            Return
        End If

        ' Get the name entered by the user
        Dim strName As String = txtName.Text

        ' Check if the user has entered a name
        If String.IsNullOrWhiteSpace(strName) Then
            ' Display a message box if the user has not entered a name
            MessageBox.Show("Please enter the name of the category!", "Edit Category")
            Return
        End If

        ' Get the Category nodes and loop through them
        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("records/category")

        For Each categoryNode As XmlNode In categoryNodes
            ' Get the name node of the current category node
            Dim nameNode As XmlNode = categoryNode.SelectSingleNode("name")

            ' Check if the inner text of the name node is equal to the selected category
            If nameNode.InnerText = cbxCategory.SelectedItem Then
                Try
                    ' Remove the name node
                    categoryNode.RemoveChild(nameNode)

                    ' Create a new name node and set the inner text of the name node to the name entered by the user and append it to the category node
                    Dim newNameNode As XmlNode = xmlDoc.CreateElement("name")

                    newNameNode.InnerText = strName
                    categoryNode.AppendChild(newNameNode)

                    ' Get the record nodes and loop through them
                    Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")

                    For Each recordNode As XmlNode In recordNodes
                        ' Get the category node of the current record node
                        Dim selectedCategoryNode As XmlNode = recordNode.SelectSingleNode("category")

                        ' Check if the inner text of the category node is equal to the selected category
                        If selectedCategoryNode.InnerText = cbxCategory.SelectedItem Then
                            Try
                                ' Remove the category node
                                recordNode.RemoveChild(selectedCategoryNode)

                                ' Create a new category node and set the inner text of the category node to the name entered by the user and append it to the record node
                                Dim newCategoryNode As XmlNode = xmlDoc.CreateElement("category")

                                newCategoryNode.InnerText = strName
                                recordNode.AppendChild(newCategoryNode)
                            Catch ex As System.NullReferenceException
                                ' Display a message box if the category node was not found in the selected record's XML record
                                MessageBox.Show("The category node was not found in the selected record's XML record. Please delete your XML file or manually add in a category node.", "Edit Category")
                            End Try
                        End If
                    Next

                    ' Save the XML file
                    xmlDoc.Save(xmlFilePath)

                    ' Populate the combo box and display the financial report
                    PopulateComboBox()
                    LoadFinancialReport("all")

                    ' Display a message box to confirm that the category has been edited
                    MessageBox.Show("Category edited successfully!", "Edit Category")
                Catch ex As System.NullReferenceException
                    ' Display a message box if the name node was not found in the selected category's XML record
                    MessageBox.Show("The name node was not found in the selected category's XML record. Please delete your XML file or manually add in a name node.", "Edit Category")
                End Try
            End If
        Next
    End Sub

    ' The btnDeleteCategory_Click event handler is executed when the Delete Category button is clicked
    Private Sub btnDeleteCategory_Click(sender As Object, e As EventArgs) Handles btnDeleteCategory.Click
        ' Check if the user has selected a category
        If cbxCategory.SelectedIndex <= 0 Then
            ' Display a message box if the user has not selected a category
            MessageBox.Show("Please select a category!", "Delete Category")
            Return
        End If

        ' Display an input box to confirm the deletion of the category
        Dim strConfirm As String = LCase(InputBox("Are you sure you want to delete the category " & cbxCategory.SelectedItem & "? (y or n):", "Delete Category", "n"))

        ' Check if the user has confirmed the deletion
        If strConfirm <> "y" Then
            ' Display a message box if the user has not confirmed the deletion
            MessageBox.Show("Operation cancelled.", "Delete Category")
            Return
        ElseIf strConfirm = "y" Then
            ' Get the category nodes, the record nodes and the root node
            Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("records/category")
            Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("records")

            ' Loop through the category nodes
            For Each categoryNode As XmlNode In categoryNodes
                ' Get the name node of the current category node
                Dim nameNode As XmlNode = categoryNode.SelectSingleNode("name")

                ' Check if the inner text of the name node is equal to the selected category
                If nameNode.InnerText = cbxCategory.SelectedItem Then
                    Try
                        ' Remove the category node
                        rootNode.RemoveChild(categoryNode)

                        ' Loop through the record nodes
                        For Each recordNode As XmlNode In recordNodes
                            ' Get the category node of the current record node
                            Dim selectedCategoryNode As XmlNode = recordNode.SelectSingleNode("category")

                            ' Check if the inner text of the category node is equal to the selected category
                            If selectedCategoryNode.InnerText = cbxCategory.SelectedItem Then
                                Try
                                    ' Remove the record node
                                    rootNode.RemoveChild(recordNode)
                                Catch ex As System.NullReferenceException
                                    ' Display a message box if the category node was not found in the selected record's XML record
                                    MessageBox.Show("The category node was not found in the selected record's XML record. Please delete your XML file or manually add in a category node.", "Delete Category")
                                End Try
                            End If
                        Next
                    Catch ex As System.NullReferenceException
                        ' Display a message box if the category node was not found in the selected record's XML record
                        MessageBox.Show("The category node was not found in the selected record's XML record. Please delete your XML file or manually add in a category node.", "Delete Category")
                    End Try
                End If
            Next

            ' Save the XML file
            xmlDoc.Save(xmlFilePath)

            ' Populate the combo box and display the financial report
            PopulateComboBox()
            LoadFinancialReport("all")

            ' Display a message box to confirm that the category has been deleted
            MessageBox.Show("Category deleted successfully!", "Delete Category")
        End If
    End Sub

    ' The btnAddIncomeExpense_Click event handler is executed when the Add Income/Expense button is clicked
    Private Sub btnAddIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnAddIncomeExpense.Click
        ' Get the name, income and expense entered by the user
        Dim strName As String = txtName.Text
        Dim strIncome As String = LCase(txtIncome.Text)
        Dim strExpense As String = LCase(txtExpense.Text)
        Dim strIncomeExpense As String

        ' Check if the user has selected a category
        If cbxCategory.SelectedIndex <= 0 Then
            ' Display a message box if the user has not selected a category
            MessageBox.Show("Please select a category!", "Add Income/Expense")
            Return
        End If

        ' Check if the user has entered a name
        If String.IsNullOrWhiteSpace(strName) Then
            ' Display a message box if the user has not entered a name
            MessageBox.Show("Please enter the name of the income/expense!", "Add Income/Expense")
            Return
        End If

        ' Check if the user has entered an income or expense
        Select Case True
            Case String.IsNullOrWhiteSpace(strIncome) And String.IsNullOrWhiteSpace(strExpense)
                ' Display a message box if the user has not entered an income or expense
                MessageBox.Show("Please enter an income or expense!", "Add Income/Expense")
                Return
            Case Not String.IsNullOrWhiteSpace(strIncome) And Not String.IsNullOrWhiteSpace(strExpense)
                ' Display a message box if the user has entered both an income and an expense
                MessageBox.Show("Please enter either an income or an expense!", "Add Income/Expense")
                Return
            Case String.IsNullOrWhiteSpace(strExpense) And Not String.IsNullOrWhiteSpace(strIncome)
                ' Set the amount entered by the user to the income/expense variable if the user has entered an income
                strIncomeExpense = strIncome
            Case String.IsNullOrWhiteSpace(strIncome) And Not String.IsNullOrWhiteSpace(strExpense)
                ' Set the amount entered by the user to the income/expense variable if the user has entered an expense
                strIncomeExpense = "-" & strExpense
            Case Else
                ' Display a message box if an error has occurred
                MessageBox.Show("An error has occurred!", "Add Income/Expense")
                Return
        End Select

        ' Check if the amount entered by the user is numeric
        If Not IsNumeric(strIncomeExpense) Then
            ' Display a message box if the user has not entered a numeric amount
            MessageBox.Show("Please enter a numeric amount!", "Add Income/Expense")
            Return
        End If

        ' Declare the integer variables to store the maximum and new IDs to be used in determining the ID of the new record
        Dim intMaxId As Integer = 0
        Dim intNewId As Integer

        ' Get the record nodes and the budgetLimit node
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")
        Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//budgetLimit")

        ' Get the limit from the limit node and declare the boolean variable to store whether the budget limit has been exceeded
        Dim dblBudgetLimit As Double = budgetLimitNode.SelectSingleNode("limit").InnerText
        Dim blnExceededLimit As Boolean = False
        Dim dblTotalExpenses As Double

        ' Load the financial report for all time periods and set it equal to the total expenses
        dblTotalExpenses = LoadFinancialReport("all")

        ' Check if the total expenses are greater than the budget limit and the budget limit is not 0
        If dblTotalExpenses > dblBudgetLimit And dblBudgetLimit <> 0 Then
            ' Set the boolean variable to true if the budget limit has been exceeded
            blnExceededLimit = True
        End If

        ' Loop through the record nodes
        For Each selectedNode As XmlNode In recordNodes
            ' Get the id node of the current record node
            Dim selectedIdNode As XmlNode = selectedNode.SelectSingleNode("id")

            ' Check if the value of the id node is greater than the maximum ID
            If Val(selectedIdNode.InnerText) > intMaxId Then
                Try
                    ' Set the maximum ID to the value of the id node
                    intMaxId = Val(selectedNode.SelectSingleNode("id").InnerText)
                Catch ex As System.NullReferenceException
                    ' Display a message box if the id node was not found in the selected record's XML record
                    MessageBox.Show("The id node was not found in the selected record's XML record. Please delete your XML file or manually add in an id node.", "Add Income/Expense")
                End Try
            End If
        Next

        ' Calculate the new ID by adding 1 to the maximum ID
        intNewId = intMaxId + 1

        ' Create the record node and the id node
        Dim recordNode As XmlNode = xmlDoc.CreateElement("record")
        Dim idNode As XmlNode = xmlDoc.CreateElement("id")

        ' Set the inner text of the id node to the new ID and append it to the record node
        idNode.InnerText = intNewId
        recordNode.AppendChild(idNode)

        ' Create the createdAt node and set the inner text of the createdAt node to the current date and time, then append it to the record node
        Dim createdAtNode As XmlNode = xmlDoc.CreateElement("createdAt")
        createdAtNode.InnerText = Date.Now()
        recordNode.AppendChild(createdAtNode)

        ' Create the updatedAt node and set the inner text of the updatedAt node to the current date and time, then append it to the record node
        Dim updatedAtNode As XmlNode = xmlDoc.CreateElement("updatedAt")
        updatedAtNode.InnerText = Date.Now()
        recordNode.AppendChild(updatedAtNode)

        ' Create the category node and set the inner text of the category node to the selected category, then append it to the record node
        Dim categoryNode As XmlNode = xmlDoc.CreateElement("category")
        categoryNode.InnerText = cbxCategory.SelectedItem
        recordNode.AppendChild(categoryNode)

        ' Create the name node and set the inner text of the name node to the name entered by the user, then append it to the record node
        Dim nameNode As XmlNode = xmlDoc.CreateElement("name")
        nameNode.InnerText = strName
        recordNode.AppendChild(nameNode)

        ' Create the amount node and set the inner text of the amount node to the amount entered by the user, then append it to the record node
        Dim amountNode As XmlNode = xmlDoc.CreateElement("amount")
        amountNode.InnerText = Val(strIncomeExpense)
        recordNode.AppendChild(amountNode)

        ' Get the root node and append the record node to it
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("records")
        rootNode.AppendChild(recordNode)

        ' Save the XML file
        xmlDoc.Save(xmlFilePath)

        ' Load the financial report for all time periods and set it equal to the total expenses
        dblTotalExpenses = LoadFinancialReport("all")

        ' Check if the budget limit has been exceeded and the total expenses are greater than the budget limit and the budget limit is not 0
        If Not blnExceededLimit And dblTotalExpenses > dblBudgetLimit And dblBudgetLimit <> 0 Then
            ' Display a message box if the budget limit has been exceeded
            MessageBox.Show("Budget limit exceeded!", "Add Income/Expense")
        ElseIf blnExceededLimit And dblTotalExpenses <= dblBudgetLimit And dblBudgetLimit <> 0 Then
            ' Display a message box if the budget limit is no longer exceeded
            MessageBox.Show("Budget limit no longer exceeded!", "Add Income/Expense")
        End If

        ' Display a message box to confirm that the income/expense has been added
        MessageBox.Show("Income/Expense added successfully!", "Add Income/Expense")
    End Sub

    ' The btnEditIncomeExpense_Click event handler is executed when the Edit Income/Expense button is clicked
    Private Sub btnEditIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnEditIncomeExpense.Click
        ' Check if the user is editing a record
        If blnEditing = False Then
            ' Check if the user has selected a record
            If String.IsNullOrWhiteSpace(txtSelectedRecord.Text) Then
                ' Display a message box if the user has not selected a record
                MessageBox.Show("Please select a record!", "Edit Income/Expense")
                Return
            End If

            ' Check if the record selected by the user is numeric
            If Not IsNumeric(txtSelectedRecord.Text) Then
                ' Display a message box if the user has not selected a numeric record
                MessageBox.Show("Please select a numeric record!", "Edit Income/Expense")
                Return
            End If

            ' Get the Record nodes
            Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")

            For Each recordNode As XmlNode In recordNodes
                ' Get the id, category, name and amount nodes of the current Record node
                Dim idNode As XmlNode = recordNode.SelectSingleNode("id")
                Dim categoryNode As XmlNode = recordNode.SelectSingleNode("category")
                Dim nameNode As XmlNode = recordNode.SelectSingleNode("name")
                Dim amountNode As XmlNode = recordNode.SelectSingleNode("amount")

                ' Check if the inner text of the id node is equal to the selected record
                If idNode.InnerText = Val(txtSelectedRecord.Text) Then
                    Try
                        ' Display an input box to confirm the editing of the record
                        Dim strConfirm As String = LCase(InputBox("Are you sure you want to edit Record " & txtSelectedRecord.Text & "? (y or n):", "Edit Income/Expense", "n"))

                        If strConfirm <> "y" Then
                            ' Display a message box if the user has not confirmed the editing of the record
                            MessageBox.Show("Operation cancelled.", "Edit Income/Expense")
                            Return
                        ElseIf strConfirm = "y" Then
                            ' Set the boolean variable to true if the user is editing a record
                            blnEditing = True

                            ' Set the text of the title label to Editing Record and the selected record
                            lblTitle.Text = "Editing Record " & txtSelectedRecord.Text

                            ' Set the text of the category combo box, name text box and amount text boxes to the category, name and amount of the selected record
                            cbxCategory.SelectedItem = categoryNode.InnerText
                            txtName.Text = nameNode.InnerText

                            If amountNode.InnerText.StartsWith("-") Then
                                txtExpense.Text = amountNode.InnerText.Substring(1)
                            Else
                                txtIncome.Text = amountNode.InnerText
                            End If

                            ' Disable the text boxes and buttons
                            txtSelectedRecord.Enabled = False
                            btnAddCategory.Enabled = False
                            btnEditCategory.Enabled = False
                            btnDeleteCategory.Enabled = False
                            btnAddIncomeExpense.Enabled = False
                            btnDeleteIncomeExpense.Enabled = False
                            btnSetBudgetLimit.Enabled = False
                            btnSortRecords.Enabled = False
                            btnSearchRecords.Enabled = False
                            btnClear.Enabled = False

                            ' Set the focus to the category combo box
                            cbxCategory.Focus()
                            Return
                        End If
                    Catch ex As System.NullReferenceException
                        ' Display a message box if the id, category, name or amount node was not found in the selected record's XML record
                        MessageBox.Show("The id, category, name or amount node was not found in the selected record's XML record. Please delete your XML file or manually add in an id, category, name or amount node.", "Edit Income/Expense")
                    End Try
                End If
            Next
        End If

        ' Check if the user is editing a record
        If blnEditing = True Then
            ' Get the name, income and expense entered by the user
            Dim strName As String = txtName.Text
            Dim strIncome As String = LCase(txtIncome.Text)
            Dim strExpense As String = LCase(txtExpense.Text)
            Dim strIncomeExpense As String

            ' Check if the user has entered a name
            If String.IsNullOrWhiteSpace(strName) Then
                ' Display a message box if the user has not entered a name
                MessageBox.Show("Please enter the name of the income/expense!", "Edit Income/Expense")
                Return
            End If

            ' Check if the user has entered an income or expense
            Select Case True
                Case String.IsNullOrWhiteSpace(strIncome) And String.IsNullOrWhiteSpace(strExpense)
                    ' Display a message box if the user has not entered an income or expense
                    MessageBox.Show("Please enter an income or expense!", "Edit Income/Expense")
                    Return
                Case Not String.IsNullOrWhiteSpace(strIncome) And Not String.IsNullOrWhiteSpace(strExpense)
                    ' Display a message box if the user has entered both an income and an expense
                    MessageBox.Show("Please enter either an income or an expense!", "Edit Income/Expense")
                    Return
                Case String.IsNullOrWhiteSpace(strExpense) And Not String.IsNullOrWhiteSpace(strIncome)
                    ' Set the amount entered by the user to the income/expense variable if the user has entered an income
                    strIncomeExpense = strIncome
                Case String.IsNullOrWhiteSpace(strIncome) And Not String.IsNullOrWhiteSpace(strExpense)
                    ' Set the amount entered by the user to the income/expense variable if the user has entered an expense
                    strIncomeExpense = "-" & strExpense
                Case Else
                    ' Display a message box if an error has occurred
                    MessageBox.Show("An error has occurred!", "Edit Income/Expense")
                    Return
            End Select

            ' Check if the amount entered by the user is numeric
            If Not IsNumeric(strIncomeExpense) Then
                ' Display a message box if the user has not entered a numeric amount
                MessageBox.Show("Please enter a numeric amount!", "Edit Income/Expense")
                Return
            End If

            ' Get the record nodes and the budgetLimit node
            Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")
            Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//budgetLimit")

            ' Get the limit from the limit node and declare the boolean variable to store whether the budget limit has been exceeded
            Dim dblBudgetLimit As Double = budgetLimitNode.SelectSingleNode("limit").InnerText
            Dim blnExceededLimit As Boolean = False
            Dim dblTotalExpenses As Double

            ' Load the financial report for all time periods and set it equal to the total expenses
            dblTotalExpenses = LoadFinancialReport("all")

            ' Check if the total expenses are greater than the budget limit and the budget limit is not 0
            If dblTotalExpenses > dblBudgetLimit And dblBudgetLimit <> 0 Then
                ' Set the boolean variable to true if the budget limit has been exceeded
                blnExceededLimit = True
            End If

            ' Loop through the record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Get the updatedAt, id, category, name and amount nodes of the current record node
                Dim idNode As XmlNode = recordNode.SelectSingleNode("id")
                Dim updatedAtNode As XmlNode = recordNode.SelectSingleNode("updatedAt")
                Dim categoryNode As XmlNode = recordNode.SelectSingleNode("category")
                Dim nameNode As XmlNode = recordNode.SelectSingleNode("name")
                Dim amountNode As XmlNode = recordNode.SelectSingleNode("amount")

                ' Check if the inner text of the id node is equal to the selected record
                If idNode.InnerText = Val(txtSelectedRecord.Text) Then
                    Try
                        ' Remove the updatedAt node
                        recordNode.RemoveChild(recordNode.SelectSingleNode("updatedAt"))

                        ' Create a new updatedAt node and set the inner text of the updatedAt node to the current date and time and append it to the record node
                        Dim newUpdatedAtNode As XmlNode = xmlDoc.CreateElement("updatedAt")
                        newUpdatedAtNode.InnerText = Date.Now()
                        recordNode.AppendChild(newUpdatedAtNode)

                        ' Remove the category node
                        recordNode.RemoveChild(recordNode.SelectSingleNode("category"))

                        ' Create a new category node and set the inner text of the category node to the category selected by the user and append it to the record node
                        Dim newCategoryNode As XmlNode = xmlDoc.CreateElement("category")
                        newCategoryNode.InnerText = cbxCategory.SelectedItem
                        recordNode.AppendChild(newCategoryNode)

                        ' Remove the name node
                        recordNode.RemoveChild(recordNode.SelectSingleNode("name"))

                        ' Create a new name node and set the inner text of the name node to the name entered by the user and append it to the record node
                        Dim newNameNode As XmlNode = xmlDoc.CreateElement("name")
                        newNameNode.InnerText = txtName.Text
                        recordNode.AppendChild(newNameNode)

                        ' Remove the amount node
                        recordNode.RemoveChild(recordNode.SelectSingleNode("amount"))

                        ' Create a new amount node and set the inner text of the amount node to the amount entered by the user and append it to the record node
                        Dim newAmountNode As XmlNode = xmlDoc.CreateElement("amount")
                        newAmountNode.InnerText = strIncomeExpense
                        recordNode.AppendChild(newAmountNode)

                        ' Save the XML file
                        xmlDoc.Save(xmlFilePath)

                        ' Load the financial report for all time periods and set it equal to the total expenses
                        dblTotalExpenses = LoadFinancialReport("all")

                        ' Check if the budget limit has been exceeded and the total expenses are greater than the budget limit and the budget limit is not 0
                        If Not blnExceededLimit And dblTotalExpenses > dblBudgetLimit And dblBudgetLimit <> 0 Then
                            ' Display a message box if the budget limit has been exceeded
                            MessageBox.Show("Budget limit exceeded!", "Edit Income/Expense")
                        ElseIf blnExceededLimit And dblTotalExpenses <= dblBudgetLimit And dblBudgetLimit <> 0 Then
                            ' Display a message box if the budget limit is no longer exceeded
                            MessageBox.Show("Budget limit no longer exceeded!", "Edit Income/Expense")
                        End If

                        ' Set the boolean variable to true if the user is editing a record
                        blnEditing = False

                        ' Set the text of the title label to Editing Record and the selected record
                        lblTitle.Text = "INCOME EXPENSES MANAGER"

                        ' Enable the text boxes and buttons
                        txtSelectedRecord.Enabled = True
                        btnAddCategory.Enabled = True
                        btnEditCategory.Enabled = True
                        btnDeleteCategory.Enabled = True
                        btnAddIncomeExpense.Enabled = True
                        btnDeleteIncomeExpense.Enabled = True
                        btnSetBudgetLimit.Enabled = True
                        btnSortRecords.Enabled = True
                        btnSearchRecords.Enabled = True
                        btnClear.Enabled = True

                        ' Display a message box to confirm that the income/expense has been edited
                        MessageBox.Show("Income/Expense edited successfully!", "Edit Income/Expense")
                    Catch ex As System.NullReferenceException
                        ' Display a message box if the id, category, name, amount or updatedAt node was not found in the selected record's XML record
                        MessageBox.Show("The id, category, name, amount or updatedAt node was not found in the selected record's XML record. Please delete your XML file or manually add in an id, category, name, amount or updatedAt node.", "Edit Income/Expense")
                    End Try
                End If
            Next
        End If
    End Sub

    ' The btnDeleteIncomeExpense_Click event handler is executed when the Delete Income/Expense button is clicked
    Private Sub btnDeleteIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnDeleteIncomeExpense.Click
        ' Get the selected record
        Dim strSelectedRecord As String = LCase(txtSelectedRecord.Text)

        ' Check if the user has selected a record
        If String.IsNullOrWhiteSpace(strSelectedRecord) Then
            ' Display a message box if the user has not selected a record
            MessageBox.Show("Please select a record!", "Delete Income/Expense")
            Return
        End If

        ' Display an input box to confirm the deletion of the record
        Dim strConfirm As String = LCase(InputBox("Are you sure you want to delete Record " & strSelectedRecord & "? (y or n):", "Delete Income/Expense", "n"))

        ' Check if the user has confirmed the deletion
        If strConfirm <> "y" Then
            ' Display a message box if the user has not confirmed the deletion
            MessageBox.Show("Operation cancelled.", "Delete Income/Expense")
            Return
        ElseIf strConfirm = "y" Then
            ' Get the record nodes, the budgetLimit node and the root node
            Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")
            Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//budgetLimit")
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("records")

            ' Get the limit from the limit node and declare the boolean variable to store whether the budget limit has been exceeded
            Dim dblBudgetLimit As Double = budgetLimitNode.SelectSingleNode("limit").InnerText
            Dim blnExceededLimit As Boolean = False
            Dim dblTotalExpenses As Double

            ' Load the financial report for all time periods and set it equal to the total expenses
            dblTotalExpenses = LoadFinancialReport("all")

            ' Check if the total expenses are greater than the budget limit and the budget limit is not 0
            If dblTotalExpenses > dblBudgetLimit And dblBudgetLimit <> 0 Then
                ' Set the boolean variable to true if the budget limit has been exceeded
                blnExceededLimit = True
            End If

            ' Loop through the record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Get the id node of the current Record node
                Dim idNode As XmlNode = recordNode.SelectSingleNode("id")

                ' Check if the inner text of the id node is equal to the selected record
                If idNode.InnerText = Val(strSelectedRecord) Then
                    Try
                        ' Remove the Record node
                        rootNode.RemoveChild(recordNode)

                        ' Save the XML file
                        xmlDoc.Save(xmlFilePath)

                        ' Load the financial report for all time periods and set it equal to the total expenses
                        dblTotalExpenses = LoadFinancialReport("all")

                        ' Check if the budget limit has been exceeded and the total expenses are greater than the budget limit and the budget limit is not 0
                        If Not blnExceededLimit And dblTotalExpenses > dblBudgetLimit And dblBudgetLimit <> 0 Then
                            ' Display a message box if the budget limit has been exceeded
                            MessageBox.Show("Budget limit exceeded!", "Delete Income/Expense")
                        ElseIf blnExceededLimit And dblTotalExpenses <= dblBudgetLimit And dblBudgetLimit <> 0 Then
                            ' Display a message box if the budget limit is no longer exceeded
                            MessageBox.Show("Budget limit no longer exceeded!", "Delete Income/Expense")
                        End If

                        ' Display a message box to confirm that the income/expense has been deleted
                        MessageBox.Show("Income/Expense deleted successfully!", "Delete Income/Expense")
                    Catch ex As System.NullReferenceException
                        ' Display a message box if the id node was not found in the selected record's XML record
                        MessageBox.Show("The id node was not found in the selected record's XML record. Please delete your XML file or manually add in an id node.", "Delete Income/Expense")
                    End Try
                End If
            Next
        End If
    End Sub

    ' The btnSetBudgetLimit_Click event handler is executed when the Set Budget Limit button is clicked
    Private Sub btnSetBudgetLimit_Click(sender As Object, e As EventArgs) Handles btnSetBudgetLimit.Click
        ' Display an input box to get the budget limit
        Dim strBudgetLimit As String = LCase(InputBox("Enter the budget limit (or type 0 to remove):", "Set Budget Limit", "0"))

        ' Check if the user has entered a budget limit
        If String.IsNullOrWhiteSpace(strBudgetLimit) Then
            ' Display a message box if the user has not entered a budget limit
            MessageBox.Show("Operation cancelled.", "Set Budget Limit")
            Return
        End If

        ' Check if the budget limit entered by the user is numeric
        If Not IsNumeric(strBudgetLimit) Then
            ' Display a message box if the user has not entered a numeric budget limit
            MessageBox.Show("Please enter a numeric budget limit!", "Set Budget Limit")
            Return
        End If

        ' Check if the budget limit entered by the user is negative
        If Val(strBudgetLimit) < 0 Then
            ' Display a message box if the user has entered a negative budget limit
            MessageBox.Show("Please enter a positive budget limit!", "Set Budget Limit")
            Return
        End If

        ' Get the root node and the budgetLimit node
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("records")
        Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//budgetLimit")

        ' Check if the budgetLimit node exists
        Try
            ' Remove the limit node
            budgetLimitNode.RemoveChild(budgetLimitNode.SelectSingleNode("limit"))

            ' Create a new limit node and set the inner text of the limit node to the budget limit entered by the user and append it to the BudgetLimit node
            Dim limitNode As XmlNode = xmlDoc.CreateElement("limit")
            limitNode.InnerText = strBudgetLimit
            budgetLimitNode.AppendChild(limitNode)

            ' Save the XML file
            xmlDoc.Save(xmlFilePath)

            ' Display a message box to confirm that the budget limit has been set
            MessageBox.Show("Budget limit set successfully!", "Set Budget Limit")
        Catch ex As System.NullReferenceException
            ' Create a new budgetLimit node and set the inner text of the limit node to the budget limit entered by the user and append it to the root node
            Dim newBudgetLimitNode As XmlNode = xmlDoc.CreateElement("budgetLimit")
            Dim newLimitNode As XmlNode = xmlDoc.CreateElement("limit")

            newLimitNode.InnerText = strBudgetLimit
            newBudgetLimitNode.AppendChild(newLimitNode)
            rootNode.AppendChild(newBudgetLimitNode)

            ' Save the XML file
            xmlDoc.Save(xmlFilePath)

            ' Display a message box to confirm that the budget limit has been set
            MessageBox.Show("Budget limit set successfully!", "Set Budget Limit")
        End Try

        ' Load the financial report for all time periods
        LoadFinancialReport("all")
    End Sub

    ' The btnSortRecords_Click event handler is executed when the Sort Records button is clicked
    Private Sub btnSortRecords_Click(sender As Object, e As EventArgs) Handles btnSortRecords.Click
        ' Load the records
        LoadRecords()

        ' Declare the integer array to store the amounts of the displayed records and the current index
        Dim intNbrsArray(intDisplayedRecords - 1) As Integer
        Dim intCurrentIndex As Integer = 0

        ' Display an input box to get the sort option for the records
        Dim strSortOption As String = LCase(InputBox("Enter the sort option for the records (id or amount):", "Sort Records", "amount"))

        ' Check if the user has entered a sort option
        If String.IsNullOrWhiteSpace(strSortOption) Then
            ' Display a message box if the user has not entered a sort option
            MessageBox.Show("Operation cancelled.", "Sort Records")
            Return
        End If

        ' Check if the sort option entered by the user is not id or amount
        If strSortOption <> "id" And strSortOption <> "amount" Then
            ' Display a message box if the user has entered an invalid sort option
            MessageBox.Show("Please enter a valid sort option (id or amount)!", "Sort Records")
            Return
        End If

        ' Display an input box to get the sort order for the records
        Dim strSortOrder As String = LCase(InputBox("Sort the records in descending order (y or n):", "Sort Records", "n"))

        ' Check if the user has entered a sort order
        If String.IsNullOrWhiteSpace(strSortOrder) Then
            ' Display a message box if the user has not entered a sort order
            MessageBox.Show("Operation cancelled.", "Sort Records")
            Return
        End If

        ' Get the record nodes
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")

        ' Loop through the record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Get the selected node based on the sort option entered by the user
            Dim selectedNode As XmlNode = recordNode.SelectSingleNode(strSortOption)

            ' Set each index of the integer array to the value of the selected node and increment the current index
            intNbrsArray(intCurrentIndex) = (Val(selectedNode.InnerText))
            intCurrentIndex = intCurrentIndex + 1
        Next

        ' Declare the integer variables to store the start and end index of the integer array
        Dim intLowerBound As Integer = LBound(intNbrsArray)
        Dim intUpperBound As Integer = UBound(intNbrsArray)

        ' Declare the integer variable to store the main loop counter for array processing
        Dim intMainLoopCounter As Integer = 0
        ' Declare the integer variable to store the index counter value to shuffle along 1 index step
        Dim intIndexCounter As Integer = 0

        ' Declare the integer variable to store the index to move the minimum value to
        Dim intMinValueIndex As Integer
        ' Declare the integer variable to store the index value that the minimum value is found at
        Dim intSwapValueIndex As Integer
        ' Declare the integer variable to store the actual value of the found minimum
        Dim intSwapValue As Integer

        ' Loop through the integer array
        For intMainLoopCounter = intLowerBound To intUpperBound
            ' Set the minimum value index to the current index
            intMinValueIndex = intMainLoopCounter
            ' Reset the saved swap values from the last pass of the array
            intSwapValue = 0
            intSwapValueIndex = 0

            ' Loop through the integer array starting from the next index
            For intIndexCounter = intMainLoopCounter + 1 To intUpperBound
                ' Search and find the lowest value remaining in the array for this sort pass of the array
                If intNbrsArray(intIndexCounter) < intNbrsArray(intMinValueIndex) Then
                    ' First check for first time finding a smaller value, then check if smallest so far
                    If intSwapValue = 0 Then
                        intSwapValue = intNbrsArray(intIndexCounter)
                        intSwapValueIndex = intIndexCounter
                    Else
                        If intNbrsArray(intIndexCounter) < intSwapValue Then
                            intSwapValue = intNbrsArray(intIndexCounter)
                            intSwapValueIndex = intIndexCounter
                        End If
                    End If
                End If
            Next

            If intSwapValue <> 0 Then
                ' Swap the value at the minimum index to where the found swap value is and move the saved found swap value to the current minimum index
                intNbrsArray(intSwapValueIndex) = intNbrsArray(intMinValueIndex)
                intNbrsArray(intMinValueIndex) = intSwapValue
            End If
        Next

        ' Check if the sort order is descending
        If strSortOrder = "y" Then
            ' Reverse the integer array if the sort order is descending
            Array.Reverse(intNbrsArray)
        End If

        ' Clear the list box
        lstDisplay.Items.Clear()

        ' Loop through the integer array
        For intArrayIndex = 0 To intNbrsArray.Length - 1
            ' Loop through the record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Check if the selected node based on the sort option is equal to the current index of the integer array
                If recordNode.SelectSingleNode(strSortOption).InnerText = intNbrsArray(intArrayIndex) Then
                    ' Add the record to the list box
                    lstDisplay.Items.Add(recordNode.SelectSingleNode("id").InnerText & ": " & recordNode.SelectSingleNode("name").InnerText & ": " & FormatCurrency(recordNode.SelectSingleNode("amount").InnerText))
                End If
            Next
        Next
    End Sub

    ' The btnSearchRecords_Click event handler is executed when the Search Records button is clicked
    Private Sub btnSearchRecords_Click(sender As Object, e As EventArgs) Handles btnSearchRecords.Click
        ' Load the records
        LoadRecords()

        ' Declare the integer array to store the amounts of the displayed records and the current index
        Dim intNbrsArray(intDisplayedRecords - 1) As Integer
        Dim intCurrentIndex As Integer = 0

        ' Display an input box to get the search request type
        Dim strSearchRequestType As String = LCase(InputBox("Enter the search request type (id or amount):", "Search Records", "amount"))

        ' Check if the user has entered a search request type
        If String.IsNullOrWhiteSpace(strSearchRequestType) Then
            ' Display a message box if the user has not entered a search request type
            MessageBox.Show("Operation cancelled.", "Search Records")
            Return
        End If

        ' Check if the search request type entered by the user is not id or amount
        If strSearchRequestType <> "id" And strSearchRequestType <> "amount" Then
            ' Display a message box if the user has entered an invalid search request type
            MessageBox.Show("Please enter a valid search request type (id or amount)!", "Search Records")
            Return
        End If

        ' Display an input box to get the search request item
        Dim strSearchRequestItem As String = LCase(InputBox("Enter the search request item:", "Search Records", "100"))

        ' Check if the user has entered a search request item
        If String.IsNullOrWhiteSpace(strSearchRequestItem) Then
            ' Display a message box if the user has not entered a search request item
            MessageBox.Show("Operation cancelled.", "Search Records")
            Return
        End If

        ' Check if the search request item entered by the user is numeric
        If Not IsNumeric(strSearchRequestItem) Then
            ' Display a message box if the user has not entered a numeric search request item
            MessageBox.Show("Please enter a numeric search request item!", "Search Records")
            Return
        End If

        ' Get the record nodes
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")

        ' Loop through the record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Get the selected node based on the search request type
            Dim selectedNode As XmlNode = recordNode.SelectSingleNode(strSearchRequestType)

            ' Set each index of the integer array to the value of the selected node and increment the current index
            intNbrsArray(intCurrentIndex) = Val(selectedNode.InnerText)
            intCurrentIndex = intCurrentIndex + 1
        Next

        ' Declare the boolean variables to store whether the item has been found and whether the end of the array has been reached
        Dim blnItemFound As Boolean = False
        Dim blnEndOfArray As Boolean = False
        Dim intSubscript As Integer = 0
        Dim intMaxSubscript As Integer = 0

        ' Determine the end of the array
        intMaxSubscript = intNbrsArray.Length - 1

        ' Use a while loop to do a linear search of the array
        While blnItemFound = False And blnEndOfArray = False
            ' Check if the search request item is equal to the current index of the integer array
            If strSearchRequestItem = intNbrsArray(intSubscript) Then
                ' Set the boolean variable to true if the item has been found
                blnItemFound = True
            ElseIf intSubscript = intMaxSubscript Then
                ' Set the boolean variable to true if the end of the array has been reached
                blnEndOfArray = True
            Else
                ' Increment the current index
                intSubscript = intSubscript + 1
            End If
        End While

        ' Check if the item has been found
        If blnItemFound = True Then
            ' Clear the list box
            lstDisplay.Items.Clear()

            ' Loop through the record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Check if the selected node based on the search request type is equal to the search request item
                If recordNode.SelectSingleNode(strSearchRequestType).InnerText = strSearchRequestItem Then
                    ' Add the record to the list box
                    lstDisplay.Items.Add(recordNode.SelectSingleNode("id").InnerText & ": " & recordNode.SelectSingleNode("name").InnerText & ": " & FormatCurrency(recordNode.SelectSingleNode("amount").InnerText))
                End If
            Next
        Else
            ' Display a message box if the item has not been found
            MessageBox.Show("Item not found!", "Search Records")
        End If
    End Sub

    ' The btnViewRecords_Click event handler is executed when the View Records button is clicked
    Private Sub btnViewRecords_Click(sender As Object, e As EventArgs) Handles btnViewRecords.Click
        ' Load the records
        LoadRecords()
    End Sub

    ' The btnViewReport_Click event handler is executed when the View Report button is clicked
    Private Sub btnViewReport_Click(sender As Object, e As EventArgs) Handles btnViewReport.Click
        ' Display an input box to get the year for the financial report
        Dim strTimePeriod As String = LCase(InputBox("Enter the year for the financial report (e.g. All, 2024):", "View Report", "All"))

        ' Check if the user has entered a year
        If String.IsNullOrWhiteSpace(strTimePeriod) Then
            ' Display a message box if the user has not entered a year
            MessageBox.Show("Operation cancelled.", "View Report")
            Return
        End If

        ' Check if the year entered by the user is numeric
        If Not IsNumeric(strTimePeriod) And strTimePeriod <> "all" Then
            ' Display a message box if the user has not entered a numeric year
            MessageBox.Show("Please enter a numeric time period or select all time periods!", "View Report")
            Return
        End If

        ' Check if the year entered by the user is negative
        If strTimePeriod <> "all" And Val(strTimePeriod) < 0 Then
            ' Display a message box if the user has entered a negative year
            MessageBox.Show("Please enter a positive numeric time period!", "View Report")
            Return
        End If

        ' Load the financial report for the year entered by the user
        LoadFinancialReport(strTimePeriod)
    End Sub

    ' The btnClear_Click event handler is executed when the Clear button is clicked
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clear the combo box, text boxes and list box
        cbxCategory.SelectedIndex = 0
        txtName.Clear()
        txtIncome.Clear()
        txtExpense.Clear()
        txtSelectedRecord.Clear()
        lstDisplay.Items.Clear()

        ' Set the focus to the category combo box
        cbxCategory.Focus()
    End Sub

    ' The subroutine PopulateComboBox is used to populate the combo box with the categories from the XML file
    Private Sub PopulateComboBox()
        ' Get the Category nodes
        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("records/category")

        ' Clear the combo box and add the "All" item
        cbxCategory.Items.Clear()
        cbxCategory.Items.Add("All")

        ' Loop through the category nodes
        For Each categoryNode As XmlNode In categoryNodes
            ' Add the name node inner text to the combo box
            cbxCategory.Items.Add(categoryNode.SelectSingleNode("name").InnerText)
        Next

        ' Set the selected index of the combo box to 0
        cbxCategory.SelectedIndex = 0
    End Sub

    ' The subroutine LoadRecords is used to load the records from the XML file based on the selected category
    Private Sub LoadRecords()
        ' Get the record nodes
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")

        ' Set the displayed records counter to 0
        intDisplayedRecords = 0

        ' Clear the list box and add the selected category to the list box
        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Category: " & cbxCategory.SelectedItem)

        ' Loop through the Record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Check if the selected category is "All" or the category node inner text is equal to the selected category
            If cbxCategory.SelectedItem = "All" Or recordNode.SelectSingleNode("category").InnerText = cbxCategory.SelectedItem Then
                ' Add the record to the list box and increment the displayed records counter
                lstDisplay.Items.Add(recordNode.SelectSingleNode("id").InnerText & ": " & recordNode.SelectSingleNode("name").InnerText & ": " & FormatCurrency(recordNode.SelectSingleNode("amount").InnerText))
                intDisplayedRecords = intDisplayedRecords + 1
            End If
        Next
    End Sub

    ' The function LoadFinancialReport is used to load the financial report from the XML file based on the selected time period
    ' strTimePeriod Parameter: The time period for the financial report
    ' Returns the net income for the selected time period
    Private Function LoadFinancialReport(ByVal strTimePeriod As String) As Double
        ' Declare the double variables to store the total income, total expenses, net income and budget limit
        Dim dblTotalIncome As Double
        Dim dblTotalExpenses As Double
        Dim dblNetIncome As Double

        ' Get the record nodes and the budgetLimit node
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//record")
        Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//budgetLimit")

        ' Get the limit from the Limit node
        Dim dblBudgetLimit As Double = budgetLimitNode.SelectSingleNode("limit").InnerText

        ' Clear the list box
        lstDisplay.Items.Clear()

        ' Loop through the record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Get the updatedAt node of the current record node
            Dim updatedAtNode As XmlNode = recordNode.SelectSingleNode("updatedAt")

            ' Check if the inner text of the updatedAt node contains the selected time period
            If updatedAtNode.InnerText.Contains(strTimePeriod) OrElse strTimePeriod = "all" Then
                Try
                    ' Check if the amount node inner text starts with a negative sign
                    If recordNode.SelectSingleNode("amount").InnerText.StartsWith("-") Then
                        ' Add the amount to the total expenses
                        dblTotalExpenses = dblTotalExpenses + Val(recordNode.SelectSingleNode("amount").InnerText)
                    Else
                        ' Add the amount to the total income
                        dblTotalIncome = dblTotalIncome + Val(recordNode.SelectSingleNode("amount").InnerText)
                    End If
                Catch ex As System.NullReferenceException
                    ' Display a message box if the id, updatedAt, category, name or amount node was not found in the selected record's XML record
                    MessageBox.Show("The id, updatedAt category, name or amount node was not found in the selected record's XML record. Please delete your XML file or manually add in an id, updatedAt, category, name or amount node.", "Load Financial Report")
                End Try
            End If
        Next

        ' Calculate the net income
        dblNetIncome = dblTotalIncome + dblTotalExpenses

        ' Add the financial report to the list box
        lstDisplay.Items.Add("Total Income: " & FormatCurrency(dblTotalIncome))
        lstDisplay.Items.Add("Total Expenses: " & FormatCurrency(Math.Abs(dblTotalExpenses)))
        lstDisplay.Items.Add("Net Income: " & FormatCurrency(dblNetIncome))

        If dblBudgetLimit = 0 Then
            lstDisplay.Items.Add("Budget Limit: None")
            lstDisplay.Items.Add("Remaining Budget: No Budget")
        ElseIf dblBudgetLimit <> 0 Then
            lstDisplay.Items.Add("Budget Limit: " & FormatCurrency(dblBudgetLimit))
            lstDisplay.Items.Add("Remaining Budget: " & FormatCurrency((dblBudgetLimit + dblTotalExpenses)))
        End If

        lstDisplay.Items.Add("Time Period: " & strTimePeriod)

        ' Return the total expenses
        Return Math.Abs(dblTotalExpenses)
    End Function
End Class
