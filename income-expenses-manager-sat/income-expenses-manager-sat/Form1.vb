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
            Dim rootNode As XmlNode = xmlDoc.CreateElement("Records")
            xmlDoc.AppendChild(rootNode)

            ' Create the BudgetLimit node and the Limit node
            Dim budgetLimitNode As XmlNode = xmlDoc.CreateElement("BudgetLimit")
            Dim limitNode As XmlNode = xmlDoc.CreateElement("Limit")

            ' Set the inner text of the Limit node to 0 and append it to the BudgetLimit node, then append the BudgetLimit node to the root node
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
        ' Display an input box to get the name of the category
        Dim strName As String = InputBox("Enter the name of the category:", "Add Category", "Category")

        ' Check if the user has entered a name
        If String.IsNullOrWhiteSpace(strName) Then
            ' Display a message box if the user has not entered a name
            MessageBox.Show("Operation cancelled.", "Add Category")
            Return
        End If

        ' Create the Category node and the Name node
        Dim categoryNode As XmlNode = xmlDoc.CreateElement("Category")
        Dim nameNode As XmlNode = xmlDoc.CreateElement("Name")

        ' Set the inner text of the Name node to the name entered by the user and append it to the Category node
        nameNode.InnerText = strName
        categoryNode.AppendChild(nameNode)

        ' Get the root node and append the Category node to it
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
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
        If cbxCategory.SelectedIndex = -1 Or cbxCategory.SelectedIndex = 0 Then
            ' Display a message box if the user has not selected a category
            MessageBox.Show("Please select a category!", "Edit Category")
            Return
        End If

        ' Display an input box to get the new name of the category
        Dim strName As String = InputBox("Enter the name of the category:", "Edit Category", "Category")

        ' Check if the user has entered a name
        If String.IsNullOrWhiteSpace(strName) Then
            ' Display a message box if the user has not entered a name
            MessageBox.Show("Operation cancelled.", "Edit Category")
            Return
        End If

        ' Get the Category nodes and loop through them
        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/Category")

        For Each categoryNode As XmlNode In categoryNodes
            ' Get the Name node of the current category node
            Dim nameNode As XmlNode = categoryNode.SelectSingleNode("Name")

            ' Check if the Name node exists and its inner text is equal to the selected category
            If nameNode IsNot Nothing And nameNode.InnerText = cbxCategory.SelectedItem Then
                ' Remove the Name node
                categoryNode.RemoveChild(nameNode)

                ' Create a new Name node and set the inner text of the Name node to the name entered by the user and append it to the Category node
                Dim newNameNode As XmlNode = xmlDoc.CreateElement("Name")

                newNameNode.InnerText = strName
                categoryNode.AppendChild(newNameNode)

                ' Get the Record nodes and loop through them
                Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

                For Each recordNode As XmlNode In recordNodes
                    ' Get the Category node of the current Record node
                    Dim selectedCategoryNode As XmlNode = recordNode.SelectSingleNode("Category")

                    ' Check if the Category node exists and its inner text is equal to the selected category
                    If selectedCategoryNode IsNot Nothing And selectedCategoryNode.InnerText = cbxCategory.SelectedItem Then
                        ' Remove the Category node
                        recordNode.RemoveChild(selectedCategoryNode)

                        ' Create a new Category node and set the inner text of the Category node to the name entered by the user and append it to the Record node
                        Dim newCategoryNode As XmlNode = xmlDoc.CreateElement("Category")

                        newCategoryNode.InnerText = strName
                        recordNode.AppendChild(newCategoryNode)
                    End If
                Next

                ' Save the XML file
                xmlDoc.Save(xmlFilePath)

                ' Populate the combo box and display the financial report
                PopulateComboBox()
                LoadFinancialReport("all")

                ' Display a message box to confirm that the category has been edited
                MessageBox.Show("Category edited successfully!", "Edit Category")
            End If
        Next
    End Sub

    ' The btnDeleteCategory_Click event handler is executed when the Delete Category button is clicked
    Private Sub btnDeleteCategory_Click(sender As Object, e As EventArgs) Handles btnDeleteCategory.Click
        ' Check if the user has selected a category
        If cbxCategory.SelectedIndex = -1 Or cbxCategory.SelectedIndex = 0 Then
            ' Display a message box if the user has not selected a category
            MessageBox.Show("Please select a category!", "Delete Category")
            Return
        End If

        ' Display an input box to confirm the deletion of the category
        Dim strConfirm As String = LCase(InputBox("Are you sure you want to delete the category " & cbxCategory.SelectedItem & "? (Y or N):", "Delete Category", "N"))

        ' Check if the user has confirmed the deletion
        If strConfirm <> "y" Then
            ' Display a message box if the user has not confirmed the deletion
            MessageBox.Show("Operation cancelled.", "Delete Category")
            Return
        ElseIf strConfirm = "y" Then
            ' Get the Category nodes, the Record nodes and the root node
            Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/Category")
            Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")

            ' Loop through the Record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Get the Category node of the current Record node
                Dim selectedCategoryNode As XmlNode = recordNode.SelectSingleNode("Category")

                ' Check if the Category node exists and its inner text is equal to the selected category
                If selectedCategoryNode IsNot Nothing And selectedCategoryNode.InnerText = cbxCategory.SelectedItem Then
                    ' Remove the Record node
                    rootNode.RemoveChild(recordNode)

                    ' Loop through the Category nodes
                    For Each categoryNode As XmlNode In categoryNodes
                        ' Get the Name node of the current Category node
                        Dim nameNode As XmlNode = categoryNode.SelectSingleNode("Name")

                        ' Check if the Name node exists and its inner text is equal to the selected category
                        If nameNode IsNot Nothing And nameNode.InnerText = cbxCategory.SelectedItem Then
                            ' Remove the Category node
                            rootNode.RemoveChild(categoryNode)
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
            Next
        End If
    End Sub

    ' The btnAddIncomeExpense_Click event handler is executed when the Add Income/Expense button is clicked
    Private Sub btnAddIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnAddIncomeExpense.Click
        ' Get the selected category and the amount entered by the user
        Dim strIncomeExpense As String = LCase(txtIncomeExpense.Text)

        ' Check if the user has selected a category
        If cbxCategory.SelectedIndex = -1 Or cbxCategory.SelectedIndex = 0 Then
            ' Display a message box if the user has not selected a category
            MessageBox.Show("Please select a category!", "Add Income/Expense")
            Return
        End If

        ' Check if the user has entered an amount
        If String.IsNullOrWhiteSpace(strIncomeExpense) Then
            ' Display a message box if the user has not entered an amount
            MessageBox.Show("Please enter an amount!", "Edit Income/Expense")
            Return
        End If

        ' Declare the integer variables to store the maximum and new IDs to be used in determining the ID of the new record
        Dim intMaxId As Integer = -1
        Dim intNewId As Integer

        ' Get the Record nodes and the BudgetLimit node
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
        Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//BudgetLimit")

        ' Get the limit from the Limit node and declare the boolean variable to store whether the budget limit has been exceeded
        Dim intBudgetLimit As Integer = budgetLimitNode.SelectSingleNode("Limit").InnerText
        Dim blnExceededLimit As Boolean = False
        Dim intNetIncome As Integer

        ' Load the financial report for all time periods and set it equal to the net income
        intNetIncome = LoadFinancialReport("all")

        ' Check if the net income is greater than the budget limit and the budget limit is not 0
        If intNetIncome > intBudgetLimit And intBudgetLimit <> 0 Then
            ' Set the boolean variable to true if the budget limit has been exceeded
            blnExceededLimit = True
        End If

        ' Loop through the Record nodes
        For Each selectedNode As XmlNode In recordNodes
            ' Get the Id node of the current Record node
            Dim selectedIdNode As XmlNode = selectedNode.SelectSingleNode("Id")

            ' Check if the Id node exists and its value is greater than the maximum ID
            If selectedIdNode IsNot Nothing And Val(selectedIdNode.InnerText) > intMaxId Then
                ' Set the maximum ID to the value of the Id node
                intMaxId = Val(selectedNode.SelectSingleNode("Id").InnerText)
            End If
        Next

        ' Calculate the new ID by adding 1 to the maximum ID
        intNewId = intMaxId + 1

        ' Create the Record node and the Id node
        Dim recordNode As XmlNode = xmlDoc.CreateElement("Record")
        Dim idNode As XmlNode = xmlDoc.CreateElement("Id")

        ' Set the inner text of the Id node to the new ID and append it to the Record node
        idNode.InnerText = intNewId
        recordNode.AppendChild(idNode)

        ' Create the Category node and set the inner text of the Category node to the selected category, then append it to the Record node
        Dim categoryNode As XmlNode = xmlDoc.CreateElement("Category")
        categoryNode.InnerText = cbxCategory.SelectedItem
        recordNode.AppendChild(categoryNode)

        ' Create the Amount node and set the inner text of the Amount node to the amount entered by the user, then append it to the Record node
        Dim amountNode As XmlNode = xmlDoc.CreateElement("Amount")
        amountNode.InnerText = Val(txtIncomeExpense.Text)
        recordNode.AppendChild(amountNode)

        ' Create the CreatedAt node and set the inner text of the CreatedAt node to the current date and time, then append it to the Record node
        Dim createdAtNode As XmlNode = xmlDoc.CreateElement("CreatedAt")
        createdAtNode.InnerText = Date.Now()
        recordNode.AppendChild(createdAtNode)

        ' Create the UpdatedAt node and set the inner text of the UpdatedAt node to the current date and time, then append it to the Record node
        Dim updatedAtNode As XmlNode = xmlDoc.CreateElement("UpdatedAt")
        updatedAtNode.InnerText = Date.Now()
        recordNode.AppendChild(updatedAtNode)

        ' Get the root node and append the Record node to it
        Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
        rootNode.AppendChild(recordNode)

        ' Save the XML file
        xmlDoc.Save(xmlFilePath)

        ' Load the financial report for all time periods and set it equal to the net income
        intNetIncome = LoadFinancialReport("all")

        ' Check if the budget limit has been exceeded and the net income is greater than the budget limit and the budget limit is not 0
        If Not blnExceededLimit And intNetIncome > intBudgetLimit And intBudgetLimit <> 0 Then
            ' Display a message box if the budget limit has been exceeded
            MessageBox.Show("Budget limit exceeded!", "Add Income/Expense")
        ElseIf blnExceededLimit And intNetIncome <= intBudgetLimit And intBudgetLimit <> 0 Then
            ' Display a message box if the budget limit is no longer exceeded
            MessageBox.Show("Budget limit no longer exceeded!", "Add Income/Expense")
        End If

        ' Display a message box to confirm that the income/expense has been added
        MessageBox.Show("Income/Expense added successfully!", "Add Income/Expense")
    End Sub

    ' The btnEditIncomeExpense_Click event handler is executed when the Edit Income/Expense button is clicked
    Private Sub btnEditIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnEditIncomeExpense.Click
        ' Get the amount entered by the user and the selected record
        Dim strIncomeExpense As String = LCase(txtIncomeExpense.Text)

        ' Check if the user has entered an amount
        If String.IsNullOrWhiteSpace(strIncomeExpense) Then
            ' Display a message box if the user has not entered an amount
            MessageBox.Show("Please enter an amount!", "Edit Income/Expense")
            Return
        End If

        ' Check if the user has selected a record
        If String.IsNullOrWhiteSpace(txtSelectedRecord.Text) Then
            ' Display a message box if the user has not selected a record
            MessageBox.Show("Please select a record!", "Edit Income/Expense")
            Return
        End If

        ' Get the Record nodes and the BudgetLimit node
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
        Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//BudgetLimit")

        ' Get the limit from the Limit node and declare the boolean variable to store whether the budget limit has been exceeded
        Dim intBudgetLimit As Integer = budgetLimitNode.SelectSingleNode("Limit").InnerText
        Dim blnExceededLimit As Boolean = False
        Dim intNetIncome As Integer

        ' Load the financial report for all time periods and set it equal to the net income
        intNetIncome = LoadFinancialReport("all")

        ' Check if the net income is greater than the budget limit and the budget limit is not 0
        If intNetIncome > intBudgetLimit And intBudgetLimit <> 0 Then
            ' Set the boolean variable to true if the budget limit has been exceeded
            blnExceededLimit = True
        End If

        ' Loop through the Record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Get the Id, amount and updatedAt nodes of the current Record node
            Dim idNode As XmlNode = recordNode.SelectSingleNode("Id")
            Dim amountNode As XmlNode = recordNode.SelectSingleNode("Amount")
            Dim updatedAtNode As XmlNode = recordNode.SelectSingleNode("UpdatedAt")

            ' Check if the Id node exists and its inner text is equal to the selected record
            If idNode IsNot Nothing And idNode.InnerText = Val(txtSelectedRecord.Text) Then
                ' Check if the amount and updatedAt nodes exist
                If amountNode IsNot Nothing And updatedAtNode IsNot Nothing Then
                    ' Remove the Amount node
                    recordNode.RemoveChild(recordNode.SelectSingleNode("Amount"))

                    ' Create a new Amount node and set the inner text of the Amount node to the amount entered by the user and append it to the Record node
                    Dim newAmountNode As XmlNode = xmlDoc.CreateElement("Amount")
                    newAmountNode.InnerText = strIncomeExpense
                    recordNode.AppendChild(newAmountNode)

                    ' Remove the UpdatedAt node
                    recordNode.RemoveChild(recordNode.SelectSingleNode("UpdatedAt"))

                    ' Create a new UpdatedAt node and set the inner text of the UpdatedAt node to the current date and time and append it to the Record node
                    Dim newUpdatedAtNode As XmlNode = xmlDoc.CreateElement("UpdatedAt")
                    newUpdatedAtNode.InnerText = Date.Now()
                    recordNode.AppendChild(newUpdatedAtNode)

                    ' Save the XML file
                    xmlDoc.Save(xmlFilePath)

                    ' Load the financial report for all time periods and set it equal to the net income
                    intNetIncome = LoadFinancialReport("all")

                    ' Check if the budget limit has been exceeded and the net income is greater than the budget limit and the budget limit is not 0
                    If Not blnExceededLimit And intNetIncome > intBudgetLimit And intBudgetLimit <> 0 Then
                        ' Display a message box if the budget limit has been exceeded
                        MessageBox.Show("Budget limit exceeded!", "Edit Income/Expense")
                    ElseIf blnExceededLimit And intNetIncome <= intBudgetLimit And intBudgetLimit <> 0 Then
                        ' Display a message box if the budget limit is no longer exceeded
                        MessageBox.Show("Budget limit no longer exceeded!", "Edit Income/Expense")
                    End If

                    ' Display a message box to confirm that the income/expense has been edited
                    MessageBox.Show("Income/Expense edited successfully!", "Edit Income/Expense")
                End If
            End If
        Next
    End Sub

    ' The btnDeleteIncomeExpense_Click event handler is executed when the Delete Income/Expense button is clicked
    Private Sub btnDeleteIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnDeleteIncomeExpense.Click
        ' Get the selected record
        Dim strSelectedRecord As String = LCase(txtSelectedRecord.Text)

        ' Check if the user has selected a record
        If String.IsNullOrWhiteSpace(strSelectedRecord) Then
            ' Display a message box if the user has not selected a record
            MessageBox.Show("Please select a record!", "Edit Income/Expense")
            Return
        End If

        ' Display an input box to confirm the deletion of the record
        Dim strConfirm As String = LCase(InputBox("Are you sure you want to delete the record " & strSelectedRecord & "? (Y or N):", "Delete Income/Expense", "N"))

        ' Check if the user has confirmed the deletion
        If strConfirm <> "y" Then
            ' Display a message box if the user has not confirmed the deletion
            MessageBox.Show("Operation cancelled.", "Delete Income/Expense")
            Return
        ElseIf strConfirm = "y" Then
            ' Get the Record nodes, the BudgetLimit node and the root node
            Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
            Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//BudgetLimit")
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")

            ' Get the limit from the Limit node and declare the boolean variable to store whether the budget limit has been exceeded
            Dim intBudgetLimit As Integer = budgetLimitNode.SelectSingleNode("Limit").InnerText
            Dim blnExceededLimit As Boolean = False
            Dim intNetIncome As Integer

            ' Load the financial report for all time periods and set it equal to the net income
            intNetIncome = LoadFinancialReport("all")

            ' Check if the net income is greater than the budget limit and the budget limit is not 0
            If intNetIncome > intBudgetLimit And intBudgetLimit <> 0 Then
                ' Set the boolean variable to true if the budget limit has been exceeded
                blnExceededLimit = True
            End If

            ' Loop through the Record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Get the Id node of the current Record node
                Dim idNode As XmlNode = recordNode.SelectSingleNode("Id")

                ' Check if the Id node exists and its inner text is equal to the selected record
                If idNode IsNot Nothing And idNode.InnerText = Val(strSelectedRecord) Then
                    ' Remove the Record node
                    rootNode.RemoveChild(recordNode)

                    ' Save the XML file
                    xmlDoc.Save(xmlFilePath)

                    ' Load the financial report for all time periods and set it equal to the net income
                    intNetIncome = LoadFinancialReport("all")

                    ' Check if the budget limit has been exceeded and the net income is greater than the budget limit and the budget limit is not 0
                    If Not blnExceededLimit And intNetIncome > intBudgetLimit And intBudgetLimit <> 0 Then
                        ' Display a message box if the budget limit has been exceeded
                        MessageBox.Show("Budget limit exceeded!", "Delete Income/Expense")
                    ElseIf blnExceededLimit And intNetIncome <= intBudgetLimit And intBudgetLimit <> 0 Then
                        ' Display a message box if the budget limit is no longer exceeded
                        MessageBox.Show("Budget limit no longer exceeded!", "Delete Income/Expense")
                    End If

                    ' Display a message box to confirm that the income/expense has been deleted
                    MessageBox.Show("Income/Expense deleted successfully!", "Delete Income/Expense")
                End If
            Next
        End If
    End Sub

    ' The btnSetBudgetLimit_Click event handler is executed when the Set Budget Limit button is clicked
    Private Sub btnSetBudgetLimit_Click(sender As Object, e As EventArgs) Handles btnSetBudgetLimit.Click
        ' Display an input box to get the budget limit
        Dim strBudgetLimit As String = LCase(InputBox("Enter the budget limit (or type None to remove):", "Set Budget Limit", "1000"))

        ' Check if the user has entered a budget limit
        If String.IsNullOrWhiteSpace(strBudgetLimit) Then
            ' Display a message box if the user has not entered a budget limit
            MessageBox.Show("Operation cancelled.", "Set Budget Limit")
            Return
        End If

        ' Check if the budget limit is not none
        If strBudgetLimit = "none" Then
            ' Get the root node and the BudgetLimit node
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
            Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//BudgetLimit")

            ' Check if the BudgetLimit node exists
            If budgetLimitNode IsNot Nothing Then
                ' Remove the Limit node
                budgetLimitNode.RemoveChild(budgetLimitNode.SelectSingleNode("Limit"))

                ' Create a new Limit node and set the inner text of the Limit node to 0 and append it to the BudgetLimit node
                Dim limitNode As XmlNode = xmlDoc.CreateElement("Limit")
                limitNode.InnerText = "0"
                budgetLimitNode.AppendChild(limitNode)

                ' Save the XML file
                xmlDoc.Save(xmlFilePath)

                ' Display a message box to confirm that the budget limit has been set
                MessageBox.Show("Budget limit set successfully!", "Set Budget Limit")
            End If
        ElseIf strBudgetLimit <> "none" Then
            ' Get the root node and the BudgetLimit node
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
            Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//BudgetLimit")

            ' Check if the BudgetLimit node exists
            If budgetLimitNode IsNot Nothing Then
                ' Remove the Limit node
                budgetLimitNode.RemoveChild(budgetLimitNode.SelectSingleNode("Limit"))

                ' Create a new Limit node and set the inner text of the Limit node to the budget limit entered by the user and append it to the BudgetLimit node
                Dim limitNode As XmlNode = xmlDoc.CreateElement("Limit")
                limitNode.InnerText = strBudgetLimit
                budgetLimitNode.AppendChild(limitNode)

                ' Save the XML file
                xmlDoc.Save(xmlFilePath)

                ' Display a message box to confirm that the budget limit has been set
                MessageBox.Show("Budget limit set successfully!", "Set Budget Limit")
            End If
        End If

        ' Load the financial report for all time periods
        LoadFinancialReport("all")
    End Sub

    ' The btnViewRecords_Click event handler is executed when the View Records button is clicked
    Private Sub btnViewRecords_Click(sender As Object, e As EventArgs) Handles btnViewRecords.Click
        ' Load the records
        LoadRecords()
    End Sub

    ' The btnViewReport_Click event handler is executed when the View Report button is clicked
    Private Sub btnViewReport_Click(sender As Object, e As EventArgs) Handles btnViewReport.Click
        ' Display an input box to get the year for the financial report
        Dim strTimePeriod As String = LCase(InputBox("Enter the year for the financial report (e.g. All, 2024):", "Financial Report", "2024"))

        ' Check if the user has entered a year
        If String.IsNullOrWhiteSpace(strTimePeriod) Then
            ' Display a message box if the user has not entered a year
            MessageBox.Show("Operation cancelled.", "View Report")
            Return
        End If

        ' Load the financial report for the year entered by the user
        LoadFinancialReport(strTimePeriod)
    End Sub

    ' The btnSortRecords_Click event handler is executed when the Sort Records button is clicked
    Private Sub btnSortRecords_Click(sender As Object, e As EventArgs) Handles btnSortRecords.Click
        ' Load the records
        LoadRecords()

        ' Declare the integer array to store the amounts of the displayed records and the current index
        Dim intNbrsArray(intDisplayedRecords - 1) As Integer
        Dim intCurrentIndex As Integer = 0

        ' Display an input box to get the sort option for the records
        Dim strSortOption As String = InputBox("Enter the sort option for the records (Id or Amount):", "Sort Records", "Amount")

        ' Check if the user has entered a sort option
        If String.IsNullOrWhiteSpace(strSortOption) Then
            ' Display a message box if the user has not entered a sort option
            MessageBox.Show("Operation cancelled.", "Sort Records")
            Return
        End If

        ' Display an input box to get the sort order for the records
        Dim strSortOrder As String = LCase(InputBox("Sort the records in descending order (Y or N):", "Sort Records", "N"))

        ' Check if the user has entered a sort order
        If String.IsNullOrWhiteSpace(strSortOrder) Then
            ' Display a message box if the user has not entered a sort order
            MessageBox.Show("Operation cancelled.", "Sort Records")
            Return
        End If

        ' Get the Record nodes
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        ' Loop through the Record nodes
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
        Dim i As Integer = 0
        ' Declare the integer variable to store the index counter value to shuffle along 1 index step
        Dim j As Integer = 0

        ' Declare the integer variable to store the index to move the minimum value to
        Dim intMinValueIndex As Integer
        ' Declare the integer variable to store the index value that the minimum value is found at
        Dim intSwapValueIndex As Integer
        ' Declare the integer variable to store the actual value of the found minimum
        Dim intSwapValue As Integer

        ' Loop through the integer array
        For i = intLowerBound To intUpperBound
            ' Set the minimum value index to the current index
            intMinValueIndex = i
            ' Reset the saved swap values from the last pass of the array
            intSwapValue = 0
            intSwapValueIndex = 0

            ' Loop through the integer array starting from the next index
            For j = i + 1 To intUpperBound
                ' Search and find the lowest value remaining in the array for this sort pass of the array
                If intNbrsArray(j) < intNbrsArray(intMinValueIndex) Then
                    ' First check for first time finding a smaller value, then check if smallest so far
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
        For c = 0 To intNbrsArray.Length - 1
            ' Loop through the Record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Check if the selected node based on the sort option is equal to the current index of the integer array
                If recordNode.SelectSingleNode(strSortOption).InnerText = intNbrsArray(c) Then
                    ' Add the record to the list box
                    lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
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
        Dim strSearchRequestType As String = InputBox("Enter the search request type (Id or Amount):", "Search Records", "Amount")

        ' Check if the user has entered a search request type
        If String.IsNullOrWhiteSpace(strSearchRequestType) Then
            ' Display a message box if the user has not entered a search request type
            MessageBox.Show("Operation cancelled.", "Search Records")
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

        ' Get the Record nodes
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        ' Loop through the Record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Get the selected node based on the search request type
            Dim selectedNode As XmlNode = recordNode.SelectSingleNode(strSearchRequestType)

            ' Set each index of the integer array to the value of the selected node and increment the current index
            intNbrsArray(intCurrentIndex) = (Val(selectedNode.InnerText))
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

            ' Loop through the Record nodes
            For Each recordNode As XmlNode In recordNodes
                ' Check if the selected node based on the search request type is equal to the search request item
                If recordNode.SelectSingleNode(strSearchRequestType).InnerText = strSearchRequestItem Then
                    ' Add the record to the list box
                    lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
                End If
            Next
        Else
            ' Display a message box if the item has not been found
            MessageBox.Show("Item not found!", "Search Records")
        End If
    End Sub

    ' The btnClear_Click event handler is executed when the Clear button is clicked
    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ' Clear the combo box, text boxes and list box
        cbxCategory.SelectedIndex = 0
        txtIncomeExpense.Clear()
        txtSelectedRecord.Clear()
        lstDisplay.Items.Clear()
        ' Set the focus to the combo box
        cbxCategory.Focus()
    End Sub

    ' The subroutine PopulateComboBox is used to populate the combo box with the categories from the XML file
    Private Sub PopulateComboBox()
        ' Get the Category nodes
        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/Category")

        ' Clear the combo box and add the "All" item
        cbxCategory.Items.Clear()
        cbxCategory.Items.Add("All")

        ' Loop through the Category nodes
        For Each categoryNode As XmlNode In categoryNodes
            ' Add the Name node inner text to the combo box
            cbxCategory.Items.Add(categoryNode.SelectSingleNode("Name").InnerText)
        Next

        ' Set the selected index of the combo box to 0
        cbxCategory.SelectedIndex = 0
    End Sub

    ' The subroutine LoadRecords is used to load the records from the XML file based on the selected category
    Private Sub LoadRecords()
        ' Get the Record nodes
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        ' Set the displayed records counter to 0
        intDisplayedRecords = 0

        ' Clear the list box and add the selected category to the list box
        lstDisplay.Items.Clear()
        lstDisplay.Items.Add("Category: " & cbxCategory.SelectedItem)

        ' Loop through the Record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Check if the selected category is "All" or the Category node inner text is equal to the selected category
            If cbxCategory.SelectedItem = "All" Or recordNode.SelectSingleNode("Category").InnerText = cbxCategory.SelectedItem Then
                ' Add the record to the list box and increment the displayed records counter
                lstDisplay.Items.Add(recordNode.SelectSingleNode("Id").InnerText & ": $" & recordNode.SelectSingleNode("Amount").InnerText)
                intDisplayedRecords = intDisplayedRecords + 1
            End If
        Next
    End Sub

    ' The function LoadFinancialReport is used to load the financial report from the XML file based on the selected time period
    ' strTimePeriod Parameter: The time period for the financial report
    ' Returns the net income for the selected time period
    Public Function LoadFinancialReport(ByVal strTimePeriod As String) As Integer
        ' Declare the integer variables to store the total income, total expenses, net income and budget limit
        Dim intTotalIncome As Integer
        Dim intTotalExpenses As Integer
        Dim intNetIncome As Integer

        ' Get the Record nodes and the BudgetLimit node
        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")
        Dim budgetLimitNode As XmlNode = xmlDoc.SelectSingleNode("//BudgetLimit")

        ' Get the limit from the Limit node
        Dim intBudgetLimit As Integer = budgetLimitNode.SelectSingleNode("Limit").InnerText

        ' Clear the list box
        lstDisplay.Items.Clear()

        ' Loop through the Record nodes
        For Each recordNode As XmlNode In recordNodes
            ' Get the UpdatedAt node of the current Record node
            Dim updatedAtNode As XmlNode = recordNode.SelectSingleNode("UpdatedAt")

            ' Check if the UpdatedAt node exists and its inner text contains the selected time period
            If updatedAtNode IsNot Nothing And (updatedAtNode.InnerText.Contains(strTimePeriod) OrElse strTimePeriod = "all") Then
                ' Check if the Amount node inner text starts with a negative sign
                If recordNode.SelectSingleNode("Amount").InnerText.StartsWith("-") Then
                    ' Add the amount to the total expenses
                    intTotalExpenses = intTotalExpenses + Val(recordNode.SelectSingleNode("Amount").InnerText)
                Else
                    ' Add the amount to the total income
                    intTotalIncome = intTotalIncome + Val(recordNode.SelectSingleNode("Amount").InnerText)
                End If
            End If
        Next

        ' Calculate the net income
        intNetIncome = intTotalIncome + intTotalExpenses

        ' Add the financial report to the list box
        lstDisplay.Items.Add("Total Income: $" & intTotalIncome)
        lstDisplay.Items.Add("Total Expenses: $" & intTotalExpenses)
        lstDisplay.Items.Add("Net Income: $" & intNetIncome)

        If intBudgetLimit = 0 Then
            lstDisplay.Items.Add("Budget Limit: None")
            lstDisplay.Items.Add("Remaining Budget: No Budget")
        ElseIf intBudgetLimit <> 0 Then
            lstDisplay.Items.Add("Budget Limit: $" & intBudgetLimit)
            lstDisplay.Items.Add("Remaining Budget: $" & (intBudgetLimit - intNetIncome))
        End If

        lstDisplay.Items.Add("Time Period: " & strTimePeriod)

        ' Return the net income
        Return intNetIncome
    End Function
End Class
