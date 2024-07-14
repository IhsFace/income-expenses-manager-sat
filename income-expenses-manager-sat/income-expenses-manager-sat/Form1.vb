Imports System.Xml

Public Class Form1
    Dim xmlDoc As New XmlDocument()
    Dim xmlFilePath As String = "IncomeExpenseRecords.xml"
    Dim blnItemsValid As Boolean = True

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If System.IO.File.Exists(xmlFilePath) Then
            xmlDoc.Load(xmlFilePath)
            PopulateComboBox()
            LoadFinancialReport()
        Else
            Dim xmlDeclaration As XmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", Nothing)
            xmlDoc.AppendChild(xmlDeclaration)
            Dim rootNode As XmlNode = xmlDoc.CreateElement("Records")
            xmlDoc.AppendChild(rootNode)
            xmlDoc.Save(xmlFilePath)
        End If
    End Sub

    Private Sub btnAddCategory_Click(sender As Object, e As EventArgs) Handles btnAddCategory.Click
        DoValidation()

        If blnItemsValid Then
            Dim strName As String = InputBox("Enter the name of the category:", "Add Category", "Category Name")

            Dim categoryNode As XmlNode = xmlDoc.CreateElement(strName)
            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records")
            rootNode.AppendChild(categoryNode)

            xmlDoc.Save(xmlFilePath)

            PopulateComboBox()
            LoadFinancialReport()

            MessageBox.Show("Category added successfully!", "Add Category")
        End If
    End Sub

    Private Sub btnAddIncomeExpense_Click(sender As Object, e As EventArgs) Handles btnAddIncomeExpense.Click
        DoValidation()

        If blnItemsValid Then
            Dim intAmount As Integer = Val(txtIncomeExpense.Text)
            Dim dteDate As Date = Date.Now()

            Dim recordNode As XmlNode = xmlDoc.CreateElement("Record")

            Dim amountNode As XmlNode = xmlDoc.CreateElement("Amount")
            amountNode.InnerText = intAmount
            recordNode.AppendChild(amountNode)

            Dim dateNode As XmlNode = xmlDoc.CreateElement("Date")
            dateNode.InnerText = dteDate
            recordNode.AppendChild(dateNode)

            Dim rootNode As XmlNode = xmlDoc.SelectSingleNode("Records/" & cbxCategory.SelectedItem)
            rootNode.AppendChild(recordNode)

            xmlDoc.Save(xmlFilePath)

            LoadFinancialReport()

        End If
    End Sub

    Private Sub btnViewReport_Click(sender As Object, e As EventArgs) Handles btnViewReport.Click
        LoadFinancialReport()
    End Sub

    Private Sub PopulateComboBox()
        Dim categoryNodes As XmlNodeList = xmlDoc.SelectNodes("Records/*")

        For Each categoryNode As XmlNode In categoryNodes
            cbxCategory.Items.Add(categoryNode.Name)
        Next

        cbxCategory.SelectedIndex = 0
    End Sub

    Private Sub LoadFinancialReport()
        Dim intTotalIncome As Integer
        Dim intTotalExpenses As Integer
        'Dim strTimePeriod As String = InputBox("Enter the time period for the financial report (e.g. 2024):", "Financial Report", "2024")

        Dim recordNodes As XmlNodeList = xmlDoc.SelectNodes("//Record")

        lstDisplay.Items.Clear()

        For Each recordNode As XmlNode In recordNodes
            If recordNode.SelectSingleNode("Date").InnerText.Contains("2024") Then
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
        lstDisplay.Items.Add("Time Period: " & "2024")
    End Sub

    Private Sub DoValidation()
    End Sub
End Class
