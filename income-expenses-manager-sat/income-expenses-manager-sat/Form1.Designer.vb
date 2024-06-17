<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        lblTitle = New Label()
        lblCategory = New Label()
        lblIncomeExpense = New Label()
        lblSelectedRecord = New Label()
        cbxCategory = New ComboBox()
        txtIncomeExpense = New TextBox()
        txtSelectedRecord = New TextBox()
        btnAddCategory = New Button()
        btnEditCategory = New Button()
        btnDeleteCategory = New Button()
        btnAddIncomeExpense = New Button()
        btnEditIncomeExpense = New Button()
        btnDeleteIncomeExpense = New Button()
        btnSetBudgetLimit = New Button()
        btnViewRecords = New Button()
        btnViewReport = New Button()
        btnSortRecords = New Button()
        btnSearchRecords = New Button()
        btnClear = New Button()
        lstDisplay = New ListBox()
        SuspendLayout()
        ' 
        ' lblTitle
        ' 
        lblTitle.AutoSize = True
        lblTitle.Font = New Font("Arial", 16.125F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblTitle.Location = New Point(50, 50)
        lblTitle.Name = "lblTitle"
        lblTitle.Size = New Size(676, 51)
        lblTitle.TabIndex = 0
        lblTitle.Text = "INCOME EXPENSES MANAGER"
        ' 
        ' lblCategory
        ' 
        lblCategory.AutoSize = True
        lblCategory.Location = New Point(50, 200)
        lblCategory.Name = "lblCategory"
        lblCategory.Size = New Size(141, 36)
        lblCategory.TabIndex = 1
        lblCategory.Text = "Category"
        ' 
        ' lblIncomeExpense
        ' 
        lblIncomeExpense.AutoSize = True
        lblIncomeExpense.Location = New Point(50, 300)
        lblIncomeExpense.Name = "lblIncomeExpense"
        lblIncomeExpense.Size = New Size(248, 36)
        lblIncomeExpense.TabIndex = 2
        lblIncomeExpense.Text = "Income/Expense"
        ' 
        ' lblSelectedRecord
        ' 
        lblSelectedRecord.AutoSize = True
        lblSelectedRecord.Location = New Point(50, 400)
        lblSelectedRecord.Name = "lblSelectedRecord"
        lblSelectedRecord.Size = New Size(246, 36)
        lblSelectedRecord.TabIndex = 3
        lblSelectedRecord.Text = "Selected Record"
        ' 
        ' cbxCategory
        ' 
        cbxCategory.FormattingEnabled = True
        cbxCategory.Location = New Point(310, 200)
        cbxCategory.Name = "cbxCategory"
        cbxCategory.Size = New Size(690, 44)
        cbxCategory.TabIndex = 4
        ' 
        ' txtIncomeExpense
        ' 
        txtIncomeExpense.Location = New Point(310, 300)
        txtIncomeExpense.Name = "txtIncomeExpense"
        txtIncomeExpense.Size = New Size(690, 44)
        txtIncomeExpense.TabIndex = 5
        ' 
        ' txtSelectedRecord
        ' 
        txtSelectedRecord.Location = New Point(310, 400)
        txtSelectedRecord.Name = "txtSelectedRecord"
        txtSelectedRecord.Size = New Size(690, 44)
        txtSelectedRecord.TabIndex = 6
        ' 
        ' btnAddCategory
        ' 
        btnAddCategory.Location = New Point(50, 550)
        btnAddCategory.Name = "btnAddCategory"
        btnAddCategory.Size = New Size(300, 100)
        btnAddCategory.TabIndex = 7
        btnAddCategory.Text = "Add Category"
        btnAddCategory.UseVisualStyleBackColor = True
        ' 
        ' btnEditCategory
        ' 
        btnEditCategory.Location = New Point(375, 550)
        btnEditCategory.Name = "btnEditCategory"
        btnEditCategory.Size = New Size(300, 100)
        btnEditCategory.TabIndex = 8
        btnEditCategory.Text = "Edit Category"
        btnEditCategory.UseVisualStyleBackColor = True
        ' 
        ' btnDeleteCategory
        ' 
        btnDeleteCategory.Location = New Point(700, 550)
        btnDeleteCategory.Name = "btnDeleteCategory"
        btnDeleteCategory.Size = New Size(300, 100)
        btnDeleteCategory.TabIndex = 9
        btnDeleteCategory.Text = "Delete Category"
        btnDeleteCategory.UseVisualStyleBackColor = True
        ' 
        ' btnAddIncomeExpense
        ' 
        btnAddIncomeExpense.Location = New Point(50, 675)
        btnAddIncomeExpense.Name = "btnAddIncomeExpense"
        btnAddIncomeExpense.Size = New Size(300, 100)
        btnAddIncomeExpense.TabIndex = 10
        btnAddIncomeExpense.Text = "Add Income/Expense"
        btnAddIncomeExpense.UseVisualStyleBackColor = True
        ' 
        ' btnEditIncomeExpense
        ' 
        btnEditIncomeExpense.Location = New Point(375, 675)
        btnEditIncomeExpense.Name = "btnEditIncomeExpense"
        btnEditIncomeExpense.Size = New Size(300, 100)
        btnEditIncomeExpense.TabIndex = 11
        btnEditIncomeExpense.Text = "Edit Income/Expense"
        btnEditIncomeExpense.UseVisualStyleBackColor = True
        ' 
        ' btnDeleteIncomeExpense
        ' 
        btnDeleteIncomeExpense.Location = New Point(700, 675)
        btnDeleteIncomeExpense.Name = "btnDeleteIncomeExpense"
        btnDeleteIncomeExpense.Size = New Size(300, 100)
        btnDeleteIncomeExpense.TabIndex = 12
        btnDeleteIncomeExpense.Text = "Delete Income/Expense"
        btnDeleteIncomeExpense.UseVisualStyleBackColor = True
        ' 
        ' btnSetBudgetLimit
        ' 
        btnSetBudgetLimit.Location = New Point(50, 800)
        btnSetBudgetLimit.Name = "btnSetBudgetLimit"
        btnSetBudgetLimit.Size = New Size(300, 100)
        btnSetBudgetLimit.TabIndex = 13
        btnSetBudgetLimit.Text = "Set Budget Limit"
        btnSetBudgetLimit.UseVisualStyleBackColor = True
        ' 
        ' btnViewRecords
        ' 
        btnViewRecords.Location = New Point(375, 800)
        btnViewRecords.Name = "btnViewRecords"
        btnViewRecords.Size = New Size(300, 100)
        btnViewRecords.TabIndex = 14
        btnViewRecords.Text = "View Records"
        btnViewRecords.UseVisualStyleBackColor = True
        ' 
        ' btnViewReport
        ' 
        btnViewReport.Location = New Point(700, 800)
        btnViewReport.Name = "btnViewReport"
        btnViewReport.Size = New Size(300, 100)
        btnViewReport.TabIndex = 15
        btnViewReport.Text = "View Report"
        btnViewReport.UseVisualStyleBackColor = True
        ' 
        ' btnSortRecords
        ' 
        btnSortRecords.Location = New Point(50, 925)
        btnSortRecords.Name = "btnSortRecords"
        btnSortRecords.Size = New Size(300, 100)
        btnSortRecords.TabIndex = 16
        btnSortRecords.Text = "Sort Records"
        btnSortRecords.UseVisualStyleBackColor = True
        ' 
        ' btnSearchRecords
        ' 
        btnSearchRecords.Location = New Point(375, 925)
        btnSearchRecords.Name = "btnSearchRecords"
        btnSearchRecords.Size = New Size(300, 100)
        btnSearchRecords.TabIndex = 17
        btnSearchRecords.Text = "Search Records"
        btnSearchRecords.UseVisualStyleBackColor = True
        ' 
        ' btnClear
        ' 
        btnClear.Location = New Point(700, 925)
        btnClear.Name = "btnClear"
        btnClear.Size = New Size(300, 100)
        btnClear.TabIndex = 18
        btnClear.Text = "Clear"
        btnClear.UseVisualStyleBackColor = True
        ' 
        ' lstDisplay
        ' 
        lstDisplay.BackColor = Color.PaleGreen
        lstDisplay.FormattingEnabled = True
        lstDisplay.ItemHeight = 36
        lstDisplay.Location = New Point(1050, 50)
        lstDisplay.Name = "lstDisplay"
        lstDisplay.Size = New Size(875, 976)
        lstDisplay.TabIndex = 19
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(18F, 36F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.ForestGreen
        ClientSize = New Size(1974, 1129)
        Controls.Add(lstDisplay)
        Controls.Add(btnClear)
        Controls.Add(btnSearchRecords)
        Controls.Add(btnSortRecords)
        Controls.Add(btnViewReport)
        Controls.Add(btnViewRecords)
        Controls.Add(btnSetBudgetLimit)
        Controls.Add(btnDeleteIncomeExpense)
        Controls.Add(btnEditIncomeExpense)
        Controls.Add(btnAddIncomeExpense)
        Controls.Add(btnDeleteCategory)
        Controls.Add(btnEditCategory)
        Controls.Add(btnAddCategory)
        Controls.Add(txtSelectedRecord)
        Controls.Add(txtIncomeExpense)
        Controls.Add(cbxCategory)
        Controls.Add(lblSelectedRecord)
        Controls.Add(lblIncomeExpense)
        Controls.Add(lblCategory)
        Controls.Add(lblTitle)
        Font = New Font("Arial", 12F, FontStyle.Regular, GraphicsUnit.Point, CByte(0))
        Margin = New Padding(4, 3, 4, 3)
        Name = "Form1"
        Text = "Income Expenses Manager"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblTitle As Label
    Friend WithEvents lblCategory As Label
    Friend WithEvents lblIncomeExpense As Label
    Friend WithEvents lblSelectedRecord As Label
    Friend WithEvents cbxCategory As ComboBox
    Friend WithEvents txtIncomeExpense As TextBox
    Friend WithEvents txtSelectedRecord As TextBox
    Friend WithEvents btnAddCategory As Button
    Friend WithEvents btnEditCategory As Button
    Friend WithEvents btnDeleteCategory As Button
    Friend WithEvents btnAddIncomeExpense As Button
    Friend WithEvents btnEditIncomeExpense As Button
    Friend WithEvents btnDeleteIncomeExpense As Button
    Friend WithEvents btnSetBudgetLimit As Button
    Friend WithEvents btnViewRecords As Button
    Friend WithEvents btnViewReport As Button
    Friend WithEvents btnSortRecords As Button
    Friend WithEvents btnSearchRecords As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents lstDisplay As ListBox

End Class
