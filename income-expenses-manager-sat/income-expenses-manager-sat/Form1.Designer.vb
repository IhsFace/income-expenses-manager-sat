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
        lblIncome = New Label()
        lblSelectedRecord = New Label()
        cbxCategory = New ComboBox()
        txtIncome = New TextBox()
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
        lblExpense = New Label()
        txtExpense = New TextBox()
        lblName = New Label()
        txtName = New TextBox()
        SuspendLayout()
        ' 
        ' lblTitle
        ' 
        lblTitle.AutoSize = True
        lblTitle.Font = New Font("Arial Black", 16.125F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblTitle.ForeColor = SystemColors.Control
        lblTitle.Location = New Point(50, 50)
        lblTitle.Name = "lblTitle"
        lblTitle.Size = New Size(744, 60)
        lblTitle.TabIndex = 0
        lblTitle.Text = "INCOME EXPENSES MANAGER"
        ' 
        ' lblCategory
        ' 
        lblCategory.AutoSize = True
        lblCategory.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblCategory.ForeColor = SystemColors.Control
        lblCategory.Location = New Point(50, 150)
        lblCategory.Name = "lblCategory"
        lblCategory.Size = New Size(155, 37)
        lblCategory.TabIndex = 1
        lblCategory.Text = "Category"
        ' 
        ' lblIncome
        ' 
        lblIncome.AutoSize = True
        lblIncome.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblIncome.ForeColor = SystemColors.Control
        lblIncome.Location = New Point(50, 300)
        lblIncome.Name = "lblIncome"
        lblIncome.Size = New Size(127, 37)
        lblIncome.TabIndex = 3
        lblIncome.Text = "Income"
        ' 
        ' lblSelectedRecord
        ' 
        lblSelectedRecord.AutoSize = True
        lblSelectedRecord.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblSelectedRecord.ForeColor = SystemColors.Control
        lblSelectedRecord.Location = New Point(50, 450)
        lblSelectedRecord.Name = "lblSelectedRecord"
        lblSelectedRecord.Size = New Size(166, 37)
        lblSelectedRecord.TabIndex = 5
        lblSelectedRecord.Text = "Record ID"
        ' 
        ' cbxCategory
        ' 
        cbxCategory.FormattingEnabled = True
        cbxCategory.Location = New Point(222, 150)
        cbxCategory.Name = "cbxCategory"
        cbxCategory.Size = New Size(778, 44)
        cbxCategory.TabIndex = 6
        ' 
        ' txtIncome
        ' 
        txtIncome.Location = New Point(222, 300)
        txtIncome.Name = "txtIncome"
        txtIncome.Size = New Size(778, 44)
        txtIncome.TabIndex = 8
        ' 
        ' txtSelectedRecord
        ' 
        txtSelectedRecord.Location = New Point(222, 450)
        txtSelectedRecord.Name = "txtSelectedRecord"
        txtSelectedRecord.Size = New Size(778, 44)
        txtSelectedRecord.TabIndex = 10
        ' 
        ' btnAddCategory
        ' 
        btnAddCategory.BackColor = Color.SpringGreen
        btnAddCategory.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnAddCategory.Location = New Point(50, 550)
        btnAddCategory.Name = "btnAddCategory"
        btnAddCategory.Size = New Size(300, 100)
        btnAddCategory.TabIndex = 11
        btnAddCategory.Text = "Add Category"
        btnAddCategory.UseVisualStyleBackColor = False
        ' 
        ' btnEditCategory
        ' 
        btnEditCategory.BackColor = Color.DeepSkyBlue
        btnEditCategory.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnEditCategory.Location = New Point(375, 550)
        btnEditCategory.Name = "btnEditCategory"
        btnEditCategory.Size = New Size(300, 100)
        btnEditCategory.TabIndex = 12
        btnEditCategory.Text = "Edit Category"
        btnEditCategory.UseVisualStyleBackColor = False
        ' 
        ' btnDeleteCategory
        ' 
        btnDeleteCategory.BackColor = Color.Red
        btnDeleteCategory.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnDeleteCategory.ForeColor = SystemColors.Control
        btnDeleteCategory.Location = New Point(700, 550)
        btnDeleteCategory.Name = "btnDeleteCategory"
        btnDeleteCategory.Size = New Size(300, 100)
        btnDeleteCategory.TabIndex = 13
        btnDeleteCategory.Text = "Delete Category"
        btnDeleteCategory.UseVisualStyleBackColor = False
        ' 
        ' btnAddIncomeExpense
        ' 
        btnAddIncomeExpense.BackColor = Color.SpringGreen
        btnAddIncomeExpense.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnAddIncomeExpense.Location = New Point(50, 675)
        btnAddIncomeExpense.Name = "btnAddIncomeExpense"
        btnAddIncomeExpense.Size = New Size(300, 100)
        btnAddIncomeExpense.TabIndex = 14
        btnAddIncomeExpense.Text = "Add Income/Expense"
        btnAddIncomeExpense.UseVisualStyleBackColor = False
        ' 
        ' btnEditIncomeExpense
        ' 
        btnEditIncomeExpense.BackColor = Color.DeepSkyBlue
        btnEditIncomeExpense.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnEditIncomeExpense.Location = New Point(375, 675)
        btnEditIncomeExpense.Name = "btnEditIncomeExpense"
        btnEditIncomeExpense.Size = New Size(300, 100)
        btnEditIncomeExpense.TabIndex = 15
        btnEditIncomeExpense.Text = "Edit Income/Expense"
        btnEditIncomeExpense.UseVisualStyleBackColor = False
        ' 
        ' btnDeleteIncomeExpense
        ' 
        btnDeleteIncomeExpense.BackColor = Color.Red
        btnDeleteIncomeExpense.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnDeleteIncomeExpense.ForeColor = SystemColors.Control
        btnDeleteIncomeExpense.Location = New Point(700, 675)
        btnDeleteIncomeExpense.Name = "btnDeleteIncomeExpense"
        btnDeleteIncomeExpense.Size = New Size(300, 100)
        btnDeleteIncomeExpense.TabIndex = 16
        btnDeleteIncomeExpense.Text = "Delete Income/Expense"
        btnDeleteIncomeExpense.UseVisualStyleBackColor = False
        ' 
        ' btnSetBudgetLimit
        ' 
        btnSetBudgetLimit.BackColor = Color.MediumSlateBlue
        btnSetBudgetLimit.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnSetBudgetLimit.ForeColor = SystemColors.Control
        btnSetBudgetLimit.Location = New Point(50, 800)
        btnSetBudgetLimit.Name = "btnSetBudgetLimit"
        btnSetBudgetLimit.Size = New Size(300, 100)
        btnSetBudgetLimit.TabIndex = 17
        btnSetBudgetLimit.Text = "Set Budget Limit"
        btnSetBudgetLimit.UseVisualStyleBackColor = False
        ' 
        ' btnViewRecords
        ' 
        btnViewRecords.BackColor = Color.PaleGreen
        btnViewRecords.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnViewRecords.Location = New Point(50, 925)
        btnViewRecords.Name = "btnViewRecords"
        btnViewRecords.Size = New Size(300, 100)
        btnViewRecords.TabIndex = 20
        btnViewRecords.Text = "View Records"
        btnViewRecords.UseVisualStyleBackColor = False
        ' 
        ' btnViewReport
        ' 
        btnViewReport.BackColor = Color.PaleGreen
        btnViewReport.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnViewReport.Location = New Point(375, 925)
        btnViewReport.Name = "btnViewReport"
        btnViewReport.Size = New Size(300, 100)
        btnViewReport.TabIndex = 21
        btnViewReport.Text = "View Report"
        btnViewReport.UseVisualStyleBackColor = False
        ' 
        ' btnSortRecords
        ' 
        btnSortRecords.BackColor = Color.Orange
        btnSortRecords.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnSortRecords.Location = New Point(375, 800)
        btnSortRecords.Name = "btnSortRecords"
        btnSortRecords.Size = New Size(300, 100)
        btnSortRecords.TabIndex = 18
        btnSortRecords.Text = "Sort Records"
        btnSortRecords.UseVisualStyleBackColor = False
        ' 
        ' btnSearchRecords
        ' 
        btnSearchRecords.BackColor = Color.Teal
        btnSearchRecords.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnSearchRecords.ForeColor = SystemColors.Control
        btnSearchRecords.Location = New Point(700, 800)
        btnSearchRecords.Name = "btnSearchRecords"
        btnSearchRecords.Size = New Size(300, 100)
        btnSearchRecords.TabIndex = 19
        btnSearchRecords.Text = "Search Records"
        btnSearchRecords.UseVisualStyleBackColor = False
        ' 
        ' btnClear
        ' 
        btnClear.BackColor = Color.Yellow
        btnClear.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        btnClear.Location = New Point(700, 925)
        btnClear.Name = "btnClear"
        btnClear.Size = New Size(300, 100)
        btnClear.TabIndex = 22
        btnClear.Text = "Clear"
        btnClear.UseVisualStyleBackColor = False
        ' 
        ' lstDisplay
        ' 
        lstDisplay.BackColor = Color.PaleGreen
        lstDisplay.FormattingEnabled = True
        lstDisplay.ItemHeight = 36
        lstDisplay.Location = New Point(1050, 50)
        lstDisplay.Name = "lstDisplay"
        lstDisplay.Size = New Size(875, 976)
        lstDisplay.TabIndex = 23
        ' 
        ' lblExpense
        ' 
        lblExpense.AutoSize = True
        lblExpense.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblExpense.ForeColor = SystemColors.Control
        lblExpense.Location = New Point(50, 375)
        lblExpense.Name = "lblExpense"
        lblExpense.Size = New Size(148, 37)
        lblExpense.TabIndex = 4
        lblExpense.Text = "Expense"
        ' 
        ' txtExpense
        ' 
        txtExpense.Location = New Point(222, 375)
        txtExpense.Name = "txtExpense"
        txtExpense.Size = New Size(778, 44)
        txtExpense.TabIndex = 9
        ' 
        ' lblName
        ' 
        lblName.AutoSize = True
        lblName.Font = New Font("Arial", 12F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblName.ForeColor = SystemColors.Control
        lblName.Location = New Point(50, 225)
        lblName.Name = "lblName"
        lblName.Size = New Size(103, 37)
        lblName.TabIndex = 2
        lblName.Text = "Name"
        ' 
        ' txtName
        ' 
        txtName.Location = New Point(222, 225)
        txtName.Name = "txtName"
        txtName.Size = New Size(778, 44)
        txtName.TabIndex = 7
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(18F, 36F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.ForestGreen
        ClientSize = New Size(1974, 1129)
        Controls.Add(txtName)
        Controls.Add(lblName)
        Controls.Add(txtExpense)
        Controls.Add(lblExpense)
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
        Controls.Add(txtIncome)
        Controls.Add(cbxCategory)
        Controls.Add(lblSelectedRecord)
        Controls.Add(lblIncome)
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
    Friend WithEvents lblIncome As Label
    Friend WithEvents lblSelectedRecord As Label
    Friend WithEvents cbxCategory As ComboBox
    Friend WithEvents txtIncome As TextBox
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
    Friend WithEvents lblExpense As Label
    Friend WithEvents txtExpense As TextBox
    Friend WithEvents lblName As Label
    Friend WithEvents txtName As TextBox

End Class
