<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Форма переопределяет dispose для очистки списка компонентов.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Является обязательной для конструктора форм Windows Forms
    Private components As System.ComponentModel.IContainer

    'Примечание: следующая процедура является обязательной для конструктора форм Windows Forms
    'Для ее изменения используйте конструктор форм Windows Form.  
    'Не изменяйте ее в редакторе исходного кода.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.PlanetsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.PlanetsDataSet1 = New Планеты.PlanetsDataSet1()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.PlanetsTableAdapter = New Планеты.PlanetsDataSet1TableAdapters.PlanetsTableAdapter()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        CType(Me.PlanetsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PlanetsDataSet1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PlanetsBindingSource
        '
        Me.PlanetsBindingSource.DataMember = "Planets"
        Me.PlanetsBindingSource.DataSource = Me.PlanetsDataSet1
        '
        'PlanetsDataSet1
        '
        Me.PlanetsDataSet1.DataSetName = "PlanetsDataSet1"
        Me.PlanetsDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 295)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(160, 52)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Обновить"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(221, 295)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(160, 52)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Добавить Планету"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(424, 295)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(160, 52)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Удалить Планету"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(628, 295)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(160, 52)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Редактировать Планету"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'PlanetsTableAdapter
        '
        Me.PlanetsTableAdapter.ClearBeforeFill = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(12, 22)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(776, 252)
        Me.DataGridView1.TabIndex = 5
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Form2"
        Me.Text = "Редактирование базы планет"
        CType(Me.PlanetsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PlanetsDataSet1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents PlanetsDataSet1 As PlanetsDataSet1
    Friend WithEvents PlanetsBindingSource As BindingSource
    Friend WithEvents PlanetsTableAdapter As PlanetsDataSet1TableAdapters.PlanetsTableAdapter
    Friend WithEvents DataGridView1 As DataGridView
End Class
