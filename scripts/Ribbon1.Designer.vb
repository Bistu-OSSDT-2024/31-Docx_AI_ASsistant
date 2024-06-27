Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Windows.Forms 类撰写设计器支持所必需的
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        '组件设计器需要此调用。
        InitializeComponent()

    End Sub

    '组件重写释放以清理组件列表。
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

    '组件设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是组件设计器所必需的
    '可使用组件设计器修改它。
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.插件 = Me.Factory.CreateRibbonTab
        Me.文档处理插件 = Me.Factory.CreateRibbonGroup
        Me.启动纠错或总结 = Me.Factory.CreateRibbonButton
        Me.插件.SuspendLayout()
        Me.文档处理插件.SuspendLayout()
        Me.SuspendLayout()
        '
        '插件
        '
        Me.插件.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.插件.Groups.Add(Me.文档处理插件)
        Me.插件.Label = "TabAddIns"
        Me.插件.Name = "插件"
        '
        '文档处理插件
        '
        Me.文档处理插件.Items.Add(Me.启动纠错或总结)
        Me.文档处理插件.Label = "文档处理插件"
        Me.文档处理插件.Name = "文档处理插件"
        '
        '启动纠错或总结
        '
        Me.启动纠错或总结.Label = "启动纠错或总结"
        Me.启动纠错或总结.Name = "启动纠错或总结"
        Me.启动纠错或总结.ShowImage = True
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Word.Document"
        Me.Tabs.Add(Me.插件)
        Me.插件.ResumeLayout(False)
        Me.插件.PerformLayout()
        Me.文档处理插件.ResumeLayout(False)
        Me.文档处理插件.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents 插件 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents 文档处理插件 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents 启动纠错或总结 As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
