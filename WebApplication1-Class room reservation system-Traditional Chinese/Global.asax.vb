Imports System.Web
Imports System.Web.SessionState

Public Class [Global]
    Inherits System.Web.HttpApplication

#Region " 元件設計工具產生的程式碼 "

    Public Sub New()
        MyBase.New()

        '此為元件設計工具所需的呼叫。
        InitializeComponent()

        '在 InitializeComponent() 呼叫之後加入所有的初始設定

    End Sub

    '為元件設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為元件設計工具所需的程序
    '您可以使用元件設計工具進行修改，
    '請勿使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 啟動應用程式時引發
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 啟動工作階段時引發
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 於每一個要求開始時引發
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 嘗試驗證使用時引發
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' 發生錯誤時引發
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 於工作階段結束時引發
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 於應用程式結束時引發
    End Sub

End Class
