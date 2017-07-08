<%@ Page Language="vb" AutoEventWireup="True" Codebehind="Admin.aspx.vb" Inherits="WebApplication1.Admin"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title></title>
		<SCRIPT language="VB" Runat="Server">

				Sub Page_Load( sender As Object, e As Eventargs )
        If Session("LV") <> "Admin" Then
           Page.Response.Redirect ("./Main.aspx")
           Exit Sub
        End If
				End Sub

				Sub hypAdmin1( sender As Object, e As Eventargs )
        HyperLink1.NavigateUrl="Admin-PhotoStudio.aspx?ID=" & Request.QueryString("ID")
    End Sub

				Sub hypAdmin2( sender As Object, e As Eventargs )
        HyperLink2.NavigateUrl="Admin-Book.aspx?ID=" & Request.QueryString("ID")
    End Sub

				Sub hypAdmin3( sender As Object, e As Eventargs )
        HyperLink3.NavigateUrl="Admin-Class.aspx?ID=" & Request.QueryString("ID")
    End Sub

				Sub hypAdmin4( sender As Object, e As Eventargs )
        HyperLink4.NavigateUrl="Admin-Teacher.aspx?ID=" & Request.QueryString("ID")
    End Sub

				Sub hypAdmin6( sender As Object, e As Eventargs )
        HyperLink6.NavigateUrl="Admin-OrderList.aspx?ID=" & Request.QueryString("ID")
    End Sub

				Sub hypAdmin7( sender As Object, e As Eventargs )
        HyperLink7.NavigateUrl="Admin-Item.aspx?ID=" & Request.QueryString("ID")
    End Sub

		</SCRIPT>
		<base target="main">
		<SCRIPT language="VB" Runat="Server">
		</SCRIPT>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
	</HEAD>
	<body MS_POSITIONING="GridLayout" background="./image/admin-bg.gif">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:Label id="Label1" style="Z-INDEX: 100; LEFT: 8px; POSITION: absolute; TOP: 20px" runat="server"
					Width="100px" Height="16px">管理者工具列</asp:Label>
				<asp:HyperLink id="HyperLink7" style="Z-INDEX: 108; LEFT: 12px; POSITION: absolute; TOP: 196px"
					runat="server" Height="20px" Width="88px" OnLoad="hypAdmin7" Target="main">使用器材</asp:HyperLink>
				<asp:HyperLink id="HyperLink1" style="Z-INDEX: 101; LEFT: 12px; POSITION: absolute; TOP: 68px"
					runat="server" Width="88px" Height="20px" Target="main" OnLoad="hypAdmin1">專業教室</asp:HyperLink>
				<asp:HyperLink id="HyperLink2" style="Z-INDEX: 102; LEFT: 12px; POSITION: absolute; TOP: 100px"
					runat="server" Height="20px" Width="88px" Target="main" OnLoad="hypAdmin2">使用內容</asp:HyperLink>
				<asp:HyperLink id="HyperLink3" style="Z-INDEX: 103; LEFT: 12px; POSITION: absolute; TOP: 132px"
					runat="server" Height="20px" Width="88px" Target="main" OnLoad="hypAdmin3">借用單位</asp:HyperLink>
				<asp:HyperLink id="HyperLink4" style="Z-INDEX: 105; LEFT: 12px; POSITION: absolute; TOP: 164px"
					runat="server" Height="20px" Width="88px" Target="main" OnLoad="hypAdmin4">教　　師</asp:HyperLink>
				<asp:HyperLink id="Hyperlink5" style="Z-INDEX: 106; LEFT: 12px; POSITION: absolute; TOP: 260px"
					runat="server" Height="20px" Width="88px" NavigateUrl="Main.aspx" Target="_parent">返回首頁</asp:HyperLink>
				<asp:HyperLink id="HyperLink6" style="Z-INDEX: 107; LEFT: 12px; POSITION: absolute; TOP: 228px"
					runat="server" Height="20px" Width="88px" OnLoad="hypAdmin6" Target="main">預約紀錄</asp:HyperLink>
			</FONT>
		</form>
	</body>
</HTML>
