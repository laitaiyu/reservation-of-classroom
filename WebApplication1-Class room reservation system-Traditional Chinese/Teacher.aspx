<%@ Page Language="vb" AutoEventWireup="True" Codebehind="Teacher.aspx.vb" Inherits="WebApplication1.Teacher_Order"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title></title>
		<SCRIPT language="VB" Runat="Server">

				Sub Page_Load( sender As Object, e As Eventargs )
        If Session("LV") <> "Teacher" Then
           Page.Response.Redirect ("./Main.aspx")
           Exit Sub
        End If
				End Sub

				Sub hypTeacher1( sender As Object, e As Eventargs )
        HyperLink1.NavigateUrl="Teacher-Order.aspx?ID=" & Request.QueryString("ID")
    End Sub

				Sub hypTeacher2( sender As Object, e As Eventargs )
        HyperLink2.NavigateUrl="Teacher-History.aspx?ID=" & Request.QueryString("ID")
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
	<body MS_POSITIONING="GridLayout" BACKGROUND="./image/Teacher-bg.gif">
		<form id="Form1" method="post" runat="server">
			<FONT face="新細明體">
				<asp:Label id="Label1" style="Z-INDEX: 100; LEFT: 16px; POSITION: absolute; TOP: 20px" runat="server"
					Width="84px" Height="20px">教師工具列</asp:Label>
				<asp:HyperLink id="HyperLink1" style="Z-INDEX: 101; LEFT: 8px; POSITION: absolute; TOP: 68px" runat="server"
					Width="100px" Height="20px" Target="main" OnLoad="hypTeacher1">預約專業教室</asp:HyperLink>
				<asp:HyperLink id="HyperLink2" style="Z-INDEX: 102; LEFT: 8px; POSITION: absolute; TOP: 100px"
					runat="server" Height="20px" Width="88px" Target="main" OnLoad="hypTeacher2">預約紀錄表</asp:HyperLink>
				<asp:HyperLink id="Hyperlink5" style="Z-INDEX: 105; LEFT: 8px; POSITION: absolute; TOP: 132px"
					runat="server" Height="20px" Width="88px" NavigateUrl="Main.aspx" Target="_parent">返回首頁</asp:HyperLink>
			</FONT>
		</form>
	</body>
</HTML>
