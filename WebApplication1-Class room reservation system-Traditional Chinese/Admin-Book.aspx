<%@ Page CodeBehind="Admin-Book.aspx.vb" Language="vb" AutoEventWireup="True" Inherits="WebApplication1.Admin_OrderList" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<!-- #include File="OleDbFunction.inc" -->
		<meta content="True" name="vs_snapToGrid">
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 10.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<Script Language="VB" Runat="Server">

  Sub BindList()
      Dim strSQL As String = "Select * From [Book]"
      myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "[Book]")
      myDataList.DataBind()
  End Sub
 
  Sub Page_Load(sender As Object, e As Eventargs)
      If Session("LV") <> "Admin" Then
         Page.Response.Redirect ("./Main.aspx")
         Exit Sub
      End If
      If Not IsPostBack Then BindList()
  End Sub

  Sub DataList_EditCommand(sender As Object, e As DataListCommandEventArgs)
      myDataList.EditItemIndex = e.Item.ItemIndex
      BindList()
  End Sub

  Sub DataList_CancelCommand(sender As Object, e As DataListCommandEventArgs)
      myDataList.EditItemIndex = -1
      BindList()
  End Sub

  Sub ExecuteSQL(strSQL As String)
      Dim objConn As New OleDbConnection()
      objConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & Server.MapPath("./DB/order.mdb")
      objConn.Open()
      Dim objCmd As New OleDbCommand(strSQL, objConn)
      objCmd.ExecuteNonQuery
      objConn.Close()
  End Sub

  Sub DataList_DeleteCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strSQL As String = "Delete From [Book] Where [" & _
        myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub DataList_UpdateCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strBook As String = Ctype(e.Item.FindControl("txtBook"), TextBox).Text
    Dim UpdateDate As DateTime = DateTime.Now.Date()
    Dim strSQL As String = "Update [Book] Set [B-Name]='" & strBook & "' " & _
                           "Where [" & myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub Book_Insert (sender As Object, e As Eventargs)
    If txtBook_Insert.Text = "" Then Exit Sub
    Dim strSQL As String = "Insert Into [Book] ([B-Name]) Values ('" & txtBook_Insert.Text & "')"
    ExecuteSQL(strSQL)
    txtBook_Insert.Text = ""
    BindList()
  End Sub
		</Script>
	</HEAD>
	<Body MS_POSITIONING="GridLayout" background="./image/admin-bg.gif">
		<form id="Form1" method="post" runat="server">
			<Asp:DataList Runat="Server" Id="myDataList" CellPadding="3" Width="489px" HorizontalAlign="Center"
				OnEditCommand="DataList_EditCommand" OnUpdateCommand="DataList_UpdateCommand" OnDeleteCommand="DataList_DeleteCommand"
				OnCancelCommand="DataList_CancelCommand" DataKeyField="B-ID" ExtractTemplateRows="True" BorderColor="#DEBA84"
				GridLines="Both" RepeatLayout="Flow" style="Z-INDEX: 101; LEFT: 40px; POSITION: absolute; TOP: 108px"
				BackColor="#DEBA84" BorderWidth="1px" BorderStyle="None" CellSpacing="2">
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#738A9C"></SelectedItemStyle>
				<HeaderTemplate>
					<Asp:Table Runat="Server" Width="400">
						<Asp:TableRow Runat="Server">
							<Asp:TableCell Runat="Server" Width="70">
                自動編號
              </Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="200">
                使用內容
              </Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="130">
                功能
              </Asp:TableCell>
						</Asp:TableRow>
					</Asp:Table>
				</HeaderTemplate>
				<EditItemStyle BackColor="Lavender"></EditItemStyle>
				<ItemStyle ForeColor="#8C4510" BackColor="#FFF7E7"></ItemStyle>
				<ItemTemplate>
					<Asp:Table Runat="Server" Width="400">
						<Asp:TableRow Runat="Server" HorizontalAlign="Center">
							<Asp:TableCell Runat="Server" Width="70">
								<%# Container.DataItem("B-ID") %>
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="200">
								<%# Container.DataItem("B-Name") %>
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="130">
								<Asp:LinkButton Runat="Server" Text="編輯" CommandName="Edit" />
							</Asp:TableCell>
						</Asp:TableRow>
					</Asp:Table>
				</ItemTemplate>
				<FooterStyle ForeColor="#8C4510" BackColor="#F7DFB5"></FooterStyle>
				<HeaderStyle Font-Bold="True" HorizontalAlign="Center" ForeColor="White" BackColor="#A55129"></HeaderStyle>
				<EditItemTemplate>
					<Asp:Table Runat="Server" Width="400">
						<Asp:TableRow Runat="Server" HorizontalAlign="Center">
							<Asp:TableCell Runat="Server" Width="70">
								<%# Container.DataItem("B-ID") %>
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="200">
								<Asp:TextBox Runat="Server" Id="txtBook" Width="200" Text='<%# Container.DataItem("B-Name") %>' />
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="130">
								<Asp:LinkButton Runat="Server" Width="18" Text="更新" CommandName="Update" />
								<Asp:LinkButton Runat="Server" Width="18" Text="刪除" CommandName="Delete" />
								<Asp:LinkButton Runat="Server" Width="18" Text="取消" CommandName="Cancel" />
							</Asp:TableCell>
						</Asp:TableRow>
					</Asp:Table>
				</EditItemTemplate>
			</Asp:DataList>
			<asp:Label id="Label2" style="Z-INDEX: 102; LEFT: 80px; POSITION: absolute; TOP: 72px" Width="76px"
				runat="server" Height="24px">使用內容：</asp:Label>
			<asp:TextBox id="txtBook_Insert" style="Z-INDEX: 103; LEFT: 164px; POSITION: absolute; TOP: 72px"
				Width="232px" runat="server"></asp:TextBox>
			<asp:Button id="btnBook_Insert" style="Z-INDEX: 104; LEFT: 404px; POSITION: absolute; TOP: 72px"
				Width="72px" runat="server" Height="24px" Text="新增" OnClick="Book_Insert"></asp:Button>
			<asp:Label id="Label1" style="Z-INDEX: 105; LEFT: 208px; POSITION: absolute; TOP: 12px" Width="136px"
				runat="server" Font-Bold="True" Height="36px" Font-Size="X-Large">使用內容</asp:Label>
		</form>
	</Body>
</HTML>
