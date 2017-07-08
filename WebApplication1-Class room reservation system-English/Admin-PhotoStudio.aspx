<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- #include File="OleDbFunction.inc" -->
<HTML>
	<head>
		<meta content="True" name="vs_snapToGrid">
		<meta content="True" name="vs_showGrid">
		<meta content="Microsoft Visual Studio .NET 10.0" name="GENERATOR">
		<meta content="Visual Basic .NET 10.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<Script Language="VB" Runat="Server">

  Sub BindList()
      Dim strSQL As String = "Select * From [PhotoStudio]"
      myDataList.DataSource = CreateDataSet(strSQL, "./DB/order.mdb", "[PhotoStudio]")
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
    Dim strSQL As String = "Delete From [PhotoStudio] Where [" & _
        myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub DataList_UpdateCommand(sender As Object, e As DataListCommandEventArgs)
    Dim strPAdds As String = Ctype(e.Item.FindControl("txtPAdds"), TextBox).Text
    Dim UpdateDate As DateTime = DateTime.Now.Date()
    Dim strSQL As String = "Update [PhotoStudio] Set [P-Adds]='" & strPAdds & "' " & _
                           "Where [" & myDataList.DataKeyField & "]=" & myDataList.DataKeys(e.Item.ItemIndex)
    ExecuteSQL(strSQL)
    myDataList.EditItemIndex = -1
    BindList()
  End Sub

  Sub PAdds_Insert (sender As Object, e As Eventargs)
    If txtPAdds_Insert.Text = "" Then Exit Sub
    Dim strSQL As String = "Insert Into [PhotoStudio] ([P-Adds]) Values ('" & txtPAdds_Insert.Text & "')"
    ExecuteSQL(strSQL)
    txtPAdds_Insert.Text = ""
    BindList()
  End Sub
		</Script>
	</head>
	<Body MS_POSITIONING="GridLayout" background="./image/admin-bg.gif">
		<form id="Form1" method="post" runat="server">
			<Asp:DataList Runat="Server" Id="myDataList" CellPadding="3" Width="489px" HorizontalAlign="Center"
				OnEditCommand="DataList_EditCommand" OnUpdateCommand="DataList_UpdateCommand" OnDeleteCommand="DataList_DeleteCommand"
				OnCancelCommand="DataList_CancelCommand" DataKeyField="P-ID" ExtractTemplateRows="True" BorderColor="#DEBA84"
				GridLines="Both" RepeatLayout="Flow" style="Z-INDEX: 101; LEFT: 40px; POSITION: absolute; TOP: 108px"
				BackColor="#DEBA84" BorderWidth="1px" BorderStyle="None" CellSpacing="2">
				<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#738A9C"></SelectedItemStyle>
				<HeaderTemplate>
					<Asp:Table Runat="Server" Width="400">
						<Asp:TableRow Runat="Server">
							<Asp:TableCell Runat="Server" Width="70">
                ID
              </Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="200">
                Place
              </Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="130">
                Function
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
								<%# Container.DataItem("P-ID") %>
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="200">
								<%# Container.DataItem("P-Adds") %>
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="130">
								<Asp:LinkButton Runat="Server" Text="Edit" CommandName="Edit" />
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
								<%# Container.DataItem("P-ID") %>
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="200">
								<Asp:TextBox Runat="Server" Id="txtPAdds" Width="200" Text='<%# Container.DataItem("P-Adds") %>' />
							</Asp:TableCell>
							<Asp:TableCell Runat="Server" Width="130">
								<Asp:LinkButton Runat="Server" Width="18" Text="Update" CommandName="Update" />
								, <Asp:LinkButton Runat="Server" Width="18" Text="Delete" CommandName="Delete" />
								, <Asp:LinkButton Runat="Server" Width="18" Text="Cancel" CommandName="Cancel" />
							</Asp:TableCell>
						</Asp:TableRow>
					</Asp:Table>
				</EditItemTemplate>
			</Asp:DataList>
			<asp:Label id="Label2" style="Z-INDEX: 102; LEFT: 80px; POSITION: absolute; TOP: 72px" Width="76px"
				runat="server" Height="24px">Place</asp:Label>
			<asp:TextBox id="txtPAdds_Insert" style="Z-INDEX: 103; LEFT: 164px; POSITION: absolute; TOP: 72px"
				Width="232px" runat="server"></asp:TextBox>
			<asp:Button id="btnPAdds_Insert" style="Z-INDEX: 104; LEFT: 404px; POSITION: absolute; TOP: 72px"
				Width="72px" runat="server" Height="24px" Text="Add" OnClick="PAdds_Insert"></asp:Button>
			<asp:Label id="Label1" style="Z-INDEX: 105; LEFT: 208px; POSITION: absolute; TOP: 12px" Width="136px"
				runat="server" Font-Bold="True" Height="36px" Font-Size="X-Large">Professional Classroom</asp:Label>
		</form>
	</Body>
</HTML>
