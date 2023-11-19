Imports System.Text.RegularExpressions
Module Module1
    Public cont, aux_id As Integer
    Public diretorio, sql, aux_cpf, aux_data, resp, tipo, aux_rec, bloqueado, ativo, nome_func As String
    Public db As New ADODB.Connection
    Public rs As New ADODB.Recordset
    Public dir_banco = Application.StartupPath & "\banco\petshop_db.mdb"

    Sub conectar_banco()
        Try
            db = CreateObject("ADODB.Connection")
            db.Open("Provider=Microsoft.JET.OLEDB.4.0; Data Source=" & dir_banco)
            'MsgBox("Conectado!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        Catch ex As Exception
            MsgBox("Erro ao conectar com o banco!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "AVISO")
        End Try
    End Sub

    Sub carregar_dados_login()
        Try
            sql = "select * from tb_funcionario order by Codigo_funcionario asc"
            rs = db.Execute(sql)
            With Form1.dgv_funcionario
                cont = 1
                .Rows.Clear()
                Do While rs.EOF = False
                    .Rows.Add(rs.Fields(1).Value, rs.Fields(0).Value, rs.Fields(2).Value)
                    rs.MoveNext()
                    cont = cont + 1
                Loop
            End With
        Catch ex As Exception
            MsgBox("ERRO", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "ERRO!")
        End Try
    End Sub

    Sub carregar_dados_pets()
        Try
            sql = "SELECT * FROM tb_pet WHERE FK_CPF='" & aux_cpf & "' ORDER BY Nome ASC"
            rs = db.Execute(sql)

            Form1.dgv_pets.Rows.Clear()

            Do While Not rs.EOF
                Form1.dgv_pets.Rows.Add(rs.Fields("FK_CPF").Value, rs.Fields("Nome").Value, rs.Fields("Peso").Value, Nothing, Nothing)
                rs.MoveNext()
            Loop
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao carregar os dados dos pets: " & ex.Message, MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Public Sub AdicionarMascaraCPF(textbox As TextBox)
        ' Adicionar o evento TextChanged para formatar o CPF
        AddHandler textbox.TextChanged, AddressOf TextBox_CPF_TextChanged
    End Sub

    Private Sub TextBox_CPF_TextChanged(sender As Object, e As EventArgs)
        Dim textbox As TextBox = DirectCast(sender, TextBox)

        ' Remover formatação atual do CPF
        Dim cpfSemFormatacao As String = Regex.Replace(textbox.Text, "[^\d]", "")

        ' Verificar se o CPF possui 11 dígitos
        If cpfSemFormatacao.Length = 11 Then
            ' Aplicar a máscara de CPF
            Dim cpfFormatado As String = Regex.Replace(cpfSemFormatacao, "(^\d{3})(\d{3})(\d{3})(\d{2}$)", "$1.$2.$3-$4")
            textbox.Text = cpfFormatado
        End If
    End Sub

    Sub carregar_cmb_funcionarios()
        Try
            sql = "select * from tb_funcionario order by Nome asc"
            rs = db.Execute(sql)
            With Form1.cmb_funcionario
                cont = 1
                Do While rs.EOF = False
                    Form1.cmb_funcionario.Items.Add(rs.Fields("Nome").Value)
                    rs.MoveNext()
                    cont = cont + 1
                Loop
            End With
        Catch ex As Exception
            MsgBox("ERRO", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "ERRO!")
        End Try
    End Sub

    Sub relat_servicos()
        Try
            sql = "Select * from tb_servico order by Data desc"
            rs = db.Execute(sql)

            Form1.dgv_relat.Rows.Clear()

            Do While Not rs.EOF
                Form1.dgv_relat.Rows.Add(rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, Nothing)
                rs.MoveNext()
            Loop
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao carregar os dados dos pets: " & ex.Message, MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Sub relat_clientes()
        Try
            sql = "SELECT * FROM tb_servico where FK_CPF_Cliente = '" & aux_cpf & "' "
            rs = db.Execute(sql)

            Form1.dgv_relat.Rows.Clear()

            Do While Not rs.EOF
                Form1.dgv_relat.Rows.Add(rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(4).Value, rs.Fields(5).Value)
                rs.MoveNext()
            Loop
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao carregar os dados", MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Sub relat_pets()
        Try
            sql = "SELECT * FROM tb_pet order by FK_CPF asc"
            rs = db.Execute(sql)

            Form1.dgv_relat.Rows.Clear()

            Do While Not rs.EOF
                Form1.dgv_relat.Rows.Add(rs.Fields(1).Value, rs.Fields(2).Value, rs.Fields(3).Value, rs.Fields(5).Value, Nothing)
                rs.MoveNext()
            Loop
        Catch ex As Exception
            MsgBox("Ocorreu um erro ao carregar os dados dos pets: " & ex.Message, MsgBoxStyle.Critical, "Erro")
        End Try
    End Sub

    Sub limpar_cmb_func()
        Try
            Form1.cmb_funcionario.Items.Clear()
        Catch ex As Exception
            MsgBox("ERRO", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "ERRO!")
        End Try
    End Sub

End Module
