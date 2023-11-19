Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        conectar_banco()
        carregar_dados_login()
        limpar_cmb_func()
        carregar_cmb_funcionarios()

        cmb_funcao.Items.Add("Atendente")
        cmb_funcao.Items.Add("Veterinario")

        cmb_servicos.Items.Add("Banho")
        cmb_servicos.Items.Add("Banho e Tosa")
        cmb_servicos.Items.Add("Veterinário")

        cmb_relat.Items.Add("Clientes")
        cmb_relat.Items.Add("Serviços")
        cmb_relat.Items.Add("Pets")

        TabControl1.TabPages.Remove(clientes)
        TabControl1.TabPages.Remove(pets)
        TabControl1.TabPages.Remove(serviço)
        TabControl1.TabPages.Remove(funcionarios)
        TabControl1.TabPages.Remove(relatorio)
        TabControl1.TabPages.Remove(cad_pets)

        carregar_dados_pets() ' Carrega os dados iniciais dos pets

        AdicionarMascaraCPF(txt_busca_cliente)
        AdicionarMascaraCPF(txt_cpf_dono)
        AdicionarMascaraCPF(txt_cpf_cliente)
        AdicionarMascaraCPF(txt_cpf_relat)

        AddHandler dgv_pets.CellContentClick, AddressOf dgv_pets_CellContentClick

    End Sub


    Private Sub btn_cad_func_Click(sender As Object, e As EventArgs) Handles btn_cad_func.Click
        ' Obter os valores inseridos nos campos
        Dim codigoFuncionario As String = txt_cod_func.Text
        Dim nomeFuncionario As String = txt_novo_func.Text
        Dim funcao As String = cmb_funcao.Text
        Dim senha As String = txt_senha_func.Text

        Try
            ' Verificar se o código do funcionário já existe no banco de dados
            sql = "SELECT * FROM tb_funcionario WHERE Codigo_funcionario = '" & codigoFuncionario & "'"
            rs = db.Execute(sql)
            If Not rs.EOF Then
                MsgBox("O código de funcionário já existe!")
            Else
                ' Realizar a inserção do novo funcionário na tabela
                sql = "INSERT INTO tb_funcionario (Codigo_funcionario, Nome, Funcao, Senha) VALUES ('" & codigoFuncionario & "', " &
                "'" & nomeFuncionario & "', '" & funcao & "', '" & senha & "')"
                rs = db.Execute(UCase(sql))
                MsgBox("Dados Gravados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

                ' Atualizar a DataGridView
                carregar_dados_login()
                limpar_cmb_func()
                carregar_cmb_funcionarios()
                'cmb_funcionario.Items.Add(nomeFuncionario)
            End If
        Catch ex As Exception
            MsgBox("Erro ao cadastrar o funcionário: " & ex.Message)
        End Try
    End Sub


    Private Sub btn_cadastrar_Click(sender As Object, e As EventArgs) Handles btn_cadastrar.Click
        Try
            ' Verificar se todos os campos foram preenchidos
            If String.IsNullOrEmpty(txt_nome_cliente.Text) OrElse
                String.IsNullOrEmpty(txt_cpf_cliente.Text) OrElse
                String.IsNullOrEmpty(txt_tel_cliente.Text) Then
                MsgBox("Por favor, preencha todos os campos.")
                Return
            End If

            ' Realizar a lógica de cadastro do cliente
            Dim nome As String = txt_nome_cliente.Text
            Dim cpf As String = txt_cpf_cliente.Text
            Dim telefone As String = txt_tel_cliente.Text

            ' Estabelecer conexão com o banco de dados
            conectar_banco()

            ' Preparar o comando SQL
            sql = "INSERT INTO tb_cliente (Nome, CPF, Telefone) VALUES ('" & nome & "', '" & cpf & "', '" & telefone & "')"
            rs = db.Execute(sql)

            ' Exibir uma mensagem de sucesso
            MsgBox("Cliente cadastrado com sucesso!")

            ' Limpar os campos de entrada após o cadastro
            txt_nome_cliente.Text = ""
            txt_cpf_cliente.Text = ""
            txt_tel_cliente.Text = ""

            ' Recarregar os dados no DataGridView
            carregar_dados_login()
            TabControl1.SelectTab(2)
        Catch ex As Exception
            MsgBox("Erro ao cadastrar cliente: " & ex.Message)
        End Try
    End Sub

    Private Sub btn_entrar_Click(sender As Object, e As EventArgs)
        Try
            Dim cpf As String = txt_busca_cliente.Text

            ' Verifique se o CPF do cliente foi fornecido
            If String.IsNullOrEmpty(cpf) Then
                MsgBox("Digite o CPF do cliente!")
                Return
            End If

            ' Estabelecer conexão com o banco de dados
            conectar_banco()

            ' Preparar o comando SQL
            sql = "SELECT * FROM tb_cliente WHERE CPF = '" & cpf & "'"
            rs = db.Execute(sql)
            aux_cpf = cpf

            If rs.EOF Then
                MsgBox("Nenhum cliente encontrado com o CPF informado.")
            Else
                ' Exibir os dados do cliente encontrado
                Dim nome As String = rs.Fields("Nome").Value
                Dim telefone As String = rs.Fields("Telefone").Value

                MsgBox("Cliente encontrado:" & vbCrLf &
                   "Nome: " & nome & vbCrLf &
                   "CPF: " & cpf & vbCrLf &
                   "Telefone: " & telefone)
                TabControl1.SelectTab(pets)

                ' Chame a sub-rotina para carregar os dados dos pets relacionados ao cliente
                carregar_dados_pets()
            End If
        Catch ex As Exception
            MsgBox("Erro ao buscar cliente: " & ex.Message)
        End Try
    End Sub


    Private Sub btn_login_Click(sender As Object, e As EventArgs) Handles btn_login.Click
        ' Obter as credenciais de login inseridas pelo usuário
        Dim usuario As String = txt_usuario.Text
        Dim senha As String = txt_senha.Text

        ' Verificar as credenciais no banco de dados
        Try
            sql = "SELECT COUNT(*) FROM tb_funcionario WHERE Nome = ? AND Senha = ?"
            Dim cmd As New ADODB.Command()
            cmd.ActiveConnection = db
            cmd.CommandText = sql
            cmd.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmd.Parameters.Append(cmd.CreateParameter("nome", ADODB.DataTypeEnum.adVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 50, usuario))
            cmd.Parameters.Append(cmd.CreateParameter("senha", ADODB.DataTypeEnum.adVarWChar, ADODB.ParameterDirectionEnum.adParamInput, 50, senha))

            rs = cmd.Execute()
            Dim result As Integer = CInt(rs.Fields(0).Value)
            If txt_usuario.Text = "admin" And txt_senha.Text = "admin" Then
                MsgBox("Login bem-sucedido!")
                TabControl1.TabPages.Add(clientes)
                TabControl1.TabPages.Add(cad_pets)
                TabControl1.TabPages.Add(pets)
                TabControl1.TabPages.Add(serviço)
                TabControl1.TabPages.Add(funcionarios)
                TabControl1.TabPages.Add(relatorio)
                TabControl1.SelectTab(clientes)
            ElseIf result > 0 Then
                MsgBox("Login bem-sucedido!")
                TabControl1.TabPages.Add(clientes)
                TabControl1.TabPages.Add(cad_pets)
                TabControl1.TabPages.Add(pets)
                TabControl1.TabPages.Add(serviço)
                TabControl1.TabPages.Add(relatorio)
                TabControl1.SelectTab(clientes)
                ' Aqui você pode redirecionar para a próxima tela ou executar outras ações desejadas após o login bem-sucedido.
            Else
                MsgBox("Credenciais inválidas. Tente novamente.")
            End If
        Catch ex As Exception
            MsgBox("Erro ao realizar o login: " & ex.Message)
        End Try
    End Sub

    Private Sub btn_cadastrar_pet_Click(sender As Object, e As EventArgs)
        Try
            ' Verifique se os campos obrigatórios foram preenchidos
            If String.IsNullOrEmpty(txt_nome_pet.Text) OrElse String.IsNullOrEmpty(txt_peso.Text) OrElse String.IsNullOrEmpty(txt_cpf_dono.Text) Then
                MsgBox("Preencha todos os campos obrigatórios!")
                Return
            End If

            ' Insira os dados na tabela tb_pet
            sql = "INSERT INTO tb_pet (Nome, Peso, Raca, Descricao, FK_CPF) VALUES ('" & txt_nome_pet.Text & "', " &
              "'" & txt_peso.Text & "', '" & txt_raca.Text & "', '" & txt_descricao.Text & "', '" & txt_cpf_dono.Text & "')"
            db.Execute(sql)

            ' Exiba uma mensagem de sucesso
            MsgBox("Dados Gravados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

            ' Limpe os campos após o cadastro
            txt_nome_pet.Text = ""
            txt_peso.Text = ""
            txt_raca.Text = ""
            txt_descricao.Text = ""
            txt_cpf_dono.Text = ""

            ' Atualize o DataGridView
            carregar_dados_pets()
        Catch ex As Exception
            MsgBox("Erro ao inserir dados na tabela tb_pet: " & ex.Message)
        End Try
    End Sub
    Private Sub dgv_pets_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        ' Verifique se o botão de exclusão foi clicado
        If e.ColumnIndex = dgv_pets.Columns("btnExcluir").Index AndAlso e.RowIndex >= 0 Then
            ' Obtenha o nome do pet a ser excluído
            Dim nomePet As String = dgv_pets.Rows(e.RowIndex).Cells("Nome").Value.ToString()

            ' Confirme se o usuário deseja excluir o pet
            Dim result As DialogResult = MessageBox.Show("Tem certeza que deseja excluir o pet " & nomePet & "?", "Confirmar exclusão", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Try
                    ' Construa a instrução SQL para excluir o pet com base no nome
                    Dim sqlExclusao As String = "DELETE FROM tb_pet WHERE Nome = '" & nomePet & "'"

                    ' Execute a instrução de exclusão
                    db.Execute(sqlExclusao)

                    ' Após a exclusão, atualize a exibição da DataGridView
                    carregar_dados_pets()
                Catch ex As Exception
                    MsgBox("Erro ao excluir o pet: " & ex.Message, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Erro!")
                End Try
            End If
        ElseIf e.ColumnIndex = dgv_pets.Columns("btn_marcar").Index AndAlso e.RowIndex >= 0 Then
            TabControl1.SelectTab(serviço)
            lbl_nome_pet.Text = dgv_pets.CurrentRow.Cells(1).Value
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim cpf As String = txt_busca_cliente.Text

            ' Verifique se o CPF do cliente foi fornecido
            If String.IsNullOrEmpty(cpf) Then
                MsgBox("Digite o CPF do cliente!")
                Return
            End If

            ' Estabelecer conexão com o banco de dados
            conectar_banco()

            ' Preparar o comando SQL
            sql = "SELECT * FROM tb_cliente WHERE CPF = '" & cpf & "'"
            rs = db.Execute(sql)
            aux_cpf = cpf

            If rs.EOF Then
                MsgBox("Nenhum cliente encontrado com o CPF informado.")
            Else
                ' Exibir os dados do cliente encontrado
                Dim nome As String = rs.Fields("Nome").Value
                Dim telefone As String = rs.Fields("Telefone").Value

                MsgBox("Cliente encontrado:" & vbCrLf &
                   "Nome: " & nome & vbCrLf &
                   "CPF: " & cpf & vbCrLf &
                   "Telefone: " & telefone)
                'TabControl1.SelectTab(2)

                ' Chame a sub-rotina para carregar os dados dos pets relacionados ao cliente
                carregar_dados_pets()
            End If
        Catch ex As Exception
            MsgBox("Erro ao buscar cliente: " & ex.Message)
        End Try
    End Sub

    Private Sub btn_cadastrar_pet_Click_1(sender As Object, e As EventArgs) Handles btn_cadastrar_pet.Click
        Try
            ' Verifique se os campos obrigatórios foram preenchidos
            If String.IsNullOrEmpty(txt_nome_pet.Text) OrElse String.IsNullOrEmpty(txt_peso.Text) OrElse String.IsNullOrEmpty(txt_cpf_dono.Text) Then
                MsgBox("Preencha todos os campos obrigatórios!")
                Return
            End If

            ' Insira os dados na tabela tb_pet
            sql = "INSERT INTO tb_pet (Nome, Peso, Raca, Descricao, FK_CPF) VALUES ('" & txt_nome_pet.Text & "', " &
              "'" & txt_peso.Text & "', '" & txt_raca.Text & "', '" & txt_descricao.Text & "', '" & txt_cpf_dono.Text & "')"
            db.Execute(sql)

            ' Exiba uma mensagem de sucesso
            MsgBox("Dados Gravados com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

            ' Limpe os campos após o cadastro
            txt_nome_pet.Text = ""
            txt_peso.Text = ""
            txt_raca.Text = ""
            txt_descricao.Text = ""
            txt_cpf_dono.Text = ""

            ' Atualize o DataGridView
            carregar_dados_pets()
        Catch ex As Exception
            MsgBox("Erro ao inserir dados na tabela tb_pet: " & ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            ' Verifique se os campos obrigatórios foram preenchidos
            If String.IsNullOrEmpty(cmb_servicos.Text) OrElse String.IsNullOrEmpty(cmb_funcionario.Text) OrElse String.IsNullOrEmpty(data.Value) Then
                MsgBox("Selecione todos os campos!")
                Return
            End If

            ' Insira os dados na tabela tb_pet
            sql = "INSERT INTO tb_servico (Data, Tipo, FK_Cod_Funcionario, FK_Nome_Pet, FK_CPF_Cliente) VALUES ('" & data.Value & "', " &
                  "'" & cmb_servicos.Text & "', '" & cmb_funcionario.Text & "', '" & lbl_nome_pet.Text & "', '" & txt_busca_cliente.Text & "')"
            db.Execute(sql)

            ' Exiba uma mensagem de sucesso
            MsgBox("Serviço Marcado com Sucesso!", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "SUCESSO!")

        Catch ex As Exception
            MsgBox("Erro ao marcar serviço")
        End Try
    End Sub

    Private Sub dgv_funcionario_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv_funcionario.CellContentClick
        ' Verifique se o botão de exclusão foi clicado
        If e.ColumnIndex = dgv_funcionario.Columns("btn_excluir").Index AndAlso e.RowIndex >= 0 Then
            ' Obtenha o nome do funcionario a ser excluído
            Dim nomeFunc As String = dgv_funcionario.Rows(e.RowIndex).Cells(0).Value.ToString()

            ' Confirme se o usuário deseja excluir o funcionario
            Dim result As DialogResult = MessageBox.Show("Tem certeza que deseja excluir o funcionário " & nomeFunc & "?", "Confirmar exclusão", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Try
                    ' Construa a instrução SQL para excluir o funcionario com base no nome
                    Dim sqlExclusao As String = "DELETE FROM tb_funcionario WHERE Nome = '" & nomeFunc & "'"

                    ' Execute a instrução de exclusão
                    db.Execute(sqlExclusao)

                    ' Após a exclusão, atualize a exibição da DataGridView
                    carregar_dados_login()
                    limpar_cmb_func()
                    carregar_cmb_funcionarios()
                Catch ex As Exception
                    MsgBox("Erro ao excluir o pet: " & ex.Message, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Erro!")
                End Try
            End If
        End If
    End Sub

    Private Sub btn_gerar_Click(sender As Object, e As EventArgs) Handles btn_gerar.Click
        Dim aux_cpf_relat As String = txt_cpf_relat.Text
        aux_cpf = aux_cpf_relat
        If cmb_relat.Text = "Serviços" Then
            relat_servicos()
        ElseIf cmb_relat.Text = "Clientes" Then
            relat_clientes()
        ElseIf cmb_relat.Text = "Pets" Then
            relat_pets()
        End If
    End Sub

    Private Sub btn_sair_Click(sender As Object, e As EventArgs) Handles btn_sair.Click
        Dim resp As DialogResult = MessageBox.Show("Deseja realmente sair", "Confirmar Saida", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If resp = DialogResult.Yes Then
            TabControl1.TabPages.Remove(clientes)
            TabControl1.TabPages.Remove(pets)
            TabControl1.TabPages.Remove(serviço)
            TabControl1.TabPages.Remove(funcionarios)
            TabControl1.TabPages.Remove(relatorio)
            TabControl1.TabPages.Remove(cad_pets)

            txt_usuario.Text = ""
            txt_senha.Text = ""
        End If
    End Sub
End Class