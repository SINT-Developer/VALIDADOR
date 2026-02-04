=== PRE-VALIDACAO ===
APROVADO COM ADVERTENCIAS

=== IMPORTACAO ===
Conectado ao SQL Server.
Planilha: DADOS PRODUTOS CRISTAIS LABONE - Copia.XLSX
Limpando tabelas com FK dependente...
  OK: RestricaoCondicaoPagamento limpa
  OK: Representantes_X_AspNetUsers limpa
Executando ExcluiPedidosCadastrosConfiguracao...
  OK: ExcluiPedidosCadastrosConfiguracao
Todos os dados foram excluidos.
--- Importando EMPRESA ---
  OK: ImportaConfiguracaoEmpresa_PKUnica1
  Registrando 82 codigos...
  Importando 82 parametros...
  Executando ConfigurarEmpresa...
  ConfigurarEmpresa: Configuração efetuada com sucesso.
  EMPRESA: 82 ok, 0 erros | 82 linhas em 0.3s
--- Importando FILIAL ---
  OK: ImportaFilial_PKUnica1
  FILIAL: 1 ok, 0 erros | 1 linhas em 0.0s
--- Importando REPR ---
  OK: ImportaRepresentante_PKUnica1
  REPR: 47 ok, 0 erros | 47 linhas em 0.2s
--- Importando PAGTO ---
  OK: ImportaCondPagamento_PKUnica1
  PAGTO: 8 ok, 0 erros | 8 linhas em 0.2s
--- Importando TRANSP ---
  OK: ImportaTransportadora_PKUnica1
  TRANSP: 1 ok, 0 erros | 1 linhas em 0.1s
--- Importando ESTADOS ---
  OK: ImportaEstado_PKUnica1
  ESTADOS: 28 ok, 0 erros | 28 linhas em 0.1s
--- Importando FAMILIAS ---
  Nenhuma linha valida em FAMILIAS
--- Importando ESTILOS ---
  Nenhuma linha valida em ESTILOS
--- Importando CLIENTES ---
  OK: ImportaCliente_PKUnica1
  ERRO CLIENTES linha 3582 PK=[8471]: <<Não foi possível importar Cliente devido a erros de banco de dados.\r\n (A instrução INSERT conflitou com a restrição do CHECK "CNPJCPFNull_". O conflito ocorreu no banco de dados "SRPP", tabela "dbo.Cliente", column \'CNPJCPF\'.)>> (50000) (SQLExecDirectW)')
  ERRO CLIENTES linha 5257 PK=[11273]: <<Não foi possível importar Cliente devido a erros de banco de dados.\r\n (A instrução INSERT conflitou com a restrição do CHECK "CNPJCPFNull_". O conflito ocorreu no banco de dados "SRPP", tabela "dbo.Cliente", column \'CNPJCPF\'.)>> (50000) (SQLExecDirectW)')
  ERRO CLIENTES linha 5262 PK=[11278]: <<Não foi possível importar Cliente devido a erros de banco de dados.\r\n (A instrução INSERT conflitou com a restrição do CHECK "CNPJCPFNull_". O conflito ocorreu no banco de dados "SRPP", tabela "dbo.Cliente", column \'CNPJCPF\'.)>> (50000) (SQLExecDirectW)')
  CLIENTES: 7065 ok, 3 erros | 7068 linhas em 51.0s
--- Importando PAGTOFILIAL ---
  Nenhuma linha valida em PAGTOFILIAL
--- Importando PRODUTOS (modo lote) ---
  73618 linhas inseridas no staging.
  PRODUTOS: 73618 ok, 0 erros | 73618 linhas em 47.5s
Tempo total: 2m 4s

=== RESUMO ===
  EMPRESA: 82 importados, OK
  FILIAL: 1 importados, OK
  REPR: 47 importados, OK
  PAGTO: 8 importados, OK
  TRANSP: 1 importados, OK
  ESTADOS: 28 importados, OK
  FAMILIAS: 0 importados, OK
  ESTILOS: 0 importados, OK
  CLIENTES: 7065 importados, 3 ERROS
  PAGTOFILIAL: 0 importados, OK
  PRODUTOS: 73618 importados, OK

TOTAL: 80850 importados, 3 erros

