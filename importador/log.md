Driver ODBC detectado: SQL Server Native Client 10.0

=== IMPORTACAO ===
Conectado ao SQL Server.
Planilha: 2026.02.04 11-04_GEOMÉTRICA IND. COM. E IMP. LTDA._IMPORTAÇÃO - Copia.xlsx
Fazendo backup do banco de dados...
  Destino: c:\Arquivos de programas\Microsoft SQL Server\MSSQL10_50.SQLEXPRESS\MSSQL\DATA\SRPP - GEOMÉTRICA IND. COM. E IMP. LTDA. - 2026-02-05 10h53m29 - Apagar Pedidos e Cadastros e Configuracoes.bak
  Backup concluido com sucesso!
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
  EMPRESA: 82 ok, 0 erros | 82 linhas em 0.2s
--- Importando FILIAL ---
  OK: ImportaFilial_PKUnica1
  FILIAL: 3 ok, 0 erros | 3 linhas em 0.1s
--- Importando REPR ---
  OK: ImportaRepresentante_PKUnica1
  REPR: 126 ok, 0 erros | 126 linhas em 0.4s
--- Importando PAGTO ---
  OK: ImportaCondPagamento_PKUnica1
  PAGTO: 5 ok, 0 erros | 5 linhas em 0.1s
--- Importando TRANSP (modo lote) ---
  93 linhas inseridas no staging.
  TRANSP: 93 ok, 0 erros | 93 linhas em 0.2s
--- Importando ESTADOS ---
  OK: ImportaEstado_PKUnica1
  ESTADOS: 28 ok, 0 erros | 28 linhas em 0.1s
--- Importando FAMILIAS ---
  OK: ImportaFamilia_PKUnica1
  FAMILIAS: 20 ok, 0 erros | 20 linhas em 0.1s
--- Importando ESTILOS ---
  OK: ImportaEstilo_PKUnica1
  ESTILOS: 3 ok, 0 erros | 3 linhas em 0.0s
--- Importando CLIENTES (modo lote) ---
  617 linhas inseridas no staging.
  CLIENTES: 617 ok, 0 erros | 617 linhas em 1.8s
--- Importando PAGTOFILIAL ---
  OK: ImportaRestricaoCondPagamento_PKUnica1
  PAGTOFILIAL: 10 ok, 0 erros | 10 linhas em 0.1s
--- Importando PRODUTOS (modo lote) ---
  432 linhas inseridas no staging.
  PRODUTOS: 432 ok, 0 erros | 432 linhas em 1.2s
Tempo total: 0m 6s

=== RESUMO ===
  EMPRESA: 82 importados, OK
  FILIAL: 3 importados, OK
  REPR: 126 importados, OK
  PAGTO: 5 importados, OK
  TRANSP: 93 importados, OK
  ESTADOS: 28 importados, OK
  FAMILIAS: 20 importados, OK
  ESTILOS: 3 importados, OK
  CLIENTES: 617 importados, OK
  PAGTOFILIAL: 10 importados, OK
  PRODUTOS: 432 importados, OK

TOTAL: 1419 importados, 0 erros

