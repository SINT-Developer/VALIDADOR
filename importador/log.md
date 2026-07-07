Driver ODBC: ODBC Driver 18 for SQL Server

=== IMPORTACAO ===
Conectado ao SQL Server.
Planilha: 2026.05.25 09-19_DAYHOME FOOD SERVICE_IMPORTAÇÃO.xlsx
  Config EMPRESA: TipoCodProduto=A, TipoCodAuxiliar=N
Fazendo backup do banco de dados...
  Destino: C:\Program Files\Microsoft SQL Server\MSSQL16.SQLERICK\MSSQL\Backup\SRPP - SEM_EMPRESA - 2026-05-25 11h27m20 - Apagar Pedidos e Cadastros e Configuracoes.bak
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
  ERRO ConfigurarEmpresa: <<Versão do Questionário de Paramerização não é compatível com a versão atual do SRPPwin.>> (50000) (SQLExecDirectW)')
  EMPRESA: 82 ok, 1 erros | 82 linhas em 0.2s
--- Importando FILIAL ---
  OK: ImportaFilial_PKUnica1
  ERRO PKUnica2 FILIAL PK=[1]: <<Não foi possível importar Filiais porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO FILIAL linha 1 PK=[1]: <<Não foi possível importar Filiais porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  FILIAL: 0 ok, 2 erros | 1 linhas em 0.0s
--- Importando REPR ---
  ERRO: ImportaRepresentante_PKUnica1 - <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[105]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[106]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[107]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[111]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[113]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[120]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[124]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[125]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[127]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[130]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[132]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[135]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[148]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[156]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[169]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[183]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[189]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[191]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[193]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[194]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[208]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[213]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[216]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[217]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[226]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[227]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[232]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[233]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[234]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[236]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[237]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[238]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[240]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[245]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[247]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[249]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[250]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[255]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[256]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[261]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[263]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[265]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[266]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[267]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[268]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[270]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[273]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[275]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[276]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[277]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[280]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[286]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[290]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[291]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[292]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[294]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[295]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[296]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[297]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[300]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[303]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[304]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[305]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[306]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[309]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[310]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[311]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[314]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[315]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[318]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[319]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[320]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[321]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[322]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[323]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[324]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[325]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[326]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[327]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[328]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[329]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[330]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[331]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[400]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[401]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[402]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[412]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[413]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[414]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[419]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[423]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 REPR PK=[428]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 1 PK=[105]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 2 PK=[106]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 3 PK=[107]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 4 PK=[111]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 5 PK=[113]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 6 PK=[120]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 7 PK=[124]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 8 PK=[125]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 9 PK=[127]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 10 PK=[130]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 11 PK=[132]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 12 PK=[135]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 13 PK=[148]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 14 PK=[156]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 15 PK=[169]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 16 PK=[183]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 17 PK=[189]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 18 PK=[191]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 19 PK=[193]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 20 PK=[194]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 21 PK=[208]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 22 PK=[213]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 23 PK=[216]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 24 PK=[217]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 25 PK=[226]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 26 PK=[227]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 27 PK=[232]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 28 PK=[233]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 29 PK=[234]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 30 PK=[236]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 31 PK=[237]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 32 PK=[238]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 33 PK=[240]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 34 PK=[245]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 35 PK=[247]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 36 PK=[249]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 37 PK=[250]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 38 PK=[255]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 39 PK=[256]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 40 PK=[261]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 41 PK=[263]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 42 PK=[265]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 43 PK=[266]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 44 PK=[267]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 45 PK=[268]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 46 PK=[270]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 47 PK=[273]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 48 PK=[275]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 49 PK=[276]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 50 PK=[277]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 51 PK=[280]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 52 PK=[286]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 53 PK=[290]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 54 PK=[291]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 55 PK=[292]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 56 PK=[294]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 57 PK=[295]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 58 PK=[296]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 59 PK=[297]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 60 PK=[300]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 61 PK=[303]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 62 PK=[304]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 63 PK=[305]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 64 PK=[306]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 65 PK=[309]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 66 PK=[310]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 67 PK=[311]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 68 PK=[314]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 69 PK=[315]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 70 PK=[318]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 71 PK=[319]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 72 PK=[320]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 73 PK=[321]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 74 PK=[322]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 75 PK=[323]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 76 PK=[324]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 77 PK=[325]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 78 PK=[326]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 79 PK=[327]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 80 PK=[328]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 81 PK=[329]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 82 PK=[330]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 83 PK=[331]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 84 PK=[400]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 85 PK=[401]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 86 PK=[402]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 87 PK=[412]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 88 PK=[413]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 89 PK=[414]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 90 PK=[419]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 91 PK=[423]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO REPR linha 92 PK=[428]: <<Não foi possível importar Representantes porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  REPR: 0 ok, 184 erros | 92 linhas em 3.4s
--- Importando PAGTO ---
  OK: ImportaCondPagamento_PKUnica1
  ERRO PKUnica2 PAGTO PK=[001]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 PAGTO PK=[002]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 PAGTO PK=[003]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 PAGTO PK=[004]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 PAGTO PK=[005]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 PAGTO PK=[720]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PAGTO linha 1 PK=[001]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PAGTO linha 2 PK=[002]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PAGTO linha 3 PK=[003]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PAGTO linha 4 PK=[004]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PAGTO linha 5 PK=[005]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PAGTO linha 6 PK=[720]: <<Não foi possível importar Condições de Pagamentos porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  PAGTO: 0 ok, 12 erros | 6 linhas em 0.3s
--- Importando TRANSP (modo lote) ---
  2151 linhas inseridas no staging.
  ERRO TRANSP | DB=SRPP | linha 0 PK=[1]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1 PK=[4]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2 PK=[5]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 3 PK=[6]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 4 PK=[7]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 5 PK=[8]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 6 PK=[10]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 7 PK=[11]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 8 PK=[12]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 9 PK=[13]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 10 PK=[14]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 11 PK=[15]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 12 PK=[16]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 13 PK=[17]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 14 PK=[18]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 15 PK=[19]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 16 PK=[20]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 17 PK=[21]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 18 PK=[22]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 19 PK=[23]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 20 PK=[24]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 21 PK=[25]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 22 PK=[27]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 23 PK=[28]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 24 PK=[29]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 25 PK=[30]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 26 PK=[31]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 27 PK=[32]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 28 PK=[33]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 29 PK=[34]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 30 PK=[35]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 31 PK=[36]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 32 PK=[37]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 33 PK=[38]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 34 PK=[39]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 35 PK=[40]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 36 PK=[41]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 37 PK=[42]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 38 PK=[43]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 39 PK=[44]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 40 PK=[45]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 41 PK=[46]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 42 PK=[48]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 43 PK=[49]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 44 PK=[50]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 45 PK=[51]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 46 PK=[52]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 47 PK=[53]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 48 PK=[54]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 49 PK=[55]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 50 PK=[56]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 51 PK=[57]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 52 PK=[58]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 53 PK=[59]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 54 PK=[60]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 55 PK=[61]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 56 PK=[63]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 57 PK=[64]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 58 PK=[65]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 59 PK=[66]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 60 PK=[67]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 61 PK=[68]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 62 PK=[69]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 63 PK=[70]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 64 PK=[71]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 65 PK=[72]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 66 PK=[73]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 67 PK=[74]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 68 PK=[75]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 69 PK=[76]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 70 PK=[77]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 71 PK=[78]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 72 PK=[80]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 73 PK=[82]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 74 PK=[83]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 75 PK=[84]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 76 PK=[86]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 77 PK=[87]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 78 PK=[88]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 79 PK=[89]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 80 PK=[90]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 81 PK=[91]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 82 PK=[92]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 83 PK=[93]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 84 PK=[94]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 85 PK=[95]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 86 PK=[96]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 87 PK=[97]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 88 PK=[98]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 89 PK=[99]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 90 PK=[100]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 91 PK=[101]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 92 PK=[102]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 93 PK=[103]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 94 PK=[105]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 95 PK=[106]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 96 PK=[107]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 97 PK=[108]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 98 PK=[109]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 99 PK=[110]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 100 PK=[111]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 101 PK=[113]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 102 PK=[114]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 103 PK=[115]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 104 PK=[116]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 105 PK=[117]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 106 PK=[118]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 107 PK=[119]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 108 PK=[120]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 109 PK=[121]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 110 PK=[123]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 111 PK=[124]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 112 PK=[125]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 113 PK=[126]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 114 PK=[127]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 115 PK=[128]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 116 PK=[129]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 117 PK=[130]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 118 PK=[132]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 119 PK=[134]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 120 PK=[135]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 121 PK=[136]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 122 PK=[137]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 123 PK=[138]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 124 PK=[139]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 125 PK=[140]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 126 PK=[141]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 127 PK=[143]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 128 PK=[144]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 129 PK=[145]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 130 PK=[147]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 131 PK=[148]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 132 PK=[149]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 133 PK=[152]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 134 PK=[153]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 135 PK=[154]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 136 PK=[155]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 137 PK=[156]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 138 PK=[157]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 139 PK=[158]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 140 PK=[159]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 141 PK=[160]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 142 PK=[162]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 143 PK=[163]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 144 PK=[164]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 145 PK=[165]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 146 PK=[166]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 147 PK=[167]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 148 PK=[168]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 149 PK=[169]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 150 PK=[170]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 151 PK=[171]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 152 PK=[172]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 153 PK=[173]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 154 PK=[174]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 155 PK=[175]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 156 PK=[176]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 157 PK=[177]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 158 PK=[178]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 159 PK=[179]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 160 PK=[180]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 161 PK=[181]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 162 PK=[182]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 163 PK=[183]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 164 PK=[184]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 165 PK=[185]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 166 PK=[186]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 167 PK=[187]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 168 PK=[188]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 169 PK=[189]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 170 PK=[190]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 171 PK=[191]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 172 PK=[192]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 173 PK=[193]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 174 PK=[194]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 175 PK=[195]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 176 PK=[196]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 177 PK=[197]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 178 PK=[198]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 179 PK=[199]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 180 PK=[200]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 181 PK=[201]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 182 PK=[202]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 183 PK=[203]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 184 PK=[204]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 185 PK=[205]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 186 PK=[207]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 187 PK=[208]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 188 PK=[209]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 189 PK=[210]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 190 PK=[211]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 191 PK=[212]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 192 PK=[213]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 193 PK=[214]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 194 PK=[215]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 195 PK=[217]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 196 PK=[218]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 197 PK=[219]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 198 PK=[220]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 199 PK=[221]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 200 PK=[222]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 201 PK=[223]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 202 PK=[224]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 203 PK=[225]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 204 PK=[227]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 205 PK=[228]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 206 PK=[229]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 207 PK=[230]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 208 PK=[231]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 209 PK=[232]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 210 PK=[233]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 211 PK=[234]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 212 PK=[235]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 213 PK=[236]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 214 PK=[237]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 215 PK=[238]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 216 PK=[239]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 217 PK=[240]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 218 PK=[241]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 219 PK=[242]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 220 PK=[243]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 221 PK=[244]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 222 PK=[245]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 223 PK=[246]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 224 PK=[247]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 225 PK=[248]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 226 PK=[249]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 227 PK=[250]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 228 PK=[251]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 229 PK=[252]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 230 PK=[253]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 231 PK=[255]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 232 PK=[256]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 233 PK=[257]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 234 PK=[258]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 235 PK=[259]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 236 PK=[260]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 237 PK=[261]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 238 PK=[262]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 239 PK=[263]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 240 PK=[264]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 241 PK=[265]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 242 PK=[266]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 243 PK=[267]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 244 PK=[268]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 245 PK=[270]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 246 PK=[271]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 247 PK=[272]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 248 PK=[273]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 249 PK=[275]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 250 PK=[276]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 251 PK=[277]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 252 PK=[278]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 253 PK=[279]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 254 PK=[280]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 255 PK=[281]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 256 PK=[282]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 257 PK=[283]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 258 PK=[284]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 259 PK=[287]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 260 PK=[288]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 261 PK=[289]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 262 PK=[290]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 263 PK=[291]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 264 PK=[292]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 265 PK=[293]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 266 PK=[294]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 267 PK=[295]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 268 PK=[296]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 269 PK=[297]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 270 PK=[298]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 271 PK=[300]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 272 PK=[301]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 273 PK=[302]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 274 PK=[303]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 275 PK=[304]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 276 PK=[305]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 277 PK=[306]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 278 PK=[307]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 279 PK=[308]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 280 PK=[310]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 281 PK=[311]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 282 PK=[312]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 283 PK=[314]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 284 PK=[315]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 285 PK=[316]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 286 PK=[317]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 287 PK=[318]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 288 PK=[319]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 289 PK=[320]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 290 PK=[321]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 291 PK=[322]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 292 PK=[323]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 293 PK=[324]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 294 PK=[326]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 295 PK=[328]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 296 PK=[329]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 297 PK=[330]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 298 PK=[331]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 299 PK=[332]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 300 PK=[333]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 301 PK=[334]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 302 PK=[335]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 303 PK=[336]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 304 PK=[337]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 305 PK=[338]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 306 PK=[339]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 307 PK=[340]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 308 PK=[341]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 309 PK=[342]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 310 PK=[343]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 311 PK=[344]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 312 PK=[345]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 313 PK=[346]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 314 PK=[347]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 315 PK=[348]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 316 PK=[349]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 317 PK=[351]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 318 PK=[352]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 319 PK=[353]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 320 PK=[354]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 321 PK=[355]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 322 PK=[356]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 323 PK=[357]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 324 PK=[358]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 325 PK=[359]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 326 PK=[360]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 327 PK=[361]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 328 PK=[362]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 329 PK=[363]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 330 PK=[365]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 331 PK=[366]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 332 PK=[367]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 333 PK=[368]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 334 PK=[369]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 335 PK=[370]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 336 PK=[371]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 337 PK=[372]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 338 PK=[373]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 339 PK=[374]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 340 PK=[375]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 341 PK=[376]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 342 PK=[377]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 343 PK=[378]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 344 PK=[379]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 345 PK=[380]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 346 PK=[381]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 347 PK=[382]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 348 PK=[383]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 349 PK=[384]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 350 PK=[385]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 351 PK=[386]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 352 PK=[387]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 353 PK=[388]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 354 PK=[389]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 355 PK=[390]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 356 PK=[392]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 357 PK=[393]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 358 PK=[394]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 359 PK=[395]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 360 PK=[396]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 361 PK=[397]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 362 PK=[398]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 363 PK=[399]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 364 PK=[400]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 365 PK=[401]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 366 PK=[402]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 367 PK=[403]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 368 PK=[404]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 369 PK=[405]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 370 PK=[407]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 371 PK=[408]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 372 PK=[409]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 373 PK=[410]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 374 PK=[411]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 375 PK=[412]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 376 PK=[413]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 377 PK=[414]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 378 PK=[415]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 379 PK=[416]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 380 PK=[417]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 381 PK=[418]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 382 PK=[419]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 383 PK=[420]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 384 PK=[421]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 385 PK=[422]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 386 PK=[423]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 387 PK=[424]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 388 PK=[425]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 389 PK=[426]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 390 PK=[428]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 391 PK=[429]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 392 PK=[430]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 393 PK=[431]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 394 PK=[432]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 395 PK=[433]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 396 PK=[434]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 397 PK=[435]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 398 PK=[436]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 399 PK=[437]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 400 PK=[438]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 401 PK=[439]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 402 PK=[440]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 403 PK=[441]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 404 PK=[442]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 405 PK=[443]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 406 PK=[444]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 407 PK=[445]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 408 PK=[447]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 409 PK=[448]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 410 PK=[449]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 411 PK=[450]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 412 PK=[451]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 413 PK=[452]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 414 PK=[453]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 415 PK=[454]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 416 PK=[455]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 417 PK=[456]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 418 PK=[457]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 419 PK=[458]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 420 PK=[459]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 421 PK=[460]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 422 PK=[461]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 423 PK=[462]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 424 PK=[463]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 425 PK=[464]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 426 PK=[465]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 427 PK=[466]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 428 PK=[467]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 429 PK=[468]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 430 PK=[469]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 431 PK=[470]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 432 PK=[471]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 433 PK=[472]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 434 PK=[473]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 435 PK=[474]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 436 PK=[475]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 437 PK=[476]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 438 PK=[477]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 439 PK=[478]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 440 PK=[479]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 441 PK=[480]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 442 PK=[481]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 443 PK=[482]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 444 PK=[483]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 445 PK=[484]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 446 PK=[485]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 447 PK=[486]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 448 PK=[487]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 449 PK=[488]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 450 PK=[489]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 451 PK=[490]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 452 PK=[491]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 453 PK=[492]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 454 PK=[493]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 455 PK=[494]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 456 PK=[495]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 457 PK=[496]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 458 PK=[497]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 459 PK=[498]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 460 PK=[499]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 461 PK=[500]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 462 PK=[501]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 463 PK=[502]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 464 PK=[503]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 465 PK=[504]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 466 PK=[505]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 467 PK=[506]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 468 PK=[507]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 469 PK=[508]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 470 PK=[509]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 471 PK=[510]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 472 PK=[511]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 473 PK=[512]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 474 PK=[513]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 475 PK=[514]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 476 PK=[515]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 477 PK=[516]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 478 PK=[517]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 479 PK=[518]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 480 PK=[519]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 481 PK=[520]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 482 PK=[521]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 483 PK=[522]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 484 PK=[523]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 485 PK=[525]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 486 PK=[526]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 487 PK=[527]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 488 PK=[528]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 489 PK=[529]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 490 PK=[530]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 491 PK=[531]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 492 PK=[532]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 493 PK=[533]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 494 PK=[534]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 495 PK=[535]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 496 PK=[536]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 497 PK=[538]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 498 PK=[539]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 499 PK=[540]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 500 PK=[541]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 501 PK=[542]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 502 PK=[543]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 503 PK=[544]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 504 PK=[545]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 505 PK=[546]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 506 PK=[547]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 507 PK=[548]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 508 PK=[549]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 509 PK=[550]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 510 PK=[551]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 511 PK=[552]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 512 PK=[553]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 513 PK=[554]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 514 PK=[555]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 515 PK=[556]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 516 PK=[557]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 517 PK=[558]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 518 PK=[559]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 519 PK=[560]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 520 PK=[561]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 521 PK=[562]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 522 PK=[563]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 523 PK=[564]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 524 PK=[565]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 525 PK=[566]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 526 PK=[567]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 527 PK=[568]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 528 PK=[569]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 529 PK=[571]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 530 PK=[572]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 531 PK=[573]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 532 PK=[574]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 533 PK=[575]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 534 PK=[576]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 535 PK=[577]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 536 PK=[578]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 537 PK=[579]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 538 PK=[580]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 539 PK=[581]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 540 PK=[582]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 541 PK=[583]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 542 PK=[584]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 543 PK=[585]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 544 PK=[586]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 545 PK=[587]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 546 PK=[588]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 547 PK=[589]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 548 PK=[590]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 549 PK=[592]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 550 PK=[593]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 551 PK=[594]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 552 PK=[595]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 553 PK=[596]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 554 PK=[597]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 555 PK=[598]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 556 PK=[599]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 557 PK=[600]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 558 PK=[601]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 559 PK=[602]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 560 PK=[603]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 561 PK=[604]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 562 PK=[605]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 563 PK=[606]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 564 PK=[607]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 565 PK=[608]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 566 PK=[609]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 567 PK=[610]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 568 PK=[611]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 569 PK=[612]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 570 PK=[613]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 571 PK=[614]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 572 PK=[616]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 573 PK=[617]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 574 PK=[618]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 575 PK=[619]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 576 PK=[620]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 577 PK=[622]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 578 PK=[623]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 579 PK=[624]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 580 PK=[625]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 581 PK=[626]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 582 PK=[627]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 583 PK=[628]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 584 PK=[629]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 585 PK=[630]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 586 PK=[631]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 587 PK=[632]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 588 PK=[633]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 589 PK=[634]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 590 PK=[635]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 591 PK=[636]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 592 PK=[637]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 593 PK=[638]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 594 PK=[639]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 595 PK=[640]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 596 PK=[641]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 597 PK=[642]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 598 PK=[643]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 599 PK=[644]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 600 PK=[645]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 601 PK=[646]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 602 PK=[647]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 603 PK=[649]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 604 PK=[650]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 605 PK=[651]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 606 PK=[652]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 607 PK=[653]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 608 PK=[654]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 609 PK=[655]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 610 PK=[656]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 611 PK=[657]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 612 PK=[658]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 613 PK=[659]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 614 PK=[660]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 615 PK=[661]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 616 PK=[662]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 617 PK=[663]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 618 PK=[664]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 619 PK=[665]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 620 PK=[666]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 621 PK=[667]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 622 PK=[668]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 623 PK=[669]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 624 PK=[670]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 625 PK=[671]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 626 PK=[672]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 627 PK=[673]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 628 PK=[674]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 629 PK=[675]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 630 PK=[676]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 631 PK=[677]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 632 PK=[678]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 633 PK=[679]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 634 PK=[680]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 635 PK=[681]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 636 PK=[682]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 637 PK=[683]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 638 PK=[684]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 639 PK=[685]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 640 PK=[686]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 641 PK=[687]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 642 PK=[688]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 643 PK=[689]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 644 PK=[690]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 645 PK=[691]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 646 PK=[692]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 647 PK=[693]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 648 PK=[694]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 649 PK=[695]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 650 PK=[696]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 651 PK=[697]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 652 PK=[698]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 653 PK=[700]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 654 PK=[701]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 655 PK=[702]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 656 PK=[703]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 657 PK=[704]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 658 PK=[705]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 659 PK=[706]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 660 PK=[707]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 661 PK=[708]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 662 PK=[709]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 663 PK=[710]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 664 PK=[711]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 665 PK=[712]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 666 PK=[713]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 667 PK=[714]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 668 PK=[715]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 669 PK=[716]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 670 PK=[717]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 671 PK=[720]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 672 PK=[721]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 673 PK=[722]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 674 PK=[723]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 675 PK=[724]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 676 PK=[725]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 677 PK=[726]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 678 PK=[727]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 679 PK=[729]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 680 PK=[730]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 681 PK=[731]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 682 PK=[732]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 683 PK=[733]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 684 PK=[734]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 685 PK=[735]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 686 PK=[736]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 687 PK=[737]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 688 PK=[738]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 689 PK=[739]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 690 PK=[740]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 691 PK=[741]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 692 PK=[742]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 693 PK=[743]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 694 PK=[744]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 695 PK=[745]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 696 PK=[746]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 697 PK=[747]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 698 PK=[748]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 699 PK=[749]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 700 PK=[750]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 701 PK=[751]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 702 PK=[752]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 703 PK=[753]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 704 PK=[754]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 705 PK=[755]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 706 PK=[756]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 707 PK=[757]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 708 PK=[758]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 709 PK=[759]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 710 PK=[760]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 711 PK=[761]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 712 PK=[762]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 713 PK=[763]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 714 PK=[764]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 715 PK=[765]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 716 PK=[766]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 717 PK=[767]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 718 PK=[768]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 719 PK=[769]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 720 PK=[770]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 721 PK=[771]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 722 PK=[772]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 723 PK=[773]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 724 PK=[774]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 725 PK=[775]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 726 PK=[776]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 727 PK=[777]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 728 PK=[778]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 729 PK=[779]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 730 PK=[780]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 731 PK=[781]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 732 PK=[782]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 733 PK=[784]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 734 PK=[785]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 735 PK=[786]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 736 PK=[787]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 737 PK=[788]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 738 PK=[789]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 739 PK=[790]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 740 PK=[791]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 741 PK=[792]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 742 PK=[793]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 743 PK=[794]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 744 PK=[795]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 745 PK=[796]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 746 PK=[799]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 747 PK=[800]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 748 PK=[801]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 749 PK=[802]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 750 PK=[803]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 751 PK=[804]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 752 PK=[805]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 753 PK=[806]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 754 PK=[807]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 755 PK=[808]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 756 PK=[809]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 757 PK=[810]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 758 PK=[811]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 759 PK=[813]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 760 PK=[814]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 761 PK=[815]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 762 PK=[816]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 763 PK=[817]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 764 PK=[818]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 765 PK=[819]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 766 PK=[820]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 767 PK=[821]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 768 PK=[822]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 769 PK=[823]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 770 PK=[824]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 771 PK=[825]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 772 PK=[826]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 773 PK=[827]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 774 PK=[828]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 775 PK=[829]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 776 PK=[830]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 777 PK=[831]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 778 PK=[832]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 779 PK=[833]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 780 PK=[834]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 781 PK=[835]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 782 PK=[836]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 783 PK=[837]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 784 PK=[838]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 785 PK=[839]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 786 PK=[840]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 787 PK=[841]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 788 PK=[842]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 789 PK=[843]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 790 PK=[844]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 791 PK=[845]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 792 PK=[846]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 793 PK=[847]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 794 PK=[848]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 795 PK=[849]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 796 PK=[850]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 797 PK=[851]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 798 PK=[852]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 799 PK=[853]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 800 PK=[854]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 801 PK=[855]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 802 PK=[856]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 803 PK=[857]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 804 PK=[858]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 805 PK=[859]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 806 PK=[860]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 807 PK=[861]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 808 PK=[862]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 809 PK=[863]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 810 PK=[864]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 811 PK=[865]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 812 PK=[867]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 813 PK=[868]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 814 PK=[869]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 815 PK=[870]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 816 PK=[871]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 817 PK=[872]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 818 PK=[873]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 819 PK=[874]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 820 PK=[875]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 821 PK=[876]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 822 PK=[877]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 823 PK=[878]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 824 PK=[879]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 825 PK=[880]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 826 PK=[881]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 827 PK=[882]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 828 PK=[883]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 829 PK=[884]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 830 PK=[885]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 831 PK=[886]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 832 PK=[887]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 833 PK=[888]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 834 PK=[890]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 835 PK=[891]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 836 PK=[893]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 837 PK=[894]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 838 PK=[895]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 839 PK=[896]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 840 PK=[897]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 841 PK=[898]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 842 PK=[899]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 843 PK=[900]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 844 PK=[901]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 845 PK=[902]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 846 PK=[903]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 847 PK=[904]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 848 PK=[905]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 849 PK=[906]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 850 PK=[907]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 851 PK=[908]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 852 PK=[909]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 853 PK=[910]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 854 PK=[911]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 855 PK=[912]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 856 PK=[913]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 857 PK=[914]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 858 PK=[915]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 859 PK=[916]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 860 PK=[917]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 861 PK=[918]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 862 PK=[919]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 863 PK=[920]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 864 PK=[921]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 865 PK=[922]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 866 PK=[923]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 867 PK=[924]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 868 PK=[925]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 869 PK=[926]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 870 PK=[927]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 871 PK=[928]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 872 PK=[929]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 873 PK=[930]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 874 PK=[931]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 875 PK=[933]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 876 PK=[934]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 877 PK=[936]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 878 PK=[937]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 879 PK=[938]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 880 PK=[939]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 881 PK=[940]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 882 PK=[941]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 883 PK=[942]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 884 PK=[943]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 885 PK=[944]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 886 PK=[945]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 887 PK=[946]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 888 PK=[947]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 889 PK=[948]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 890 PK=[949]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 891 PK=[950]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 892 PK=[951]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 893 PK=[952]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 894 PK=[953]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 895 PK=[954]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 896 PK=[955]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 897 PK=[956]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 898 PK=[957]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 899 PK=[958]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 900 PK=[959]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 901 PK=[960]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 902 PK=[961]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 903 PK=[962]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 904 PK=[963]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 905 PK=[964]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 906 PK=[965]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 907 PK=[966]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 908 PK=[967]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 909 PK=[968]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 910 PK=[969]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 911 PK=[970]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 912 PK=[971]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 913 PK=[972]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 914 PK=[973]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 915 PK=[974]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 916 PK=[975]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 917 PK=[976]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 918 PK=[977]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 919 PK=[978]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 920 PK=[979]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 921 PK=[980]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 922 PK=[981]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 923 PK=[982]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 924 PK=[983]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 925 PK=[985]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 926 PK=[986]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 927 PK=[987]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 928 PK=[988]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 929 PK=[989]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 930 PK=[990]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 931 PK=[991]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 932 PK=[992]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 933 PK=[993]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 934 PK=[994]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 935 PK=[995]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 936 PK=[996]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 937 PK=[997]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 938 PK=[998]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 939 PK=[999]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 940 PK=[1000]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 941 PK=[1001]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 942 PK=[1002]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 943 PK=[1003]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 944 PK=[1004]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 945 PK=[1005]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 946 PK=[1006]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 947 PK=[1007]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 948 PK=[1008]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 949 PK=[1009]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 950 PK=[1010]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 951 PK=[1011]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 952 PK=[1012]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 953 PK=[1013]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 954 PK=[1014]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 955 PK=[1015]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 956 PK=[1016]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 957 PK=[1017]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 958 PK=[1018]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 959 PK=[1019]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 960 PK=[1020]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 961 PK=[1021]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 962 PK=[1022]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 963 PK=[1023]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 964 PK=[1024]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 965 PK=[1025]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 966 PK=[1026]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 967 PK=[1027]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 968 PK=[1028]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 969 PK=[1029]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 970 PK=[1030]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 971 PK=[1031]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 972 PK=[1032]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 973 PK=[1033]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 974 PK=[1034]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 975 PK=[1035]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 976 PK=[1036]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 977 PK=[1037]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 978 PK=[1038]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 979 PK=[1039]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 980 PK=[1040]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 981 PK=[1041]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 982 PK=[1042]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 983 PK=[1043]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 984 PK=[1044]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 985 PK=[1045]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 986 PK=[1046]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 987 PK=[1047]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 988 PK=[1048]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 989 PK=[1049]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 990 PK=[1050]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 991 PK=[1051]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 992 PK=[1052]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 993 PK=[1053]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 994 PK=[1054]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 995 PK=[1055]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 996 PK=[1056]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 997 PK=[1057]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 998 PK=[1058]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 999 PK=[1059]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1000 PK=[1060]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1001 PK=[1061]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1002 PK=[1062]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1003 PK=[1063]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1004 PK=[1064]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1005 PK=[1065]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1006 PK=[1066]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1007 PK=[1067]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1008 PK=[1068]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1009 PK=[1069]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1010 PK=[1070]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1011 PK=[1071]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1012 PK=[1072]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1013 PK=[1073]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1014 PK=[1074]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1015 PK=[1075]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1016 PK=[1076]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1017 PK=[1077]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1018 PK=[1078]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1019 PK=[1079]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1020 PK=[1080]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1021 PK=[1081]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1022 PK=[1083]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1023 PK=[1084]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1024 PK=[1085]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1025 PK=[1086]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1026 PK=[1087]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1027 PK=[1088]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1028 PK=[1089]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1029 PK=[1090]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1030 PK=[1091]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1031 PK=[1092]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1032 PK=[1093]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1033 PK=[1094]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1034 PK=[1095]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1035 PK=[1096]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1036 PK=[1097]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1037 PK=[1098]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1038 PK=[1099]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1039 PK=[1100]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1040 PK=[1101]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1041 PK=[1102]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1042 PK=[1103]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1043 PK=[1104]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1044 PK=[1105]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1045 PK=[1106]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1046 PK=[1107]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1047 PK=[1108]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1048 PK=[1109]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1049 PK=[1110]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1050 PK=[1111]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1051 PK=[1112]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1052 PK=[1113]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1053 PK=[1114]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1054 PK=[1115]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1055 PK=[1116]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1056 PK=[1117]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1057 PK=[1118]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1058 PK=[1119]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1059 PK=[1120]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1060 PK=[1121]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1061 PK=[1122]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1062 PK=[1123]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1063 PK=[1124]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1064 PK=[1125]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1065 PK=[1126]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1066 PK=[1127]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1067 PK=[1128]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1068 PK=[1129]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1069 PK=[1130]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1070 PK=[1131]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1071 PK=[1132]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1072 PK=[1133]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1073 PK=[1134]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1074 PK=[1135]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1075 PK=[1136]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1076 PK=[1137]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1077 PK=[1138]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1078 PK=[1139]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1079 PK=[1140]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1080 PK=[1141]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1081 PK=[1142]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1082 PK=[1143]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1083 PK=[1144]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1084 PK=[1145]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1085 PK=[1146]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1086 PK=[1147]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1087 PK=[1148]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1088 PK=[1149]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1089 PK=[1150]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1090 PK=[1151]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1091 PK=[1152]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1092 PK=[1153]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1093 PK=[1154]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1094 PK=[1155]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1095 PK=[1156]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1096 PK=[1157]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1097 PK=[1158]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1098 PK=[1159]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1099 PK=[1160]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1100 PK=[1161]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1101 PK=[1162]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1102 PK=[1163]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1103 PK=[1164]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1104 PK=[1165]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1105 PK=[1166]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1106 PK=[1167]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1107 PK=[1168]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1108 PK=[1169]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1109 PK=[1170]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1110 PK=[1171]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1111 PK=[1172]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1112 PK=[1173]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1113 PK=[1174]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1114 PK=[1175]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1115 PK=[1176]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1116 PK=[1177]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1117 PK=[1178]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1118 PK=[1179]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1119 PK=[1180]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1120 PK=[1181]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1121 PK=[1182]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1122 PK=[1183]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1123 PK=[1184]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1124 PK=[1185]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1125 PK=[1186]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1126 PK=[1187]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1127 PK=[1188]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1128 PK=[1189]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1129 PK=[1190]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1130 PK=[1191]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1131 PK=[1192]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1132 PK=[1193]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1133 PK=[1194]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1134 PK=[1195]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1135 PK=[1196]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1136 PK=[1197]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1137 PK=[1198]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1138 PK=[1199]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1139 PK=[1200]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1140 PK=[1201]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1141 PK=[1202]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1142 PK=[1203]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1143 PK=[1204]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1144 PK=[1205]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1145 PK=[1206]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1146 PK=[1207]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1147 PK=[1208]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1148 PK=[1209]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1149 PK=[1210]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1150 PK=[1211]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1151 PK=[1212]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1152 PK=[1213]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1153 PK=[1214]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1154 PK=[1215]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1155 PK=[1216]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1156 PK=[1217]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1157 PK=[1218]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1158 PK=[1219]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1159 PK=[1220]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1160 PK=[1221]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1161 PK=[1222]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1162 PK=[1223]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1163 PK=[1224]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1164 PK=[1225]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1165 PK=[1226]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1166 PK=[1227]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1167 PK=[1228]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1168 PK=[1229]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1169 PK=[1230]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1170 PK=[1231]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1171 PK=[1232]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1172 PK=[1233]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1173 PK=[1234]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1174 PK=[1235]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1175 PK=[1236]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1176 PK=[1237]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1177 PK=[1238]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1178 PK=[1239]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1179 PK=[1240]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1180 PK=[1241]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1181 PK=[1242]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1182 PK=[1243]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1183 PK=[1244]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1184 PK=[1245]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1185 PK=[1246]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1186 PK=[1247]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1187 PK=[1248]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1188 PK=[1249]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1189 PK=[1250]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1190 PK=[1251]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1191 PK=[1252]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1192 PK=[1253]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1193 PK=[1254]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1194 PK=[1255]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1195 PK=[1256]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1196 PK=[1257]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1197 PK=[1258]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1198 PK=[1259]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1199 PK=[1260]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1200 PK=[1261]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1201 PK=[1263]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1202 PK=[1264]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1203 PK=[1265]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1204 PK=[1266]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1205 PK=[1267]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1206 PK=[1268]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1207 PK=[1269]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1208 PK=[1270]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1209 PK=[1271]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1210 PK=[1272]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1211 PK=[1273]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1212 PK=[1274]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1213 PK=[1275]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1214 PK=[1276]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1215 PK=[1277]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1216 PK=[1278]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1217 PK=[1279]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1218 PK=[1280]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1219 PK=[1281]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1220 PK=[1282]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1221 PK=[1283]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1222 PK=[1284]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1223 PK=[1285]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1224 PK=[1286]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1225 PK=[1287]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1226 PK=[1288]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1227 PK=[1289]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1228 PK=[1290]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1229 PK=[1291]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1230 PK=[1292]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1231 PK=[1293]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1232 PK=[1294]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1233 PK=[1295]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1234 PK=[1296]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1235 PK=[1297]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1236 PK=[1298]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1237 PK=[1299]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1238 PK=[1300]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1239 PK=[1301]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1240 PK=[1303]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1241 PK=[1304]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1242 PK=[1305]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1243 PK=[1306]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1244 PK=[1307]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1245 PK=[1308]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1246 PK=[1309]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1247 PK=[1310]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1248 PK=[1311]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1249 PK=[1312]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1250 PK=[1313]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1251 PK=[1314]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1252 PK=[1315]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1253 PK=[1316]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1254 PK=[1317]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1255 PK=[1318]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1256 PK=[1319]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1257 PK=[1320]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1258 PK=[1321]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1259 PK=[1322]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1260 PK=[1323]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1261 PK=[1324]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1262 PK=[1325]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1263 PK=[1326]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1264 PK=[1327]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1265 PK=[1328]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1266 PK=[1329]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1267 PK=[1330]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1268 PK=[1331]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1269 PK=[1332]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1270 PK=[1333]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1271 PK=[1334]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1272 PK=[1335]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1273 PK=[1336]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1274 PK=[1337]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1275 PK=[1338]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1276 PK=[1339]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1277 PK=[1340]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1278 PK=[1341]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1279 PK=[1342]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1280 PK=[1343]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1281 PK=[1344]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1282 PK=[1345]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1283 PK=[1346]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1284 PK=[1347]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1285 PK=[1348]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1286 PK=[1349]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1287 PK=[1350]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1288 PK=[1351]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1289 PK=[1352]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1290 PK=[1353]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1291 PK=[1354]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1292 PK=[1355]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1293 PK=[1356]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1294 PK=[1357]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1295 PK=[1358]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1296 PK=[1359]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1297 PK=[1360]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1298 PK=[1361]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1299 PK=[1362]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1300 PK=[1363]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1301 PK=[1364]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1302 PK=[1365]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1303 PK=[1366]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1304 PK=[1367]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1305 PK=[1368]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1306 PK=[1369]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1307 PK=[1370]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1308 PK=[1371]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1309 PK=[1372]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1310 PK=[1373]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1311 PK=[1374]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1312 PK=[1375]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1313 PK=[1376]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1314 PK=[1377]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1315 PK=[1378]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1316 PK=[1379]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1317 PK=[1380]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1318 PK=[1381]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1319 PK=[1382]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1320 PK=[1383]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1321 PK=[1384]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1322 PK=[1385]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1323 PK=[1386]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1324 PK=[1387]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1325 PK=[1388]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1326 PK=[1389]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1327 PK=[1390]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1328 PK=[1391]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1329 PK=[1392]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1330 PK=[1393]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1331 PK=[1394]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1332 PK=[1395]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1333 PK=[1396]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1334 PK=[1397]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1335 PK=[1398]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1336 PK=[1399]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1337 PK=[1400]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1338 PK=[1401]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1339 PK=[1402]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1340 PK=[1403]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1341 PK=[1404]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1342 PK=[1406]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1343 PK=[1407]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1344 PK=[1408]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1345 PK=[1409]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1346 PK=[1410]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1347 PK=[1411]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1348 PK=[1412]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1349 PK=[1413]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1350 PK=[1414]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1351 PK=[1415]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1352 PK=[1416]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1353 PK=[1417]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1354 PK=[1418]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1355 PK=[1419]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1356 PK=[1420]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1357 PK=[1421]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1358 PK=[1422]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1359 PK=[1423]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1360 PK=[1424]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1361 PK=[1425]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1362 PK=[1426]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1363 PK=[1427]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1364 PK=[1428]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1365 PK=[1429]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1366 PK=[1430]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1367 PK=[1431]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1368 PK=[1432]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1369 PK=[1433]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1370 PK=[1434]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1371 PK=[1435]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1372 PK=[1437]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1373 PK=[1438]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1374 PK=[1439]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1375 PK=[1440]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1376 PK=[1441]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1377 PK=[1442]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1378 PK=[1443]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1379 PK=[1444]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1380 PK=[1445]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1381 PK=[1446]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1382 PK=[1447]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1383 PK=[1448]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1384 PK=[1449]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1385 PK=[1450]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1386 PK=[1451]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1387 PK=[1452]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1388 PK=[1453]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1389 PK=[1454]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1390 PK=[1455]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1391 PK=[1456]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1392 PK=[1457]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1393 PK=[1458]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1394 PK=[1459]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1395 PK=[1460]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1396 PK=[1461]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1397 PK=[1462]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1398 PK=[1463]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1399 PK=[1464]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1400 PK=[1465]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1401 PK=[1466]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1402 PK=[1467]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1403 PK=[1468]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1404 PK=[1469]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1405 PK=[1470]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1406 PK=[1471]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1407 PK=[1472]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1408 PK=[1473]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1409 PK=[1474]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1410 PK=[1475]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1411 PK=[1476]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1412 PK=[1477]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1413 PK=[1478]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1414 PK=[1479]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1415 PK=[1480]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1416 PK=[1481]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1417 PK=[1482]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1418 PK=[1483]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1419 PK=[1484]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1420 PK=[1485]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1421 PK=[1486]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1422 PK=[1487]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1423 PK=[1488]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1424 PK=[1489]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1425 PK=[1490]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1426 PK=[1491]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1427 PK=[1492]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1428 PK=[1493]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1429 PK=[1494]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1430 PK=[1495]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1431 PK=[1496]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1432 PK=[1497]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1433 PK=[1498]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1434 PK=[1499]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1435 PK=[1500]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1436 PK=[1501]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1437 PK=[1502]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1438 PK=[1503]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1439 PK=[1504]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1440 PK=[1505]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1441 PK=[1506]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1442 PK=[1507]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1443 PK=[1508]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1444 PK=[1509]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1445 PK=[1510]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1446 PK=[1511]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1447 PK=[1512]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1448 PK=[1513]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1449 PK=[1514]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1450 PK=[1515]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1451 PK=[1516]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1452 PK=[1517]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1453 PK=[1518]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1454 PK=[1519]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1455 PK=[1520]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1456 PK=[1521]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1457 PK=[1522]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1458 PK=[1523]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1459 PK=[1524]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1460 PK=[1525]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1461 PK=[1526]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1462 PK=[1527]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1463 PK=[1528]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1464 PK=[1529]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1465 PK=[1530]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1466 PK=[1531]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1467 PK=[1532]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1468 PK=[1533]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1469 PK=[1535]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1470 PK=[1536]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1471 PK=[1537]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1472 PK=[1538]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1473 PK=[1539]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1474 PK=[1540]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1475 PK=[1541]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1476 PK=[1542]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1477 PK=[1543]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1478 PK=[1544]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1479 PK=[1545]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1480 PK=[1546]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1481 PK=[1547]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1482 PK=[1548]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1483 PK=[1549]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1484 PK=[1550]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1485 PK=[1551]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1486 PK=[1552]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1487 PK=[1553]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1488 PK=[1554]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1489 PK=[1555]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1490 PK=[1557]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1491 PK=[1558]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1492 PK=[1559]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1493 PK=[1560]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1494 PK=[1562]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1495 PK=[1563]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1496 PK=[1564]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1497 PK=[1565]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1498 PK=[1566]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1499 PK=[1567]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1500 PK=[1568]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1501 PK=[1569]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1502 PK=[1570]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1503 PK=[1571]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1504 PK=[1572]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1505 PK=[1573]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1506 PK=[1574]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1507 PK=[1575]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1508 PK=[1576]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1509 PK=[1577]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1510 PK=[1578]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1511 PK=[1579]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1512 PK=[1580]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1513 PK=[1581]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1514 PK=[1582]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1515 PK=[1583]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1516 PK=[1584]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1517 PK=[1585]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1518 PK=[1586]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1519 PK=[1587]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1520 PK=[1588]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1521 PK=[1589]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1522 PK=[1590]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1523 PK=[1591]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1524 PK=[1592]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1525 PK=[1593]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1526 PK=[1595]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1527 PK=[1596]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1528 PK=[1597]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1529 PK=[1598]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1530 PK=[1599]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1531 PK=[1600]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1532 PK=[1601]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1533 PK=[1602]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1534 PK=[1603]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1535 PK=[1604]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1536 PK=[1605]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1537 PK=[1606]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1538 PK=[1607]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1539 PK=[1608]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1540 PK=[1609]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1541 PK=[1610]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1542 PK=[1611]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1543 PK=[1612]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1544 PK=[1613]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1545 PK=[1614]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1546 PK=[1615]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1547 PK=[1616]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1548 PK=[1617]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1549 PK=[1618]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1550 PK=[1619]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1551 PK=[1620]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1552 PK=[1621]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1553 PK=[1622]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1554 PK=[1623]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1555 PK=[1624]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1556 PK=[1625]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1557 PK=[1626]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1558 PK=[1627]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1559 PK=[1628]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1560 PK=[1629]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1561 PK=[1630]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1562 PK=[1631]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1563 PK=[1632]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1564 PK=[1633]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1565 PK=[1634]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1566 PK=[1635]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1567 PK=[1636]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1568 PK=[1637]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1569 PK=[1638]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1570 PK=[1639]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1571 PK=[1640]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1572 PK=[1641]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1573 PK=[1642]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1574 PK=[1643]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1575 PK=[1644]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1576 PK=[1645]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1577 PK=[1646]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1578 PK=[1647]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1579 PK=[1648]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1580 PK=[1649]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1581 PK=[1650]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1582 PK=[1652]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1583 PK=[1653]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1584 PK=[1654]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1585 PK=[1655]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1586 PK=[1656]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1587 PK=[1657]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1588 PK=[1658]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1589 PK=[1659]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1590 PK=[1660]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1591 PK=[1661]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1592 PK=[1662]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1593 PK=[1663]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1594 PK=[1665]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1595 PK=[1666]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1596 PK=[1667]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1597 PK=[1668]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1598 PK=[1669]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1599 PK=[1670]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1600 PK=[1671]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1601 PK=[1672]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1602 PK=[1673]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1603 PK=[1674]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1604 PK=[1675]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1605 PK=[1676]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1606 PK=[1677]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1607 PK=[1678]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1608 PK=[1679]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1609 PK=[1680]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1610 PK=[1681]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1611 PK=[1682]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1612 PK=[1683]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1613 PK=[1684]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1614 PK=[1685]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1615 PK=[1686]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1616 PK=[1687]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1617 PK=[1688]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1618 PK=[1689]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1619 PK=[1690]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1620 PK=[1691]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1621 PK=[1693]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1622 PK=[1694]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1623 PK=[1695]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1624 PK=[1696]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1625 PK=[1697]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1626 PK=[1698]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1627 PK=[1699]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1628 PK=[1700]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1629 PK=[1701]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1630 PK=[1702]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1631 PK=[1703]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1632 PK=[1704]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1633 PK=[1705]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1634 PK=[1706]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1635 PK=[1707]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1636 PK=[1708]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1637 PK=[1709]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1638 PK=[1710]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1639 PK=[1711]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1640 PK=[1712]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1641 PK=[1713]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1642 PK=[1714]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1643 PK=[1715]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1644 PK=[1716]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1645 PK=[1717]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1646 PK=[1718]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1647 PK=[1719]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1648 PK=[1720]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1649 PK=[1721]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1650 PK=[1722]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1651 PK=[1723]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1652 PK=[1725]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1653 PK=[1726]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1654 PK=[1727]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1655 PK=[1728]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1656 PK=[1729]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1657 PK=[1730]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1658 PK=[1731]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1659 PK=[1732]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1660 PK=[1733]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1661 PK=[1734]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1662 PK=[1735]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1663 PK=[1736]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1664 PK=[1738]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1665 PK=[1739]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1666 PK=[1740]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1667 PK=[1741]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1668 PK=[1742]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1669 PK=[1743]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1670 PK=[1745]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1671 PK=[1746]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1672 PK=[1747]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1673 PK=[1748]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1674 PK=[1749]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1675 PK=[1750]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1676 PK=[1751]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1677 PK=[1752]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1678 PK=[1753]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1679 PK=[1754]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1680 PK=[1755]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1681 PK=[1756]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1682 PK=[1757]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1683 PK=[1758]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1684 PK=[1759]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1685 PK=[1760]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1686 PK=[1761]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1687 PK=[1762]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1688 PK=[1763]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1689 PK=[1765]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1690 PK=[1766]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1691 PK=[1767]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1692 PK=[1768]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1693 PK=[1769]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1694 PK=[1770]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1695 PK=[1771]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1696 PK=[1772]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1697 PK=[1773]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1698 PK=[1774]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1699 PK=[1775]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1700 PK=[1776]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1701 PK=[1777]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1702 PK=[1778]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1703 PK=[1779]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1704 PK=[1780]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1705 PK=[1781]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1706 PK=[1782]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1707 PK=[1783]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1708 PK=[1784]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1709 PK=[1785]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1710 PK=[1786]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1711 PK=[1787]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1712 PK=[1788]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1713 PK=[1789]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1714 PK=[1790]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1715 PK=[1791]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1716 PK=[1792]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1717 PK=[1793]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1718 PK=[1794]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1719 PK=[1795]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1720 PK=[1796]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1721 PK=[1797]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1722 PK=[1798]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1723 PK=[1799]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1724 PK=[1800]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1725 PK=[1801]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1726 PK=[1802]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1727 PK=[1803]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1728 PK=[1804]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1729 PK=[1805]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1730 PK=[1806]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1731 PK=[1807]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1732 PK=[1809]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1733 PK=[1810]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1734 PK=[1811]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1735 PK=[1812]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1736 PK=[1813]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1737 PK=[1814]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1738 PK=[1815]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1739 PK=[1816]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1740 PK=[1817]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1741 PK=[1819]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1742 PK=[1820]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1743 PK=[1821]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1744 PK=[1822]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1745 PK=[1824]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1746 PK=[1825]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1747 PK=[1826]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1748 PK=[1827]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1749 PK=[1828]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1750 PK=[1829]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1751 PK=[1830]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1752 PK=[1831]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1753 PK=[1832]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1754 PK=[1833]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1755 PK=[1834]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1756 PK=[1835]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1757 PK=[1836]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1758 PK=[1837]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1759 PK=[1838]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1760 PK=[1839]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1761 PK=[1840]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1762 PK=[1841]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1763 PK=[1842]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1764 PK=[1843]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1765 PK=[1844]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1766 PK=[1845]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1767 PK=[1846]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1768 PK=[1847]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1769 PK=[1848]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1770 PK=[1849]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1771 PK=[1850]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1772 PK=[1851]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1773 PK=[1852]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1774 PK=[1853]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1775 PK=[1854]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1776 PK=[1855]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1777 PK=[1856]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1778 PK=[1857]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1779 PK=[1858]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1780 PK=[1859]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1781 PK=[1860]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1782 PK=[1861]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1783 PK=[1862]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1784 PK=[1863]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1785 PK=[1864]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1786 PK=[1865]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1787 PK=[1866]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1788 PK=[1867]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1789 PK=[1868]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1790 PK=[1869]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1791 PK=[1870]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1792 PK=[1871]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1793 PK=[1872]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1794 PK=[1873]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1795 PK=[1874]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1796 PK=[1875]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1797 PK=[1876]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1798 PK=[1877]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1799 PK=[1878]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1800 PK=[1879]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1801 PK=[1880]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1802 PK=[1881]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1803 PK=[1882]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1804 PK=[1883]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1805 PK=[1884]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1806 PK=[1885]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1807 PK=[1886]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1808 PK=[1887]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1809 PK=[1888]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1810 PK=[1889]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1811 PK=[1890]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1812 PK=[1891]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1813 PK=[1892]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1814 PK=[1893]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1815 PK=[1894]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1816 PK=[1895]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1817 PK=[1896]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1818 PK=[1897]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1819 PK=[1899]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1820 PK=[1900]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1821 PK=[1901]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1822 PK=[1902]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1823 PK=[1903]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1824 PK=[1904]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1825 PK=[1905]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1826 PK=[1906]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1827 PK=[1907]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1828 PK=[1908]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1829 PK=[1909]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1830 PK=[1910]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1831 PK=[1911]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1832 PK=[1912]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1833 PK=[1913]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1834 PK=[1914]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1835 PK=[1915]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1836 PK=[1916]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1837 PK=[1917]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1838 PK=[1918]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1839 PK=[1919]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1840 PK=[1920]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1841 PK=[1921]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1842 PK=[1922]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1843 PK=[1923]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1844 PK=[1924]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1845 PK=[1925]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1846 PK=[1926]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1847 PK=[1927]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1848 PK=[1928]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1849 PK=[1929]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1850 PK=[1930]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1851 PK=[1931]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1852 PK=[1932]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1853 PK=[1933]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1854 PK=[1934]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1855 PK=[1935]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1856 PK=[1936]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1857 PK=[1937]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1858 PK=[1938]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1859 PK=[1939]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1860 PK=[1940]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1861 PK=[1941]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1862 PK=[1942]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1863 PK=[1943]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1864 PK=[1944]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1865 PK=[1945]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1866 PK=[1946]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1867 PK=[1947]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1868 PK=[1948]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1869 PK=[1949]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1870 PK=[1950]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1871 PK=[1951]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1872 PK=[1952]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1873 PK=[1953]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1874 PK=[1954]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1875 PK=[1956]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1876 PK=[1957]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1877 PK=[1958]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1878 PK=[1959]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1879 PK=[1960]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1880 PK=[1961]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1881 PK=[1962]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1882 PK=[1963]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1883 PK=[1964]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1884 PK=[1965]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1885 PK=[1966]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1886 PK=[1967]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1887 PK=[1968]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1888 PK=[1969]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1889 PK=[1970]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1890 PK=[1971]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1891 PK=[1972]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1892 PK=[1973]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1893 PK=[1974]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1894 PK=[1975]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1895 PK=[1976]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1896 PK=[1977]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1897 PK=[1978]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1898 PK=[1979]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1899 PK=[1980]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1900 PK=[1981]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1901 PK=[1982]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1902 PK=[1983]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1903 PK=[1984]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1904 PK=[1985]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1905 PK=[1986]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1906 PK=[1987]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1907 PK=[1988]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1908 PK=[1989]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1909 PK=[1990]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1910 PK=[1991]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1911 PK=[1992]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1912 PK=[1993]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1913 PK=[1994]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1914 PK=[1995]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1915 PK=[1996]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1916 PK=[1997]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1917 PK=[1998]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1918 PK=[1999]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1919 PK=[2000]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1920 PK=[2001]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1921 PK=[2002]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1922 PK=[2003]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1923 PK=[2004]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1924 PK=[2005]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1925 PK=[2006]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1926 PK=[2007]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1927 PK=[2008]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1928 PK=[2009]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1929 PK=[2010]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1930 PK=[2011]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1931 PK=[2012]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1932 PK=[2013]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1933 PK=[2014]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1934 PK=[2015]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1935 PK=[2016]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1936 PK=[2017]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1937 PK=[2018]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1938 PK=[2019]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1939 PK=[2020]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1940 PK=[2021]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1941 PK=[2022]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1942 PK=[2023]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1943 PK=[2024]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1944 PK=[2025]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1945 PK=[2026]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1946 PK=[2027]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1947 PK=[2028]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1948 PK=[2029]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1949 PK=[2030]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1950 PK=[2031]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1951 PK=[2032]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1952 PK=[2033]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1953 PK=[2034]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1954 PK=[2035]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1955 PK=[2036]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1956 PK=[2037]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1957 PK=[2038]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1958 PK=[2039]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1959 PK=[2040]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1960 PK=[2041]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1961 PK=[2042]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1962 PK=[2043]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1963 PK=[2044]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1964 PK=[2045]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1965 PK=[2046]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1966 PK=[2047]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1967 PK=[2048]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1968 PK=[2049]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1969 PK=[2051]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1970 PK=[2052]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1971 PK=[2053]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1972 PK=[2054]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1973 PK=[2055]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1974 PK=[2056]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1975 PK=[2057]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1976 PK=[2058]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1977 PK=[2059]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1978 PK=[2060]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1979 PK=[2061]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1980 PK=[2062]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1981 PK=[2063]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1982 PK=[2064]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1983 PK=[2065]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1984 PK=[2066]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1985 PK=[2067]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1986 PK=[2068]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1987 PK=[2069]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1988 PK=[2070]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1989 PK=[2071]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1990 PK=[2072]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1991 PK=[2073]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1992 PK=[2074]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1993 PK=[2075]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1994 PK=[2076]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1995 PK=[2077]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1996 PK=[2078]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1997 PK=[2079]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1998 PK=[2080]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 1999 PK=[2081]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2000 PK=[2082]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2001 PK=[2083]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2002 PK=[2084]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2003 PK=[2085]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2004 PK=[2086]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2005 PK=[2087]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2006 PK=[2088]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2007 PK=[2089]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2008 PK=[2090]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2009 PK=[2091]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2010 PK=[2092]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2011 PK=[2093]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2012 PK=[2094]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2013 PK=[2095]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2014 PK=[2096]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2015 PK=[2097]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2016 PK=[2098]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2017 PK=[2099]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2018 PK=[2100]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2019 PK=[2101]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2020 PK=[2102]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2021 PK=[2103]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2022 PK=[2104]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2023 PK=[2105]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2024 PK=[2106]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2025 PK=[2107]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2026 PK=[2108]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2027 PK=[2109]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2028 PK=[2110]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2029 PK=[2111]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2030 PK=[2112]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2031 PK=[2113]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2032 PK=[2114]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2033 PK=[2115]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2034 PK=[2116]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2035 PK=[2117]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2036 PK=[2118]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2037 PK=[2119]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2038 PK=[2120]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2039 PK=[2121]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2040 PK=[2122]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2041 PK=[2123]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2042 PK=[2124]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2043 PK=[2125]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2044 PK=[2127]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2045 PK=[2128]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2046 PK=[2129]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2047 PK=[2130]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2048 PK=[2131]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2049 PK=[2132]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2050 PK=[2133]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2051 PK=[2134]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2052 PK=[2135]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2053 PK=[2136]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2054 PK=[2137]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2055 PK=[2138]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2056 PK=[2139]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2057 PK=[2140]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2058 PK=[2141]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2059 PK=[2142]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2060 PK=[2143]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2061 PK=[2144]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2062 PK=[2145]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2063 PK=[2146]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2064 PK=[2147]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2065 PK=[2148]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2066 PK=[2149]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2067 PK=[2150]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2068 PK=[2151]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2069 PK=[2152]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2070 PK=[2153]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2071 PK=[2154]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2072 PK=[2155]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2073 PK=[2156]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2074 PK=[2157]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2075 PK=[2158]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2076 PK=[2159]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2077 PK=[2160]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2078 PK=[2161]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2079 PK=[2162]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2080 PK=[2163]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2081 PK=[2164]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2082 PK=[2165]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2083 PK=[2166]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2084 PK=[2167]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2085 PK=[2168]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2086 PK=[2169]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2087 PK=[2170]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2088 PK=[2171]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2089 PK=[2172]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2090 PK=[2173]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2091 PK=[2174]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2092 PK=[2175]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2093 PK=[2176]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2094 PK=[2177]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2095 PK=[2178]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2096 PK=[2179]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2097 PK=[2181]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2098 PK=[2182]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2099 PK=[2183]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2100 PK=[2184]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2101 PK=[2185]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2102 PK=[2186]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2103 PK=[2187]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2104 PK=[2188]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2105 PK=[2189]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2106 PK=[2190]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2107 PK=[2191]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2108 PK=[2192]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2109 PK=[2193]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2110 PK=[2194]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2111 PK=[2195]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2112 PK=[2196]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2113 PK=[2197]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2114 PK=[2198]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2115 PK=[2199]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2116 PK=[2200]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2117 PK=[2201]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2118 PK=[2202]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2119 PK=[2203]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2120 PK=[2204]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2121 PK=[2205]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2122 PK=[2206]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2123 PK=[2207]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2124 PK=[2208]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2125 PK=[2209]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2126 PK=[2210]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2127 PK=[2211]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2128 PK=[2212]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2129 PK=[2213]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2130 PK=[2214]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2131 PK=[2215]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2132 PK=[2216]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2133 PK=[2217]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2134 PK=[2218]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2135 PK=[2219]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2136 PK=[2220]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2137 PK=[2221]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2138 PK=[2222]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2139 PK=[2223]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2140 PK=[2224]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2141 PK=[2225]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2142 PK=[2226]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2143 PK=[2227]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2144 PK=[2229]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2145 PK=[2230]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2146 PK=[2231]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2147 PK=[2232]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2148 PK=[2233]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2149 PK=[2234]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  ERRO TRANSP | DB=SRPP | linha 2150 PK=[2235]: Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.
  TRANSP: 0 ok, 2151 erros | 2151 linhas em 32.8s
--- Importando ESTADOS ---
  OK: ImportaEstado_PKUnica1
  ERRO PKUnica2 ESTADOS PK=[AC]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[AL]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[AM]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[AP]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[BA]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[CE]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[DF]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[ES]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[EX]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[GO]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[MA]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[MG]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[MS]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[MT]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[PA]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[PB]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[PE]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[PI]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[PR]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[RJ]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[RN]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[RO]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[RR]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[RS]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[SC]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[SE]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[SP]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO PKUnica2 ESTADOS PK=[TO]: <<Não foi possível importar Estado porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 1 PK=[AC]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 2 PK=[AL]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 3 PK=[AM]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 4 PK=[AP]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 5 PK=[BA]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 6 PK=[CE]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 7 PK=[DF]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 8 PK=[ES]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 9 PK=[EX]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 10 PK=[GO]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 11 PK=[MA]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 12 PK=[MG]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 13 PK=[MS]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 14 PK=[MT]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 15 PK=[PA]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 16 PK=[PB]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 17 PK=[PE]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 18 PK=[PI]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 19 PK=[PR]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 20 PK=[RJ]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 21 PK=[RN]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 22 PK=[RO]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 23 PK=[RR]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 24 PK=[RS]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 25 PK=[SC]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 26 PK=[SE]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 27 PK=[SP]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ERRO ESTADOS linha 28 PK=[TO]: <<Não foi possível importar Estados porque Empresa não foi configurada.>> (50000) (SQLExecDirectW)')
  ESTADOS: 0 ok, 56 erros | 28 linhas em 1.4s
--- Importando FAMILIAS ---
  Nenhuma linha valida em FAMILIAS
--- Importando ESTILOS ---
  Nenhuma linha valida em ESTILOS
--- Importando CLIENTES (modo lote) ---

