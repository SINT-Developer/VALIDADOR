"""
Motor de importacao - executa as stored procedures no SQL Server.
"""

import os
import time
from datetime import datetime
import pyodbc
from openpyxl import load_workbook
from mapeamento import MAPA_ABAS, ORDEM_IMPORTACAO
from sql_batch import (
    SQL_CREATE_STAGING,
    SQL_CREATE_PROCEDURE,
    SQL_CREATE_PROCEDURE_BODY,
    SQL_INSERT_STAGING,
    # Transportadora
    SQL_CREATE_STAGING_TRANSPORTADORA,
    SQL_CREATE_PROCEDURE_TRANSPORTADORA,
    SQL_CREATE_PROCEDURE_BODY_TRANSPORTADORA,
    SQL_INSERT_STAGING_TRANSPORTADORA,
    # Cliente
    SQL_CREATE_STAGING_CLIENTE,
    SQL_CREATE_PROCEDURE_CLIENTE,
    SQL_CREATE_PROCEDURE_BODY_CLIENTE,
    SQL_INSERT_STAGING_CLIENTE,
)


class PlanilhaImportador:

    # Drivers que suportam fast_executemany sem problemas
    DRIVERS_FAST_EXECUTEMANY = [
        "ODBC Driver 18",
        "ODBC Driver 17",
        "ODBC Driver 13",
        "ODBC Driver 11",
    ]

    def __init__(self, connection_string, progress_callback=None, log_callback=None):
        self.connection_string = connection_string
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.conn = None
        self.cancelar = False
        self._fast_executemany_ok = self._verificar_fast_executemany()

    def _verificar_fast_executemany(self):
        """Verifica se o driver suporta fast_executemany."""
        for driver in self.DRIVERS_FAST_EXECUTEMANY:
            if driver in self.connection_string:
                return True
        return False

    def _log(self, msg):
        if self.log_callback:
            self.log_callback(msg)
        print(msg)

    def _progresso(self, percentual, mensagem):
        if self.progress_callback:
            self.progress_callback(percentual, mensagem)

    # ----------------------------------------------------------
    # Conexao
    # ----------------------------------------------------------

    def conectar(self):
        self.conn = pyodbc.connect(self.connection_string, autocommit=True)
        cursor = self.conn.cursor()
        cursor.execute("SELECT DB_NAME()")
        self.database_name = cursor.fetchone()[0]
        cursor.close()


    def desconectar(self):
        if self.conn:
            try:
                self.conn.close()
            except Exception:
                pass
            self.conn = None

    @staticmethod
    def testar_conexao(connection_string):
        """Testa conexao e retorna (ok, mensagem)."""
        try:
            conn = pyodbc.connect(connection_string, autocommit=True, timeout=10)
            cursor = conn.cursor()
            cursor.execute("SELECT @@SERVERNAME")
            nome = cursor.fetchone()[0]
            conn.close()
            return True, f"Conectado a: {nome}"
        except Exception as e:
            return False, str(e)

    # ----------------------------------------------------------
    # Helpers SQL
    # ----------------------------------------------------------

    def _converter_valor(self, valor, tipo_sql):
        """Converte valor da celula Excel para tipo SQL."""
        if valor is None:
            return None
        val_str = str(valor).strip()
        if val_str == "":
            return None
        try:
            if tipo_sql in ("int", "smallint", "tinyint"):
                return int(float(val_str))
            elif tipo_sql == "decimal":
                return float(val_str.replace(",", "."))
            elif tipo_sql == "datetime":
                return val_str
            else:
                return val_str
        except (ValueError, TypeError):
            return val_str

    def _formatar_cnpjcpf(self, valor):
        """Formata CNPJ para formato SRPP (NNN.NNN.NNN/NNNN-NN, 19 chars).
        CPF nao e suportado pelo banco (constraint aceita apenas CNPJ ou NULL).
        """
        if valor is None:
            return None
        val = str(valor).strip()
        if not val:
            return None
        # Ja formatado 19 chars (formato SRPP)
        if len(val) == 19 and '/' in val:
            return val
        # Formatado padrao brasileiro 18 chars -> pad para 19
        if len(val) == 18 and '/' in val:
            return '0' + val
        # Extrair apenas digitos
        digitos = ''.join(c for c in val if c.isdigit())
        # CNPJ 14 digitos -> formatar para SRPP
        if len(digitos) == 14:
            d = digitos.zfill(15)
            return f"{d[0:3]}.{d[3:6]}.{d[6:9]}/{d[9:13]}-{d[13:15]}"
        # CPF (11 digitos) -> pad para 14 e formatar como CNPJ SRPP
        if len(digitos) == 11:
            digitos = digitos.zfill(14)
            d = digitos.zfill(15)
            return f"{d[0:3]}.{d[3:6]}.{d[6:9]}/{d[9:13]}-{d[13:15]}"
        # Nao e CNPJ/CPF reconhecido, retorna como veio
        return val

    # Campos onde valor 0 deve ser convertido para NULL (constraint > 0 OR NULL)
    _ZERO_TO_NULL = {"@ddd", "@codtransportadora", "@codrepresentante", "@qtdetabela1", "@qtdetabela2", "@qtdetabela3", "@codfamilia", "@codestilo", "@qtdemultipla", "@qtdeminima", "@qtdeetiquetas"}

    def _formatar_cep(self, valor):
        """Formata CEP para formato SRPP (NNNNN-NNN, 9 chars) ou None."""
        if valor is None:
            return None
        val = str(valor).strip()
        if not val:
            return None
        # Ja formatado 9 chars (NNNNN-NNN)
        if len(val) == 9 and val[5] == '-':
            return val
        # Extrair apenas digitos
        digitos = ''.join(c for c in val if c.isdigit())
        if not digitos:
            return None
        # Pad para 8 digitos e formatar
        digitos = digitos.zfill(8)
        return f"{digitos[0:5]}-{digitos[5:8]}"

    def _exec_sem_params(self, nome_proc):
        """Executa procedure sem parametros."""
        cursor = self.conn.cursor()
        try:
            cursor.execute(f"SET NOCOUNT ON; EXEC {nome_proc}")
            # Consumir todos os result sets para capturar erros deferidos
            while cursor.nextset():
                pass
            self._log(f"  OK: {nome_proc}")
            return True, ""
        except pyodbc.Error as e:
            msg = self._extrair_msg_erro(e)
            self._log(f"  ERRO: {nome_proc} - {msg}")
            return False, msg
        finally:
            cursor.close()

    def _exec_com_output(self, nome_proc, params):
        """
        Executa procedure com @msgretorno OUTPUT.
        params: lista de (nome_param, valor)
        Retorna (sucesso, msgretorno).
        """
        cursor = self.conn.cursor()
        try:
            placeholders = ", ".join(f"{nome}=?" for nome, _ in params)
            sql = (
                f"SET NOCOUNT ON; "
                f"DECLARE @msgretorno varchar(250); "
                f"EXEC {nome_proc} @msgretorno=@msgretorno OUTPUT, {placeholders}; "
                f"SELECT @msgretorno;"
            )
            valores = [v for _, v in params]
            cursor.execute(sql, valores)
            row = cursor.fetchone()
            return True, (row[0] if row else "")
        except pyodbc.Error as e:
            return False, self._extrair_msg_erro(e)
        finally:
            cursor.close()

    @staticmethod
    def _extrair_msg_erro(exc):
        """Extrai mensagem legivel de um pyodbc.Error."""
        msg = str(exc)
        # pyodbc retorna formato: "[codigo][driver] mensagem real"
        if "]" in msg:
            msg = msg.split("]")[-1].strip()
        return msg

    # ----------------------------------------------------------
    # Leitura da planilha
    # ----------------------------------------------------------

    def _ler_linhas_aba(self, wb, nome_aba):
        """Retorna (header_map, rows) de uma aba."""
        if nome_aba not in wb.sheetnames:
            return None, []
        sheet = wb[nome_aba]
        header = [cell.value for cell in sheet[1]]
        header_map = {name: idx for idx, name in enumerate(header) if name is not None}
        rows = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if all(v is None or str(v).strip() == "" for v in row):
                continue
            rows.append(row)
        return header_map, rows

    # ----------------------------------------------------------
    # Importacao: EMPRESA (formato especial chave-valor)
    # ----------------------------------------------------------

    def _importar_empresa(self, wb, sobreescreve):
        t0 = time.time()
        self._log("--- Importando EMPRESA ---")
        config = MAPA_ABAS["EMPRESA"]

        if "EMPRESA" not in wb.sheetnames:
            self._log("  Aba EMPRESA nao encontrada!")
            return 0, 1

        sheet = wb["EMPRESA"]
        linhas = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[0] is None:
                continue
            try:
                codigo = int(row[0])
            except (ValueError, TypeError):
                continue
            descricao = str(row[1]).strip() if row[1] is not None else ""
            configuracao = str(row[2]).strip() if row[2] is not None else ""
            linhas.append((codigo, descricao, configuracao))

        if not linhas:
            self._log("  Nenhuma linha de configuracao encontrada")
            return 0, 1

        total = len(linhas)
        erros = 0
        sucesso = 0

        # Passo 1: limpar
        self._exec_sem_params(config["pkunica1"])

        # Passo 2: PKUnica2
        self._log(f"  Registrando {total} codigos...")
        for codigo, _, _ in linhas:
            if self.cancelar:
                return sucesso, erros
            params = [("@sobreescreve", sobreescreve), ("@codigo", codigo)]
            ok, msg = self._exec_com_output(config["pkunica2"], params)
            if not ok:
                self._log(f"  ERRO PKUnica2 cod={codigo}: {msg}")
                erros += 1

        # Passo 3: importar cada parametro
        self._log(f"  Importando {total} parametros...")
        for i, (codigo, descricao, configuracao) in enumerate(linhas):
            if self.cancelar:
                return sucesso, erros
            params = [
                ("@sobreescreve", sobreescreve),
                ("@Codigo", codigo), ("@Codigo_rec", 1),
                ("@Descricao", descricao), ("@Descricao_rec", 1),
                ("@Configuracao", configuracao), ("@Configuracao_rec", 1),
            ]
            ok, msg = self._exec_com_output(config["procedure"], params)
            if ok:
                sucesso += 1
            else:
                erros += 1
                self._log(f"  ERRO cod={codigo}: {msg}")

        # Passo 4: ConfigurarEmpresa
        self._log("  Executando ConfigurarEmpresa...")
        cursor = self.conn.cursor()
        try:
            sql = "DECLARE @msgretorno varchar(250); EXEC ConfigurarEmpresa @msgretorno=@msgretorno OUTPUT; SELECT @msgretorno;"
            cursor.execute(sql)
            row = cursor.fetchone()
            self._log(f"  ConfigurarEmpresa: {row[0] if row else 'OK'}")
        except pyodbc.Error as e:
            erros += 1
            self._log(f"  ERRO ConfigurarEmpresa: {self._extrair_msg_erro(e)}")
        finally:
            cursor.close()

        elapsed = time.time() - t0
        self._log(f"  EMPRESA: {sucesso} ok, {erros} erros | {total} linhas em {elapsed:.1f}s")
        return sucesso, erros

    # ----------------------------------------------------------
    # Importacao: aba padrao (fluxo PKUnica1 -> PKUnica2 -> Proc)
    # ----------------------------------------------------------

    def _importar_aba_padrao(self, wb, nome_aba, sobreescreve, prog_base, prog_range):
        t0 = time.time()
        self._log(f"--- Importando {nome_aba} ---")
        config = MAPA_ABAS[nome_aba]
        colunas = config["colunas"]
        pk_cols = config["pk_columns"]

        header_map, rows = self._ler_linhas_aba(wb, nome_aba)
        if header_map is None:
            self._log(f"  Aba {nome_aba} nao encontrada!")
            return 0, 1

        # Filtrar linhas com PK preenchida
        pk0_name = pk_cols[0][0]
        pk0_idx = header_map.get(pk0_name)
        if pk0_idx is None:
            self._log(f"  Coluna PK '{pk0_name}' nao encontrada em {nome_aba}")
            return 0, 1

        linhas = [r for r in rows if pk0_idx < len(r) and r[pk0_idx] is not None and str(r[pk0_idx]).strip()]
        total = len(linhas)
        if total == 0:
            self._log(f"  Nenhuma linha valida em {nome_aba}")
            return 0, 0

        erros = 0
        sucesso = 0

        # Passo 1: limpar
        self._exec_sem_params(config["pkunica1"])

        # Passo 2: PKUnica2
        self._progresso(prog_base, f"Registrando PKs de {nome_aba}...")
        for row in linhas:
            if self.cancelar:
                return sucesso, erros
            params = [("@sobreescreve", sobreescreve)]
            for col_excel, param_sql, tipo_sql in pk_cols:
                idx = header_map.get(col_excel)
                val = row[idx] if idx is not None and idx < len(row) else None
                params.append((param_sql, self._converter_valor(val, tipo_sql)))
            ok, msg = self._exec_com_output(config["pkunica2"], params)
            if not ok:
                pk_str = self._pk_str(row, pk_cols, header_map)
                self._log(f"  ERRO PKUnica2 {nome_aba} PK=[{pk_str}]: {msg}")
                erros += 1

        # Passo 3: importacao principal
        for i, row in enumerate(linhas):
            if self.cancelar:
                return sucesso, erros

            if total > 0:
                pct = prog_base + int((i / total) * prog_range)
                self._progresso(pct, f"Importando {nome_aba}... {i+1}/{total}")

            params = [("@sobreescreve", sobreescreve)]
            for col_excel, param_sql, tipo_sql in colunas:
                idx = header_map.get(col_excel)
                if idx is not None and idx < len(row):
                    valor = self._converter_valor(row[idx], tipo_sql)
                    if param_sql.lower() == "@cnpjcpf":
                        valor = self._formatar_cnpjcpf(valor)
                    elif param_sql.lower() == "@cep":
                        valor = self._formatar_cep(valor)
                    # Campos com constraint > 0 OR NULL: converter 0 para NULL
                    if param_sql.lower() in self._ZERO_TO_NULL and valor == 0:
                        valor = None
                    # Strings vazias -> NULL (constraints NaoVazioNull)
                    if isinstance(valor, str) and valor.strip() == "":
                        valor = None
                    rec = 1  # coluna mapeada
                else:
                    valor = None
                    rec = None  # coluna nao mapeada
                params.append((param_sql, valor))
                params.append((f"{param_sql}_rec", rec))

            ok, msg = self._exec_com_output(config["procedure"], params)
            if ok:
                sucesso += 1
            else:
                erros += 1
                pk_str = self._pk_str(row, pk_cols, header_map)
                self._log(f"  ERRO {nome_aba} linha {i+1} PK=[{pk_str}]: {msg}")

        elapsed = time.time() - t0
        self._log(f"  {nome_aba}: {sucesso} ok, {erros} erros | {total} linhas em {elapsed:.1f}s")
        return sucesso, erros

    # ----------------------------------------------------------
    # Importacao em lote: PRODUTOS (staging + procedure batch)
    # ----------------------------------------------------------

    def _setup_batch_objects(self):
        """Cria staging table e procedure de lote se nao existirem."""
        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_STAGING)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_PROCEDURE)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_PROCEDURE_BODY)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

    def _importar_produtos_lote(self, wb, sobreescreve, prog_base, prog_range):
        """Importa PRODUTOS usando staging table + procedure batch."""
        t0 = time.time()
        self._log("--- Importando PRODUTOS (modo lote) ---")

        config = MAPA_ABAS["PRODUTOS"]
        colunas = config["colunas"]

        header_map, rows = self._ler_linhas_aba(wb, "PRODUTOS")
        if header_map is None:
            self._log("  Aba PRODUTOS nao encontrada!")
            return 0, 1

        # Filtrar linhas com PK preenchida (CodProduto)
        pk0_idx = header_map.get("CodProduto")
        if pk0_idx is None:
            self._log("  Coluna PK 'CodProduto' nao encontrada em PRODUTOS")
            return 0, 1

        linhas = [r for r in rows if pk0_idx < len(r) and r[pk0_idx] is not None and str(r[pk0_idx]).strip()]
        total = len(linhas)
        if total == 0:
            self._log("  Nenhuma linha valida em PRODUTOS")
            return 0, 0

        # 1. Criar objetos SQL se necessario
        self._progresso(prog_base, "Preparando importacao em lote...")
        self._setup_batch_objects()

        # 2. Limpar staging
        cursor = self.conn.cursor()
        try:
            cursor.execute("TRUNCATE TABLE ImportaProdutoAmbos_Staging")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        # Resetar IDENTITY para que 'linha' comece em 1
        cursor = self.conn.cursor()
        try:
            cursor.execute("DBCC CHECKIDENT ('ImportaProdutoAmbos_Staging', RESEED, 0)")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        # 3. Montar parametros para bulk insert
        self._progresso(prog_base + 2, f"Carregando {total} produtos no staging...")

        # Ordem das colunas no INSERT: mesma de SQL_INSERT_STAGING
        col_order = [
            "CodProduto", "CodAuxiliarProduto", "Produto",
            "PrecoTabela1", "PrecoTabela2", "PrecoTabela3",
            "QtdeEstoqueAtual", "QtdeEstoqueFuturo", "DtEstoqueFuturo",
            "LimiteDescIndividual", "AliquotaIPI", "MultiploGrade", "DescontoGrade",
            "PathFotografia", "CodFilial", "PrecoPromocional", "TipoVendaSemEstoque",
            "CodFamilia", "CodEstilo", "QtdeMultipla", "QtdeMinima",
            "QtdeTabela1", "QtdeTabela2", "QtdeTabela3", "QtdeEtiquetas",
        ]

        # Map excel column -> (index_in_col_order, tipo_sql)
        col_type = {c[0]: c[2] for c in colunas}
        
        # NOVO: Mapear NomeExcel para @NomeParametroSQL para checar a lista _ZERO_TO_NULL
        col_sql_param = {c[0]: c[1].lower() for c in colunas} 

        batch_rows = []
        for row in linhas:
            vals = []
            for col_name in col_order:
                idx = header_map.get(col_name)
                if idx is not None and idx < len(row):
                    valor = self._converter_valor(row[idx], col_type.get(col_name, "varchar"))
                    
                    # --- CORRECAO AQUI ---
                    # Se for string vazia, vira None
                    if isinstance(valor, str) and valor.strip() == "":
                        valor = None
                    
                    # Se estiver na lista de Zeros Proibidos e for 0, vira None (NULL no banco)
                    param_name = col_sql_param.get(col_name, "")
                    if param_name in self._ZERO_TO_NULL and valor == 0:
                        valor = None
                    # ---------------------

                else:
                    valor = None
                vals.append(valor)
            batch_rows.append(vals)

        # 4. Bulk insert (fast_executemany se driver suportar)
        cursor = self.conn.cursor()
        if self._fast_executemany_ok:
            cursor.fast_executemany = True
        try:
            cursor.executemany(SQL_INSERT_STAGING, batch_rows)
        finally:
            cursor.close()

        self._log(f"  {total} linhas inseridas no staging.")
        self._progresso(prog_base + int(prog_range * 0.3), "Executando validacao e importacao em lote...")

        # 5. Chamar a procedure batch
        cursor = self.conn.cursor()
        try:
            cursor.execute(f"SET NOCOUNT ON; EXEC ImportaProdutoAmbos_Lote @sobreescreve={sobreescreve}")
            results = cursor.fetchall()
        finally:
            cursor.close()

        # 6. Processar resultados
        sucesso = 0
        erros = 0
        for row in results:
            linha_num = row[0]
            codprod = row[1]
            codaux = row[2]
            status = row[3]
            mensagem = row[4]

            if status == "OK":
                sucesso += 1
            else:
                erros += 1
                self._log(
    f"  ERRO PRODUTOS | DB={self.database_name} | PROC=ImportaProdutoAmbos_Lote "
    f"| linha {linha_num} PK=[{codprod}, {codaux}]: {mensagem}"
)


        # Limpar staging table após processamento
        cursor = self.conn.cursor()
        try:
            cursor.execute("TRUNCATE TABLE ImportaProdutoAmbos_Staging")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        elapsed = time.time() - t0
        self._log(f"  PRODUTOS: {sucesso} ok, {erros} erros | {total} linhas em {elapsed:.1f}s")
        self._progresso(prog_base + prog_range, f"PRODUTOS concluido: {sucesso} ok, {erros} erros")
        return sucesso, erros

    # ----------------------------------------------------------
    # Importacao em lote: TRANSPORTADORA
    # ----------------------------------------------------------

    def _setup_batch_objects_transportadora(self):
        """Cria staging table e procedure de lote para transportadora."""
        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_STAGING_TRANSPORTADORA)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_PROCEDURE_TRANSPORTADORA)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_PROCEDURE_BODY_TRANSPORTADORA)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

    def _importar_transportadoras_lote(self, wb, sobreescreve, prog_base, prog_range):
        """Importa TRANSP usando staging table + procedure batch."""
        t0 = time.time()
        self._log("--- Importando TRANSP (modo lote) ---")

        config = MAPA_ABAS["TRANSP"]
        header_map, rows = self._ler_linhas_aba(wb, "TRANSP")
        if header_map is None:
            self._log("  Aba TRANSP nao encontrada!")
            return 0, 1

        pk0_idx = header_map.get("CodTransportadora")
        if pk0_idx is None:
            self._log("  Coluna PK 'CodTransportadora' nao encontrada em TRANSP")
            return 0, 1

        linhas = [r for r in rows if pk0_idx < len(r) and r[pk0_idx] is not None and str(r[pk0_idx]).strip()]
        total = len(linhas)
        if total == 0:
            self._log("  Nenhuma linha valida em TRANSP")
            return 0, 0

        # 1. Setup
        self._progresso(prog_base, "Preparando importacao em lote (TRANSP)...")
        self._setup_batch_objects_transportadora()

        # 2. Limpar staging
        cursor = self.conn.cursor()
        try:
            cursor.execute("TRUNCATE TABLE ImportaTransportadora_Staging")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute("DBCC CHECKIDENT ('ImportaTransportadora_Staging', RESEED, 0)")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        # 3. Montar dados
        self._progresso(prog_base + 2, f"Carregando {total} transportadoras no staging...")
        col_order = ["CodTransportadora", "Transportadora", "TransportadoraPadrao"]

        batch_rows = []
        for row in linhas:
            vals = []
            for col_name in col_order:
                idx = header_map.get(col_name)
                if idx is not None and idx < len(row):
                    valor = self._converter_valor(row[idx], "varchar")
                    if isinstance(valor, str) and valor.strip() == "":
                        valor = None
                else:
                    valor = None
                vals.append(valor)
            batch_rows.append(vals)

        # 4. Bulk insert (fast_executemany se driver suportar)
        cursor = self.conn.cursor()
        if self._fast_executemany_ok:
            cursor.fast_executemany = True
        try:
            cursor.executemany(SQL_INSERT_STAGING_TRANSPORTADORA, batch_rows)
        finally:
            cursor.close()

        self._log(f"  {total} linhas inseridas no staging.")
        self._progresso(prog_base + int(prog_range * 0.3), "Executando importacao em lote (TRANSP)...")

        # 5. Executar procedure
        cursor = self.conn.cursor()
        try:
            cursor.execute(f"SET NOCOUNT ON; EXEC ImportaTransportadora_Lote @sobreescreve={sobreescreve}")
            results = cursor.fetchall()
        finally:
            cursor.close()

        # 6. Processar resultados
        sucesso = 0
        erros = 0
        for row in results:
            linha_num = row[0]
            codtransp = row[1]
            status = row[2]
            mensagem = row[3]

            if status == "OK":
                sucesso += 1
            else:
                erros += 1
                self._log(f"  ERRO TRANSP | DB={self.database_name} | linha {linha_num} PK=[{codtransp}]: {mensagem}")

        # 7. Limpar staging
        cursor = self.conn.cursor()
        try:
            cursor.execute("TRUNCATE TABLE ImportaTransportadora_Staging")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        elapsed = time.time() - t0
        self._log(f"  TRANSP: {sucesso} ok, {erros} erros | {total} linhas em {elapsed:.1f}s")
        self._progresso(prog_base + prog_range, f"TRANSP concluido: {sucesso} ok, {erros} erros")
        return sucesso, erros

    # ----------------------------------------------------------
    # Importacao em lote: CLIENTE
    # ----------------------------------------------------------

    def _setup_batch_objects_cliente(self):
        """Cria staging table e procedure de lote para cliente."""
        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_STAGING_CLIENTE)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_PROCEDURE_CLIENTE)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute(SQL_CREATE_PROCEDURE_BODY_CLIENTE)
            while cursor.nextset():
                pass
        finally:
            cursor.close()

    def _importar_clientes_lote(self, wb, sobreescreve, prog_base, prog_range):
        """Importa CLIENTES usando staging table + procedure batch."""
        t0 = time.time()
        self._log("--- Importando CLIENTES (modo lote) ---")

        config = MAPA_ABAS["CLIENTES"]
        header_map, rows = self._ler_linhas_aba(wb, "CLIENTES")
        if header_map is None:
            self._log("  Aba CLIENTES nao encontrada!")
            return 0, 1

        pk0_idx = header_map.get("CodCliente")
        if pk0_idx is None:
            self._log("  Coluna PK 'CodCliente' nao encontrada em CLIENTES")
            return 0, 1

        linhas = [r for r in rows if pk0_idx < len(r) and r[pk0_idx] is not None and str(r[pk0_idx]).strip()]
        total = len(linhas)
        if total == 0:
            self._log("  Nenhuma linha valida em CLIENTES")
            return 0, 0

        # 1. Setup
        self._progresso(prog_base, "Preparando importacao em lote (CLIENTES)...")
        self._setup_batch_objects_cliente()

        # 2. Limpar staging
        cursor = self.conn.cursor()
        try:
            cursor.execute("TRUNCATE TABLE ImportaCliente_Staging")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute("DBCC CHECKIDENT ('ImportaCliente_Staging', RESEED, 0)")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        # 3. Montar dados
        self._progresso(prog_base + 2, f"Carregando {total} clientes no staging...")
        col_order = [
            "CodCliente", "CodRepresentante", "NomeFantasia", "RazaoSocial",
            "CNPJCPF", "IERG", "Logradouro", "Bairro", "Cidade", "UF", "CEP",
            "DDD", "Telefone1", "Telefone2", "FAX", "NomeContato",
            "NomeTransportadora", "Observacao", "EMail", "PrecoTabela", "CodTransportadora"
        ]

        batch_rows = []
        for row in linhas:
            vals = []
            for col_name in col_order:
                idx = header_map.get(col_name)
                if idx is not None and idx < len(row):
                    tipo = "int" if col_name in ("CodCliente", "CodRepresentante", "DDD", "PrecoTabela", "CodTransportadora") else "varchar"
                    valor = self._converter_valor(row[idx], tipo)
                    if isinstance(valor, str) and valor.strip() == "":
                        valor = None
                    # CNPJCPF e CEP: formatar
                    if col_name == "CNPJCPF" and valor:
                        valor = self._formatar_cnpjcpf(valor)
                    elif col_name == "CEP" and valor:
                        valor = self._formatar_cep(valor)
                    # CodRepresentante e CodTransportadora: 0 -> NULL
                    if col_name in ("CodRepresentante", "CodTransportadora", "DDD") and valor == 0:
                        valor = None
                else:
                    valor = None
                vals.append(valor)
            batch_rows.append(vals)

        # 4. Bulk insert (fast_executemany se driver suportar)
        cursor = self.conn.cursor()
        if self._fast_executemany_ok:
            cursor.fast_executemany = True
        try:
            cursor.executemany(SQL_INSERT_STAGING_CLIENTE, batch_rows)
        finally:
            cursor.close()

        self._log(f"  {total} linhas inseridas no staging.")
        self._progresso(prog_base + int(prog_range * 0.3), "Executando importacao em lote (CLIENTES)...")

        # 5. Executar procedure
        cursor = self.conn.cursor()
        try:
            cursor.execute(f"SET NOCOUNT ON; EXEC ImportaCliente_Lote @sobreescreve={sobreescreve}")
            results = cursor.fetchall()
        finally:
            cursor.close()

        # 6. Processar resultados
        sucesso = 0
        erros = 0
        for row in results:
            linha_num = row[0]
            codcliente = row[1]
            status = row[2]
            mensagem = row[3]

            if status == "OK":
                sucesso += 1
            else:
                erros += 1
                self._log(f"  ERRO CLIENTES | DB={self.database_name} | linha {linha_num} PK=[{codcliente}]: {mensagem}")

        # 7. Limpar staging
        cursor = self.conn.cursor()
        try:
            cursor.execute("TRUNCATE TABLE ImportaCliente_Staging")
            while cursor.nextset():
                pass
        finally:
            cursor.close()

        elapsed = time.time() - t0
        self._log(f"  CLIENTES: {sucesso} ok, {erros} erros | {total} linhas em {elapsed:.1f}s")
        self._progresso(prog_base + prog_range, f"CLIENTES concluido: {sucesso} ok, {erros} erros")
        return sucesso, erros

    def _pk_str(self, row, pk_cols, header_map):
        """Monta string com valores das PKs para log."""
        vals = []
        for col_excel, _, _ in pk_cols:
            idx = header_map.get(col_excel)
            val = row[idx] if idx is not None and idx < len(row) else "?"
            vals.append(str(val))
        return ", ".join(vals)

    # ----------------------------------------------------------
    # Backup
    # ----------------------------------------------------------

    def _fazer_backup(self, operacao="Apagar Pedidos e Cadastros e Configuracoes"):
        """
        Faz backup do banco antes de operacoes destrutivas.
        Salva no caminho padrao da instancia SQL Server.
        Formato: {BANCO} - {EMPRESA} - {DATA} {HORA} - {OPERACAO}.bak
        """
        self._log("Fazendo backup do banco de dados...")

        backup_path = None
        cursor = self.conn.cursor()
        try:
            # 1. Tentar obter caminho de backup da instancia (SQL 2012+)
            cursor.execute("SELECT SERVERPROPERTY('InstanceDefaultBackupPath')")
            row = cursor.fetchone()
            backup_path = row[0] if row and row[0] else None
        except Exception:
            pass
        finally:
            cursor.close()

        # Fallback: usar diretorio do arquivo de dados do banco
        if not backup_path:
            cursor = self.conn.cursor()
            try:
                cursor.execute(f"""
                    SELECT LEFT(physical_name, LEN(physical_name) - CHARINDEX('\\', REVERSE(physical_name)))
                    FROM sys.master_files
                    WHERE database_id = DB_ID('{self.database_name}') AND type = 0
                """)
                row = cursor.fetchone()
                backup_path = row[0] if row else None
            except Exception:
                pass
            finally:
                cursor.close()

        if not backup_path:
            self._log("  AVISO: Nao foi possivel obter caminho de backup")
            return False

        cursor = self.conn.cursor()
        try:
            # 2. Buscar nome da empresa
            cursor.execute("SELECT Empresa FROM empresa WHERE codempresa = 1")
            row = cursor.fetchone()
            empresa_nome = row[0] if row else "SEM_EMPRESA"
            # Limpar caracteres invalidos para nome de arquivo
            empresa_nome = empresa_nome.replace("/", "-").replace("\\", "-").replace(":", "-")
            empresa_nome = empresa_nome.replace("*", "-").replace("?", "-").replace("\"", "-")
            empresa_nome = empresa_nome.replace("<", "-").replace(">", "-").replace("|", "-")
        finally:
            cursor.close()

        # 3. Gerar nome do arquivo
        agora = datetime.now()
        data_hora = agora.strftime("%Y-%m-%d %Hh%Mm%S")
        nome_arquivo = f"{self.database_name} - {empresa_nome} - {data_hora} - {operacao}.bak"
        caminho_completo = os.path.join(backup_path, nome_arquivo)

        self._log(f"  Destino: {caminho_completo}")

        cursor = self.conn.cursor()
        try:
            # 4. Executar backup
            sql = f"BACKUP DATABASE [{self.database_name}] TO DISK = N'{caminho_completo}' WITH NOFORMAT, INIT, NAME = N'{self.database_name}-Full', SKIP, NOREWIND, NOUNLOAD, STATS = 10"
            cursor.execute(sql)
            # Consumir todos os result sets (mensagens de progresso)
            while cursor.nextset():
                pass
            self._log(f"  Backup concluido com sucesso!")
            return True
        except pyodbc.Error as e:
            self._log(f"  ERRO no backup: {self._extrair_msg_erro(e)}")
            return False
        finally:
            cursor.close()

    # ----------------------------------------------------------
    # Limpeza
    # ----------------------------------------------------------

    def excluir_tudo(self):
        # Fazer backup antes de excluir
        self._fazer_backup("Apagar Pedidos e Cadastros e Configuracoes")

        # Limpar tabelas que a procedure nao deleta mas que tem FK para as que ela deleta
        self._log("Limpando tabelas com FK dependente...")
        cursor = self.conn.cursor()
        try:
            cursor.execute("SET NOCOUNT ON; DELETE FROM RestricaoCondicaoPagamento")
            while cursor.nextset():
                pass
            self._log("  OK: RestricaoCondicaoPagamento limpa")
        except pyodbc.Error as e:
            self._log(f"  AVISO: RestricaoCondicaoPagamento - {self._extrair_msg_erro(e)}")
        finally:
            cursor.close()

        cursor = self.conn.cursor()
        try:
            cursor.execute("SET NOCOUNT ON; DELETE FROM Representantes_X_AspNetUsers")
            while cursor.nextset():
                pass
            self._log("  OK: Representantes_X_AspNetUsers limpa")
        except pyodbc.Error as e:
            self._log(f"  AVISO: Representantes_X_AspNetUsers - {self._extrair_msg_erro(e)}")
        finally:
            cursor.close()

        self._log("Executando ExcluiPedidosCadastrosConfiguracao...")
        ok, msg = self._exec_sem_params("ExcluiPedidosCadastrosConfiguracao")
        if ok:
            self._log("Todos os dados foram excluidos.")
        return ok

    def limpar_auxiliares(self, abas):
        self._log("Limpando tabelas auxiliares...")
        for aba in abas:
            config = MAPA_ABAS.get(aba)
            if config:
                self._exec_sem_params(config["pkunica1"])

    # ----------------------------------------------------------
    # Fluxo principal
    # ----------------------------------------------------------

    def importar(self, arquivo_excel, abas_selecionadas, sobreescreve,
                 excluir_tudo=False, limpar_auxiliares=False):
        """
        Executa importacao completa.
        Retorna dict {aba: {"sucesso": N, "erros": N}} ou {"ERRO_GERAL": msg}.
        """
        resultados = {}
        self.cancelar = False
        t_total = time.time()

        try:
            self._progresso(0, "Conectando ao SQL Server...")
            self.conectar()
            self._log("Conectado ao SQL Server.")

            self._progresso(2, "Carregando planilha...")
            wb = load_workbook(arquivo_excel, data_only=True)
            self._log(f"Planilha: {os.path.basename(arquivo_excel)}")

            if excluir_tudo:
                self._progresso(3, "Excluindo todos os dados...")
                if not self.excluir_tudo():
                    return {"ERRO_GERAL": "Falha ao excluir dados"}

            if limpar_auxiliares:
                self._progresso(4, "Limpando tabelas auxiliares...")
                self.limpar_auxiliares(abas_selecionadas)

            abas_ordenadas = [a for a in ORDEM_IMPORTACAO if a in abas_selecionadas]
            total_abas = len(abas_ordenadas)

            for idx, nome_aba in enumerate(abas_ordenadas):
                if self.cancelar:
                    self._log("IMPORTACAO CANCELADA")
                    break

                prog_base = 5 + int((idx / total_abas) * 90)
                prog_range = int(90 / total_abas)
                self._progresso(prog_base, f"Importando {nome_aba}...")

                if nome_aba == "EMPRESA":
                    s, e = self._importar_empresa(wb, sobreescreve)
                elif nome_aba == "PRODUTOS":
                    s, e = self._importar_produtos_lote(wb, sobreescreve, prog_base, prog_range)
                elif nome_aba == "TRANSP":
                    s, e = self._importar_transportadoras_lote(wb, sobreescreve, prog_base, prog_range)
                elif nome_aba == "CLIENTES":
                    s, e = self._importar_clientes_lote(wb, sobreescreve, prog_base, prog_range)
                else:
                    s, e = self._importar_aba_padrao(wb, nome_aba, sobreescreve, prog_base, prog_range)

                resultados[nome_aba] = {"sucesso": s, "erros": e}

            elapsed_total = time.time() - t_total
            minutos, segundos = divmod(int(elapsed_total), 60)
            self._log(f"Tempo total: {minutos}m {segundos}s")
            self._progresso(100, "Importacao concluida!")
            wb.close()

        except Exception as e:
            self._log(f"ERRO FATAL: {e}")
            import traceback
            traceback.print_exc()
            resultados["ERRO_GERAL"] = str(e)
        finally:
            self.desconectar()

        return resultados
