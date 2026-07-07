"""
SQL para importacao em lote.
- ImportaProdutoAmbos_Lote: ~34.000 roundtrips -> 1 bulk insert + 1 call
- ImportaTransportadora_Lote: importacao em lote de transportadoras
- ImportaCliente_Lote: importacao em lote de clientes
"""

# ----------------------------------------------------------------
# Remove triggers de SyncGeral (obsoleta em novos bancos)
# Dropa qualquer trigger DML que referencie SyncGeral, em qualquer tabela.
# Seguro: so executa se a tabela SyncGeral existir.
# ----------------------------------------------------------------
SQL_REMOVE_SYNC_TRIGGERS = """
IF OBJECT_ID('dbo.SyncGeral', 'U') IS NOT NULL
BEGIN
    DECLARE @sql NVARCHAR(MAX) = N''
    SELECT @sql = @sql + N'IF OBJECT_ID(N''[dbo].[' + t.name + N']'', N''TR'') IS NOT NULL DROP TRIGGER [dbo].[' + t.name + N'];'
    FROM sys.triggers t
    INNER JOIN sys.sql_modules m ON t.object_id = m.object_id
    WHERE m.definition LIKE N'%SyncGeral%'
      AND t.type = 'TR'
    IF LEN(@sql) > 0
        EXEC sp_executesql @sql
END
"""

# ----------------------------------------------------------------
# Staging table: recebe os dados do Python via fast_executemany
# ----------------------------------------------------------------
SQL_CREATE_STAGING = """
IF OBJECT_ID('dbo.ImportaProdutoAmbos_Staging','U') IS NULL
BEGIN
    CREATE TABLE dbo.ImportaProdutoAmbos_Staging (
        linha               int IDENTITY(1,1) PRIMARY KEY,
        codproduto          varchar(20),
        codauxiliarproduto  varchar(20),
        produto             varchar(45),
        precotabela1        decimal(8,2),
        precotabela2        decimal(8,2),
        precotabela3        decimal(8,2),
        qtdeestoqueatual    int,
        qtdeestoquefuturo   int,
        dtestoquefuturo     datetime,
        limitedescindividual decimal(5,2),
        aliquotaipi         decimal(5,2),
        multiplograde       int,
        descontograde       decimal(5,2),
        pathfotografia      varchar(200),
        codfilial           int,
        precopromocional    char(1),
        tipovendasemestoque char(1),
        codfamilia          int,
        codestilo           int,
        qtdemultipla        int,
        qtdeminima          int,
        qtdetabela1         int,
        qtdetabela2         int,
        qtdetabela3         int,
        qtdeetiquetas       smallint
    );
END
"""

# ----------------------------------------------------------------
# Procedure em lote: valida e importa todos os produtos de uma vez
# ----------------------------------------------------------------
SQL_CREATE_PROCEDURE = """
IF OBJECT_ID('dbo.ImportaProdutoAmbos_Lote','P') IS NOT NULL
    DROP PROCEDURE dbo.ImportaProdutoAmbos_Lote;
"""

SQL_CREATE_PROCEDURE_BODY = """
CREATE PROCEDURE dbo.ImportaProdutoAmbos_Lote
    @sobreescreve tinyint = 2
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @tipocodproduto char(1)
    DECLARE @tipocodauxproduto char(1)
    DECLARE @controladtentrega char(1)

    SELECT
        @tipocodproduto = tipocodproduto,
        @tipocodauxproduto = tipocodauxproduto,
        @controladtentrega = controladtentrega
    FROM empresa WHERE codempresa = 1

    -- ============================================================
    -- Work table com coluna de erro
    -- ============================================================
    SELECT
        linha, codproduto, codauxiliarproduto, produto,
        precotabela1, precotabela2, precotabela3,
        qtdeestoqueatual, qtdeestoquefuturo, dtestoquefuturo,
        limitedescindividual, aliquotaipi, multiplograde, descontograde,
        pathfotografia, codfilial, precopromocional, tipovendasemestoque,
        codfamilia, codestilo, qtdemultipla, qtdeminima,
        qtdetabela1, qtdetabela2, qtdetabela3, qtdeetiquetas,
        CAST(NULL AS varchar(250)) AS erro
    INTO #W
    FROM ImportaProdutoAmbos_Staging

    CREATE CLUSTERED INDEX IX_W_linha ON #W (linha)
    CREATE NONCLUSTERED INDEX IX_W_codproduto ON #W (codproduto)
    CREATE NONCLUSTERED INDEX IX_W_codaux ON #W (codauxiliarproduto)

    -- Empresa nao configurada: aborta tudo
    IF @tipocodproduto IS NULL
    BEGIN
        SELECT linha, codproduto, codauxiliarproduto, 'ERRO' AS status,
            'Nao foi possivel importar Produtos porque Empresa nao foi configurada.' AS mensagem
        FROM #W ORDER BY linha
        RETURN
    END

    -- ============================================================
    -- Aplicar defaults (mesmo comportamento da procedure original)
    -- ============================================================
    UPDATE #W SET codfilial = 1 WHERE codfilial IS NULL
    UPDATE #W SET precopromocional = 'N' WHERE precopromocional IS NULL
    UPDATE #W SET tipovendasemestoque = 'C' WHERE tipovendasemestoque IS NULL
    IF @sobreescreve = 2
        UPDATE #W SET precotabela1 = 0.00 WHERE precotabela1 IS NULL

    -- ============================================================
    -- Popula ProdutoAmbos_PKUnica (equivale PKUnica1 + PKUnica2)
    -- ============================================================
    DELETE FROM ProdutoAmbos_PKUnica

    INSERT INTO ProdutoAmbos_PKUnica (codproduto, codauxiliarproduto)
    SELECT codproduto, codauxiliarproduto
    FROM #W
    WHERE codproduto IS NOT NULL

    -- ============================================================
    -- VALIDACOES (cada UPDATE so marca linhas sem erro anterior)
    -- ============================================================

    -- 1. CodProduto obrigatorio
    UPDATE #W SET erro = 'CodProduto nao pode ser null.'
    WHERE codproduto IS NULL AND erro IS NULL

    -- 2. CodProduto != CodAuxiliarProduto
    UPDATE #W SET erro = 'CodProduto e CodAuxiliarProduto devem ser diferentes.'
    WHERE codauxiliarproduto IS NOT NULL
      AND codproduto = codauxiliarproduto
      AND erro IS NULL

    -- 3. Duplicidade de CodProduto no arquivo
    ;WITH dups AS (
        SELECT codproduto, COUNT(*)-1 AS qtd
        FROM #W WHERE codproduto IS NOT NULL
        GROUP BY codproduto HAVING COUNT(*) > 1
    )
    UPDATE w SET erro = 'Produto ' + w.codproduto + ' tem ' +
        CONVERT(varchar(5), d.qtd) +
        CASE WHEN d.qtd = 1 THEN ' duplicidade' ELSE ' duplicidades' END +
        ' no arquivo.'
    FROM #W w INNER JOIN dups d ON w.codproduto = d.codproduto
    WHERE w.erro IS NULL

    -- 4. Duplicidade de CodAuxiliarProduto no arquivo
    ;WITH dups AS (
        SELECT codauxiliarproduto, COUNT(*)-1 AS qtd
        FROM #W WHERE codauxiliarproduto IS NOT NULL
        GROUP BY codauxiliarproduto HAVING COUNT(*) > 1
    )
    UPDATE w SET erro = 'O Codigo Auxiliar ' + w.codauxiliarproduto +
        ' do Produto ' + w.codproduto + ' tem ' +
        CONVERT(varchar(5), d.qtd) +
        CASE WHEN d.qtd = 1 THEN ' duplicidade' ELSE ' duplicidades' END +
        ' no arquivo.'
    FROM #W w INNER JOIN dups d ON w.codauxiliarproduto = d.codauxiliarproduto
    WHERE w.erro IS NULL

    -- 5. CodAuxiliarProduto usado como CodProduto em outra linha
    IF @tipocodproduto = 'A' AND @tipocodauxproduto = 'A'
    BEGIN
        UPDATE w1 SET erro = 'O Codigo Auxiliar ' + w1.codauxiliarproduto +
            ' do Produto ' + w1.codproduto +
            ' nao pode ser utilizado como Codigo Principal em nenhum outro produto.'
        FROM #W w1
        WHERE w1.codauxiliarproduto IS NOT NULL
          AND EXISTS (SELECT 1 FROM #W w2
                      WHERE w2.codproduto = w1.codauxiliarproduto
                        AND w2.linha != w1.linha)
          AND w1.erro IS NULL
    END
    ELSE IF @tipocodproduto = 'N' AND @tipocodauxproduto = 'N'
    BEGIN
        UPDATE w1 SET erro = 'O Codigo Auxiliar ' + w1.codauxiliarproduto +
            ' do Produto ' + w1.codproduto +
            ' nao pode ser utilizado como Codigo Principal em nenhum outro produto.'
        FROM #W w1
        WHERE w1.codauxiliarproduto IS NOT NULL
          AND EXISTS (SELECT 1 FROM #W w2
                      WHERE CONVERT(bigint, w2.codproduto) = CONVERT(bigint, w1.codauxiliarproduto)
                        AND w2.linha != w1.linha)
          AND w1.erro IS NULL
    END
    ELSE IF @tipocodproduto = 'N' AND @tipocodauxproduto = 'A'
    BEGIN
        UPDATE w1 SET erro = 'O Codigo Auxiliar ' + w1.codauxiliarproduto +
            ' do Produto ' + w1.codproduto +
            ' nao pode ser utilizado como Codigo Principal em nenhum outro produto.'
        FROM #W w1
        WHERE w1.codauxiliarproduto IS NOT NULL
          AND EXISTS (SELECT 1 FROM #W w2
                      WHERE CONVERT(bigint, w2.codproduto) = dbo.ToBigInt(w1.codauxiliarproduto)
                        AND w2.linha != w1.linha)
          AND w1.erro IS NULL
    END
    ELSE IF @tipocodproduto = 'A' AND @tipocodauxproduto = 'N'
    BEGIN
        UPDATE w1 SET erro = 'O Codigo Auxiliar ' + w1.codauxiliarproduto +
            ' do Produto ' + w1.codproduto +
            ' nao pode ser utilizado como Codigo Principal em nenhum outro produto.'
        FROM #W w1
        WHERE w1.codauxiliarproduto IS NOT NULL
          AND EXISTS (SELECT 1 FROM #W w2
                      WHERE dbo.ToBigInt(w2.codproduto) = CONVERT(bigint, w1.codauxiliarproduto)
                        AND w2.linha != w1.linha)
          AND w1.erro IS NULL
    END

    -- 6. Modo 1: produto nao pode existir
    IF @sobreescreve = 1
    BEGIN
        UPDATE w SET erro = 'Produto ' + RTRIM(w.codproduto) + ' ja existe no Banco de Dados.'
        FROM #W w INNER JOIN Produto p ON w.codproduto = p.codproduto
        WHERE w.erro IS NULL
    END
    -- Modo 3: produto deve existir
    ELSE IF @sobreescreve = 3
    BEGIN
        UPDATE w SET erro = 'Produto ' + RTRIM(w.codproduto) + ' nao existe no Banco de Dados.'
        FROM #W w LEFT JOIN Produto p ON w.codproduto = p.codproduto
        WHERE p.codproduto IS NULL AND w.erro IS NULL
    END

    -- 7. Produto (descricao) obrigatorio
    UPDATE #W SET erro = 'Produto nao pode ser null.'
    WHERE produto IS NULL AND erro IS NULL

    -- ============================================================
    -- Validacoes de QtdeMultipla / QtdeMinima
    -- ============================================================

    -- 8. QtdeMultipla range
    UPDATE #W SET erro = 'QtdeMultipla deve estar entre 0 e 999999.'
    WHERE qtdemultipla IS NOT NULL
      AND (qtdemultipla < 0 OR qtdemultipla > 999999)
      AND erro IS NULL

    -- 9. QtdeMultipla > 0: QtdeMinima obrigatoria e em range
    UPDATE #W SET erro = 'QtdeMinima deve estar entre a QtdeMultipla e 999999.'
    WHERE qtdemultipla IS NOT NULL AND qtdemultipla > 0
      AND (qtdeminima IS NULL OR qtdeminima < 0 OR qtdeminima > 999999)
      AND erro IS NULL

    -- 10. QtdeMinima >= QtdeMultipla
    UPDATE #W SET erro = 'QtdeMinima deve ser maior ou igual a QtdeMultipla.'
    WHERE qtdemultipla IS NOT NULL AND qtdemultipla > 0
      AND qtdeminima IS NOT NULL AND qtdeminima < qtdemultipla
      AND erro IS NULL

    -- 11. QtdeMinima divisivel por QtdeMultipla
    UPDATE #W SET erro = 'QtdeMinima deve ser divisivel pela QtdeMultipla.'
    WHERE qtdemultipla IS NOT NULL AND qtdemultipla > 0
      AND qtdeminima IS NOT NULL AND qtdeminima % qtdemultipla != 0
      AND erro IS NULL

    -- 12. QtdeMinima range (quando QtdeMultipla = 0 ou null)
    UPDATE #W SET erro = 'QtdeMinima deve estar entre 0 e 999999.'
    WHERE COALESCE(qtdemultipla, 0) = 0
      AND qtdeminima IS NOT NULL
      AND (qtdeminima < 0 OR qtdeminima > 999999)
      AND erro IS NULL

    -- ============================================================
    -- Validacoes de QtdeTabela / PrecoTabela (caso QtdeTabela1 > 0)
    -- ============================================================

    -- 13. QtdeTabela1 range
    UPDATE #W SET erro = 'QtdeTabela1 deve estar entre 0 e 999999.'
    WHERE qtdetabela1 IS NOT NULL
      AND (qtdetabela1 < 0 OR qtdetabela1 > 999999)
      AND erro IS NULL

    -- 14. QtdeTabela1 > 0: QtdeTabela1 >= QtdeMinima
    UPDATE #W SET erro = 'QtdeTabela1 deve ser maior ou igual a QtdeMinima.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND COALESCE(qtdeminima, 0) > 0
      AND qtdetabela1 < qtdeminima
      AND erro IS NULL

    -- 15. QtdeTabela1 > 0: QtdeTabela1 divisivel por QtdeMultipla
    UPDATE #W SET erro = 'QtdeTabela1 deve ser divisivel por QtdeMultipla.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND COALESCE(qtdemultipla, 0) > 0
      AND qtdetabela1 % qtdemultipla != 0
      AND erro IS NULL

    -- 16. QtdeTabela1 > 0: QtdeTabela2 obrigatoria
    UPDATE #W SET erro = 'QtdeTabela2 deve ser informada quando QtdeTabela1 foi informada.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND COALESCE(qtdetabela2, 0) = 0
      AND erro IS NULL

    -- 17. QtdeTabela1 > 0: QtdeTabela2 > QtdeTabela1 e <= 999999
    UPDATE #W SET erro = 'QtdeTabela2 deve ser maior que QtdeTabela1 e menor ou igual a 999999.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND qtdetabela2 IS NOT NULL
      AND (qtdetabela2 <= qtdetabela1 OR qtdetabela2 > 999999)
      AND erro IS NULL

    -- 18. QtdeTabela1 > 0: QtdeTabela2 divisivel por QtdeMultipla
    UPDATE #W SET erro = 'QtdeTabela2 deve ser divisivel por QtdeMultipla.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND qtdetabela2 IS NOT NULL
      AND COALESCE(qtdemultipla, 0) > 0
      AND qtdetabela2 % qtdemultipla != 0
      AND erro IS NULL

    -- 19. QtdeTabela1 > 0: QtdeTabela3 > QtdeTabela2 e <= 999999 (so valida se preenchido)
    UPDATE #W SET erro = 'QtdeTabela3 deve ser maior que QtdeTabela2 e menor ou igual a 999999.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND qtdetabela2 IS NOT NULL
      AND qtdetabela3 IS NOT NULL
      AND (qtdetabela3 <= qtdetabela2 OR qtdetabela3 > 999999)
      AND erro IS NULL

    -- 20. QtdeTabela1 > 0: QtdeTabela3 divisivel por QtdeMultipla
    UPDATE #W SET erro = 'QtdeTabela3 deve ser divisivel por QtdeMultipla.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND qtdetabela3 IS NOT NULL
      AND COALESCE(qtdemultipla, 0) > 0
      AND qtdetabela3 % qtdemultipla != 0
      AND erro IS NULL

    -- 21. QtdeTabela1 > 0: PrecoTabela1 range
    UPDATE #W SET erro = 'PrecoTabela1 deve estar entre 0,01 e 999999,99.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND (precotabela1 IS NULL OR precotabela1 < 0.01 OR precotabela1 > 999999.99)
      AND erro IS NULL

    -- 22. QtdeTabela1 > 0: PrecoTabela2 range
    UPDATE #W SET erro = 'PrecoTabela2 deve estar entre 0,00 e 999999,99.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND (precotabela2 IS NULL OR precotabela2 < 0.00 OR precotabela2 > 999999.99)
      AND erro IS NULL

    -- 23. QtdeTabela1 > 0: PrecoTabela2 < PrecoTabela1
    UPDATE #W SET erro = 'PrecoTabela2 deve ser menor que PrecoTabela1.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND precotabela2 IS NOT NULL AND precotabela2 >= precotabela1
      AND erro IS NULL

    -- 24. QtdeTabela1 > 0: PrecoTabela3 range
    UPDATE #W SET erro = 'PrecoTabela3 deve estar entre 0,00 e 999999,99.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND precotabela3 IS NOT NULL
      AND (precotabela3 < 0.00 OR precotabela3 > 999999.99)
      AND erro IS NULL

    -- 25. QtdeTabela1 > 0: PrecoTabela3 < PrecoTabela2
    UPDATE #W SET erro = 'PrecoTabela3 deve ser menor que PrecoTabela2.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND precotabela3 IS NOT NULL AND precotabela2 IS NOT NULL
      AND precotabela3 >= precotabela2
      AND erro IS NULL

    -- 26. QtdeTabela1 > 0: PrecoTabela3 sem QtdeTabela3
    UPDATE #W SET erro = 'Para informar PrecoTabela3, QtdeTabela3 tambem deve ser informada.'
    WHERE qtdetabela1 IS NOT NULL AND qtdetabela1 > 0
      AND COALESCE(qtdetabela3, 0) = 0
      AND COALESCE(precotabela3, 0.00) != 0.00
      AND erro IS NULL

    -- ============================================================
    -- Validacoes quando QtdeTabela1 = 0 ou NULL
    -- ============================================================

    -- 27. QtdeTabela2 sem QtdeTabela1
    UPDATE #W SET erro = 'Para informar QtdeTabela2, QtdeTabela1 tambem deve ser informada.'
    WHERE COALESCE(qtdetabela1, 0) = 0
      AND COALESCE(qtdetabela2, 0) != 0
      AND erro IS NULL

    -- 28. QtdeTabela3 sem QtdeTabela1
    UPDATE #W SET erro = 'Para informar QtdeTabela3, QtdeTabela1 e QtdeTabela2 devem ser informadas.'
    WHERE COALESCE(qtdetabela1, 0) = 0
      AND COALESCE(qtdetabela3, 0) != 0
      AND erro IS NULL

    -- 29. PrecoTabela1 range (sem QtdeTabela)
    UPDATE #W SET erro = 'PrecoTabela1 deve estar entre 0,01 e 999999,99.'
    WHERE COALESCE(qtdetabela1, 0) = 0
      AND (precotabela1 IS NULL OR precotabela1 < 0.01 OR precotabela1 > 999999.99)
      AND erro IS NULL

    -- 30. PrecoTabela2 range (sem QtdeTabela)
    UPDATE #W SET erro = 'PrecoTabela2 deve estar entre 0,00 e 999999,99.'
    WHERE COALESCE(qtdetabela1, 0) = 0
      AND precotabela2 IS NOT NULL
      AND (precotabela2 < 0.00 OR precotabela2 > 999999.99)
      AND erro IS NULL

    -- 31. PrecoTabela3 range (sem QtdeTabela)
    UPDATE #W SET erro = 'PrecoTabela3 deve estar entre 0,00 e 999999,99.'
    WHERE COALESCE(qtdetabela1, 0) = 0
      AND precotabela3 IS NOT NULL
      AND (precotabela3 < 0.00 OR precotabela3 > 999999.99)
      AND erro IS NULL

    -- 32. PrecoTabela3 sem PrecoTabela2
    UPDATE #W SET erro = 'Para informar PrecoTabela3, PrecoTabela2 tambem deve ser informado.'
    WHERE COALESCE(qtdetabela1, 0) = 0
      AND COALESCE(precotabela2, 0.00) = 0.00
      AND COALESCE(precotabela3, 0.00) != 0.00
      AND erro IS NULL

    -- ============================================================
    -- Validacoes de estoque
    -- ============================================================

    -- 33. QtdeEstoqueAtual range
    UPDATE #W SET erro = 'QtdeEstoqueAtual deve estar entre 0 e 999999.'
    WHERE qtdeestoqueatual IS NOT NULL
      AND (qtdeestoqueatual < 0 OR qtdeestoqueatual > 999999)
      AND erro IS NULL

    -- 34. QtdeEstoqueFuturo range
    UPDATE #W SET erro = 'QtdeEstoqueFuturo deve estar entre 0 e 999999.'
    WHERE qtdeestoquefuturo IS NOT NULL
      AND (qtdeestoquefuturo < 0 OR qtdeestoquefuturo > 999999)
      AND erro IS NULL

    -- 35. DtEstoqueFuturo obrigatoria quando QtdeEstoqueFuturo > 0 e ControlaDtEntrega = S
    IF @controladtentrega = 'S'
    BEGIN
        UPDATE #W SET erro = 'DtEstoqueFuturo deve ser informada quando Controle de Data de Entrega esta ativo.'
        WHERE qtdeestoquefuturo IS NOT NULL AND qtdeestoquefuturo > 0
          AND dtestoquefuturo IS NULL
          AND erro IS NULL
    END

    -- ============================================================
    -- Validacoes de FK
    -- ============================================================

    -- 36. CodFilial existe
    UPDATE w SET erro = 'CodFilial ' + CONVERT(varchar(6), w.codfilial) + ' nao cadastrado.'
    FROM #W w
    WHERE NOT EXISTS (SELECT 1 FROM Filial f WHERE f.codfilial = w.codfilial)
      AND w.erro IS NULL

    -- 37. PrecoPromocional valido
    UPDATE #W SET erro = 'PrecoPromocional deve ser S ou N.'
    WHERE precopromocional NOT IN ('S', 'N')
      AND erro IS NULL

    -- 38. TipoVendaSemEstoque valido
    UPDATE #W SET erro = 'TipoVendaSemEstoque deve ser C, L ou B.'
    WHERE tipovendasemestoque NOT IN ('C', 'L', 'B')
      AND erro IS NULL

    -- 39. CodFamilia range e FK
    UPDATE #W SET erro = 'CodFamilia deve ser null ou estar entre 1 e 999999.'
    WHERE codfamilia IS NOT NULL
      AND (codfamilia < 0 OR codfamilia > 999999)
      AND erro IS NULL

    UPDATE w SET erro = 'CodFamilia ' + CONVERT(varchar(6), w.codfamilia) + ' nao cadastrado.'
    FROM #W w
    WHERE w.codfamilia IS NOT NULL
      AND NOT EXISTS (SELECT 1 FROM Familia f WHERE f.codfamilia = w.codfamilia)
      AND w.erro IS NULL

    -- 40. CodEstilo range e FK
    UPDATE #W SET erro = 'CodEstilo deve ser null ou estar entre 1 e 999999.'
    WHERE codestilo IS NOT NULL
      AND (codestilo < 0 OR codestilo > 999999)
      AND erro IS NULL

    UPDATE w SET erro = 'CodEstilo ' + CONVERT(varchar(6), w.codestilo) + ' nao cadastrado.'
    FROM #W w
    WHERE w.codestilo IS NOT NULL
      AND NOT EXISTS (SELECT 1 FROM Estilo e WHERE e.codestilo = w.codestilo)
      AND w.erro IS NULL

    -- 41. QtdeEtiquetas range
    UPDATE #W SET erro = 'QtdeEtiquetas deve ser null ou estar entre 1 e 999.'
    WHERE qtdeetiquetas IS NOT NULL
      AND (qtdeetiquetas < 0 OR qtdeetiquetas > 999)
      AND erro IS NULL

    -- ============================================================
    -- IMPORTACAO: INSERT / UPDATE para linhas validas
    -- ============================================================
    BEGIN TRY
        BEGIN TRANSACTION

        -- UPDATE produtos existentes (modo 2 ou 3)
        IF @sobreescreve IN (2, 3)
        BEGIN
            UPDATE p SET
                codauxiliarproduto  = w.codauxiliarproduto,
                produto             = w.produto,
                precotabela1        = w.precotabela1,
                precotabela2        = w.precotabela2,
                precotabela3        = w.precotabela3,
                qtdeestoqueatual    = w.qtdeestoqueatual,
                qtdeestoquefuturo   = w.qtdeestoquefuturo,
                dtestoquefuturo     = w.dtestoquefuturo,
                limitedescindividual= w.limitedescindividual,
                aliquotaipi         = w.aliquotaipi,
                multiplograde       = w.multiplograde,
                descontograde       = w.descontograde,
                pathfotografia      = w.pathfotografia,
                codfilial           = w.codfilial,
                precopromocional    = w.precopromocional,
                tipovendasemestoque = w.tipovendasemestoque,
                codfamilia          = w.codfamilia,
                codestilo           = w.codestilo,
                qtdemultipla        = w.qtdemultipla,
                qtdeminima          = w.qtdeminima,
                qtdetabela1         = w.qtdetabela1,
                qtdetabela2         = w.qtdetabela2,
                qtdetabela3         = w.qtdetabela3,
                qtdeetiquetas       = w.qtdeetiquetas
            FROM Produto p
            INNER JOIN #W w ON p.codproduto = w.codproduto
            WHERE w.erro IS NULL
        END

        -- INSERT produtos novos (modo 1 ou 2)
        IF @sobreescreve IN (1, 2)
        BEGIN
            INSERT INTO Produto (
                codproduto, codauxiliarproduto, produto,
                precotabela1, precotabela2, precotabela3,
                qtdeestoqueatual, qtdeestoquefuturo, dtestoquefuturo,
                limitedescindividual, aliquotaipi, multiplograde, descontograde,
                pathfotografia, codfilial, precopromocional, tipovendasemestoque,
                codfamilia, codestilo, qtdemultipla, qtdeminima,
                qtdetabela1, qtdetabela2, qtdetabela3, qtdeetiquetas)
            SELECT
                w.codproduto, w.codauxiliarproduto, w.produto,
                w.precotabela1, w.precotabela2, w.precotabela3,
                w.qtdeestoqueatual, w.qtdeestoquefuturo, w.dtestoquefuturo,
                w.limitedescindividual, w.aliquotaipi, w.multiplograde, w.descontograde,
                w.pathfotografia, w.codfilial, w.precopromocional, w.tipovendasemestoque,
                w.codfamilia, w.codestilo, w.qtdemultipla, w.qtdeminima,
                w.qtdetabela1, w.qtdetabela2, w.qtdetabela3, w.qtdeetiquetas
            FROM #W w
            LEFT JOIN Produto p ON w.codproduto = p.codproduto
            WHERE p.codproduto IS NULL AND w.erro IS NULL
        END

        -- ========================================================
        -- Produto_Imp: espelho para controle de importacoes
        -- ========================================================

        -- UPDATE existentes
        UPDATE imp SET
            codauxiliarproduto      = w.codauxiliarproduto,
            produto                 = w.produto,
            precotabela1            = w.precotabela1,
            precotabela2            = w.precotabela2,
            precotabela3            = w.precotabela3,
            qtdeestoqueatual        = w.qtdeestoqueatual,
            qtdeestoquefuturo       = w.qtdeestoquefuturo,
            dtestoquefuturo         = w.dtestoquefuturo,
            limitedescindividual    = w.limitedescindividual,
            aliquotaipi             = w.aliquotaipi,
            multiplograde           = w.multiplograde,
            descontograde           = w.descontograde,
            pathfotografia          = w.pathfotografia,
            codfilial               = w.codfilial,
            precopromocional        = w.precopromocional,
            tipovendasemestoque     = w.tipovendasemestoque,
            codfamilia              = w.codfamilia,
            codestilo               = w.codestilo,
            qtdemultipla            = w.qtdemultipla,
            qtdeminima              = w.qtdeminima,
            qtdetabela1             = w.qtdetabela1,
            qtdetabela2             = w.qtdetabela2,
            qtdetabela3             = w.qtdetabela3,
            qtdeetiquetas           = w.qtdeetiquetas,
            codauxiliarproduto_rec  = 1,
            produto_rec             = 1,
            precotabela1_rec        = 1,
            precotabela2_rec        = 1,
            precotabela3_rec        = 1,
            qtdeestoqueatual_rec    = 1,
            qtdeestoquefuturo_rec   = 1,
            dtestoquefuturo_rec     = 1,
            limitedescindividual_rec= 1,
            aliquotaipi_rec         = 1,
            multiplograde_rec       = 1,
            descontograde_rec       = 1,
            pathfotografia_rec      = 1,
            codfilial_rec           = 1,
            precopromocional_rec    = 1,
            tipovendasemestoque_rec = 1,
            codfamilia_rec          = 1,
            codestilo_rec           = 1,
            qtdemultipla_rec        = 1,
            qtdeminima_rec          = 1,
            qtdetabela1_rec         = 1,
            qtdetabela2_rec         = 1,
            qtdetabela3_rec         = 1,
            qtdeetiquetas_rec       = 1
        FROM Produto_Imp imp
        INNER JOIN #W w ON imp.codproduto = w.codproduto
        WHERE w.erro IS NULL

        -- INSERT novos
        INSERT INTO Produto_Imp (
            codproduto, codauxiliarproduto, produto,
            precotabela1, precotabela2, precotabela3,
            qtdeestoqueatual, qtdeestoquefuturo, dtestoquefuturo,
            limitedescindividual, aliquotaipi, multiplograde, descontograde,
            pathfotografia, codfilial, precopromocional, tipovendasemestoque,
            codfamilia, codestilo, qtdemultipla, qtdeminima,
            qtdetabela1, qtdetabela2, qtdetabela3, qtdeetiquetas,
            codauxiliarproduto_rec, produto_rec,
            precotabela1_rec, precotabela2_rec, precotabela3_rec,
            qtdeestoqueatual_rec, qtdeestoquefuturo_rec, dtestoquefuturo_rec,
            limitedescindividual_rec, aliquotaipi_rec, multiplograde_rec, descontograde_rec,
            pathfotografia_rec, codfilial_rec, precopromocional_rec, tipovendasemestoque_rec,
            codfamilia_rec, codestilo_rec, qtdemultipla_rec, qtdeminima_rec,
            qtdetabela1_rec, qtdetabela2_rec, qtdetabela3_rec, qtdeetiquetas_rec)
        SELECT
            w.codproduto, w.codauxiliarproduto, w.produto,
            w.precotabela1, w.precotabela2, w.precotabela3,
            w.qtdeestoqueatual, w.qtdeestoquefuturo, w.dtestoquefuturo,
            w.limitedescindividual, w.aliquotaipi, w.multiplograde, w.descontograde,
            w.pathfotografia, w.codfilial, w.precopromocional, w.tipovendasemestoque,
            w.codfamilia, w.codestilo, w.qtdemultipla, w.qtdeminima,
            w.qtdetabela1, w.qtdetabela2, w.qtdetabela3, w.qtdeetiquetas,
            1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1
        FROM #W w
        LEFT JOIN Produto_Imp imp ON w.codproduto = imp.codproduto
        WHERE imp.codproduto IS NULL AND w.erro IS NULL

        COMMIT TRANSACTION
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION

        -- Marca todas linhas validas como erro de banco
        UPDATE #W SET erro = 'Erro de banco de dados: ' +
            LEFT(ERROR_MESSAGE(), 200)
        WHERE erro IS NULL
    END CATCH

    -- ============================================================
    -- Retorna resultado por linha
    -- ============================================================
    SELECT
        linha,
        codproduto,
        codauxiliarproduto,
        CASE WHEN erro IS NULL THEN 'OK' ELSE 'ERRO' END AS status,
        ISNULL(erro, 'Importacao efetuada com sucesso.') AS mensagem
    FROM #W
    ORDER BY linha
    OPTION (RECOMPILE)
END
"""

# SQL para inserir dados na staging de produtos
SQL_INSERT_STAGING = """
INSERT INTO ImportaProdutoAmbos_Staging (
    codproduto, codauxiliarproduto, produto,
    precotabela1, precotabela2, precotabela3,
    qtdeestoqueatual, qtdeestoquefuturo, dtestoquefuturo,
    limitedescindividual, aliquotaipi, multiplograde, descontograde,
    pathfotografia, codfilial, precopromocional, tipovendasemestoque,
    codfamilia, codestilo, qtdemultipla, qtdeminima,
    qtdetabela1, qtdetabela2, qtdetabela3, qtdeetiquetas
) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
"""

# ================================================================
# TRANSPORTADORA - Importacao em lote
# ================================================================

SQL_CREATE_STAGING_TRANSPORTADORA = """
IF OBJECT_ID('dbo.ImportaTransportadora_Staging','U') IS NULL
BEGIN
    CREATE TABLE dbo.ImportaTransportadora_Staging (
        linha                   int IDENTITY(1,1) PRIMARY KEY,
        codtransportadora       smallint,
        transportadora          varchar(20),
        transportadorapadrao    char(1)
    );
END
"""

SQL_CREATE_PROCEDURE_TRANSPORTADORA = """
IF OBJECT_ID('dbo.ImportaTransportadora_Lote','P') IS NOT NULL
    DROP PROCEDURE dbo.ImportaTransportadora_Lote;
"""

SQL_CREATE_PROCEDURE_BODY_TRANSPORTADORA = """
CREATE PROCEDURE dbo.ImportaTransportadora_Lote
    @sobreescreve tinyint = 2
AS
BEGIN
    SET NOCOUNT ON;

    -- Verifica empresa configurada
    IF NOT EXISTS (SELECT 1 FROM empresa WHERE codempresa = 1)
    BEGIN
        SELECT linha, codtransportadora, 'ERRO' AS status,
            'Nao foi possivel importar Transportadoras porque Empresa nao foi configurada.' AS mensagem
        FROM ImportaTransportadora_Staging ORDER BY linha
        RETURN
    END

    -- Work table com coluna de erro
    SELECT
        linha, codtransportadora, transportadora, transportadorapadrao,
        CAST(NULL AS varchar(250)) AS erro
    INTO #W
    FROM ImportaTransportadora_Staging

    CREATE CLUSTERED INDEX IX_W_linha ON #W (linha)
    CREATE NONCLUSTERED INDEX IX_W_codtransp ON #W (codtransportadora)

    -- Aplicar defaults
    UPDATE #W SET transportadorapadrao = 'N' WHERE transportadorapadrao IS NULL

    -- Popula PKUnica
    DELETE FROM Transportadora_PKUnica
    INSERT INTO Transportadora_PKUnica (codtransportadora)
    SELECT codtransportadora FROM #W WHERE codtransportadora IS NOT NULL

    -- ============================================================
    -- VALIDACOES
    -- ============================================================

    -- 1. CodTransportadora obrigatorio e range
    UPDATE #W SET erro = 'CodTransportadora nao pode ser null e deve estar entre 1 e 32767.'
    WHERE (codtransportadora IS NULL OR codtransportadora < 1 OR codtransportadora > 32767)
      AND erro IS NULL

    -- 2. Duplicidade no arquivo
    ;WITH dups AS (
        SELECT codtransportadora, COUNT(*)-1 AS qtd
        FROM #W WHERE codtransportadora IS NOT NULL
        GROUP BY codtransportadora HAVING COUNT(*) > 1
    )
    UPDATE w SET erro = 'Transportadora ' + CONVERT(varchar(5), w.codtransportadora) + ' tem ' +
        CONVERT(varchar(5), d.qtd) +
        CASE WHEN d.qtd = 1 THEN ' duplicidade' ELSE ' duplicidades' END +
        ' no arquivo.'
    FROM #W w INNER JOIN dups d ON w.codtransportadora = d.codtransportadora
    WHERE w.erro IS NULL

    -- 3. Modo 1: nao pode existir
    IF @sobreescreve = 1
    BEGIN
        UPDATE w SET erro = 'Transportadora ' + CONVERT(varchar(5), w.codtransportadora) + ' ja existe no Banco de Dados.'
        FROM #W w INNER JOIN Transportadora t ON w.codtransportadora = t.codtransportadora
        WHERE w.erro IS NULL
    END
    -- Modo 3: deve existir
    ELSE IF @sobreescreve = 3
    BEGIN
        UPDATE w SET erro = 'Transportadora ' + CONVERT(varchar(5), w.codtransportadora) + ' nao existe no Banco de Dados.'
        FROM #W w LEFT JOIN Transportadora t ON w.codtransportadora = t.codtransportadora
        WHERE t.codtransportadora IS NULL AND w.erro IS NULL
    END

    -- 4. Transportadora (nome) obrigatorio
    UPDATE #W SET erro = 'Transportadora nao pode ser null.'
    WHERE transportadora IS NULL AND erro IS NULL

    -- 5. TransportadoraPadrao valido
    UPDATE #W SET erro = 'TransportadoraPadrao deve ser S ou N.'
    WHERE transportadorapadrao NOT IN ('S', 'N') AND erro IS NULL

    -- ============================================================
    -- IMPORTACAO
    -- ============================================================
    BEGIN TRY
        BEGIN TRANSACTION

        -- UPDATE existentes (modo 2 ou 3)
        IF @sobreescreve IN (2, 3)
        BEGIN
            UPDATE t SET
                transportadora = w.transportadora,
                transportadorapadrao = w.transportadorapadrao
            FROM Transportadora t
            INNER JOIN #W w ON t.codtransportadora = w.codtransportadora
            WHERE w.erro IS NULL
        END

        -- INSERT novos (modo 1 ou 2)
        IF @sobreescreve IN (1, 2)
        BEGIN
            INSERT INTO Transportadora (codtransportadora, transportadora, transportadorapadrao)
            SELECT w.codtransportadora, w.transportadora, w.transportadorapadrao
            FROM #W w
            LEFT JOIN Transportadora t ON w.codtransportadora = t.codtransportadora
            WHERE t.codtransportadora IS NULL AND w.erro IS NULL
        END

        -- Transportadora_Imp: espelho
        UPDATE imp SET
            transportadora = w.transportadora,
            transportadorapadrao = w.transportadorapadrao,
            transportadora_rec = 1,
            transportadorapadrao_rec = 1
        FROM Transportadora_Imp imp
        INNER JOIN #W w ON imp.codtransportadora = w.codtransportadora
        WHERE w.erro IS NULL

        INSERT INTO Transportadora_Imp (codtransportadora, transportadora, transportadorapadrao,
            transportadora_rec, transportadorapadrao_rec)
        SELECT w.codtransportadora, w.transportadora, w.transportadorapadrao, 1, 1
        FROM #W w
        LEFT JOIN Transportadora_Imp imp ON w.codtransportadora = imp.codtransportadora
        WHERE imp.codtransportadora IS NULL AND w.erro IS NULL

        COMMIT TRANSACTION
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION
        UPDATE #W SET erro = 'Erro de banco de dados: ' + LEFT(ERROR_MESSAGE(), 200)
        WHERE erro IS NULL
    END CATCH

    -- Retorna resultado
    SELECT
        linha, codtransportadora,
        CASE WHEN erro IS NULL THEN 'OK' ELSE 'ERRO' END AS status,
        ISNULL(erro, 'Importacao efetuada com sucesso.') AS mensagem
    FROM #W ORDER BY linha
    OPTION (RECOMPILE)
END
"""

SQL_INSERT_STAGING_TRANSPORTADORA = """
INSERT INTO ImportaTransportadora_Staging (codtransportadora, transportadora, transportadorapadrao)
VALUES (?, ?, ?)
"""

# ================================================================
# CLIENTE - Importacao em lote
# ================================================================

SQL_CREATE_STAGING_CLIENTE = """
IF OBJECT_ID('dbo.ImportaCliente_Staging','U') IS NULL
BEGIN
    CREATE TABLE dbo.ImportaCliente_Staging (
        linha               int IDENTITY(1,1) PRIMARY KEY,
        codcliente          int,
        codrepresentante    smallint,
        nomefantasia        varchar(20),
        razaosocial         varchar(40),
        cnpjcpf             varchar(19),
        ierg                varchar(19),
        logradouro          varchar(40),
        bairro              varchar(20),
        cidade              varchar(20),
        uf                  char(2),
        cep                 varchar(9),
        ddd                 tinyint,
        telefone1           varchar(9),
        telefone2           varchar(9),
        fax                 varchar(9),
        nomecontato         varchar(40),
        nometransportadora  varchar(20),
        observacao          varchar(20),
        email               varchar(40),
        precotabela         tinyint,
        codtransportadora   smallint
    );
END
"""

SQL_CREATE_PROCEDURE_CLIENTE = """
IF OBJECT_ID('dbo.ImportaCliente_Lote','P') IS NOT NULL
    DROP PROCEDURE dbo.ImportaCliente_Lote;
"""

SQL_CREATE_PROCEDURE_BODY_CLIENTE = """
CREATE PROCEDURE dbo.ImportaCliente_Lote
    @sobreescreve tinyint = 2
AS
BEGIN
    SET NOCOUNT ON;

    -- Verifica empresa configurada
    IF NOT EXISTS (SELECT 1 FROM empresa WHERE codempresa = 1)
    BEGIN
        SELECT linha, codcliente, 'ERRO' AS status,
            'Nao foi possivel importar Clientes porque Empresa nao foi configurada.' AS mensagem
        FROM ImportaCliente_Staging ORDER BY linha
        RETURN
    END

    -- Work table com coluna de erro
    SELECT
        linha, codcliente, codrepresentante, nomefantasia, razaosocial,
        cnpjcpf, ierg, logradouro, bairro, cidade, uf, cep,
        ddd, telefone1, telefone2, fax, nomecontato,
        nometransportadora, observacao, email, precotabela, codtransportadora,
        CAST(NULL AS varchar(250)) AS erro
    INTO #W
    FROM ImportaCliente_Staging

    CREATE CLUSTERED INDEX IX_W_linha ON #W (linha)
    CREATE NONCLUSTERED INDEX IX_W_codcliente ON #W (codcliente)

    -- CodRepresentante 0 -> NULL
    UPDATE #W SET codrepresentante = NULL WHERE codrepresentante = 0

    -- Popula PKUnica
    DELETE FROM Cliente_PKUnica
    INSERT INTO Cliente_PKUnica (codcliente)
    SELECT codcliente FROM #W WHERE codcliente IS NOT NULL

    -- ============================================================
    -- VALIDACOES
    -- ============================================================

    -- 1. CodCliente obrigatorio e range
    UPDATE #W SET erro = 'CodCliente nao pode ser null e deve estar entre 1 e 9999999.'
    WHERE (codcliente IS NULL OR codcliente < 1 OR codcliente > 9999999)
      AND erro IS NULL

    -- 2. Duplicidade no arquivo
    ;WITH dups AS (
        SELECT codcliente, COUNT(*)-1 AS qtd
        FROM #W WHERE codcliente IS NOT NULL
        GROUP BY codcliente HAVING COUNT(*) > 1
    )
    UPDATE w SET erro = 'Cliente ' + CONVERT(varchar(7), w.codcliente) + ' tem ' +
        CONVERT(varchar(5), d.qtd) +
        CASE WHEN d.qtd = 1 THEN ' duplicidade' ELSE ' duplicidades' END +
        ' no arquivo.'
    FROM #W w INNER JOIN dups d ON w.codcliente = d.codcliente
    WHERE w.erro IS NULL

    -- 3. Modo 1: nao pode existir
    IF @sobreescreve = 1
    BEGIN
        UPDATE w SET erro = 'Cliente ' + CONVERT(varchar(7), w.codcliente) + ' ja existe no Banco de Dados.'
        FROM #W w INNER JOIN Cliente c ON w.codcliente = c.codcliente
        WHERE w.erro IS NULL
    END
    -- Modo 3: deve existir
    ELSE IF @sobreescreve = 3
    BEGIN
        UPDATE w SET erro = 'Cliente ' + CONVERT(varchar(7), w.codcliente) + ' nao existe no Banco de Dados.'
        FROM #W w LEFT JOIN Cliente c ON w.codcliente = c.codcliente
        WHERE c.codcliente IS NULL AND w.erro IS NULL
    END

    -- 4. CodRepresentante FK
    UPDATE w SET erro = 'Representante ' + CONVERT(varchar(5), w.codrepresentante) + ' nao cadastrado no Banco de Dados.'
    FROM #W w
    WHERE w.codrepresentante IS NOT NULL
      AND NOT EXISTS (SELECT 1 FROM Representante r WHERE r.codrepresentante = w.codrepresentante)
      AND w.erro IS NULL

    -- 5. NomeFantasia obrigatorio
    UPDATE #W SET erro = 'NomeFantasia nao pode ser null.'
    WHERE nomefantasia IS NULL AND erro IS NULL

    -- 6. PrecoTabela range
    UPDATE #W SET erro = 'Tabela de Preco deve ser null ou estar entre 0 e 3.'
    WHERE precotabela IS NOT NULL AND (precotabela < 0 OR precotabela > 3)
      AND erro IS NULL

    -- 7. CodTransportadora FK (se informado)
    UPDATE w SET erro = 'Transportadora ' + CONVERT(varchar(5), w.codtransportadora) + ' nao cadastrada no Banco de Dados.'
    FROM #W w
    WHERE w.codtransportadora IS NOT NULL
      AND NOT EXISTS (SELECT 1 FROM Transportadora t WHERE t.codtransportadora = w.codtransportadora)
      AND w.erro IS NULL

    -- 8. CNPJCPF formato (deve ser NNN.NNN.NNN/NNNN-NN com 19 chars, ou NULL)
    -- Corrige para NULL se formato invalido (evita CHECK constraint CNPJCPFNull_)
    UPDATE #W SET cnpjcpf = NULL
    WHERE cnpjcpf IS NOT NULL
      AND (LEN(cnpjcpf) != 19
           OR cnpjcpf NOT LIKE '[0-9][0-9][0-9].[0-9][0-9][0-9].[0-9][0-9][0-9]/[0-9][0-9][0-9][0-9]-[0-9][0-9]')
      AND erro IS NULL

    -- 9. CEP formato (deve ser NNNNN-NNN com 9 chars, ou NULL)
    UPDATE #W SET cep = NULL
    WHERE cep IS NOT NULL
      AND (LEN(cep) != 9
           OR cep NOT LIKE '[0-9][0-9][0-9][0-9][0-9]-[0-9][0-9][0-9]')
      AND erro IS NULL

    -- 10. UF valida (2 chars, sigla de estado brasileira, ou NULL)
    UPDATE #W SET uf = NULL
    WHERE uf IS NOT NULL
      AND UPPER(LTRIM(RTRIM(uf))) NOT IN (
          'AC','AL','AM','AP','BA','CE','DF','ES','EX','GO','MA','MG','MS','MT',
          'PA','PB','PE','PI','PR','RJ','RN','RO','RR','RS','SC','SE','SP','TO')
      AND erro IS NULL

    -- 11. DDD range (deve ser > 0 ou NULL)
    UPDATE #W SET ddd = NULL
    WHERE ddd IS NOT NULL AND ddd <= 0
      AND erro IS NULL

    -- 12. CodRepresentante range (deve ser > 0 ou NULL)
    UPDATE #W SET codrepresentante = NULL
    WHERE codrepresentante IS NOT NULL AND codrepresentante <= 0
      AND erro IS NULL

    -- ============================================================
    -- IMPORTACAO
    -- ============================================================
    BEGIN TRY
        BEGIN TRANSACTION

        -- UPDATE existentes (modo 2 ou 3)
        IF @sobreescreve IN (2, 3)
        BEGIN
            UPDATE c SET
                codrepresentante = w.codrepresentante,
                nomefantasia = w.nomefantasia,
                razaosocial = w.razaosocial,
                cnpjcpf = w.cnpjcpf,
                ierg = w.ierg,
                logradouro = w.logradouro,
                bairro = w.bairro,
                cidade = w.cidade,
                uf = w.uf,
                cep = w.cep,
                ddd = w.ddd,
                telefone1 = w.telefone1,
                telefone2 = w.telefone2,
                fax = w.fax,
                nomecontato = w.nomecontato,
                observacao = w.observacao,
                email = w.email,
                clientenovo = 'N',
                precotabela = w.precotabela,
                codtransportadora = w.codtransportadora,
                codatualizacao = 1
            FROM Cliente c
            INNER JOIN #W w ON c.codcliente = w.codcliente
            WHERE w.erro IS NULL
        END

        -- INSERT novos (modo 1 ou 2)
        IF @sobreescreve IN (1, 2)
        BEGIN
            INSERT INTO Cliente (
                codcliente, codrepresentante, nomefantasia, razaosocial,
                cnpjcpf, ierg, logradouro, bairro, cidade, uf, cep,
                ddd, telefone1, telefone2, fax, nomecontato, observacao, email,
                clientenovo, precotabela, codtransportadora, codatualizacao)
            SELECT
                w.codcliente, w.codrepresentante, w.nomefantasia, w.razaosocial,
                w.cnpjcpf, w.ierg, w.logradouro, w.bairro, w.cidade, w.uf, w.cep,
                w.ddd, w.telefone1, w.telefone2, w.fax, w.nomecontato, w.observacao, w.email,
                'N', w.precotabela, w.codtransportadora, 1
            FROM #W w
            LEFT JOIN Cliente c ON w.codcliente = c.codcliente
            WHERE c.codcliente IS NULL AND w.erro IS NULL
        END

        -- Cliente_Imp: espelho
        UPDATE imp SET
            codrepresentante = w.codrepresentante,
            nomefantasia = w.nomefantasia,
            razaosocial = w.razaosocial,
            cnpjcpf = w.cnpjcpf,
            ierg = w.ierg,
            logradouro = w.logradouro,
            bairro = w.bairro,
            cidade = w.cidade,
            uf = w.uf,
            cep = w.cep,
            ddd = w.ddd,
            telefone1 = w.telefone1,
            telefone2 = w.telefone2,
            fax = w.fax,
            nomecontato = w.nomecontato,
            observacao = w.observacao,
            email = w.email,
            precotabela = w.precotabela,
            codtransportadora = w.codtransportadora,
            codrepresentante_rec = 1,
            nomefantasia_rec = 1,
            razaosocial_rec = 1,
            cnpjcpf_rec = 1,
            ierg_rec = 1,
            logradouro_rec = 1,
            bairro_rec = 1,
            cidade_rec = 1,
            uf_rec = 1,
            cep_rec = 1,
            ddd_rec = 1,
            telefone1_rec = 1,
            telefone2_rec = 1,
            fax_rec = 1,
            nomecontato_rec = 1,
            observacao_rec = 1,
            email_rec = 1,
            precotabela_rec = 1,
            codtransportadora_rec = 1
        FROM Cliente_Imp imp
        INNER JOIN #W w ON imp.codcliente = w.codcliente
        WHERE w.erro IS NULL

        INSERT INTO Cliente_Imp (
            codcliente, codrepresentante, nomefantasia, razaosocial,
            cnpjcpf, ierg, logradouro, bairro, cidade, uf, cep,
            ddd, telefone1, telefone2, fax, nomecontato, observacao, email,
            precotabela, codtransportadora,
            codrepresentante_rec, nomefantasia_rec, razaosocial_rec, cnpjcpf_rec, ierg_rec,
            logradouro_rec, bairro_rec, cidade_rec, uf_rec, cep_rec, ddd_rec,
            telefone1_rec, telefone2_rec, fax_rec, nomecontato_rec, observacao_rec, email_rec,
            precotabela_rec, codtransportadora_rec)
        SELECT
            w.codcliente, w.codrepresentante, w.nomefantasia, w.razaosocial,
            w.cnpjcpf, w.ierg, w.logradouro, w.bairro, w.cidade, w.uf, w.cep,
            w.ddd, w.telefone1, w.telefone2, w.fax, w.nomecontato, w.observacao, w.email,
            w.precotabela, w.codtransportadora,
            1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1
        FROM #W w
        LEFT JOIN Cliente_Imp imp ON w.codcliente = imp.codcliente
        WHERE imp.codcliente IS NULL AND w.erro IS NULL

        COMMIT TRANSACTION
    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION
        UPDATE #W SET erro = 'Erro de banco de dados: ' + LEFT(ERROR_MESSAGE(), 200)
        WHERE erro IS NULL
    END CATCH

    -- Retorna resultado
    SELECT
        linha, codcliente,
        CASE WHEN erro IS NULL THEN 'OK' ELSE 'ERRO' END AS status,
        ISNULL(erro, 'Importacao efetuada com sucesso.') AS mensagem
    FROM #W ORDER BY linha
    OPTION (RECOMPILE)
END
"""

SQL_INSERT_STAGING_CLIENTE = """
INSERT INTO ImportaCliente_Staging (
    codcliente, codrepresentante, nomefantasia, razaosocial,
    cnpjcpf, ierg, logradouro, bairro, cidade, uf, cep,
    ddd, telefone1, telefone2, fax, nomecontato,
    nometransportadora, observacao, email, precotabela, codtransportadora
) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
"""
