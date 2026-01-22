# Validador de Planilhas - SINT

Ferramenta para validacao de planilhas de importacao no padrao SINT.

## Instalacao

Baixe o executavel `Validador SINT.exe` na aba [Releases](../../releases) e execute.

Nao requer instalacao adicional.

## Como usar

1. Execute o `Validador SINT.exe`
2. Clique em "Procurar" e selecione a planilha (.xlsx, .xls, .xlsm, .xlsb)
3. Clique em "Validar Planilha"
4. Aguarde o processamento
5. O arquivo validado sera salvo na mesma pasta do executavel

## Saida

O validador gera:
- Planilha validada com marcacoes de erro/advertencia em cada celula
- Planilha de etiquetas (quando aplicavel)

## Status de validacao

| Status | Descricao |
|--------|-----------|
| APROVADO | Todas as validacoes passaram |
| APROVADO COM ADVERTENCIAS | Validacoes passaram com avisos |
| REPROVADO | Erros criticos encontrados |

## Abas validadas

- EMPRESA
- FILIAL
- REPR
- PAGTO
- PAGTOFILIAL
- TRANSP
- ESTADOS
- CLIENTES
- FAMILIAS
- ESTILOS
- PRODUTOS

---

## Desenvolvimento

### Requisitos

- Python 3.8+
- openpyxl >= 3.1.2

### Instalacao local

```bash
pip install -r requirements.txt
python validador_standalone.py
```

### Modo desenvolvedor

Para gerar relatorio de performance no console:

```bash
python validador_standalone.py --dev
```

### Gerar executavel

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "Validador SINT" validador_standalone.py
```

O executavel sera gerado em `dist/Validador SINT.exe`.

## Estrutura do projeto

```
├── validador_standalone.py   # Interface grafica (Tkinter)
├── planilha_validator.py     # Logica de validacao
├── requirements.txt          # Dependencias
└── README.md
```

## Abas validadas

- EMPRESA
- FILIAL
- REPR
- PAGTO
- PAGTOFILIAL
- TRANSP
- ESTADOS
- CLIENTES
- FAMILIAS
- ESTILOS
- PRODUTOS

## Status de validacao

- **APROVADO**: Todas as validacoes passaram
- **APROVADO COM ADVERTENCIAS**: Validacoes passaram com avisos
- **REPROVADO**: Erros criticos encontrados

## Saida

O validador gera:
- Planilha validada com marcacoes de erro/advertencia
- Planilha de etiquetas (quando aplicavel)

---

Desenvolvido por Erick
