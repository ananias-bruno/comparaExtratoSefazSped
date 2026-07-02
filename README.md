# 🔍 Comparador NF-e — Excel vs XML / SEFAZ

Ferramenta desktop para conferência fiscal de notas fiscais eletrônicas, comparando os dados do SPED (C100) com os XMLs das NF-e e com o extrato do portal SEFAZ.

---

## 📋 Funcionalidades

### Aba 1 — Excel vs XML
Compara, nota a nota, os valores do relatório SPED com os XMLs das NF-e:

| Campo comparado | SPED (C100) | XML NF-e |
|---|---|---|
| Base de Cálculo ICMS | `vl_bc_icms` | `<vBC>` |
| Valor do ICMS | `vl_icms` | `<vICMS>` |
| Valor total do documento | `vl_doc` | `<vNF>` |

- Destaca em vermelho cada célula com divergência
- Mostra rodapé com totais de todas as colunas de valores
- Exporta as divergências para `.xlsx`
- Gera script SQL para atualização dos registros

### Aba 2 — Excel vs SEFAZ
Compara os dados do SPED com o extrato exportado do portal SEFAZ:

| Campo comparado | SPED (C100) | SEFAZ |
|---|---|---|
| Base de Cálculo ICMS | `vl_bc_icms` | `Total_BC_ICMS` |
| Valor do ICMS | `vl_icms` | `Total_ICMS` |
| Valor total da nota | `vl_doc` | `Total_NF-e` |

Além disso, identifica **notas que estão no extrato SEFAZ mas não foram lançadas no SPED**, exibindo número, chave de acesso, valores e emitente.

---

## 🖥️ Interface

![Interface dark com duas abas: Excel vs XML e Excel vs SEFAZ](screenshot.png)

- Tema escuro
- Rodapé de totais sincronizado com a grade
- Status por linha: ✔ OK / ✘ Divergente / ⚠ Ausente
- Resumo no topo com contagem de OK, divergentes e ausentes

---

## ⚙️ Requisitos

- Python **3.9+**
- Bibliotecas:

```
pip install pandas openpyxl PyQt5
```

---

## 🚀 Como usar

```bash
python resumo_nfe_sped_update.py
```

### Aba Excel vs XML

1. Clique em **Procurar** ao lado de *Planilha Excel (C100)* e selecione o `.xlsx` exportado do SPED
2. Clique em **Procurar** ao lado de *Pasta com XMLs* e selecione a pasta com os arquivos `.xml` das NF-e
3. Clique em **▶ Comparar**

### Aba Excel vs SEFAZ

1. Clique em **Procurar** ao lado de *Planilha Excel (C100)* e selecione o mesmo `.xlsx` do SPED
2. Clique em **Procurar** ao lado de *Extrato SEFAZ* e selecione o arquivo `.xlsx` exportado do portal SEFAZ
3. Clique em **▶ Comparar com SEFAZ**

---

## 📂 Formato dos arquivos

### Planilha C100 (SPED)
Arquivo `.xlsx` com pelo menos as seguintes colunas (nomes em minúsculo):

| Coluna | Descrição |
|---|---|
| `chv_nfe` | Chave de acesso da NF-e (44 dígitos) |
| `num_doc` | Número do documento fiscal |
| `vl_doc` | Valor total do documento |
| `vl_bc_icms` | Base de cálculo do ICMS |
| `vl_icms` | Valor do ICMS |

Aceita valores no formato brasileiro (`19,66`) ou americano (`19.66`).

### XMLs das NF-e
Arquivos `.xml` no padrão da NF-e (namespace `http://www.portalfiscal.inf.br/nfe`).  
O sistema localiza cada XML pela chave de acesso contida no nome do arquivo ou dentro da tag `<chNFe>`.

### Extrato SEFAZ
Arquivo `.xlsx` exportado diretamente do portal SEFAZ RS (ou formato equivalente), contendo:

| Coluna | Descrição |
|---|---|
| `Chave_NF-e` | Chave de acesso (44 dígitos) |
| `Numero` | Número da nota |
| `Total_NF-e` | Valor total da nota |
| `Total_BC_ICMS` | Base de cálculo do ICMS |
| `Total_ICMS` | Valor do ICMS |
| `dt_Emit` | Data de emissão |
| `Razao_Social_Emit` | Nome do emitente |

O sistema detecta automaticamente os nomes das colunas, tolerando variações.

---

## 🎨 Legenda de status

| Ícone | Cor | Significado |
|---|---|---|
| ✔ OK | 🟢 Verde | Valores idênticos (tolerância de R$ 0,01) |
| ✘ Divergente | 🔴 Vermelho | Diferença acima de R$ 0,01 em algum campo |
| ⚠ Ausente | 🟡 Amarelo | Nota presente no C100 mas não encontrada no SEFAZ/XML, ou vice-versa |

---

## 📤 Exportações

- **Exportar divergências (aba XML):** gera `.xlsx` com todas as notas que têm diferença entre SPED e XML
- **Exportar divergências (aba SEFAZ):** gera `.xlsx` com notas com diferença de valores entre SPED e SEFAZ
- **Exportar ausentes no SPED:** gera `.xlsx` com notas encontradas no SEFAZ mas não lançadas no SPED
- **Gerar SQL:** gera script `.sql` para atualizar os valores no banco de dados

---

## 🗂️ Estrutura do projeto

```
comparador-nfe/
├── resumo_nfe_sped_update.py   # Aplicação principal
├── README.md                   # Este arquivo
└── exemplos/
    ├── c100Exemplo.xlsx         # Exemplo de planilha SPED
    └── extrato_exemplo.xlsx     # Exemplo de extrato SEFAZ
```

---

## 🔧 Configuração

No início do arquivo `resumo_nfe_sped_update.py`:

```python
TOLERANCIA = 0.01  # Diferença mínima (R$) para considerar divergência
```

---

## 📝 Observações

- A comparação de chaves ignora formatação (pontos, traços, prefixo `NFe`), trabalhando apenas com os 44 dígitos numéricos
- Valores monetários são aceitos tanto no padrão BR (`1.234,56`) quanto no padrão EN (`1,234.56`), detectado automaticamente por coluna
- O rodapé de totais é sincronizado com o scroll e redimensionamento de colunas da grade principal