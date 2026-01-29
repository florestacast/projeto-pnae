# Gerador de Projeto de Venda PNAE

Sistema automatizado para geração de documentos PDF do Projeto de Venda de Gêneros Alimentícios da Agricultura Familiar para Alimentação Escolar/PNAE.

**Desenvolvido por:** Floresta Cast LTDA  
**Localização:** Eunápolis/BA  
**Data:** Janeiro de 2026

---

## 📋 Descrição

Este sistema gera automaticamente documentos PDF completos para projetos PNAE, incluindo:

- ✅ Identificação dos Fornecedores (Grupo Formal)
- ✅ Identificação da Unidade Executora
- ✅ Relação de Produtos com Sazonalidade
- ✅ Declarações e Envelopes
- ✅ Capas Personalizadas
- ✅ Marca d'água automatizada

---

## 🔧 Requisitos

### Dependências Python
```bash
pip install pandas reportlab openpyxl pillow
```

### Arquivos Necessários

1. **projeto_venda.xlsx** - Planilha principal com as seguintes abas:
   - `administracao` - Dados do proponente
   - `produtor` - Dados dos produtores (opcional)
   - `edital` - Dados da entidade executora
   - `estoque` - Controle de estoque (opcional)
   - `envelope` - Declarações para envelopes
   - `capa` - Dados das capas
   - `alimentos` - Lista de produtos e sazonalidade

2. **cabecalho.png** - Imagem do cabeçalho (7" x 1")
3. **marca_dagua.png** - Marca d'água (4" x 4")

---

## 📊 Estrutura da Planilha Excel

### Aba: administracao
| Coluna | Descrição |
|--------|-----------|
| `status_representante` | "Ativo" ou "Inativo" |
| `proponente` | Nome da associação/cooperativa |
| `cnpj_proponente` | CNPJ do proponente |
| `endereco_proponente` | Endereço completo |
| `municipio_proponente` | Nome do município |
| `uf_proponente` | Sigla do estado (ex: BA) |
| `e-mailp` | E-mail do proponente |
| `celular_proponente` | Telefone de contato |
| `cep_proponente` | CEP |
| `caf_juridica` | Número DAP/CAF Jurídica |
| `banco_proponente` | Nome do banco |
| `agencia_proponente` | Número da agência |
| `conta_proponente` | Número da conta |
| `representante_proponente` | Nome do representante legal |
| `cpf_proponente` | CPF do representante |
| `rg_proponente` | RG do representante |

### Aba: edital
| Coluna | Descrição |
|--------|-----------|
| `chamada_publica` | Número da chamada pública |
| `fim_edital` | Data de fim do edital (formato: DD/MM/AAAA) |
| `nome_executora` | Nome da entidade executora |
| `cnpj_executora` | CNPJ da executora |
| `municipio_executora` | Município da executora |
| `uf_executora` | UF da executora |
| `endereco_executora` | Endereço da executora |
| `gestor_executora` | Nome do gestor |
| `e-mail_r_ex` | E-mail do gestor |
| `cpf_executora` | CPF do gestor |

### Aba: alimentos
| Coluna | Descrição |
|--------|-----------|
| `itens` | Número do item (ex: 1., 2., ...) |
| `produto` | Nome e descrição do produto |
| `unidade` | Unidade de medida (KG, UN, LT, etc.) |
| `quantidade` | Quantidade numérica |
| `preco` | Preço unitário (formato: R$ 11,56) |
| `total` | Valor total (calculado) |
| `sazonalidade` | Período de disponibilidade |
| `status_alimentos` | "Ativo" ou "Inativo" |

### Aba: envelope
| Coluna | Descrição |
|--------|-----------|
| `status_envelope` | "SIM" ou "NÃO" |
| `anexo_envelope` | Número do anexo |
| `assunto` | Assunto da declaração |
| `declaracao` | Texto completo da declaração |

### Aba: capa
| Coluna | Descrição |
|--------|-----------|
| `status_capa` | "Ativo" ou "Inativo" |
| `capa` | Título da capa |
| `titulo_capa` | Subtítulo (opcional) |

---

## 🚀 Como Usar

1. **Prepare os arquivos:**
   ```
   projeto_venda.xlsx
   cabecalho.png
   marca_dagua.png
   projeto_venda2.py
   ```

2. **Execute o script:**
   ```bash
   python projeto_venda2.py
   ```

3. **Resultado:**
   - Será gerado o arquivo `Projeto_Venda_Escola.pdf`
   - O PDF conterá todas as seções formatadas conforme PNAE

---

## 📝 Observações Importantes

### Datas de Assinatura
Todas as datas de assinatura são baseadas na coluna `fim_edital` da aba `edital`, formatadas por extenso (ex: "29 de janeiro de 2026").

### Sazonalidade
A sazonalidade dos produtos é lida diretamente da coluna `sazonalidade` da aba `alimentos`. Preencha conforme a região.

### Status
Apenas registros com status "Ativo" ou "SIM" são incluídos no PDF:
- `status_representante = "Ativo"`
- `status_alimentos = "Ativo"`
- `status_envelope = "SIM"`
- `status_capa = "Ativo"`

### Campos Vazios
Se um campo na planilha estiver vazio, ele aparecerá vazio no PDF. Preencha todos os campos necessários.

---

## 🎨 Personalização

### Marca d'água
- Ajuste a opacidade editando o valor em `opacity=0.2` (linha ~67)
- Valores: 0.1 (10%) a 0.5 (50%)

### Tamanhos de Fonte
- Títulos: `fontSize=12` ou `fontSize=16`
- Células: `fontSize=7`
- Ajuste conforme necessário nas linhas de ParagraphStyle

---

## 📞 Suporte

**Floresta Cast LTDA**  
Eunápolis/BA  
Email: florestacast@outlook.com  
Telefone: (73) 99911-0708

---

## 📄 Licença

© 2026 Floresta Cast LTDA. Todos os direitos reservados.

---

## 🔄 Histórico de Versões

### v1.0 (Janeiro 2026)
- ✅ Geração automática de PDF PNAE
- ✅ Leitura de dados de múltiplas abas Excel
- ✅ Sazonalidade customizável por produto
- ✅ Datas baseadas em fim_edital
- ✅ Marca d'água centralizada
- ✅ Formatação conforme padrão PNAE

---

**Desenvolvido com ❤️ pela Floresta Cast LTDA**
