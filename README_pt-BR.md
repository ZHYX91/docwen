[English](README.md) | [简体中文](README_zh-CN.md) | [繁體中文](README_zh-TW.md) | [Deutsch](README_de-DE.md) | [Français](README_fr-FR.md) | [Русский](README_ru-RU.md) | [Português](README_pt-BR.md) | [日本語](README_ja-JP.md) | [한국어](README_ko-KR.md) | [Español](README_es-ES.md) | [Tiếng Việt](README_vi-VN.md)

# DocWen

Uma ferramenta de conversão de formatos de documentos e gráficos que suporta conversão bidirecional Word/Markdown/Excel. Executa completamente localmente, garantindo a segurança e confiabilidade dos dados.

## 📖 Contexto do Projeto

Este software foi originalmente projetado para o trabalho diário do escritório de impressão para resolver os seguintes problemas:
- Os formatos de documentos enviados por vários departamentos são caóticos e precisam ser organizados em formatos padronizados.
- Existem muitos tipos de documentos, cada um com diferentes requisitos de formato fixo.
- Precisa rodar offline, adaptando-se a ambientes de intranet e equipamentos legados.

**Filosofia de Design**: Este software posiciona-se como uma ferramenta leve e à prova de falhas. Embora não possa ser comparado com ferramentas profissionais como LaTeX ou Pandoc em termos de profissionalismo e integridade funcional, ele se destaca pelo custo zero de aprendizado e usabilidade imediata, tornando-o adequado para cenários de escritório diários onde os requisitos de formato não são extremamente rigorosos.

## ✨ Funcionalidades Principais

- **📄 Conversão de Formato de Documento** - Conversão bidirecional Word ↔ Markdown. Suporta conversão de fórmulas matemáticas e conversão bidirecional de separadores (três tipos de separadores do Markdown vs. quebras de página, quebras de seção e linhas horizontais do Word). Suporta formatos como DOCX/DOC/WPS/RTF/ODT.
- **📊 Conversão de Formato de Planilha** - Conversão bidirecional Excel ↔ Markdown. Suporta formatos XLSX/XLS/ET/ODS/CSV. Inclui ferramentas de resumo de tabelas.
- **📑 PDF e Arquivos de Layout** - Conversão de PDF/XPS/OFD para Markdown ou DOCX. Suporta fusão, divisão e outras operações de PDF.
- **🖼️ Processamento de Imagem** - Suporta conversão bidirecional e compressão de formatos JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **🔍 Reconhecimento de Texto OCR** - RapidOCR integrado para extrair texto de imagens e PDFs.
- **✏️ Revisão de Texto** - Verifica erros de digitação, pontuação, símbolos e palavras sensíveis com base em dicionários personalizados. As regras podem ser editadas na interface de configurações.
- **📝 Sistema de Modelos** - Mecanismo de modelo flexível que suporta formatos personalizados de documentos e relatórios.
- **💻 Operação em Modo Duplo** - Interface Gráfica do Usuário (GUI) + Interface de Linha de Comando (CLI).
- **🔒 Operação Completamente Local** - Executa offline, garantindo a segurança dos dados com mecanismos integrados de isolamento de rede.
- **🔗 Operação de Instância Única** - Gerencia automaticamente instâncias do programa e suporta integração com o plugin Obsidian acompanhante.

## Registro de Alterações

### v0.6.0 (2025-01-20)

- Suporte completo à internacionalização (GUI e CLI suportam 11 idiomas).
- Substituição do PaddleOCR pelo RapidOCR para melhor compatibilidade.
- Adição de modelos Word/Excel multilíngues.
- Detecção e injeção automática de estilos de modelo.
- Outras otimizações e correções.

### v0.5.1 (2025-01-01)

- Adicionada conversão bidirecional de fórmulas matemáticas (Word OMML ↔ Markdown LaTeX).
- Adicionada conversão bidirecional de notas de rodapé/notas de fim.
- Adicionados estilos de caracteres e parágrafos para código, citações, etc.
- Processamento de lista aprimorado (aninhamento multinível, numeração automática).
- Funções de tabela aprimoradas (detecção/injeção de estilo, tabelas de três linhas, etc.).
- Otimizada a limpeza e adição de números de subtítulos.
- Interação de interface e vinculação de configurações melhoradas.

### v0.4.1 (2025-12-05)

- CLI refatorada para melhorar a experiência do usuário.
- Adicionado suporte para mais tipos de documentos.
- Implementadas mais opções configuráveis.

## 🚀 Início Rápido

### Iniciar Programa

Clique duas vezes em `DocWen.exe` para iniciar a interface gráfica.

### Guia de Início Rápido

1.  **Prepare um Arquivo Markdown**:

    ```markdown
    ---
    Título: Documento de Teste
    ---
    
    ## Título de Teste
    
    Este é o conteúdo do corpo de teste.
    ```

2.  **Conversão Arrastar e Soltar**:
    - Inicie o programa.
    - Arraste o arquivo `.md` para a janela.
    - Selecione um modelo.
    - Clique em "Converter para DOCX".

3.  **Obter Resultados**:
    - Um documento Word padronizado será gerado no mesmo diretório.

**Dica**: Você pode usar os arquivos de exemplo no diretório `samples/` para experimentar rapidamente as funcionalidades do software.

## 📝 Convenções de Sintaxe Markdown

### Mapeamento de Nível de Cabeçalho

Para facilitar a memorização por colegas sem conhecimento prévio, os cabeçalhos Markdown neste software correspondem **um-para-um** com os cabeçalhos do Word:
- O título e subtítulo do documento são colocados nos metadados YAML.
- Markdown `# Cabeçalho 1` corresponde ao Word "Cabeçalho 1".
- Markdown `## Cabeçalho 2` corresponde ao Word "Cabeçalho 2".
- E assim por diante, suportando até 9 níveis de cabeçalhos.

**Dica**: Se você prefere usar o cabeçalho de primeiro nível do Markdown (`#`) como o título do documento, começando com cabeçalhos de segundo nível (`##`) para os subtítulos do corpo, você pode estilizar "Título 1" no modelo Word para parecer com um título de documento (por exemplo, centralizado, negrito, tamanho de fonte maior), e selecionar um esquema de numeração que ignore a numeração de cabeçalhos de primeiro nível nas configurações. Desta forma, seus cabeçalhos de primeiro nível aparecerão como títulos de documento.

### Quebras de Linha e Parágrafos

**Regra Básica**: Cada linha não vazia é tratada como um parágrafo separado por padrão.

**Parágrafos Mistos**: Quando um subtítulo precisa ser misturado com o corpo do texto no mesmo parágrafo, as seguintes condições devem ser atendidas:
1.  O subtítulo termina com um sinal de pontuação de fim (suporta pontuação multilíngue, incluindo pontos, pontos de interrogação, pontos de exclamação e outros sinais de pontuação de fim comuns).
2.  O corpo do texto está localizado na **linha imediatamente seguinte** do subtítulo.
3.  A linha do corpo do texto não pode ser um elemento Markdown especial (como cabeçalhos, blocos de código, tabelas, listas, citações, blocos de fórmula, separadores, etc.).

**Exemplo**:
```markdown
## I. Requisitos de Trabalho.
Esta reunião exige que todas as unidades implementem seriamente...
```
As duas linhas acima serão mescladas no mesmo parágrafo, onde "I. Requisitos de Trabalho." mantém o formato de subtítulo, e "Esta reunião..." mantém o formato de corpo de texto.

**Nota**:
- Não pode haver uma linha vazia entre o subtítulo e o corpo do texto; caso contrário, serão reconhecidos como parágrafos separados.
- Se o subtítulo não terminar com um sinal de pontuação e não houver linha vazia antes do corpo do texto, o corpo do texto será mesclado na linha do cabeçalho com formatação ajustada.

### Conversão Bidirecional de Separadores

Suporta conversão bidirecional entre separadores Markdown e quebras de página/quebras de seção/linhas horizontais do Word:

-   **DOCX → MD**: Quebras de página, quebras de seção e linhas horizontais do Word são automaticamente convertidas para separadores Markdown.
-   **MD → DOCX**: Markdown `---`, `***`, `___` são automaticamente convertidos para elementos correspondentes do Word.
-   **Configurável**: Relações de mapeamento específicas podem ser personalizadas na interface de configurações.

## 📖 Guia de Uso Detalhado

### Word para Markdown

1.  Arraste o arquivo `.docx` para a janela do programa.
2.  O programa analisa automaticamente a estrutura do documento.
3.  Gera um arquivo `.md` contendo metadados YAML.

**Formatos Suportados**:
-   `.docx` - Documento Word Padrão.
-   `.doc` - Convertido automaticamente para DOCX para processamento.
-   `.wps` - Documento WPS convertido automaticamente.

**Opções de Exportação**:

| Opção | Descrição |
| :--- | :--- |
| **Extrair Imagens** | Se marcado, as imagens no documento são extraídas para a pasta de saída, e links de imagem são inseridos no arquivo MD. |
| **OCR de Imagem** | Se marcado, realiza OCR em imagens e cria um arquivo `.md` de imagem (contendo texto reconhecido). |
| **Limpar Números de Subtítulo** | Se marcado, remove números antes dos subtítulos (por exemplo, "一、", "（一）", "1.", etc.) e os converte em texto de título puro. |
| **Adicionar Números de Subtítulo** | Se marcado, adiciona automaticamente números com base nos níveis de cabeçalho (esquema de numeração pode ser configurado nas configurações). |

### Markdown para Word

1.  Prepare um arquivo `.md` com um cabeçalho YAML.
2.  Arraste-o para a janela do programa e selecione o modelo Word correspondente.
3.  O programa preenche automaticamente o modelo e gera o documento.

**Opções de Conversão**:

| Opção | Descrição |
| :--- | :--- |
| **Limpar Números de Subtítulo** | Se marcado, remove números antes dos subtítulos. |
| **Adicionar Números de Subtítulo** | Se marcado, adiciona automaticamente números com base nos níveis de cabeçalho. |

**Nota**: Se houver parágrafos onde subtítulos e corpo de texto são misturados no documento, quebras de linha estritas devem ser mantidas no arquivo MD (veja "Quebras de Linha e Parágrafos" acima).

### Processamento Automático de Estilo de Modelo

O conversor detecta e processa automaticamente estilos de modelo durante a conversão Markdown → DOCX:

#### Classificação de Estilo

**Estilo de Parágrafo**: Aplicado a todo o parágrafo.

| Estilo | Comportamento de Detecção | Injeção quando Ausente | Fonte |
| :--- | :--- | :--- | :--- |
| Cabeçalho (1~9) | Detecta estilo de parágrafo | Estilos de cabeçalho de modelo | Word Embutido |
| Bloco de Código | Detecta estilo de parágrafo | Fonte Consolas + Fundo cinza | Definido pelo Software |
| Citação (1~9) | Detecta estilo de parágrafo | Fundo cinza + Borda esquerda | Definido pelo Software |
| Bloco de Fórmula | Detecta estilo de parágrafo | Estilo específico de fórmula | Definido pelo Software |
| Separador (1~3) | Detecta estilo de parágrafo | Estilo de parágrafo de borda inferior | Definido pelo Software |

**Estilo de Caractere**: Aplicado ao texto selecionado.

| Estilo | Comportamento de Detecção | Injeção quando Ausente | Fonte |
| :--- | :--- | :--- | :--- |
| Código em Linha | Detecta estilo de caractere | Fonte Consolas + Sombreamento cinza | Definido pelo Software |
| Fórmula em Linha | Detecta estilo de caractere | Estilo específico de fórmula | Definido pelo Software |

**Estilo de Tabela**: Aplicado a toda a tabela.

| Estilo | Comportamento de Detecção | Injeção quando Ausente | Fonte |
| :--- | :--- | :--- | :--- |
| Tabela de Três Linhas | Prioridade de configuração do usuário | Definição de estilo de tabela de três linhas | Definido pelo Software |
| Tabela de Grade | Prioridade de configuração do usuário | Definição de estilo de tabela de grade | Definido pelo Software |

**Definição de Numeração**: Usado para formatos de lista.

| Tipo | Comportamento de Detecção | Manuseio quando Ausente |
| :--- | :--- | :--- |
| Numeração de Lista | Verifica definições de lista ordenada/não ordenada existentes no modelo | Usa predefinição decimal/marcador |

#### Internacionalização de Nome de Estilo

-   **Estilos Embutidos do Word** (cabeçalho 1~9):
    -   Nomes de estilo usam nomes ingleses padrão do Word (por exemplo, `heading 1`).
    -   O Word exibe automaticamente nomes localizados com base no idioma do sistema (por exemplo, "Título 1" em sistemas em português).
-   **Estilos Definidos pelo Software** (Bloco de Código, Citação, Fórmula, Separador, Tabela, etc.):
    -   Injeta nomes de estilo de idioma correspondentes com base na configuração de idioma da interface do software.
    -   Interface em Chinês: Injeta "代码块", "引用 1", "三线表", etc.
    -   Interface em Inglês: Injeta "Code Block", "Quote 1", "Three Line Table", etc.

**Sugestão**: Depois de personalizar estilos no modelo, o conversor usará automaticamente seus estilos; se não estiverem presentes no modelo, usará estilos predefinidos embutidos.

### Processamento de Arquivo de Planilha

1.  **Excel/CSV para Markdown**: Arraste arquivos `.xlsx` ou `.csv` para converter automaticamente em tabelas Markdown.
2.  **Markdown para Excel**: Prepare um arquivo MD e selecione um modelo Excel para conversão.

**Formatos Suportados**:
-   `.xlsx` - Documento Excel Padrão.
-   `.xls` - Convertido automaticamente para XLSX para processamento.
-   `.et` - Planilha WPS convertida automaticamente.
-   `.csv` - Tabela de texto CSV.

### Função de Revisão de Texto

O programa fornece quatro regras de revisão personalizáveis:

1.  **Verificação de Emparelhamento de Pontuação** - Detecta se pontuações emparelhadas como parênteses e aspas correspondem.
2.  **Revisão de Símbolos** - Detecta uso misto de pontuação chinesa e inglesa.
3.  **Verificação de Erros de Digitação** - Verifica erros de digitação comuns com base em um dicionário personalizado.
4.  **Detecção de Palavras Sensíveis** - Detecta palavras sensíveis com base em um dicionário personalizado.

**Dicionários Personalizados**: Edite visualmente dicionários de erros de digitação e palavras sensíveis na interface "Configurações".

**Uso**:
1.  Arraste o documento Word a ser revisado para o programa.
2.  Marque as regras de revisão necessárias.
3.  Clique no botão "Revisão de Texto".
4.  Os resultados da revisão são exibidos como comentários no documento.

## 🛠️ Sistema de Modelos

### Usando Modelos Existentes

O programa vem com vários modelos, incluindo versões multilíngues. Você pode selecionar e usar conforme necessário. Os arquivos de modelo estão localizados no diretório `templates/`.

### Modelos Personalizados

1.  Crie um arquivo de modelo usando Word ou WPS.
2.  Consulte modelos existentes e insira espaços reservados como `{{Title}}`, etc., onde o preenchimento é necessário.
3.  No modelo, estilos embutidos Título 1 ~ Título 5 precisam ser modificados manualmente.
4.  Salve o modelo no diretório `templates/`.
5.  Reinicie o programa, e o novo modelo será carregado automaticamente.

Você também pode copiar um modelo existente, modificá-lo e renomeá-lo.

### Uso de Espaço Reservado

#### Espaços Reservados de Modelo Word

**Espaços Reservados de Campo YAML**: Use o formato `{{Nome do Campo}}` no modelo, que será substituído pelo valor correspondente no cabeçalho YAML do arquivo Markdown durante a conversão.

| Espaço Reservado | Descrição |
| :--- | :--- |
| `{{Título}}` | Título do documento (Regras de recuperação veja abaixo) |
| `{{Corpo}}` | Posição de inserção do conteúdo do corpo Markdown |
| Outros | Suporta qualquer campo personalizado |

**Prioridade de Recuperação de Título**:

| Prioridade | Fonte | Descrição |
| :--- | :--- | :--- |
| 1 | Campo YAML `Title` | Maior prioridade |
| 2 | Campo YAML `aliases` | Pega o primeiro elemento da lista, ou valor da string |
| 3 | Nome do arquivo | Nome do arquivo sem extensão `.md` |

**Suporte multilíngue**: Os espaços reservados título e corpo suportam múltiplos idiomas, ex: título pode ser `{{Título}}`, `{{title}}`, `{{标题}}`, etc., corpo pode ser `{{Corpo}}`, `{{body}}`, `{{正文}}`, etc.

#### Espaços Reservados de Modelo Excel

Modelos Excel suportam três tipos de espaços reservados:

**1. Espaço Reservado de Campo YAML** `{{Nome do Campo}}`

Usado para preencher um único valor do cabeçalho YAML do arquivo Markdown:

```markdown
---
ReportName: Estatísticas de Vendas Anuais 2024
Unit: Depto de Vendas
---
```

`{{ReportName}}`, `{{Unit}}` no modelo serão substituídos pelos valores correspondantes. O campo de título também segue as regras de prioridade.

**2. Espaço Reservado de Preenchimento de Coluna** `{{↓Nome do Campo}}`

Extrai dados da tabela Markdown e preenche **para baixo** linha por linha a partir da posição do espaço reservado:

```markdown
| ProductName | Quantity |
|:--- |:--- |
| Produto A | 100 |
| Produto B | 200 |
```

`{{↓ProductName}}` no modelo Excel será substituído por "Produto A", e a próxima linha será preenchida com "Produto B".

**3. Espaço Reservado de Preenchimento de Linha** `{{→Nome do Campo}}`

Extrai dados da tabela Markdown e preenche **para a direita** coluna por coluna a partir da posição do espaço reservado:

```markdown
| Month |
|:--- |
| Jan |
| Fev |
| Mar |
```

`{{→Month}}` no modelo Excel será preenchido com "Jan", "Fev", "Mar" sequencialmente para a direita.

**Manuseio de Células Mescladas**: O programa pula automaticamente células não primárias de células mescladas para garantir o preenchimento correto dos dados.

**Mesclagem de Dados Multi-tabela**: Se houver várias tabelas no Markdown usando o mesmo nome de cabeçalho, os dados serão mesclados em ordem e preenchidos sequencialmente.

## 🖥️ Uso da Interface Gráfica

A maioria dos usuários usa este software através da interface gráfica. Aqui está o guia de operação detalhado.

### Visão Geral da Interface

O programa usa um **layout adaptável de três colunas**:

| Área | Descrição | Tempo de Exibição |
| :--- | :--- | :--- |
| **Coluna Central (Área Principal)** | Área de arrastar e soltar arquivos, painel de operação, barra de status | Sempre exibido |
| **Coluna Direita** | Seletor de modelo / Painel de conversão de formato | Expande automaticamente após selecionar um arquivo |
| **Coluna Esquerda** | Lista de arquivos em lote (agrupados por tipo) | Exibido ao mudar para modo de lote |

### Fluxo de Operação Básico

1.  **Iniciar Programa**: Clique duas vezes em `DocWen.exe`.
2.  **Importar Arquivo**:
    -   Método 1: Arraste e solte arquivos diretamente na janela.
    -   Método 2: Clique no botão "Adicionar" na área de arrastar e soltar para selecionar arquivos.
3.  **Selecionar Modelo** (se a conversão for necessária): O painel de modelo direito expande automaticamente; selecione um modelo adequado.
4.  **Configurar Opções**: Marque as opções de conversão/exportação necessárias no painel de operação.
5.  **Executar Operação**: Clique no botão de função correspondente (por exemplo, "Exportar MD", "Converter para DOCX", etc.).
6.  **Ver Resultado**: A barra de status mostra o progresso e os resultados; clique no ícone 📍 para localizar o arquivo de saída.

### Modo de Arquivo Único vs. Modo de Lote

O programa suporta dois modos de processamento, alternáveis via botão de alternância na área de arrastar e soltar arquivos:

**Modo de Arquivo Único** (Padrão):
-   Processa um arquivo de cada vez.
-   Interface simples, adequada para uso diário.

**Modo de Lote**:
-   Importa vários arquivos simultaneamente.
-   Coluna esquerda mostra lista de arquivos categorizada (agrupados por documento/planilha/imagem, etc.).
-   Suporta adição, remoção e classificação em lote.
-   Clicar em um arquivo na lista muda o alvo da operação atual.

### Funções do Painel de Operação

O painel de operação ajusta automaticamente as opções disponíveis com base no tipo de arquivo:

| Tipo de Arquivo | Operações Disponíveis |
| :--- | :--- |
| Documento Word | Exportar MD, Converter PDF, Revisão de Texto, OCR |
| Markdown | Converter DOCX, Converter PDF |
| Planilha Excel | Exportar MD, Converter PDF, Resumo de Tabela |
| Arquivo PDF | Exportar MD, Mesclar, Dividir, OCR |
| Arquivo de Imagem | Conversão de Formato, Compressão, OCR |

### Interface de Configurações

Clique no botão ⚙️ no canto inferior direito da janela para abrir as configurações:

-   **Geral**: Tema da interface, idioma, opacidade da janela.
-   **Conversão**: Valores padrão para várias opções de conversão.
-   **Saída**: Diretório de saída padrão, regras de nomeação de arquivos.
-   **Revisão**: Edite dicionários de erros de digitação e palavras sensíveis.
-   **Estilo**: Configurações de estilo de bloco de código, citação, tabela.

### Atalhos

-   **Arrastar Arquivo Externo**: Arraste diretamente para a janela para importar.
-   **Duplo clique no Resultado da Barra de Status**: Abre rapidamente o diretório do arquivo de saída.
-   **Clique com o botão direito no Item de Modelo**: Abre a localização do arquivo de modelo.

---

## 🔧 Uso da Linha de Comando

Além da GUI, o programa fornece uma Interface de Linha de Comando (CLI), adequada para scripts de automação e cenários de processamento em lote.

### Modos de Execução

-   **Modo Interativo**: Exibe um guia de menu após passar um arquivo, semelhante à operação GUI.
-   **Modo Headless**: Executa diretamente adicionando o parâmetro `--action`, adequado para invocação de script.

### Exemplos Comuns

```bash
# Modo Interativo
Docwen.exe documento.docx

# Exportar Word para Markdown (Extrair Imagens + OCR)
Docwen.exe relatorio.docx --action export_md --extract-img --ocr

# Markdown para Word (Especificar Modelo)
Docwen.exe documento.md --action convert --target docx --template "Nome do Modelo"

# Conversão em Lote (Pular confirmação, continuar em caso de erro)
Docwen.exe *.docx --action export_md --batch --yes --continue-on-error

# Revisão de Documento
Docwen.exe documento.docx --action validate --check-typo --check-punct

# Mesclar/Dividir PDF
Docwen.exe *.pdf --action merge_pdfs
Docwen.exe relatorio.pdf --action split_pdf --pages "1-3,5,7-10"
```

### Argumentos Principais

| Argumento | Descrição |
| :--- | :--- |
| `--action` | Tipo de operação: `export_md`, `convert`, `validate`, `merge_pdfs`, `split_pdf` |
| `--target` | Formato alvo: `pdf`, `docx`, `xlsx`, `md` |
| `--template` | Nome do modelo (por exemplo, `Nome do Modelo`) |
| `--extract-img` | Extrair imagens durante a exportação |
| `--ocr` | Habilitar reconhecimento OCR |
| `--batch` | Modo de processamento em lote |
| `--yes` / `-y` | Pular prompts de confirmação |
| `--continue-on-error` | Continuar processando o próximo item em caso de erro |
| `--json` | Resultado de saída em formato JSON |
| `--quiet` / `-q` | Modo silencioso, reduzir saída |

## 🔌 Plugin Obsidian

O projeto inclui um plugin Obsidian correspondente para alcançar integração com o conversor:

### Funcionalidades Principais

-   **🚀 Lançamento em Um Clique** - Ícone da barra lateral para iniciar rapidamente o conversor.
-   **📂 Transferência Automática** - Passa automaticamente o caminho do arquivo aberto atualmente.
-   **🔄 Gerenciamento de Instância Única** - Envia automaticamente o arquivo se o programa já estiver em execução, sem necessidade de reiniciar.
-   **💪 Recuperação de Falhas** - Detecta automaticamente o status do processo e limpa automaticamente arquivos residuais.

### Princípio de Funcionamento

O plugin interage com o conversor via IPC baseado em sistema de arquivos:

1.  **Primeiro Clique** → Inicia o conversor e passa o arquivo atual.
2.  **Clique Novamente (Com Arquivo)** → Substitui pelo novo arquivo (Modo de Arquivo Único).
3.  **Clique Novamente (Sem Arquivo)** → Ativa a janela do conversor.

### Instalação

O plugin foi lançado em um repositório separado. Visite [docwen-obsidian](https://github.com/ZHYX91/docwen-obsidian) para instruções de instalação e a versão mais recente.

## ❓ Perguntas Frequentes

### E se a conversão falhar?

-   Verifique se o arquivo está ocupado por outro programa.
-   Confirme se o formato do arquivo está correto.
-   Verifique os logs de erro no diretório `logs/`.

### Modelo não aparece?

-   Confirme se os arquivos de modelo estão no diretório `templates/`.
-   Verifique se o arquivo de modelo está corrompido.
-   Reinicie o programa para recarregar os modelos.

### Função de revisão não funciona?

-   Confirme se o documento está no formato .docx.
-   Verifique se o documento contém texto editável.
-   Confirme se as regras de revisão estão habilitadas nas configurações.

### Formato de saída não conforme esperado?

-   O programa gera documentos com base nos estilos de modelo. Para ajustar o formato de saída, modifique as definições de estilo diretamente no arquivo de modelo.
-   Os arquivos de modelo estão localizados no diretório `templates/`.
-   Após modificar os estilos de modelo, todos os documentos convertidos com esse modelo aplicarão os novos estilos.

### Células de fórmula ficam vazias após a conversão de Excel para Markdown?

Este é o comportamento esperado. O programa lê os **valores em cache** das células em vez das próprias fórmulas.

**Razão técnica**:
-   Em arquivos Excel, as células de fórmula armazenam tanto a fórmula quanto o último resultado calculado (valor em cache).
-   O programa usa o modo `data_only=True`, que recupera apenas valores em cache.
-   Se o arquivo nunca foi aberto no Excel (por exemplo, gerado por um programa), ou foi editado mas não salvo novamente, o valor em cache estará vazio.

**Solução**:
1.  Abra o arquivo no Excel.
2.  Aguarde a conclusão do cálculo das fórmulas.
3.  Salve o arquivo.
4.  Converta novamente.

## 🔒 Recursos de Segurança

-   **Operação Completamente Local**: Todo o processamento é feito localmente, sem dependência de rede.
-   **Isolamento de Rede**: Mecanismo de isolamento de rede integrado evita vazamento de dados.
-   **Sem Upload de Dados**: Arquivos de usuário nunca são carregados para nenhum servidor.

## 📜 Licença

Este projeto é licenciado sob a **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Este projeto usa PyMuPDF (licenciado sob AGPL-3.0), portanto, todo o projeto também é licenciado sob AGPL-3.0.
-   Você é livre para usar, modificar e distribuir este software.
-   Se você modificar este software e fornecer serviços através de uma rede, você deve fornecer o código-fonte modificado aos usuários.
-   Para informações detalhadas sobre a licença, consulte o arquivo [LICENSE](LICENSE).

### Contato

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Autor de Contato**: zhengyx91@hotmail.com

---

**Autor**: ZhengYX
