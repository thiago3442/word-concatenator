# Word Concatenator

Script Python para consolidar múltiplos arquivos .docx em um único documento final, preservando a formatação básica.

## 📋 Descrição

Este projeto consolida múltiplos arquivos Word (.docx) de uma pasta em um único documento, mantendo:
- Formatação de títulos e parágrafos
- Estilos de texto (negrito, itálico, sublinhado)
- Tamanho e fonte do texto
- Quebras de página entre documentos
- Nome do arquivo original como título de cada seção

## 🚀 Funcionalidades

### Consolidação de Documentos
- ✅ Leitura automática de todos os arquivos .docx da pasta `input/`
- ✅ Consolidação em ordem alfabética
- ✅ Preservação de formatação básica
- ✅ Quebra de página entre documentos
- ✅ Título centralizado com o nome de cada arquivo
- ✅ Tratamento robusto de erros
- ✅ Suporte para grande volume de arquivos (testado com 90 arquivos)

### Formatação como Livro de Poemas
- ✅ Página de título profissional
- ✅ Índice automático com todos os poemas
- ✅ Formatação elegante com fonte Georgia
- ✅ Títulos centralizados e decorados
- ✅ Espaçamento otimizado entre estrofes
- ✅ Numeração de páginas no rodapé
- ✅ Margens ajustadas para impressão

## 📦 Instalação

### Pré-requisitos

- Python 3.7 ou superior
- pip (gerenciador de pacotes Python)

### Instalação de Dependências

```bash
pip install -r requirements.txt
```

Ou manualmente:

```bash
pip install python-docx==1.1.2
```

## 🎯 Uso

### 1. Consolidar Arquivos

1. Coloque todos os arquivos .docx que deseja consolidar na pasta `input/`
2. Execute o script de consolidação:

```bash
python src/consolidate_docs.py
```

3. O arquivo consolidado será criado em `output/consolidado.docx`

### 2. Formatar como Livro de Poemas

Para transformar o arquivo consolidado em um livro de poemas profissional:

```bash
python src/format_as_poetry_book.py
```

O livro formatado será criado em `output/livro_de_poemas.docx` com:
- Página de título elegante
- Índice completo
- Formatação profissional para cada poema
- Numeração de páginas

### Uso Programático

Você também pode importar e usar as funções do script em seu próprio código:

```python
from src.consolidate_docs import consolidate_docx_files

# Consolidar arquivos
output_file = consolidate_docx_files(
    input_folder="input",
    output_folder="output",
    output_filename="meu_consolidado.docx",
    add_filename_titles=True  # Adiciona nome dos arquivos como títulos
)

print(f"Arquivo criado: {output_file}")
```

## 📁 Estrutura do Projeto

```
word-concatenator/
├── input/                  # Pasta com arquivos .docx de entrada
│   ├── arquivo1.docx
│   ├── arquivo2.docx
│   └── ...
├── output/                 # Pasta com arquivo consolidado (criada automaticamente)
│   └── consolidado.docx
├── src/
│   ├── __init__.py
│   ├── consolidate_docs.py      # Script de consolidação
│   └── format_as_poetry_book.py # Script de formatação como livro
├── requirements.txt        # Dependências do projeto
├── Makefile               # Comandos úteis
└── README.md              # Este arquivo
```

## ⚙️ Configuração

### Script de Consolidação (`src/consolidate_docs.py`)

```python
INPUT_FOLDER = "input"              # Pasta de entrada
OUTPUT_FOLDER = "output"            # Pasta de saída
OUTPUT_FILENAME = "consolidado.docx" # Nome do arquivo final
ADD_TITLES = True                    # Adicionar títulos com nomes dos arquivos
```

### Script de Formatação (`src/format_as_poetry_book.py`)

```python
INPUT_FILE = "output/consolidado.docx"      # Arquivo consolidado
OUTPUT_FILE = "output/livro_de_poemas.docx" # Arquivo formatado
BOOK_TITLE = "Coletânea de Poemas"          # Título do livro
AUTHOR = ""                                  # Nome do autor (opcional)
```

## 🛠️ Comandos Make

Se você tiver o `make` instalado, pode usar os seguintes comandos:

```bash
make install    # Instala as dependências
make run        # Executa o script de consolidação
make clean      # Limpa arquivos temporários
```

## 📖 Formatação do Livro de Poemas

O script `format_as_poetry_book.py` cria um livro profissional com:

### Estrutura
1. **Página de Título**: Com título do livro, subtítulo e autor (opcional)
2. **Índice**: Lista completa de todos os poemas
3. **Poemas**: Cada poema em página individual com:
   - Título centralizado e em negrito
   - Linha decorativa (• • •)
   - Conteúdo do poema centralizado
   - Espaçamento adequado entre estrofes

### Formatação
- **Fonte**: Georgia (elegante e apropriada para poesia)
- **Tamanho**: 
  - Título do livro: 24pt
  - Títulos de poemas: 14pt
  - Texto dos poemas: 11pt
- **Margens**: 1.25" laterais, 1" superior/inferior
- **Alinhamento**: Centralizado
- **Numeração**: Páginas numeradas no rodapé

## 📝 Funcionalidades Técnicas

### Ordem de Consolidação

Os arquivos são processados em **ordem alfabética** dos nomes. Exemplos:
- `A documento.docx` → primeiro
- `B documento.docx` → segundo
- `documento 01.docx` → terceiro

### Formatação Preservada

- **Estilos de parágrafo**: Títulos, subtítulos, texto normal
- **Formatação de texto**: Negrito, itálico, sublinhado
- **Fontes**: Nome e tamanho da fonte
- **Alinhamento**: Esquerda, centro, direita, justificado

### Tratamento de Erros

- Ignora arquivos que não sejam .docx
- Continua processamento se houver erro em um arquivo específico
- Mensagens de erro claras e informativas
- Não interrompe a consolidação por erros individuais

## 🧪 Testado Com

- ✅ 90 arquivos .docx simultâneos
- ✅ Documentos com formatação complexa
- ✅ Diferentes estilos e fontes
- ✅ Windows 11 / Python 3.12

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para:
- Reportar bugs
- Sugerir novas funcionalidades
- Enviar pull requests

## 📄 Licença

Este projeto é de código aberto e está disponível sob a licença MIT.

## 👤 Autor

Desenvolvido para consolidação de documentos Word de forma automatizada e eficiente.

## 🔍 Solução de Problemas

### Erro: "Pasta não encontrada"
- Certifique-se de que a pasta `input/` existe
- Verifique se você está executando o script do diretório raiz do projeto

### Erro: "Nenhum arquivo .docx encontrado"
- Verifique se há arquivos .docx na pasta `input/`
- Certifique-se de que os arquivos têm a extensão correta (.docx, não .doc)

### Problemas de formatação
- O script preserva formatação básica, mas algumas formatações avançadas podem não ser copiadas
- Tabelas, imagens e objetos incorporados podem não ser incluídos

## 📊 Exemplo de Saída

```
Encontrados 90 arquivos .docx para consolidar
Processando [1/90]: A ESCOLA DOS MEUS SONHOS.docx
Processando [2/90]: A JANELA E O ESPELHO.docx
...
Processando [90/90]: ZANGADO.docx

✓ Arquivo consolidado criado com sucesso: output\consolidado.docx
✓ Total de documentos consolidados: 90

============================================================
CONSOLIDAÇÃO CONCLUÍDA COM SUCESSO!
============================================================
Arquivo gerado: output\consolidado.docx
