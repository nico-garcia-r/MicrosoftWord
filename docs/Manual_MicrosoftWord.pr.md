# Microsoft Word

Módulo para trabalhar com Microsoft Word 

*Read this in other languages: [English](Manual_MicrosoftWord.md), [Portugues](Manual_MicrosoftWord.pr.md), [Español](Manual_MicrosoftWord.es.md).*
  
![banner](/docs/imgs/Banner_C:\Users\nicog\Desktop\Rocketbot\modules\MicrosoftWord.png)
## Como instalar este módulo
  
__Baixe__ e __instale__ o conteúdo na pasta 'modules' no caminho do Rocketbot  



## Descrição do comando

### Novo Documento
  
Criar um novo documento Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Abrir documento
  
Abra um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Arquivo|Abra o documento especificado|arquivo.docx|
|Sessão|sessão de arquivo|Word1|

### Ler documento
  
Extraia o texto do documento Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Resultado|Armazenar o resultado em uma variável|Variável|
|Sessão|sessão de arquivo|Word1|
|Adicionar detalhes|||

### Copie e cole o texto
  
Copie o texto entre os intervalos no documento do Word e cole-o em outro documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Início do intervalo|Início do intervalo|0|
|fim do intervalo|Fim do intervalo|40|
|Sessão do arquivo a ser copiado|sessão de arquivo|Word1|
|Arquivo|Escolha o documento|arquivo.docx|

### Copiar texto
  
Copiar texto para prancheta entre intervalos no documento Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Início do intervalo|Início do intervalo|0|
|Fim do intevalo|Fim do intervalo|40|
|Sessão|sessão de arquivo|Word1|

### Colar texto
  
Colar texto da prancheta para documento Word
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Contar caracteres
  
Contar caracteres em um parágrafo específico
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Parágrafo|Parágrafo para contar caracteres|1|
|Result|Armazenar o resultado em uma variável|Variável|

### Adicionar tabela
  
Adicionar tabela em um documento do Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Numero de linhas|Número de linhas que a tabela terá|3 |
|Numero de colunas|Número de colunas que a tabela terá|4 |
|Estilo da tabela|Estilo de borda da tabela||
|Sessão|sessão de arquivo|Word1|
|Estilos de borda|||

### Ler tabelas
  
Extraia os dados das tabelas no documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Tabela para ler|Número da tabela a partir da qual o conteúdo será lido|1|
|Sessão|sessão de arquivo|Word1|
|Result|Armazenar o resultado em uma variável|Variável|

### Editar tabela
  
Editar uma tabela em um documento Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Número da tabela|Número da tabela a ser editada|1|
|Sessão|sessão de arquivo|Word1|
|Digite o número da linha para excluir|Número da linha a ser removida da tabela| |
|Digite o número da coluna para excluir|Número da coluna a ser removida da tabela| |
|Inserir linha|Se selecionado, adiciona uma linha ao final da tabela||
|Inserir coluna|Se selecionado, adiciona uma coluna ao final da tabela||
|Largura da coluna|Largura que cada coluna da tabela terá|140|
|Altura da linha|Altura que cada linha da tabela terá|25|

### Salvar documento
  
Salve o documento Word aberto
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Salvar arquivo|Salve o arquivo no caminho especificado|arquivo.docx|

### Escrever no documento
  
Escreva em um documento Word.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Escrever texto|Texto a ser escrito no documento|Lorem ipsum |
|Tipo de texto|Seletor de tipo de texto||
|Nível||1-9|
|Tamanho da fonte||12|
|Alinhamento|||
|Negrito|||
|Itálico|||
|Sublinhar|||

### Fechar documento
  
Feche o documento que está sendo executado
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Inserir página
  
Inserir uma nova página no documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|

### Adicionar imagem
  
Adicionar uma imagem ao documento
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Rota da imagem|Direção da imagem que será adicionado abaixo do último parágrafo|imagem.jpg|

### Converter para PDF
  
Converter documento Word para PDF.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Salvar arquivo|Caminho do arquivo onde o PDF será criado|arquivo.pdf|

### Localizar texto no parágrafo
  
Localize o parágrafo onde se encontra o texto indicado.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Texto para pesquisar|Texto que será usado para localizar o parágrafo|Olá mundo|
|Nome variável|Armazenar o resultado em uma variável|Variável|

### Contar parágrafos
  
Conte o número de parágrafos no documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Nome variável|Armazenar o resultado em uma variável|Variável|

### Substituir texto no parágrafo
  
Substituir o texto de um parágrafo.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Texto para pesquisar||Olá mundo|
|Texto a substituir|Texto a ser substituído|Olá mundo|
|Lista de parágrafos|Parágrafos onde o texto especificado será pesquisado|Exemplo ',' separado por vírgula: 1,2|

### Excluir parágrafo
  
Excluir um parágrafo do documento.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Número do parágrafo|Número do parágrafo a ser excluído|1|
|Nome da variável onde o parágrafo excluído será salvo|Variável onde será salvo o texto que incluiu o parágrafo excluído|Variável|

### Adicionar texto a um bookmark
  
Adicionar texto a um bookmark.
|Parâmetros|Descrição|exemplo|
| --- | --- | --- |
|Sessão|sessão de arquivo|Word1|
|Texto a adicionar||Olá mundo|
|Nome do marcador||Marcador 1|