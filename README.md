# Conversor PDF e DOCX

Aplicativo independente em Python e Streamlit, com toda a interface e as duas
conversões no arquivo único `pdf-docx.py`. Não altera o aplicativo PDF/Excel hospedado.

## O que o aplicativo faz

| Aba | Entrada | Saída | Mecanismo utilizado | Precisa de LibreOffice? |
| --- | --- | --- | --- | --- |
| PDF → DOCX | `.pdf` | Word editável (`.docx`) | Biblioteca Python `pdf2docx` | Não |
| DOCX → PDF | `.docx` | `.pdf` | LibreOffice Writer em modo sem interface gráfica | Sim |

O navegador apresenta a interface do Streamlit. A conversão acontece na máquina
em que o programa está executando, sem utilizar APIs de conversão.

## Requisitos

- **Python 3.11 ou 3.12**, com `pip`.
- As bibliotecas de `requirements.txt`.
- Um navegador atualizado.
- **LibreOffice com Writer instalado**, para utilizar DOCX → PDF.
- Permissão para criar arquivos no diretório temporário do sistema.
- Espaço em disco e memória suficientes para os documentos.

É necessária conexão à internet para baixar os programas e as bibliotecas na
instalação. Depois de instalados, as conversões não dependem de uma API ou de
conexão com um serviço externo. **Microsoft Word não é necessário.**

## Conteúdo

- `pdf-docx.py`: código completo, com uma aba PDF → DOCX e outra DOCX → PDF;
  inclui inspeção, senha, seleção de páginas, validação e exportação com LibreOffice.
- `requirements.txt`: bibliotecas Python.
- `README.md`: instruções de instalação e execução.

O aplicativo não importa `pdf_to_docx.py` nem `docx_to_pdf.py`.
Não são necessários outros arquivos de código Python para executá-lo.

## 1. Instalar o Python e as bibliotecas

Use **Python 3.11 ou 3.12**. No Windows, instale-o pelo python.org e habilite
“Add Python to PATH”. No macOS, use o instalador do python.org ou
`brew install python@3.12`. No Ubuntu/Debian, use
`sudo apt install python3 python3-pip python3-venv`, verificando que a versão
disponível é 3.11 ou 3.12. Extraia o ZIP e abra um terminal
na pasta extraída (a mesma de `pdf-docx.py`).

Você pode executar diretamente os comandos abaixo se já tiver um ambiente
Python preparado.

### Windows — instalação e execução direta

```powershell
py -3.11 -m pip install -r requirements.txt
py -3.11 -m streamlit run pdf-docx.py
```

Se instalou 3.12, substitua `-3.11` por `-3.12`.

### macOS/Linux — instalação e execução direta

```sh
python3 -m pip install -r requirements.txt
python3 -m streamlit run pdf-docx.py
```

### Opção recomendada no seu computador: ambiente isolado

Um ambiente isolado evita misturar estas dependências com as de outros
programas. Execute os comandos na pasta extraída.

**Windows:**

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run pdf-docx.py
```

**macOS/Linux:**

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m streamlit run pdf-docx.py
```

Estes exemplos não exigem ativar o ambiente: utilizam diretamente seu
interpretador Python. Se o sistema bloquear instalação global (PEP 668),
use esta opção ou o gerenciador de ambientes da sua máquina.

O ZIP não contém ambientes nem bibliotecas pré-instaladas.
Abra o endereço local mostrado pelo Streamlit (normalmente
`http://localhost:8501`). Para encerrar, use Ctrl+C no terminal.

**Importante:** instalar as bibliotecas Python não substitui a instalação do
LibreOffice explicada a seguir. Se você iniciou o app antes de instalá-lo,
encerre o programa e inicie novamente depois da instalação.

## 2. Instalar e configurar o LibreOffice

### Por que o LibreOffice é necessário?

Para exportar DOCX → PDF, o programa precisa de um mecanismo que abra o
documento Word e renderize suas páginas. Neste aplicativo, esse trabalho é
feito pelo **LibreOffice Writer**, executado com a opção `--headless`, ou seja,
sem abrir a janela do editor.

O Python valida o DOCX, inicia o LibreOffice, aguarda a exportação e devolve
o PDF para download. Cada conversão utiliza seu próprio diretório temporário
e perfil, sem reutilizar uma sessão do editor.

O `pip` **não instala LibreOffice**. Ele só é necessário para DOCX → PDF;
a aba PDF → DOCX funciona sem ele.

### Instalação por sistema operacional

- **Windows:** baixe e instale LibreOffice em https://www.libreoffice.org/download/download-libreoffice/.
  Caminho usual: `C:\Program Files\LibreOffice\program\soffice.exe`.
- **macOS:** instale pelo mesmo site ou `brew install --cask libreoffice`.
  Caminho usual: `/Applications/LibreOffice.app/Contents/MacOS/soffice`.
- **Linux (Ubuntu/Debian):** `sudo apt update && sudo apt install libreoffice-writer`.
  Em outras distribuições, instale LibreOffice Writer pelo gerenciador local.
  Teste com `libreoffice --version` ou `soffice --version`.

Os caminhos acima são os mais comuns, mas podem variar conforme a instalação.
O pacote não inclui o instalador nem os arquivos do LibreOffice.

### Confirmar a instalação

No Windows, teste no PowerShell:

```powershell
& "C:\Program Files\LibreOffice\program\soffice.exe" --version
```

No macOS:

```sh
/Applications/LibreOffice.app/Contents/MacOS/soffice --version
```

No Linux:

```sh
libreoffice --version
```

Uma versão exibida sem erro confirma que o executável pode ser iniciado.
Isso, isoladamente, não garante que todo documento possa ser convertido.

### Detecção automática e configuração manual

O app procura `libreoffice` e `soffice` no PATH e depois nos caminhos usuais
do Windows/macOS. Para uma instalação diferente, preencha **Caminho do
executável (opcional)** na barra lateral com o caminho completo de
`soffice`, `soffice.exe` ou `libreoffice`, sem argumentos.
Reinicie o app após instalar LibreOffice ou mudar o PATH.
Não é necessário fechar o LibreOffice aberto: cada conversão usa um perfil isolado.

**Exemplo de preenchimento no Windows:**

```text
C:\Program Files\LibreOffice\program\soffice.exe
```

Informe somente o executável, sem adicionar `--headless` ou outros argumentos.
O programa acrescenta os argumentos necessários automaticamente.

### Se o aplicativo estiver hospedado em um servidor

O LibreOffice precisa estar instalado **no servidor onde o Python executa**.
Instalá-lo apenas no computador da pessoa que acessa o navegador não permite
que um servidor remoto faça DOCX → PDF.

O caminho configurado na barra lateral também deve ser um caminho válido
na máquina que executa o aplicativo. Se não houver LibreOffice nessa máquina,
PDF → DOCX continuará funcionando, mas DOCX → PDF apresentará uma mensagem
solicitando a instalação ou configuração.

## 3. Usar as duas abas

**PDF → DOCX**

1. Envie um PDF digital; informe a senha se ele estiver protegido.
2. Confira a quantidade de páginas.
3. Escolha documento inteiro, intervalo inclusivo, ou páginas específicas.
   Exemplo: `1-3, 5, 8-10`. As páginas são únicas e convertidas na ordem original.
4. Converta e clique em **Baixar documento DOCX**.

O arquivo baixado terá o nome `<nome_original>_convertido.docx`.
No modo **Intervalo**, a página inicial e a página final são incluídas.
No modo **Páginas específicas**, páginas repetidas são removidas e a ordem
original do documento é mantida.

**DOCX → PDF**

1. Envie um `.docx` sem senha. `.doc`, `.docm` e arquivos com macros não são suportados.
2. Clique em **Converter para PDF** e acompanhe o estado da exportação.
3. Baixe o PDF, com nome `<nome_original>_convertido.pdf`.

As abas funcionam independentemente. A falta de arquivo, senha ou LibreOffice,
ou uma falha em uma aba, não bloqueia a outra. Downloads permanecem na sessão;
trocar arquivo ou opções da respectiva conversão remove o resultado anterior.
Não há aba de download de código: este pacote é a entrega do código.

Baixe os resultados antes de encerrar a sessão. Fechar o servidor, reiniciar
o aplicativo ou abrir uma nova sessão pode fazer com que seja necessário
enviar e converter o documento novamente.

## Privacidade e limites

- Documentos **não são enviados a terceiros** ou a APIs de conversão.
  São processados no computador/servidor onde você executa o app. Se hospedá-lo,
  o upload vai a esse servidor, não ao computador do visitante.
- Uploads, senha e resultados ficam na memória da sessão do Streamlit,
  sem gravação permanente pelo aplicativo. A senha é removida ao trocar o PDF.
  Os arquivos temporários e o perfil LibreOffice são removidos tanto no sucesso
  como na falha. Não use com documentos não confiáveis em servidor público sem
  isolamento de sistema operacional; perfil isolado não é sandbox de segurança.
- Recursos externos vinculados (como imagens e modelos remotos) são rejeitados.
  Incorpore-os antes de enviar. Links clicáveis comuns são permitidos.
- Fontes ausentes, tabelas, imagens, quebras, campos e layouts complexos podem
  mudar entre Word, LibreOffice e PDF. Instale as fontes do documento, respeitando
  suas licenças. Não existe garantia de reprodução perfeita da formatação.
- PDFs escaneados não recebem OCR; páginas podem ficar como imagens ou sem
  texto editável. Arquivos danificados ou com proteção não suportada podem falhar.
- O upload usa o limite padrão do Streamlit (200 MB); DOCX também é limitado
  a 200 MB descompactados e 10.000 entradas internas.
- DOCX → PDF tem limite de **120 segundos**. Um timeout encerra o processo de
  conversão e limpa os temporários. Simplifique documentos muito grandes.
- Mensagens de executável ausente: instale LibreOffice ou configure seu caminho.
  Mensagens de permissões/espaço: verifique o diretório temporário do sistema.
- Use apenas localmente por padrão: não publique uma porta de rede para arquivos
  confidenciais sem autenticação e medidas de segurança apropriadas.

## Solução de problemas

| Problema | O que conferir ou fazer |
| --- | --- |
| `python`, `python3` ou `py` não encontrado | Instale o Python e confira o PATH. No Windows, tente `py -3.11`; no macOS/Linux, `python3`. |
| `No module named streamlit`, `pdf2docx`, `pypdf` ou `docx` | Instale `requirements.txt` com o mesmo interpretador que executará o programa. |
| O comando `streamlit` não é reconhecido | Use `python -m streamlit run pdf-docx.py`, ajustando o interpretador conforme seu sistema. |
| `File does not exist: pdf-docx.py` | Abra o terminal na pasta extraída do ZIP. Verifique que o arquivo não foi renomeado para `pdf-docx.py.txt`. |
| `No module named pdf_to_docx` ou `docx_to_pdf` | Você está executando uma versão antiga. Use o novo `pdf-docx.py`, que contém toda a lógica. |
| LibreOffice não encontrado | Instale LibreOffice com Writer, reinicie o app ou preencha o caminho do executável na barra lateral. |
| Caminho do executável inválido | Informe `soffice.exe`, `soffice` ou `libreoffice`, não apenas a pasta da instalação. |
| O DOCX é recusado | Use `.docx` sem senha e sem macros; incorpore recursos externos, como imagens vinculadas. `.doc` e `.docm` não são suportados. |
| O PDF pede senha | Informe a senha correta do documento. Proteções não suportadas podem impedir a conversão. |
| Seleção de páginas inválida | Use números e intervalos dentro do total de páginas, por exemplo `1-3, 5`. Não utilize página zero. |
| DOCX → PDF excede 120 segundos | Tente um documento menor ou simplifique seu conteúdo. |
| PDF → DOCX não produz texto editável | O PDF pode ser escaneado. Aplique OCR em outra ferramenta antes de convertê-lo. |
| Fontes ou layout mudam | Instale as fontes usadas no documento e revise o resultado. Word, PDF e LibreOffice podem renderizar documentos de formas diferentes. |
| Falha de permissões ou falta de espaço | Verifique o diretório temporário da máquina que executa o app. |

## Perguntas frequentes

**Preciso instalar o Microsoft Office?**  
Não. A conversão DOCX → PDF usa o LibreOffice, não o Microsoft Word.

**Posso usar PDF → DOCX sem instalar LibreOffice?**  
Sim, desde que as bibliotecas Python estejam instaladas.

**O aplicativo calcula ou preserva exatamente tudo do Word?**  
Não há garantia de reprodução perfeita. Campos, fontes, imagens, tabelas e
layouts complexos devem ser conferidos no arquivo convertido.

**Todos os arquivos de código precisam ficar juntos?**  
O único código Python do aplicativo é `pdf-docx.py`. O `requirements.txt`
serve para instalar dependências, e este README explica a configuração.
Não são necessários os módulos antigos `pdf_to_docx.py` e `docx_to_pdf.py`.

**O aplicativo já fica público quando eu o inicio?**  
Os comandos deste README servem para iniciar o programa. Para uso local, abra
o endereço indicado pelo Streamlit. Uma instalação em servidor exige configuração
própria de acesso e segurança; não exponha documentos confidenciais sem proteção.

## Dependências e compatibilidade

O `requirements.txt` especifica:

```text
streamlit==1.62.0
pdf2docx==0.5.13
pypdf[crypto]==6.16.2
python-docx==1.2.0
```

- **Streamlit:** interface no navegador, uploads, abas e downloads.
- **pdf2docx:** conversão de PDF para DOCX.
- **pypdf[crypto]:** leitura de PDFs, contagem de páginas e suporte a documentos
  protegidos, incluindo AES.
- **python-docx:** leitura e validação dos documentos DOCX.
- **PyMuPDF e lxml:** bibliotecas utilizadas pelos mecanismos de conversão;
  são instaladas junto com as dependências.
- **LibreOffice Writer:** programa externo para DOCX → PDF; não faz parte do
  `requirements.txt`.

Nenhuma chave de API é necessária. As instruções incluem caminhos usuais de
Windows, macOS e Linux; isso não significa que todas as combinações de sistema,
fontes e versões do LibreOffice tenham sido testadas.

