# Conversor PDF para DOCX

Aplicação web feita com Streamlit para converter arquivos PDF em documentos
DOCX editáveis usando `pdf2docx`.

## Executar localmente

```bash
streamlit --server.port 5000
```

A aplicação permite:

- enviar arquivos PDF pelo navegador;
- detectar PDFs protegidos por senha;
- converter o documento inteiro, um intervalo ou páginas específicas;
- baixar o DOCX gerado com um nome baseado no PDF original.

## Arquivos principais

- `app.py`: interface web Streamlit.
- `pdf_to_docx.py`: inspeção, validação, seleção de páginas e conversão.
- `pyproject.toml`: dependências Python do projeto.

Os arquivos enviados são usados em arquivos temporários durante a conversão e
removidos ao final do processamento.

## Architecture decisions

- O processamento usa arquivos temporários para que o conversor `pdf2docx` trabalhe com caminhos de arquivo, sem manter documentos persistentes no projeto.
- A validação com `pypdf` acontece antes da conversão para detectar senha inválida e evitar mensagens genéricas.
- A seleção de páginas é convertida para índices zero-based apenas no módulo de conversão; a interface mostra páginas começando em 1.

## O produto

O usuário envia um PDF, desbloqueia o arquivo quando necessário, escolhe páginas
e baixa o DOCX convertido.


## Notas

- PDFs escaneados podem exigir OCR e PDFs com layouts muito complexos podem não converter perfeitamente.

## Apontadores

- See the `pnpm-workspace` skill for workspace structure, TypeScript setup, and package details
