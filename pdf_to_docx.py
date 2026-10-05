"""Conversor PDF ↔ DOCX completo em um único arquivo Streamlit."""

from __future__ import annotations

from hashlib import sha256
from io import BytesIO
import os
from pathlib import Path, PurePath
import re
import shutil
import signal
import subprocess
import tempfile
from typing import Iterable
import zipfile

import streamlit as st
from docx import Document
from lxml import etree
from pdf2docx import Converter
from pypdf import PdfReader

class PdfConversionError(Exception):
    """Falha de leitura, seleção ou conversão do PDF."""


class PasswordRequiredError(PdfConversionError):
    """O PDF precisa de senha."""


class InvalidPasswordError(PdfConversionError):
    """A senha não desbloqueou o PDF."""


def inspect_pdf(pdf_bytes: bytes, password: str | None = None) -> tuple[int, bool]:
    """Retorna quantidade de páginas e se o documento estava criptografado."""
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        encrypted = reader.is_encrypted
        if encrypted:
            if not password:
                raise PasswordRequiredError(
                    "Este PDF está protegido. Informe a senha para continuar."
                )
            if not reader.decrypt(password):
                raise InvalidPasswordError("A senha informada não é válida.")
        count = len(reader.pages)
        if not count:
            raise PdfConversionError("O PDF não contém páginas para converter.")
        return count, encrypted
    except PdfConversionError:
        raise
    except Exception as exc:
        raise PdfConversionError(
            "Não foi possível ler o PDF. Ele pode estar danificado ou usar "
            "uma proteção não suportada."
        ) from exc


def parse_page_selection(selection: str, page_count: int) -> list[int]:
    """Converte páginas humanas (1-3, 5) em índices únicos, em ordem crescente."""
    if page_count < 1:
        raise PdfConversionError("O documento não contém páginas.")
    if not selection.strip():
        raise PdfConversionError("Informe pelo menos uma página para converter.")
    pages: set[int] = set()
    for chunk in selection.split(","):
        chunk = chunk.strip()
        match = re.fullmatch(r"([0-9]+)(?:\s*-\s*([0-9]+))?", chunk)
        if not match:
            raise PdfConversionError(
                f"Formato inválido: '{chunk}'. Use números e intervalos, como 1-3, 5."
            )
        try:
            start = int(match.group(1))
            end = int(match.group(2) or match.group(1))
        except ValueError as exc:
            raise PdfConversionError("Número de página inválido.") from exc
        if start > end:
            raise PdfConversionError("A página inicial deve ser menor ou igual à final.")
        if start < 1 or end > page_count:
            raise PdfConversionError(f"As páginas devem estar entre 1 e {page_count}.")
        pages.update(range(start - 1, end))
    return sorted(pages)


def convert_pdf_to_docx(
    pdf_bytes: bytes, *, password: str | None = None,
    pages: Iterable[int] | None = None,
) -> bytes:
    """Converte com pdf2docx e remove os temporários mesmo em caso de falha."""
    count, _ = inspect_pdf(pdf_bytes, password)
    selected = None if pages is None else list(pages)
    if selected is not None:
        if not selected or any(
            type(page) is not int or page < 0 or page >= count for page in selected
        ):
            raise PdfConversionError(f"Selecione páginas válidas entre 1 e {count}.")
        selected = sorted(set(selected))
    try:
        with tempfile.TemporaryDirectory(prefix="pdf-to-docx-") as directory:
            source = Path(directory) / "entrada.pdf"
            target = Path(directory) / "saida.docx"
            source.write_bytes(pdf_bytes)
            converter = None
            try:
                converter = Converter(str(source), password=password or None)
                converter.convert(str(target), pages=selected, ignore_page_error=False)
            finally:
                if converter is not None:
                    converter.close()
            result = target.read_bytes()
            Document(BytesIO(result))
            return result
    except Exception as exc:
        raise PdfConversionError(
            "A conversão não foi concluída. O PDF pode estar danificado ou conter "
            "elementos não suportados. Verifique também o espaço temporário disponível."
        ) from exc


class DocxConversionError(Exception):
    """Erro apresentado ao usuário sem interromper a outra aba."""


class LibreOfficeNotFoundError(DocxConversionError):
    """O executável local não foi encontrado."""


def validate_docx(data: bytes) -> None:
    """Rejeita arquivos corrompidos, outros formatos e recursos externos ativos."""
    try:
        with zipfile.ZipFile(BytesIO(data)) as archive:
            entries = archive.infolist()
            # Limite para evitar descompressão excessiva de uploads pequenos.
            if len(entries) > 10000 or sum(e.file_size for e in entries) > 200 * 1024**2:
                raise DocxConversionError(
                    "O DOCX excede o limite de 200 MB descompactados ou 10.000 entradas."
                )
            names = archive.namelist()
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise ValueError("Pacote não é DOCX")
            if any("vbaproject" in name.lower() for name in names):
                raise DocxConversionError("Documentos com macros não são suportados.")
            parser = etree.XMLParser(resolve_entities=False, no_network=True)
            for name in names:
                if name.endswith(".rels"):
                    root = etree.fromstring(archive.read(name), parser)
                    for relationship in root:
                        if (
                            relationship.get("TargetMode") == "External"
                            and not relationship.get("Type", "").endswith("/hyperlink")
                        ):
                            raise DocxConversionError(
                                "Este DOCX contém recursos externos vinculados. "
                                "Incorpore as imagens ou recursos no Word e envie novamente."
                            )
            if archive.testzip() is not None:
                raise ValueError("ZIP corrompido")
        Document(BytesIO(data))
    except DocxConversionError:
        raise
    except Exception as exc:
        raise DocxConversionError(
            "Não foi possível ler um DOCX válido. Envie um arquivo .docx "
            "sem senha, salvo pelo Word ou LibreOffice; .doc não é suportado."
        ) from exc


def find_libreoffice(executable: str | None = None) -> str:
    """Usa caminho configurado, PATH ou instalações comuns de Windows/macOS."""
    if executable and executable.strip():
        configured = executable.strip().strip('"')
        found = shutil.which(configured)
        if found:
            return found
        raise LibreOfficeNotFoundError(
            "O caminho informado não aponta para um executável do LibreOffice. "
            "Informe o caminho completo de soffice ou soffice.exe."
        )
    for command in ("libreoffice", "soffice"):
        found = shutil.which(command)
        if found:
            return found
    candidates = [
        Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
        Path("C:/Program Files/LibreOffice/program/soffice.exe"),
        Path("C:/Program Files (x86)/LibreOffice/program/soffice.exe"),
    ]
    for candidate in candidates:
        if candidate.is_file() and os.access(candidate, os.X_OK):
            return str(candidate)
    raise LibreOfficeNotFoundError(
        "LibreOffice não encontrado. Instale-o e reinicie o app, ou informe "
        "o caminho do executável na barra lateral. PDF → DOCX continua disponível."
    )


def _kill_process_tree(process: subprocess.Popen) -> None:
    """Encerra também o processo filho do launcher após um timeout."""
    if os.name == "posix":
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    else:
        # taskkill é uma ferramenta nativa do Windows; não utiliza shell.
        try:
            subprocess.run(
                ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10,
            )
        except (OSError, subprocess.TimeoutExpired):
            pass
        if process.poll() is None:
            process.kill()
    process.communicate()


def convert_docx_to_pdf(
    data: bytes, *, executable: str | None = None, timeout: float = 120,
) -> bytes:
    """Exporta o documento inteiro com perfil e diretório únicos por chamada."""
    validate_docx(data)
    program = find_libreoffice(executable)
    if timeout <= 0:
        raise DocxConversionError("O tempo limite deve ser maior que zero.")
    try:
        with tempfile.TemporaryDirectory(prefix="docx-to-pdf-") as directory:
            work = Path(directory)
            source = work / "documento.docx"
            source.write_bytes(data)
            output = work / "saida"
            output.mkdir()
            profile = work / "perfil"
            profile.mkdir()
            # Perfil sem macros, atualização de vínculos ou restauração de sessão.
            (profile / "user").mkdir()
            (profile / "user" / "registrymodifications.xcu").write_text(
                '<?xml version="1.0" encoding="UTF-8"?>'
                '<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
                '<item oor:path="/org.openoffice.Office.Common/Security/Scripting">'
                '<prop oor:name="MacroSecurityLevel" oor:op="fuse">'
                '<value>3</value></prop></item>'
                '<item oor:path="/org.openoffice.Office.Writer/Content/Update">'
                '<prop oor:name="Link" oor:op="fuse"><value>2</value></prop>'
                '</item></oor:items>', encoding="utf-8",
            )
            command = [
                program, f"-env:UserInstallation={profile.as_uri()}",
                "--headless", "--nologo", "--nodefault", "--norestore",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", str(output), str(source),
            ]
            process = subprocess.Popen(
                command, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                start_new_session=os.name == "posix", shell=False,
            )
            try:
                process.communicate(timeout=timeout)
            except subprocess.TimeoutExpired as exc:
                _kill_process_tree(process)
                raise DocxConversionError(
                    f"A conversão ultrapassou o limite de {timeout:g} segundos. "
                    "Tente um documento menor ou simplifique seu conteúdo."
                ) from exc
            if process.returncode != 0:
                raise DocxConversionError(
                    "O LibreOffice não conseguiu converter o documento. "
                    "Abra e salve o DOCX no Word ou LibreOffice e tente novamente."
                )
            target = output / "documento.pdf"
            if not target.is_file() or target.stat().st_size == 0:
                raise DocxConversionError("O LibreOffice não gerou o PDF esperado.")
            result = target.read_bytes()
            if not result.startswith(b"%PDF-") or not len(PdfReader(BytesIO(result)).pages):
                raise DocxConversionError("O PDF gerado não é válido ou não contém páginas.")
            return result
    except DocxConversionError:
        raise
    except OSError as exc:
        raise DocxConversionError(
            "Não foi possível executar o LibreOffice ou criar os arquivos temporários. "
            "Verifique o caminho, as permissões e o espaço disponível."
        ) from exc
    except Exception as exc:
        raise DocxConversionError("Falha ao validar o PDF gerado pelo LibreOffice.") from exc


def output_name(uploaded_name: str, extension: str = "docx") -> str:
    stem = PurePath(uploaded_name.replace("\\", "/")).stem or "documento"
    return f"{stem}_convertido.{extension}"


def clear_result(direction: str) -> None:
    st.session_state.pop(f"{direction}_result", None)


def reset_pdf() -> None:
    clear_result("pdf")
    for key in ("pdf_password", "pdf_mode", "pdf_first", "pdf_last", "pdf_selection"):
        st.session_state.pop(key, None)


def render_pdf_tab() -> None:
    uploaded = st.file_uploader(
        "Selecione o PDF", type=["pdf"], key="pdf_upload", on_change=reset_pdf,
    )
    if uploaded is None:
        clear_result("pdf")
        st.info("Envie um PDF para começar.")
        return  # Retorna apenas desta aba, nunca de main().
    data = uploaded.getvalue()
    password = st.text_input(
        "Senha do PDF (opcional)", type="password", key="pdf_password",
        on_change=clear_result, args=("pdf",),
        help="Usada nesta sessão, sem gravação em disco; removida ao trocar o PDF.",
    )
    try:
        count, encrypted = inspect_pdf(data, password or None)
    except PasswordRequiredError as exc:
        clear_result("pdf")
        st.warning(str(exc))
        return
    except (InvalidPasswordError, PdfConversionError) as exc:
        clear_result("pdf")
        st.error(str(exc))
        return

    left, right = st.columns(2)
    left.metric("Páginas encontradas", count)
    right.metric("Tamanho do PDF", f"{len(data) / 1024**2:.2f} MB")
    st.success("PDF protegido desbloqueado." if encrypted else "PDF pronto para conversão.")
    mode = st.radio(
        "Páginas para converter",
        ["Documento inteiro", "Intervalo", "Páginas específicas"],
        horizontal=True, key="pdf_mode", on_change=clear_result, args=("pdf",),
    )
    pages = None
    valid = True
    if mode == "Intervalo":
        left, right = st.columns(2)
        first = left.number_input(
            "Página inicial", min_value=1, max_value=count, value=1, step=1,
            key="pdf_first", on_change=clear_result, args=("pdf",),
        )
        last = right.number_input(
            "Página final", min_value=1, max_value=count, value=count, step=1,
            key="pdf_last", on_change=clear_result, args=("pdf",),
        )
        valid = first <= last
        if valid:
            pages = list(range(first - 1, last))
        else:
            st.warning("A página inicial deve ser menor ou igual à página final.")
    elif mode == "Páginas específicas":
        selection = st.text_input(
            "Páginas", value=f"1-{min(count, 3)}", key="pdf_selection",
            help="Exemplo: 1-3, 5, 8-10. As páginas são convertidas em ordem crescente.",
            on_change=clear_result, args=("pdf",),
        )
        try:
            pages = parse_page_selection(selection, count)
            st.caption(f"{len(pages)} página(s) selecionada(s).")
        except PdfConversionError as exc:
            valid = False
            st.warning(str(exc))
    fingerprint = sha256(
        data + repr((uploaded.name, password, mode, pages, valid)).encode()
    ).hexdigest()
    if not valid:
        clear_result("pdf")
    if st.button(
        "Converter para DOCX", key="pdf_convert", type="primary",
        use_container_width=True, disabled=not valid,
    ):
        clear_result("pdf")
        with st.spinner("Convertendo o PDF para DOCX..."):
            try:
                result = convert_pdf_to_docx(data, password=password or None, pages=pages)
                st.session_state.pdf_result = (fingerprint, result)
            except PdfConversionError as exc:
                st.error(str(exc))
    stored = st.session_state.get("pdf_result")
    if stored and stored[0] == fingerprint:
        st.success("Conversão concluída.")
        st.download_button(
            "Baixar documento DOCX", data=stored[1],
            file_name=output_name(uploaded.name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="pdf_download", type="primary", use_container_width=True,
        )


def render_docx_tab(executable: str) -> None:
    uploaded = st.file_uploader(
        "Selecione o DOCX", type=["docx"], key="docx_upload",
        on_change=clear_result, args=("docx",),
    )
    if uploaded is None:
        clear_result("docx")
        st.info("Envie um documento .docx para começar.")
        return
    data = uploaded.getvalue()
    try:
        validate_docx(data)
    except DocxConversionError as exc:
        clear_result("docx")
        st.error(str(exc))
        return
    st.caption(
        f"Tamanho do DOCX: {len(data) / 1024**2:.2f} MB. "
        "O documento inteiro será convertido; o limite é de 120 segundos."
    )
    fingerprint = sha256(data + repr((uploaded.name, executable)).encode()).hexdigest()
    if st.button(
        "Converter para PDF", key="docx_convert", type="primary", use_container_width=True,
    ):
        clear_result("docx")
        with st.status("Convertendo DOCX para PDF...", expanded=True) as status:
            st.write("Documento validado. Exportando com o LibreOffice local...")
            try:
                result = convert_docx_to_pdf(data, executable=executable or None)
                st.session_state.docx_result = (fingerprint, result)
                status.update(label="Conversão concluída.", state="complete", expanded=False)
            except DocxConversionError as exc:
                status.update(label="A conversão não foi concluída.", state="error")
                st.error(str(exc))
    stored = st.session_state.get("docx_result")
    if stored and stored[0] == fingerprint:
        st.success("PDF pronto para download.")
        st.download_button(
            "Baixar documento PDF", data=stored[1],
            file_name=output_name(uploaded.name, "pdf"), mime="application/pdf",
            key="docx_download", type="primary", use_container_width=True,
        )


def main() -> None:
    st.set_page_config(page_title="Conversor PDF e DOCX", layout="centered")
    st.title("Conversor PDF e DOCX")
    st.write("Converta PDF em Word editável ou exporte um DOCX para PDF.")
    st.caption(
        "Processamento local, sem serviços externos de conversão. "
        "A formatação pode variar conforme o layout e as fontes instaladas."
    )
    with st.sidebar:
        st.header("Como usar")
        st.markdown(
            "1. Escolha uma aba e envie o documento.\n"
            "2. No PDF, informe a senha e selecione as páginas se necessário.\n"
            "3. Clique em converter e baixe o arquivo gerado."
        )
        st.divider()
        st.caption("PDFs escaneados precisam de OCR; este app não inclui OCR.")
        st.subheader("LibreOffice")
        executable = st.text_input(
            "Caminho do executável (opcional)", key="lo_executable",
            help="Deixe vazio para detecção automática. Informe apenas o caminho, sem argumentos.",
            on_change=clear_result, args=("docx",),
        )
        st.caption(
            "Necessário somente para DOCX → PDF. Não há envio a terceiros. "
            "Os temporários são removidos ao finalizar cada conversão; "
            "uploads e resultados permanecem na memória da sessão."
        )
    pdf_tab, docx_tab = st.tabs(["PDF → DOCX", "DOCX → PDF"])
    with pdf_tab:
        render_pdf_tab()
    with docx_tab:
        render_docx_tab(executable)


if __name__ == "__main__":
    main()
