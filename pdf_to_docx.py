"""Funções de inspeção e conversão de PDF para DOCX."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
import tempfile
from typing import Iterable

from pdf2docx import Converter
from pypdf import PdfReader


class PdfConversionError(Exception):
    """Erro esperado durante a inspeção ou conversão do PDF."""


class PasswordRequiredError(PdfConversionError):
    """O documento está protegido e precisa de senha."""


class InvalidPasswordError(PdfConversionError):
    """A senha informada não desbloqueou o documento."""


def inspect_pdf(pdf_bytes: bytes, password: str | None = None) -> tuple[int, bool]:
    """Retorna quantidade de páginas e se o arquivo estava criptografado."""

    try:
        reader = PdfReader(BytesIO(pdf_bytes))
    except Exception as exc:
        raise PdfConversionError(
            "Não foi possível ler este arquivo como um PDF válido."
        ) from exc

    was_encrypted = reader.is_encrypted
    if was_encrypted:
        if not password:
            raise PasswordRequiredError(
                "Este PDF está protegido por senha. Informe a senha para continuar."
            )
        try:
            decrypted = reader.decrypt(password)
        except Exception as exc:
            raise InvalidPasswordError("A senha informada não é válida.") from exc
        if not decrypted:
            raise InvalidPasswordError("A senha informada não é válida.")

    try:
        page_count = len(reader.pages)
    except Exception as exc:
        raise PdfConversionError(
            "O PDF foi encontrado, mas não foi possível acessar suas páginas."
        ) from exc

    if page_count == 0:
        raise PdfConversionError("O PDF não contém páginas para converter.")

    return page_count, was_encrypted


def parse_page_selection(selection: str, page_count: int) -> list[int]:
    """Converte '1-3, 5, 8-9' em índices de páginas começando em zero."""

    if not selection.strip():
        raise PdfConversionError("Informe pelo menos uma página para converter.")

    pages: set[int] = set()
    chunks = [chunk.strip() for chunk in selection.split(",") if chunk.strip()]

    for chunk in chunks:
        if re.fullmatch(r"\d+", chunk):
            start = end = int(chunk)
        elif match := re.fullmatch(r"(\d+)\s*-\s*(\d+)", chunk):
            start, end = int(match.group(1)), int(match.group(2))
            if start > end:
                raise PdfConversionError(
                    f"O intervalo '{chunk}' precisa começar antes de terminar."
                )
        else:
            raise PdfConversionError(
                f"Formato inválido: '{chunk}'. Use números e intervalos, como 1-3, 5."
            )

        if start < 1 or end > page_count:
            raise PdfConversionError(
                f"As páginas devem estar entre 1 e {page_count}."
            )

        pages.update(range(start - 1, end))

    return sorted(pages)


def convert_pdf_to_docx(
    pdf_bytes: bytes,
    *,
    password: str | None = None,
    pages: Iterable[int] | None = None,
) -> bytes:
    """Converte o PDF recebido em memória e retorna o DOCX também em memória."""

    inspect_pdf(pdf_bytes, password=password)

    with tempfile.TemporaryDirectory(prefix="pdf-to-docx-") as temp_dir:
        temp_path = Path(temp_dir)
        pdf_path = temp_path / "entrada.pdf"
        docx_path = temp_path / "saida.docx"
        pdf_path.write_bytes(pdf_bytes)

        converter = Converter(str(pdf_path), password=password or None)
        try:
            if pages is None:
                converter.convert(str(docx_path))
            else:
                selected_pages = list(pages)
                if not selected_pages:
                    raise PdfConversionError(
                        "Selecione pelo menos uma página para converter."
                    )
                converter.convert(str(docx_path), pages=selected_pages)
        except PdfConversionError:
            raise
        except Exception as exc:
            raise PdfConversionError(
                "A conversão não foi concluída. Este PDF pode conter elementos "
                "complexos ou estar danificado."
            ) from exc
        finally:
            converter.close()

        if not docx_path.exists() or docx_path.stat().st_size == 0:
            raise PdfConversionError(
                "A biblioteca não gerou um DOCX válido para este arquivo."
            )

        return docx_path.read_bytes()