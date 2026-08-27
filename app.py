"""Aplicação Streamlit para converter arquivos PDF em documentos DOCX."""

from __future__ import annotations

from pathlib import PurePath

import streamlit as st

from pdf_to_docx import (
    InvalidPasswordError,
    PasswordRequiredError,
    PdfConversionError,
    convert_pdf_to_docx,
    inspect_pdf,
    parse_page_selection,
)


st.set_page_config(
    page_title="PDF para DOCX",
    page_icon="PDF",
    layout="centered",
    initial_sidebar_state="expanded",
)


def output_name(uploaded_name: str) -> str:
    """Gera um nome de download seguro e previsível."""

    stem = PurePath(uploaded_name).stem or "documento"
    return f"{stem}_convertido.docx"


def main() -> None:
    st.title("Conversor PDF para DOCX")
    st.write(
        "Transforme um PDF em um documento editável do Word. "
        "O arquivo é processado temporariamente e fica disponível para download."
    )

    with st.sidebar:
        st.header("Como usar")
        st.markdown(
            """
            1. Envie um arquivo PDF.
            2. Informe a senha, se necessário.
            3. Escolha as páginas.
            4. Clique em **Converter para DOCX**.
            """
        )
        st.divider()
        st.caption(
            "A conversão funciona melhor com PDFs digitais. "
            "Documentos escaneados podem precisar de OCR antes da conversão."
        )

    uploaded_file = st.file_uploader(
        "Selecione o PDF",
        type=["pdf"],
        help="Envie um arquivo com extensão .pdf.",
    )

    if uploaded_file is None:
        st.info("Envie um PDF para começar.")
        return

    pdf_bytes = uploaded_file.getvalue()
    password = st.text_input(
        "Senha do PDF (opcional)",
        type="password",
        help="A senha é usada somente durante esta conversão e não é armazenada.",
    )

    page_count: int | None = None
    is_encrypted = False
    try:
        page_count, is_encrypted = inspect_pdf(pdf_bytes, password=password or None)
    except PasswordRequiredError:
        st.warning(
            "Este PDF está protegido. Informe a senha acima para visualizar "
            "as opções de conversão."
        )
    except InvalidPasswordError as exc:
        st.error(str(exc))
    except PdfConversionError as exc:
        st.error(str(exc))

    if page_count is None:
        return

    col1, col2 = st.columns(2)
    with col1:
        st.metric("Páginas encontradas", page_count)
    with col2:
        st.metric("Tamanho do PDF", f"{len(pdf_bytes) / 1024 / 1024:.2f} MB")

    if is_encrypted:
        st.success("PDF protegido desbloqueado com sucesso.")
    else:
        st.success("PDF pronto para conversão.")

    st.subheader("Páginas para converter")
    selection_mode = st.radio(
        "Escolha uma opção",
        ["Documento inteiro", "Intervalo", "Páginas específicas"],
        horizontal=True,
        label_visibility="collapsed",
    )

    pages: list[int] | None = None
    if selection_mode == "Intervalo":
        start_page, end_page = st.columns(2)
        with start_page:
            first_page = st.number_input(
                "Página inicial",
                min_value=1,
                max_value=page_count,
                value=1,
                step=1,
            )
        with end_page:
            last_page = st.number_input(
                "Página final",
                min_value=1,
                max_value=page_count,
                value=page_count,
                step=1,
            )
        if first_page > last_page:
            st.warning("A página inicial deve ser menor ou igual à página final.")
        else:
            pages = list(range(first_page - 1, last_page))
    elif selection_mode == "Páginas específicas":
        selection = st.text_input(
            "Páginas",
            value=f"1-{min(page_count, 3)}",
            help="Use números separados por vírgula e intervalos, por exemplo: 1-3, 5, 8-10.",
        )
        try:
            pages = parse_page_selection(selection, page_count)
            st.caption(f"{len(pages)} página(s) selecionada(s).")
        except PdfConversionError as exc:
            st.warning(str(exc))

    st.divider()
    if st.button("Converter para DOCX", type="primary", use_container_width=True):
        if selection_mode == "Intervalo" and (first_page > last_page):
            st.error("Corrija o intervalo de páginas antes de converter.")
            return
        if selection_mode == "Páginas específicas" and pages is None:
            st.error("Corrija a seleção de páginas antes de converter.")
            return

        with st.spinner("Convertendo o PDF para DOCX..."):
            try:
                docx_bytes = convert_pdf_to_docx(
                    pdf_bytes,
                    password=password or None,
                    pages=pages,
                )
            except InvalidPasswordError as exc:
                st.error(str(exc))
                return
            except PasswordRequiredError as exc:
                st.error(str(exc))
                return
            except PdfConversionError as exc:
                st.error(str(exc))
                return

        st.success("Conversão concluída.")
        st.download_button(
            "Baixar documento DOCX",
            data=docx_bytes,
            file_name=output_name(uploaded_file.name),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
