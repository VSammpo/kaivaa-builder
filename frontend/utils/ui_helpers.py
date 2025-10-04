# frontend/utils/ui_helpers.py
from __future__ import annotations
import streamlit as st
from typing import Iterable, Tuple

def page_header(title: str, *, icon: str | None = None, crumbs: Iterable[Tuple[str,str]] = ()):
    """
    Affiche un titre + un fil d'Ariane cohérent.
    - title : texte du titre (sans emoji pour éviter la double icône)
    - icon  : (option) emoji/icone à mettre avant le titre (si tu le veux vraiment)
    - crumbs: liste de tuples (label, page_path) ; si page_path == "" -> segment courant
    """
    st.markdown(
        ("#" if icon is None else f"# {icon}") + f" {title}" if not title.startswith("#") else title
    )
    if crumbs:
        parts = []
        for lbl, path in crumbs:
            if path:
                parts.append(f"[{lbl}]({path})")
            else:
                parts.append(f"**{lbl}**")
        st.caption(" · ".join(parts))
    st.divider()
