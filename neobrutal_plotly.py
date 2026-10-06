"""Template Plotly bergaya neo-brutalism untuk Dashboard CPL.

Dipanggil sekali oleh ``load_custom_css()``. Jika Plotly menolak salah satu properti
(misalnya versi lama), grafik tetap tampil dengan tema bawaan dan app tidak error.
"""

from __future__ import annotations

INK = "#14120f"
CARD = "#fffdf6"
GRID = "rgba(20, 18, 15, 0.12)"
FONT = "Plus Jakarta Sans, system-ui, -apple-system, Segoe UI, sans-serif"
COLORWAY = ["#3a4fd7", "#f6c453", "#3fbf84", "#f4806b", "#cdbef7", "#5c564b"]


def _axis() -> dict:
    return dict(
        showline=True,
        linecolor=INK,
        linewidth=2,
        mirror=True,
        ticks="outside",
        tickcolor=INK,
        gridcolor=GRID,
        zeroline=False,
    )


def apply_plotly_theme() -> None:
    try:
        import plotly.graph_objects as go
        import plotly.io as pio

        template = go.layout.Template(
            layout=dict(
                font=dict(family=FONT, color=INK, size=13),
                paper_bgcolor="rgba(0,0,0,0)",
                plot_bgcolor=CARD,
                colorway=COLORWAY,
                title=dict(font=dict(family=FONT, color=INK, size=18)),
                xaxis=_axis(),
                yaxis=_axis(),
                legend=dict(bgcolor="#ffffff", bordercolor=INK, borderwidth=2),
                hoverlabel=dict(
                    bgcolor=CARD,
                    bordercolor=INK,
                    font=dict(family=FONT, color=INK),
                ),
                polar=dict(
                    bgcolor=CARD,
                    radialaxis=dict(gridcolor=GRID, linecolor=INK),
                    angularaxis=dict(gridcolor=GRID, linecolor=INK),
                ),
            ),
            data=dict(bar=[go.Bar(marker=dict(line=dict(color=INK, width=2)))]),
        )
        pio.templates["neobrutal"] = template
        pio.templates.default = "neobrutal"
    except Exception:
        return
