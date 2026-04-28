from io import BytesIO

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
from PIL import Image, ImageDraw, ImageFont


def render_html_frame(html_content, height="content", width="stretch"):
    """Render inline HTML with Streamlit's supported components API."""
    if height == "content":
        height = 240
    if isinstance(height, int) and height < 1:
        height = 1
    component_width = None if width in (None, "stretch") else width
    components.html(str(html_content), width=component_width, height=height, scrolling=True)


def table_to_png_bytes(table_data, title=None):
    """Render table rows as a PNG image and return the bytes."""
    try:
        font = ImageFont.load_default()
    except Exception:
        font = None

    if not table_data:
        table_data = [["No data"]]

    normalized_table = [[str(cell) for cell in row] for row in table_data]
    max_columns = max(len(row) for row in normalized_table)
    normalized_table = [row + [""] * (max_columns - len(row)) for row in normalized_table]

    padding = 10
    row_height = 22
    col_padding = 18
    col_widths = []
    for col_idx in range(max_columns):
        col_width = max(len(row[col_idx]) for row in normalized_table) * 7 + col_padding
        col_widths.append(max(80, min(col_width, 360)))

    width = sum(col_widths) + padding * 2
    height = row_height * len(normalized_table) + padding * 2
    if title:
        height += row_height

    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    y = padding
    if title:
        draw.text((padding, y), title, fill="black", font=font)
        y += row_height

    for row in normalized_table:
        x = padding
        for col_idx, cell in enumerate(row):
            draw.text((x, y), str(cell)[:48], fill="black", font=font)
            x += col_widths[col_idx]
        y += row_height

    output = BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def image_bytes_to_png_bytes(image_bytes):
    """Convert an uploaded image to PNG bytes."""
    with Image.open(BytesIO(image_bytes)) as image:
        png_buffer = BytesIO()
        image.save(png_buffer, format="PNG")
        return png_buffer.getvalue()


def dataframe_to_table_rows(df):
    safe_df = df.fillna("")
    rows = [list(map(str, safe_df.columns.tolist()))]
    rows.extend([list(map(str, row)) for row in safe_df.values.tolist()])
    return rows


def plot_pie_chart(counts, title):
    labels, values = list(counts.keys()), list(counts.values())
    fig = go.Figure(go.Pie(labels=labels[:50], values=values[:50], textinfo="label+value", textposition="outside"))
    fig.update_layout(title=title, margin=dict(t=80, b=80, l=80, r=80), height=700)
    return fig


def plot_bar_chart(counts, title, horizontal=False):
    labels, values = list(counts.keys()), list(counts.values())
    fig = px.bar(x=values, y=labels, orientation="h", text=values) if horizontal else px.bar(x=labels, y=values, text=values)
    fig.update_traces(texttemplate="%{text}", textposition="outside", marker_color="skyblue")
    fig.update_layout(title=title, margin=dict(t=80, b=150 if not horizontal else 80), height=700)
    return fig


def show_current_sidebar_selection(selected_files):
    st.caption("Selected in sidebar")
    if selected_files:
        st.write(", ".join(selected_files))
    else:
        st.info("No files selected yet.")


def render_file_context_card(title, available_files, active_files=None):
    active_files = active_files or []
    st.markdown(f"**{title}**")
    st.caption(f"{len(active_files)} active of {len(available_files)} available files")
