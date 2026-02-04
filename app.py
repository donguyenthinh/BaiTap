import io
from typing import Optional

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Đọc Excel", page_icon="📄", layout="wide")
st.title("App đọc file Excel (Streamlit)")
st.caption("Tải lên file `.xlsx`/`.xls`, chọn sheet và xem dữ liệu.")


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes: bytes) -> list[str]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return list(map(str, xls.sheet_names))


@st.cache_data(show_spinner=False)
def read_sheet(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


uploaded = st.file_uploader("Chọn file Excel", type=["xlsx", "xls"])

if not uploaded:
    st.info("Hãy tải lên một file Excel để bắt đầu.")
    st.stop()

file_bytes = uploaded.getvalue()

try:
    sheet_names = list_sheets(file_bytes)
except Exception as e:
    st.error("Không thể đọc file Excel. Vui lòng kiểm tra file có hợp lệ không.")
    st.exception(e)
    st.stop()

left, right = st.columns([1, 2], gap="large")

with left:
    st.subheader("Thiết lập")
    sheet = st.selectbox("Sheet", sheet_names, index=0)
    nrows: Optional[int] = st.number_input("Giới hạn số dòng (0 = tất cả)", min_value=0, value=0, step=100)

with right:
    try:
        df = read_sheet(file_bytes, sheet)
    except Exception as e:
        st.error("Đọc sheet thất bại.")
        st.exception(e)
        st.stop()

    st.subheader("Dữ liệu")
    st.write(f"**Số dòng/cột:** {len(df):,} / {len(df.columns):,}")

    cols = st.multiselect("Chọn cột để hiển thị (bỏ trống = tất cả)", list(df.columns))
    view_df = df[cols] if cols else df

    if nrows and nrows > 0:
        view_df = view_df.head(int(nrows))

    st.dataframe(view_df, use_container_width=True, height=520)

    st.download_button(
        "Tải xuống CSV (từ dữ liệu đang hiển thị)",
        data=to_csv_bytes(view_df),
        file_name=f"{uploaded.name.rsplit('.', 1)[0]}_{sheet}.csv",
        mime="text/csv",
        use_container_width=True,
    )

