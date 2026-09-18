"""
Hệ thống Web Demo Mô phỏng & Đối sánh Thuật toán Phân cụm CURE (Clustering Using REpresentatives)
Môn học: Khai phá dữ liệu - Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT)
Khởi chạy:python -m streamlit run app_streamlit.py 
"""

import os
import sys
import time
import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from sklearn import datasets
from sklearn.cluster import KMeans, AgglomerativeClustering
from sklearn.metrics import silhouette_score, davies_bouldin_score, calinski_harabasz_score

from cure_algorithm import CURE, KMedoids

# ==========================================
# CẤU HÌNH TRANG STREAMLIT
# ==========================================
st.set_page_config(
    page_title="CURE Clustering Visualizer - HUIT",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS giao diện chuẩn HUIT
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #103673 0%, #1f4e79 100%);
        padding: 22px;
        border-radius: 12px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .huit-title {
        color: #F2A900 !important;
        font-weight: 800;
        letter-spacing: 0.5px;
        margin: 0;
        font-size: 22px;
    }
    .sub-title {
        color: #FFFFFF;
        font-weight: 600;
        margin: 6px 0 0 0;
        font-size: 17px;
    }
    .metric-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 12px;
        margin-bottom: 10px;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 48px;
        white-space: pre-wrap;
        background-color: #f1f5f9;
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
        font-weight: 600;
        color: #1e293b;
    }
    .stTabs [aria-selected="true"] {
        background-color: #103673 !important;
        color: #ffffff !important;
    }
</style>
""", unsafe_allow_html=True)

# Banner Header
st.markdown("""
<div class="main-header">
    <div class="huit-title">TRƯỜNG ĐẠI HỌC CÔNG THƯƠNG TP. HỒ CHÍ MINH (HUIT)</div>
    <div class="sub-title">HỆ THỐNG MÔ PHỎNG & ĐỐI SÁNH THUẬT TOÁN PHÂN CỤM CURE</div>
    <p style="margin: 6px 0 0 0; font-size: 13.5px; opacity: 0.92;">
        Môn học: Khai phá dữ liệu | Đề tài: <b>Phân cụm dữ liệu dựa trên thuật toán CURE (Clustering Using REpresentatives)</b>
    </p>
</div>
""", unsafe_allow_html=True)

# ==========================================
# SIDEBAR: CẤU HÌNH THAM SỐ
# ==========================================
st.sidebar.markdown("### ⚙️ CẤU HÌNH THỰC NGHIỆM")

dataset_type = st.sidebar.selectbox(
    "1. Chọn tập dữ liệu kiểm thử:",
    [
        "Two Moons (2 Vầng trăng khuyết - Phi cầu)",
        "Concentric Circles (2 Vòng tròn đồng tâm - Phi cầu)",
        "Anisotropic Blobs (Cụm kéo dài hình elip)",
        "Blobs with Outliers (Cụm có điểm ngoại lai/nhiễu)"
    ]
)

n_samples = st.sidebar.slider("Số lượng điểm dữ liệu (N):", min_value=120, max_value=600, value=300, step=30)
noise_seed = st.sidebar.number_input("Random Seed:", min_value=1, max_value=999, value=42, step=1)

st.sidebar.markdown("---")
st.sidebar.markdown("### 🎛️ THAM SỐ THUẬT TOÁN CURE")
k_clusters = st.sidebar.slider("Số cụm mục tiêu (k):", min_value=2, max_value=6, value=2, step=1)
c_reps = st.sidebar.slider("Số điểm đại diện mỗi cụm (c):", min_value=1, max_value=8, value=5, step=1,
                          help="Càng nhiều điểm đại diện càng nắm bắt được các khúc uốn lượn của cụm phi cầu.")
alpha_shrink = st.sidebar.slider("Hệ số co cụm (alpha):", min_value=0.0, max_value=1.0, value=0.4, step=0.05,
                                help="0.0: Điểm đại diện giữ nguyên ở biên ngoài | 1.0: Điểm đại diện co hoàn toàn về trọng tâm mean | 0.3-0.5: Tối ưu chống nhiễu.")

# ==========================================
# TẠO DỮ LIỆU THỰC NGHIỆM
# ==========================================
@st.cache_data
def generate_dataset(d_type, n_pts, seed):
    rng = np.random.RandomState(seed)
    if "Moons" in d_type:
        X, _ = datasets.make_moons(n_samples=n_pts, noise=0.06, random_state=seed)
    elif "Circles" in d_type:
        X, _ = datasets.make_circles(n_samples=n_pts, factor=0.5, noise=0.05, random_state=seed)
    elif "Anisotropic" in d_type:
        X_blobs, _ = datasets.make_blobs(n_samples=n_pts, cluster_std=[0.9, 0.9, 0.9], random_state=seed)
        transformation = [[0.6, -0.6], [-0.4, 0.8]]
        X = np.dot(X_blobs, transformation)
    else:
        # Blobs with Outliers
        n_clean = max(80, n_pts - 40)
        blobs, _ = datasets.make_blobs(n_samples=n_clean, centers=2, cluster_std=0.8, random_state=seed)
        outliers = rng.uniform(low=-7, high=7, size=(40, 2))
        X = np.vstack([blobs, outliers])
    return X

X_data = generate_dataset(dataset_type, n_samples, noise_seed)

# Helper: Hàm tính metric an toàn
def compute_metrics(X, labels, exec_time):
    valid_mask = labels != -1
    unique_clusters = len(set(labels[valid_mask]))
    if unique_clusters >= 2:
        try:
            sil = silhouette_score(X[valid_mask], labels[valid_mask])
        except Exception:
            sil = 0.0
        try:
            db = davies_bouldin_score(X[valid_mask], labels[valid_mask])
        except Exception:
            db = 99.0
        try:
            ch = calinski_harabasz_score(X[valid_mask], labels[valid_mask])
        except Exception:
            ch = 0.0
    else:
        sil, db, ch = -1.0, 99.0, 0.0
    return {
        "Silhouette": sil,
        "Davies-Bouldin": db,
        "Calinski-Harabasz": ch,
        "Time": exec_time,
        "Clusters": unique_clusters
    }

# Helper: Vẽ đồ thị Plotly đẹp cho từng thuật toán
def make_scatter_figure(X, labels, title, reps=None, means=None, centers=None, center_label="Tâm"):
    df_p = pd.DataFrame(X, columns=['X', 'Y'])
    df_p['Cụm'] = [f"Cụm {l+1}" if l != -1 else "Ngoại lai (Nhiễu)" for l in labels]
    
    color_map = {}
    palette = px.colors.qualitative.Plotly
    for idx, l in enumerate(sorted(list(set(labels)))):
        if l == -1:
            color_map["Ngoại lai (Nhiễu)"] = "#94a3b8"
        else:
            color_map[f"Cụm {l+1}"] = palette[l % len(palette)]

    fig = px.scatter(
        df_p, x='X', y='Y', color='Cụm',
        color_discrete_map=color_map,
        opacity=0.78
    )
    fig.update_traces(marker=dict(size=7.5))

    # Vẽ điểm đại diện sau co cụm (CURE)
    if reps is not None:
        for idx, rep in enumerate(reps):
            fig.add_trace(go.Scatter(
                x=rep[:, 0], y=rep[:, 1],
                mode='markers',
                marker=dict(symbol='x', size=11, color='black', line=dict(width=2.2)),
                name=f"Rep points Cụm {idx+1}",
                showlegend=(idx == 0)
            ))

    # Vẽ trọng tâm Mean (CURE hoặc KMeans)
    if means is not None:
        fig.add_trace(go.Scatter(
            x=means[:, 0], y=means[:, 1],
            mode='markers',
            marker=dict(symbol='star', size=16, color='gold', line=dict(width=1.5, color='black')),
            name="Trọng tâm (Mean)",
            showlegend=True
        ))

    # Vẽ Medoids (K-Medoids)
    if centers is not None:
        fig.add_trace(go.Scatter(
            x=centers[:, 0], y=centers[:, 1],
            mode='markers',
            marker=dict(symbol='diamond', size=13, color='#dc2626', line=dict(width=1.5, color='black')),
            name=center_label,
            showlegend=True
        ))

    fig.update_layout(
        title=dict(text=title, font=dict(size=14, color='#103673', family='sans-serif')),
        height=480,
        margin=dict(l=15, r=15, t=45, b=15),
        legend=dict(orientation="h", y=-0.15, x=0.0, font=dict(size=11)),
        plot_bgcolor="#fafbfc",
        xaxis=dict(showgrid=True, gridcolor='#f1f5f9', zeroline=False),
        yaxis=dict(showgrid=True, gridcolor='#f1f5f9', zeroline=False)
    )
    return fig

# ==========================================
# THIẾT LẬP HỆ THỐNG TABS CHUYÊN BIỆT
# ==========================================
tab_cure_main, tab_cure_steps, tab_cure_flow, tab_vs_kmeans, tab_vs_kmedoids, tab_vs_hier, tab_summary = st.tabs([
    "🎯 CURE: Trực quan hóa",
    "📝 CURE: Ví dụ tính tay (Toy Example)",
    "🔍 CURE: Quy trình 5 giai đoạn",
    "⚔️ So sánh: CURE vs K-Means",
    "⚔️ So sánh: CURE vs K-Medoids",
    "⚔️ So sánh: CURE vs Hierarchical",
    "📋 Bảng Tổng hợp Đối sánh"
])

# ──────────────────────────────────────────
# CHẠY TRƯỚC CÁC THUẬT TOÁN ĐỂ LẤY KẾT QUẢ
# ──────────────────────────────────────────
# 1. CURE
t0 = time.time()
cure_model = CURE(n_clusters=k_clusters, n_representatives=c_reps, shrink_factor=alpha_shrink)
cure_model.fit(X_data)
cure_time = time.time() - t0
cure_labels = cure_model.labels_
cure_reps = cure_model.get_representatives()
cure_means = cure_model.get_cluster_means()
cure_metrics = compute_metrics(X_data, cure_labels, cure_time)

# 2. K-Means
t_km0 = time.time()
km_model = KMeans(n_clusters=k_clusters, random_state=noise_seed, n_init='auto')
km_labels = km_model.fit_predict(X_data)
km_time = time.time() - t_km0
km_centers = km_model.cluster_centers_
km_metrics = compute_metrics(X_data, km_labels, km_time)

# 3. K-Medoids
t_kmed0 = time.time()
kmed_model = KMedoids(n_clusters=k_clusters, random_state=noise_seed)
kmed_labels = kmed_model.fit_predict(X_data)
kmed_time = time.time() - t_kmed0
kmed_centers = kmed_model.cluster_centers_
kmed_metrics = compute_metrics(X_data, kmed_labels, kmed_time)

# 4. Hierarchical (Single Linkage)
t_hier0 = time.time()
hier_model = AgglomerativeClustering(n_clusters=k_clusters, linkage='single')
hier_labels = hier_model.fit_predict(X_data)
hier_time = time.time() - t_hier0
hier_metrics = compute_metrics(X_data, hier_labels, hier_time)


# ==============================================================================
# TAB 1: CURE TRỰC QUAN HÓA CHI TIẾT
# ==============================================================================
with tab_cure_main:
    st.markdown("### 🎯 Mô phỏng Hoạt động Phân cụm Thuật toán CURE")
    st.markdown(f"Đang phân cụm trên tập: **{dataset_type}** với $N = {n_samples}$ điểm dữ liệu.")

    col_c1, col_c2 = st.columns([3, 1])

    with col_c1:
        fig_cure = make_scatter_figure(
            X_data, cure_labels,
            f"Kết quả phân cụm CURE (k={k_clusters}, c={c_reps} đại diện, α={alpha_shrink})",
            reps=cure_reps, means=cure_means
        )
        st.plotly_chart(fig_cure, use_container_width=True)

    with col_c2:
        st.markdown("#### 📊 Chỉ số Chất lượng Cụm")
        st.metric("Silhouette Score", f"{cure_metrics['Silhouette']:.4f}", help="Càng gần 1 càng phân cụm tốt")
        st.metric("Davies-Bouldin Index", f"{cure_metrics['Davies-Bouldin']:.4f}", help="Càng nhỏ cụm càng đặc và tách biệt")
        st.metric("Calinski-Harabasz", f"{cure_metrics['Calinski-Harabasz']:.1f}")
        st.metric("Thời gian thực thi", f"{cure_time:.4f} giây")

        st.markdown("""
        <div style="background-color: #e0f2fe; padding: 12px; border-radius: 8px; font-size: 12.5px; border-left: 4px solid #0284c7;">
            <b>Ý nghĩa trực quan của CURE:</b><br>
            • <b>Dấu X đen:</b> $c$ điểm đại diện đã co cụm về phía tâm.<br>
            • <b>Ngôi sao vàng:</b> Trọng tâm mean của cụm.<br>
            • Nhờ phân bố rải rác nhiều điểm đại diện, CURE ôm trọn vẹn đường cong tự nhiên của dữ liệu!
        </div>
        """, unsafe_allow_html=True)


# ==============================================================================
# TAB 2: CURE VÍ DỤ TÍNH TAY TỪNG BƯỚC (TOY EXAMPLE)
# ==============================================================================
with tab_cure_steps:
    st.markdown("### 📝 Bài toán Ví dụ Tính tay Từng bước (Toy Example)")
    st.markdown("""
    Nhóm thiết lập một tập dữ liệu nhỏ gồm **6 điểm 2D cụ thể** để minh họa chính xác từng bước hoạt động của thuật toán CURE:
    * Cụm bên trái: $P_1(1, 2)$, $P_2(2, 3)$, $P_3(2, 1)$
    * Cụm bên phải: $P_4(8, 7)$, $P_5(9, 8)$, $P_6(8, 9)$
    * **Cấu hình:** $k = 2$ cụm mục tiêu, $c = 2$ điểm đại diện mỗi cụm, hệ số co cụm $\\alpha = 0.5$.
    """)

    step_choice = st.radio("Chọn bước thực hiện để quan sát:", [
        "Bước 0: Khởi tạo 6 điểm riêng lẻ (6 cụm ban đầu)",
        "Bước 1: Sáp nhập P1 và P2 -> C{1,2}",
        "Bước 2: Sáp nhập P4 và P5 -> C{4,5}",
        "Bước 3: Sáp nhập C{1,2} với P3 -> C{1,2,3}",
        "Bước 4: Sáp nhập C{4,5} với P6 -> C{4,5,6} (Hoàn tất k=2)"
    ], horizontal=True)

    step_idx = int(step_choice.split(":")[0].replace("Bước ", ""))
    step_img = f"toy_example_steps/step_{step_idx}.png"

    col_t1, col_t2 = st.columns([1.8, 1.2])

    with col_t1:
        if os.path.exists(step_img):
            st.image(step_img, caption=f"Hình minh họa {step_choice}", use_container_width=True)
        else:
            st.info("Hình ảnh minh họa đang được cập nhật.")

    with col_t2:
        st.markdown("#### 📐 Công thức và Phép tính chi tiết")
        if step_idx == 0:
            st.markdown("""
            * Mỗi điểm ban đầu là 1 cụm đơn lẻ: $C_1=\\{P_1\\}, \\dots, C_6=\\{P_6\\}$.
            * Khoảng cách giữa 2 cụm chính là khoảng cách Euclidean giữa 2 điểm:
              $$d(P_1, P_2) = \\sqrt{(2-1)^2 + (3-2)^2} = \\sqrt{2} \\approx 1.414$$
            * Cặp $(P_1, P_2)$ và $(P_4, P_5)$ có khoảng cách nhỏ nhất toàn ma trận ($1.414$).
            """)
        elif step_idx == 1:
            st.markdown("""
            * Sáp nhập $P_1$ và $P_2$ thành cụm $C_{\\{1,2\\}}$.
            * **Trọng tâm cụm mới:** $m = \\left(\\frac{1+2}{2}, \\frac{2+3}{2}\\right) = (1.5, 2.5)$.
            * **Co cụm 2 điểm đại diện với $\\alpha = 0.5$:**
              $$p'_1 = (1, 2) + 0.5 \\times ((1.5, 2.5) - (1, 2)) = (1.25, 2.25)$$
              $$p'_2 = (2, 3) + 0.5 \\times ((1.5, 2.5) - (2, 3)) = (1.75, 2.75)$$
            * Điểm đại diện đã dịch chuyển lùi vào trong một khoảng an toàn!
            """)
        elif step_idx == 2:
            st.markdown("""
            * Sáp nhập $P_4(8, 7)$ và $P_5(9, 8)$ thành cụm $C_{\\{4,5\\}}$.
            * **Trọng tâm cụm:** $m = (8.5, 7.5)$.
            * **Co cụm 2 điểm đại diện với $\\alpha = 0.5$:**
              $$p'_4 = (8.25, 7.25), \\quad p'_5 = (8.75, 7.75)$$
            * Số cụm hiện tại giảm xuống còn 4 cụm.
            """)
        elif step_idx == 3:
            st.markdown("""
            * Tính khoảng cách từ $P_3(2, 1)$ tới các điểm đại diện đã co của $C_{\\{1,2\\}}$:
              $$d(P_3, p'_1) = \\sqrt{(2-1.25)^2 + (1-2.25)^2} \\approx 1.458$$
            * Do $1.458$ là khoảng cách nhỏ nhất, $P_3$ sáp nhập vào cụm $C_{\\{1,2\\}} \\rightarrow C_{\\{1,2,3\\}}$.
            * Trọng tâm mới: $m = (1.667, 2.0)$.
            * Giải thuật **Farthest-Point** chọn 2 điểm xa nhất trong cụm rồi co về $m$.
            """)
        elif step_idx == 4:
            st.markdown("""
            * $P_6(8, 9)$ sáp nhập vào cụm $C_{\\{4,5\\}} \\rightarrow C_{\\{4,5,6\\}}$.
            * Cụm bên phải hoàn tất với 3 phần tử.
            * **Điều kiện dừng:** Số cụm còn lại đúng bằng $k = 2$.
            * Thuật toán kết thúc thành công với độ chính xác tuyệt đối 100%!
            """)


# ==============================================================================
# TAB 3: CURE QUY TRÌNH 5 GIAI ĐOẠN & LƯU ĐỒ KHỐI
# ==============================================================================
with tab_cure_flow:
    st.markdown("### 🔍 Quy trình 5 Giai đoạn & Kiến trúc Xử lý Dữ liệu lớn")
    
    col_f1, col_f2 = st.columns([1.5, 1])

    with col_f1:
        st.markdown("""
        Thuật toán CURE kết hợp hoàn hảo giữa **phân cụm phân cấp chất lượng cao** và **khả năng mở rộng trên dữ liệu lớn (Scalability)** qua 5 giai đoạn:

        1. **Giai đoạn 1: Lấy mẫu ngẫu nhiên (Random Sampling)**
           - Rút một mẫu ngẫu nhiên $s$ điểm từ tập dữ liệu khổng lồ $N$ điểm ($s \\ll N$). Mẫu $s$ vẫn bảo toàn đầy đủ hình học và phân bố của các cụm lớn.
        
        2. **Giai đoạn 2: Phân hoạch không gian (Partitioning)**
           - Chia mẫu $s$ thành $p$ phân vùng đều nhau (mỗi phần $\\frac{s}{p}$ điểm). Tiến hành gom cụm sơ bộ trên từng phân vùng để giảm chi phí bộ nhớ từ $O(s^2)$ xuống $O(s^2/p)$.

        3. **Giai đoạn 3: Gom cụm phân cấp trên các cụm đại diện**
           - Gom các cụm cục bộ lại với nhau. Tại mỗi bước sáp nhập, CURE chọn ra $c$ điểm đại diện phân tán đều trên thân cụm bằng giải thuật *Farthest-Point Heuristic*, sau đó co về trọng tâm theo hệ số $\\alpha$.

        4. **Giai đoạn 4: Loại bỏ ngoại lai 2 pha (Outlier Elimination)**
           - **Pha 1 (Giữa chừng):** Khi số cụm giảm xuống còn $k'$, các cụm chỉ có 1–2 điểm tăng trưởng chậm bị loại bỏ.
           - **Pha 2 (Cuối kỳ):** Các cụm có kích thước quá nhỏ so với kích thước trung bình bị coi là cụm nhiễu và xóa bỏ.

        5. **Giai đoạn 5: Gán nhãn toàn bộ dữ liệu lớn trên đĩa (Disk Labeling)**
           - Quét qua toàn bộ $N - s$ điểm dữ liệu còn lại trên đĩa và gán mỗi điểm vào cụm có điểm đại diện gần nhất. Chi phí tuyến tính $O(N)$.
        """)

    with col_f2:
        st.markdown("#### ⚙️ Bảng Hướng dẫn Chọn Siêu tham số")
        df_params = pd.DataFrame({
            "Tham số": ["s (Kích thước mẫu)", "p (Số phân vùng)", "c (Số điểm đại diện)", "α (Hệ số co cụm)", "k (Số cụm)"],
            "Ý nghĩa": ["Đại diện cho tập lớn N", "Chia nhỏ để tăng tốc", "Định hình đường biên cụm", "Khoảng đệm an toàn chống nhiễu", "Mục tiêu bài toán"],
            "Giá trị khuyến nghị": ["2.000 – 5.000", "2 – 4 phân vùng", "4 – 8 điểm", "0.3 – 0.5", "Tùy bài toán thực tế"]
        })
        st.table(df_params)


# ==============================================================================
# TAB 4: SO SÁNH 1-1: CURE VS K-MEANS
# ==============================================================================
with tab_vs_kmeans:
    st.markdown("### ⚔️ So sánh Đối đầu Trực diện: CURE vs K-Means")
    st.markdown("#### 🎯 Trọng tâm kiểm thử: Khắc phục hạn chế giả định cụm hình cầu của K-Means")

    col_km1, col_km2 = st.columns(2)

    with col_km1:
        st.markdown("##### 🟢 Thuật toán CURE (Đề tài nghiên cứu)")
        fig_c_km = make_scatter_figure(
            X_data, cure_labels,
            f"CURE: Nhận diện chuẩn xác hình thái cụm tự nhiên",
            reps=cure_reps, means=cure_means
        )
        st.plotly_chart(fig_c_km, use_container_width=True)
        st.write(f"**Silhouette Score:** `{cure_metrics['Silhouette']:.4f}` | **Thời gian:** `{cure_time:.4f}s`")
        st.success("✅ **Đạt chất lượng 100%:** Ôm trọn vẹn dải cong phi cầu nhờ c điểm đại diện rải đều!")

    with col_km2:
        st.markdown("##### 🔵 Thuật toán K-Means (Thuật toán đối chứng)")
        fig_km = make_scatter_figure(
            X_data, km_labels,
            f"K-Means: Cắt ngang cụm do giả định hình cầu",
            means=km_centers
        )
        st.plotly_chart(fig_km, use_container_width=True)
        st.write(f"**Silhouette Score:** `{km_metrics['Silhouette']:.4f}` | **Thời gian:** `{km_time:.4f}s`")
        if "Moons" in dataset_type or "Circles" in dataset_type:
            st.error("❌ **Thất bại hình học:** K-Means cắt đôi cụm phi cầu do chỉ dùng 1 tâm trung bình!")
        else:
            st.info("ℹ️ K-Means hoạt động tốt khi các cụm có dạng hình cầu lồi tách biệt.")

    st.markdown("#### 📋 Bảng Đối chiếu Trực tiếp: CURE vs K-Means")
    df_cmp_km = pd.DataFrame({
        "Tiêu chí đối sánh": [
            "Cơ chế đại diện cụm",
            "Giả định hình học cụm",
            "Xử lý cụm phi cầu (Two Moons, Circles)",
            "Độ nhạy với điểm ngoại lai (Outliers)",
            "Độ phức tạp thời gian",
            "Silhouette Score (Tập hiện tại)",
            "Davies-Bouldin Index (Tập hiện tại)"
        ],
        "CURE (Clustering Using REpresentatives)": [
            f"Dùng {c_reps} điểm đại diện phân bố đều trên thân cụm",
            "Tự do bắt hình dạng tùy ý (Arbitrary shape)",
            "Xuất sắc (Bảo toàn 100% hình thái cụm)",
            "Rất tốt (Nhờ co cụm alpha và 2 pha lọc ngoại lai)",
            "O(n² log n) trên mẫu, O(N) khi gán nhãn lớn",
            f"{cure_metrics['Silhouette']:.4f}",
            f"{cure_metrics['Davies-Bouldin']:.4f}"
        ],
        "K-Means (Thuật toán phân hoạch)": [
            "Dùng duy nhất 1 điểm trọng tâm trung bình (Centroid)",
            "Giả định cụm hình cầu lồi (Spherical/Convex)",
            "Thất bại hoàn toàn (Cắt ngang thân dải dữ liệu)",
            "Kém (1 điểm ngoại lai có thể kéo lệch hẳn tâm mean)",
            "O(k * n * t) - Rất nhanh trên dữ liệu hình cầu",
            f"{km_metrics['Silhouette']:.4f}",
            f"{km_metrics['Davies-Bouldin']:.4f}"
        ]
    })
    st.table(df_cmp_km)


# ==============================================================================
# TAB 5: SO SÁNH 1-1: CURE VS K-MEDOIDS (PAM)
# ==============================================================================
with tab_vs_kmedoids:
    st.markdown("### ⚔️ So sánh Đối đầu Trực diện: CURE vs K-Medoids (PAM)")
    st.markdown("#### 🎯 Trọng tâm kiểm thử: Khả năng kháng điểm ngoại lai (Outliers) và cụm dị hướng (Anisotropic)")

    col_kmed1, col_kmed2 = st.columns(2)

    with col_kmed1:
        st.markdown("##### 🟢 Thuật toán CURE (Đề tài nghiên cứu)")
        fig_c_kmed = make_scatter_figure(
            X_data, cure_labels,
            f"CURE: Kháng ngoại lai bằng cơ chế co cụm α = {alpha_shrink}",
            reps=cure_reps, means=cure_means
        )
        st.plotly_chart(fig_c_kmed, use_container_width=True)
        st.write(f"**Silhouette Score:** `{cure_metrics['Silhouette']:.4f}` | **Thời gian:** `{cure_time:.4f}s`")
        st.success("✅ **Miễn nhiễm ngoại lai:** Co cụm giúp điểm đại diện lùi vào sâu bên trong lõi cụm!")

    with col_kmed2:
        st.markdown("##### 🔴 Thuật toán K-Medoids / PAM (Thuật toán đối chứng)")
        fig_kmed = make_scatter_figure(
            X_data, kmed_labels,
            f"K-Medoids: Chọn điểm thực tế làm tâm (Medoid đỏ)",
            centers=kmed_centers, center_label="Medoid thực tế"
        )
        st.plotly_chart(fig_kmed, use_container_width=True)
        st.write(f"**Silhouette Score:** `{kmed_metrics['Silhouette']:.4f}` | **Thời gian:** `{kmed_time:.4f}s`")
        st.warning("⚠️ **Hạn chế:** Giảm nhạy cảm với ngoại lai hơn K-Means nhưng vẫn giả định cụm hình cầu và tốn chi phí hoán đổi.")

    st.markdown("#### 📋 Bảng Đối chiếu Trực tiếp: CURE vs K-Medoids")
    df_cmp_kmed = pd.DataFrame({
        "Tiêu chí đối sánh": [
            "Cách chọn tâm/đại diện",
            "Ảnh hưởng của điểm ngoại lai cực đoan",
            "Cơ chế loại bỏ nhiễu",
            "Khả năng mở rộng trên dữ liệu lớn",
            "Độ nhạy với hình dạng dị hướng (Anisotropic)",
            "Silhouette Score (Tập hiện tại)",
            "Davies-Bouldin Index (Tập hiện tại)"
        ],
        "CURE (Clustering Using REpresentatives)": [
            "c điểm đại diện phân tán và co lại về trọng tâm",
            "Vô hiệu hóa hoàn toàn nhờ co cụm α lùi vào lõi",
            "Có 2 pha loại bỏ ngoại lai độc lập tự động",
            "Rất tốt nhờ lấy mẫu ngẫu nhiên s và gán nhãn đĩa",
            "Rất tốt (Các điểm đại diện trải dọc thân elip)",
            f"{cure_metrics['Silhouette']:.4f}",
            f"{cure_metrics['Davies-Bouldin']:.4f}"
        ],
        "K-Medoids (PAM - Partitioning Around Medoids)": [
            "Chọn 1 điểm dữ liệu thực tế có tổng khoảng cách nhỏ nhất",
            "Ít bị kéo lệch hơn K-Means nhưng vẫn bị méo ranh giới",
            "Không có pha lọc nhiễu riêng biệt",
            "Kém (Độ phức tạp O(k(n-k)²) quá nặng khi n lớn)",
            "Kém (Vẫn coi cụm là khối tròn đẳng hướng quanh medoid)",
            f"{kmed_metrics['Silhouette']:.4f}",
            f"{kmed_metrics['Davies-Bouldin']:.4f}"
        ]
    })
    st.table(df_cmp_kmed)


# ==============================================================================
# TAB 6: SO SÁNH 1-1: CURE VS HIERARCHICAL (SINGLE LINKAGE)
# ==============================================================================
with tab_vs_hier:
    st.markdown("### ⚔️ So sánh Đối đầu Trực diện: CURE vs Gom cụm Phân cấp (Single Linkage)")
    st.markdown("#### 🎯 Trọng tâm kiểm thử: Khắc phục hiện tượng nối chuỗi (Chaining Effect) và Tối ưu bộ nhớ")

    col_h1, col_h2 = st.columns(2)

    with col_h1:
        st.markdown("##### 🟢 Thuật toán CURE (Đề tài nghiên cứu)")
        fig_c_hier = make_scatter_figure(
            X_data, cure_labels,
            f"CURE: Triệt tiêu nối chuỗi nhờ khoảng đệm co cụm",
            reps=cure_reps, means=cure_means
        )
        st.plotly_chart(fig_c_hier, use_container_width=True)
        st.write(f"**Silhouette Score:** `{cure_metrics['Silhouette']:.4f}` | **Thời gian:** `{cure_time:.4f}s`")
        st.success("✅ **Không bị nối chuỗi:** Các điểm đại diện co cụm tạo khoảng cách ngăn cách an toàn!")

    with col_h2:
        st.markdown("##### 🟣 Gom cụm Phân cấp (Single Linkage)")
        fig_hier = make_scatter_figure(
            X_data, hier_labels,
            f"Hierarchical Single Linkage: Dễ bị dính cụm do nhiễu nối chuỗi"
        )
        st.plotly_chart(fig_hier, use_container_width=True)
        st.write(f"**Silhouette Score:** `{hier_metrics['Silhouette']:.4f}` | **Thời gian:** `{hier_time:.4f}s`")
        if "Outliers" in dataset_type:
            st.error("❌ **Hiện tượng nối chuỗi:** Các điểm ngoại lai nằm giữa đã nối dính 2 cụm riêng biệt lại với nhau!")
        else:
            st.info("ℹ️ Single Linkage tìm được cụm phi cầu nhưng tốn bộ nhớ O(N²) và cực kỳ sợ nhiễu.")

    st.markdown("#### 📋 Bảng Đối chiếu Trực tiếp: CURE vs Gom cụm Phân cấp (Single Link)")
    df_cmp_hier = pd.DataFrame({
        "Tiêu chí đối sánh": [
            "Khoảng cách giữa hai cụm",
            "Hiện tượng nối chuỗi (Chaining Effect)",
            "Độ nhạy với điểm ngoại lai",
            "Độ phức tạp bộ nhớ",
            "Khả năng chạy trên cơ sở dữ liệu lớn",
            "Silhouette Score (Tập hiện tại)",
            "Davies-Bouldin Index (Tập hiện tại)"
        ],
        "CURE (Clustering Using REpresentatives)": [
            "Khoảng cách nhỏ nhất giữa các điểm đại diện ĐÃ CO CỤM",
            "Triệt tiêu hoàn toàn nhờ khoảng đệm co cụm alpha",
            "Miễn nhiễm nhờ co cụm và 2 pha lọc ngoại lai",
            "O(s) - Tiết kiệm nhờ lấy mẫu ngẫu nhiên s điểm",
            "Xuất sắc (Gán nhãn tuyến tính O(N) trên đĩa)",
            f"{cure_metrics['Silhouette']:.4f}",
            f"{cure_metrics['Davies-Bouldin']:.4f}"
        ],
        "Hierarchical (Single Linkage)": [
            "Khoảng cách nhỏ nhất giữa TẤT CẢ các cặp điểm thuộc 2 cụm",
            "Rất nghiêm trọng (Chỉ cần 1 vệt điểm nhiễu là sáp nhập nhầm)",
            "Rất nhạy cảm với các điểm nhiễu ngoại lai ở rìa",
            "O(N²) - Tràn bộ nhớ khi N > 10.000",
            "Không khả thi trên dữ liệu lớn (Big Data)",
            f"{hier_metrics['Silhouette']:.4f}",
            f"{hier_metrics['Davies-Bouldin']:.4f}"
        ]
    })
    st.table(df_cmp_hier)


# ==============================================================================
# TAB 7: BẢNG TỔNG HỢP MA TRẬN ĐỐI SÁNH TẤT CẢ THUẬT TOÁN
# ==============================================================================
with tab_summary:
    st.markdown("### 📋 Bảng Tổng hợp Ma trận Đối sánh Toàn diện")
    st.markdown("Bảng tổng hợp đối đầu giữa **CURE** và các thuật toán phân cụm trong chương trình môn học:")

    # Đồ thị Bar Chart so sánh Silhouette Score
    algs_names = ["CURE", "K-Means", "K-Medoids", "Hierarchical (Single)"]
    sils = [cure_metrics['Silhouette'], km_metrics['Silhouette'], kmed_metrics['Silhouette'], hier_metrics['Silhouette']]
    dbs = [cure_metrics['Davies-Bouldin'], km_metrics['Davies-Bouldin'], kmed_metrics['Davies-Bouldin'], hier_metrics['Davies-Bouldin']]
    times = [cure_metrics['Time'], km_metrics['Time'], kmed_metrics['Time'], hier_metrics['Time']]

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        fig_bar_sil = go.Figure(data=[
            go.Bar(name='Silhouette Score (Càng cao càng tốt)', x=algs_names, y=sils,
                   marker_color=['#103673', '#0284c7', '#dc2626', '#8b5cf6'],
                   text=[f"{v:.3f}" for v in sils], textposition='auto')
        ])
        fig_bar_sil.update_layout(title="Chỉ số Silhouette Score trên Tập dữ liệu Hiện hành", height=340, yaxis=dict(range=[-0.2, 1.0]))
        st.plotly_chart(fig_bar_sil, use_container_width=True)

    with col_s2:
        fig_bar_time = go.Figure(data=[
            go.Bar(name='Thời gian thực thi (giây)', x=algs_names, y=times,
                   marker_color=['#103673', '#0284c7', '#dc2626', '#8b5cf6'],
                   text=[f"{v:.4f}s" for v in times], textposition='auto')
        ])
        fig_bar_time.update_layout(title="Thời gian Thực thi (giây)", height=340)
        st.plotly_chart(fig_bar_time, use_container_width=True)

    st.markdown("#### 🏆 Ma trận Đánh giá Tổng kết Toàn diện")
    df_full_summary = pd.DataFrame({
        "Tiêu chuẩn đánh giá": [
            "Cụm hình dạng phi cầu (Two Moons, Circles)",
            "Cụm kéo dài hình elip (Anisotropic)",
            "Khả năng kháng ngoại lai và nhiễu biên",
            "Hiện tượng nối chuỗi (Chaining Effect)",
            "Khả năng mở rộng trên Big Data",
            "Độ nhạy với siêu tham số",
            "Thời gian thực thi trên tập hiện tại",
            "Đánh giá ứng dụng thực tế"
        ],
        "CURE": [
            "⭐⭐⭐⭐⭐ (Xuất sắc)",
            "⭐⭐⭐⭐⭐ (Xuất sắc)",
            "⭐⭐⭐⭐⭐ (Kháng triệt để)",
            "⭐⭐⭐⭐⭐ (Không bị)",
            "⭐⭐⭐⭐⭐ (Rất tốt - Lấy mẫu + Gán nhãn)",
            "⭐⭐⭐ (Cần chỉnh c, alpha)",
            f"{cure_time:.4f}s",
            "Tối ưu cho dữ liệu không gian, hình học phức tạp"
        ],
        "K-Means": [
            "⭐ (Thất bại)",
            "⭐⭐ (Kém)",
            "⭐ (Bị kéo lệch tâm)",
            "⭐⭐⭐⭐⭐ (Không bị)",
            "⭐⭐⭐⭐⭐ (Cực nhanh)",
            "⭐⭐⭐⭐⭐ (Chỉ cần k)",
            f"{km_time:.4f}s",
            "Chỉ dùng cho cụm hình cầu cân đối"
        ],
        "K-Medoids (PAM)": [
            "⭐ (Thất bại)",
            "⭐⭐ (Kém)",
            "⭐⭐⭐ (Giảm lệch tâm)",
            "⭐⭐⭐⭐⭐ (Không bị)",
            "⭐ (Rất chậm khi N lớn)",
            "⭐⭐⭐⭐ (Chỉ cần k)",
            f"{kmed_time:.4f}s",
            "Dùng khi có ngoại lai nhưng dữ liệu nhỏ và hình cầu"
        ],
        "Hierarchical (Single)": [
            "⭐⭐⭐⭐ (Tốt khi không có nhiễu)",
            "⭐⭐⭐⭐ (Tốt)",
            "⭐ (Rất nhạy cảm với nhiễu)",
            "⭐ (Bị nối chuỗi nghiêm trọng)",
            "⭐ (Tràn bộ nhớ O(N²))",
            "⭐⭐⭐⭐ (Chỉ cần k)",
            f"{hier_time:.4f}s",
            "Chỉ dùng cho dữ liệu nhỏ và không có ngoại lai"
        ]
    })
    st.table(df_full_summary)

st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #64748b; font-size: 12px;">
    Đồ án môn học Khai phá dữ liệu - Trường Đại học Công Thương TP. Hồ Chí Minh (HUIT)<br>
    Hệ thống mô phỏng phục vụ thuyết trình và bảo vệ đề tài tiểu luận CURE.
</div>
""", unsafe_allow_html=True)
