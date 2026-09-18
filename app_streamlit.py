"""
Web Demo Tương tác Thuật toán CURE (Clustering Using REpresentatives)
Phục vụ Báo cáo Đồ án môn Khai phá dữ liệu - ĐH Công Thương TP.HCM (HUIT)
Khởi chạy: python -m streamlit run app_streamlit.py
"""

import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from sklearn import datasets
from sklearn.cluster import KMeans, DBSCAN
from sklearn.metrics import silhouette_score, davies_bouldin_score
import time

from cure_algorithm import CURE

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="CURE Clustering Demo - HUIT",
    page_icon="🔬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Header giao diện
st.markdown("""
<div style="background: linear-gradient(135deg, #103673 0%, #1f4e79 100%); padding: 20px; border-radius: 10px; color: white; margin-bottom: 25px;">
    <h2 style="margin: 0; color: #F2A900;">TRƯỜNG ĐẠI HỌC CÔNG THƯƠNG TP. HỒ CHÍ MINH (HUIT)</h2>
    <h3 style="margin: 5px 0 0 0; color: white;">HỆ THỐNG MÔ PHỎNG & ĐỐI SÁNH THUẬT TOÁN PHÂN CỤM CURE</h3>
    <p style="margin: 5px 0 0 0; font-size: 14px; opacity: 0.9;">Đề tài: Phân cụm dữ liệu dựa trên thuật toán CURE (Clustering Using REpresentatives)</p>
</div>
""", unsafe_allow_html=True)

# SIDEBAR: Điều khiển tham số
st.sidebar.header("⚙️ CẤU HÌNH THAM SỐ")

dataset_type = st.sidebar.selectbox(
    "1. Chọn tập dữ liệu thử nghiệm:",
    [
        "Two Moons (2 Vầng trăng khuyết - Phi cầu)",
        "Concentric Circles (2 Vòng tròn đồng tâm)",
        "Anisotropic Blobs (Cụm kéo dài hình elip)",
        "Blobs with Outliers (Cụm có điểm ngoại lai/nhiễu)"
    ]
)

n_samples = st.sidebar.slider("Số lượng điểm dữ liệu:", min_value=100, max_value=600, value=300, step=50)

st.sidebar.markdown("---")
st.sidebar.subheader("Tham số thuật toán CURE:")
k_clusters = st.sidebar.slider("Số cụm mục tiêu (k):", min_value=2, max_value=6, value=2, step=1)
c_reps = st.sidebar.slider("Số điểm đại diện mỗi cụm (c):", min_value=1, max_value=8, value=4, step=1)
alpha_shrink = st.sidebar.slider("Hệ số co cụm (alpha):", min_value=0.0, max_value=1.0, value=0.4, step=0.05,
                                help="0.0: Điểm đại diện giữ nguyên ở biên | 1.0: Điểm đại diện co hoàn toàn về trọng tâm mean")

# Tạo dữ liệu dựa trên lựa chọn
@st.cache_data
def get_data(d_type, n_pts):
    if "Moons" in d_type:
        X, y = datasets.make_moons(n_samples=n_pts, noise=0.07, random_state=42)
    elif "Circles" in d_type:
        X, y = datasets.make_circles(n_samples=n_pts, factor=0.5, noise=0.05, random_state=42)
    elif "Anisotropic" in d_type:
        X_blobs, y = datasets.make_blobs(n_samples=n_pts, cluster_std=[0.9, 0.9, 0.9], random_state=42)
        trans = [[0.6, -0.6], [-0.4, 0.8]]
        X = np.dot(X_blobs, trans)
    else:
        blobs, y_b = datasets.make_blobs(n_samples=n_pts-40, centers=2, cluster_std=0.8, random_state=42)
        rng = np.random.RandomState(42)
        outliers = rng.uniform(low=-7, high=7, size=(40, 2))
        X = np.vstack([blobs, outliers])
    return X

X_data = get_data(dataset_type, n_samples)

# Các Tabs chính
tab1, tab2, tab3, tab4 = st.tabs([
    "🎯 Trực quan hóa CURE", 
    "⚔️ So sánh đối đầu (CURE vs K-Means vs DBSCAN)", 
    "📝 Ví dụ tính tay từng bước (Toy Example)", 
    "📚 Lý thuyết & Tài liệu"
])

with tab1:
    st.subheader("1. Kết quả phân cụm thuật toán CURE")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        # Chạy CURE
        t0 = time.time()
        cure_model = CURE(n_clusters=k_clusters, n_representatives=c_reps, shrink_factor=alpha_shrink)
        cure_model.fit(X_data)
        cure_time = time.time() - t0
        labels = cure_model.labels_
        reps = cure_model.get_representatives()
        means = cure_model.get_cluster_means()
        
        # Vẽ Plotly interactive
        df_plot = pd.DataFrame(X_data, columns=['X', 'Y'])
        df_plot['Cụm'] = [f"Cụm {l+1}" for l in labels]
        
        fig = px.scatter(
            df_plot, x='X', y='Y', color='Cụm',
            title=f"Kết quả phân cụm CURE (k={k_clusters}, c={c_reps}, α={alpha_shrink})",
            color_discrete_sequence=px.colors.qualitative.Plotly,
            opacity=0.75
        )
        fig.update_traces(marker=dict(size=8))
        
        # Thêm các điểm đại diện đã co cụm (Dấu X đen)
        for idx, rep in enumerate(reps):
            fig.add_trace(go.Scatter(
                x=rep[:, 0], y=rep[:, 1],
                mode='markers',
                marker=dict(symbol='x', size=11, color='black', line=dict(width=2)),
                name=f"Rep points Cụm {idx+1}"
            ))
            
        # Thêm trọng tâm mean (Hình sao vàng)
        fig.add_trace(go.Scatter(
            x=means[:, 0], y=means[:, 1],
            mode='markers',
            marker=dict(symbol='star', size=16, color='gold', line=dict(width=1.5, color='black')),
            name="Trọng tâm (Mean)"
        ))
        
        fig.update_layout(height=520, margin=dict(l=20, r=20, t=40, b=20), legend=dict(orientation="h", y=-0.15))
        st.plotly_chart(fig, use_container_width=True)
        
    with col2:
        st.markdown("#### 📊 Chỉ số đánh giá")
        sil = silhouette_score(X_data, labels)
        db = davies_bouldin_score(X_data, labels)
        
        st.metric("Silhouette Score", f"{sil:.4f}", help="Càng gần 1 càng phân tách cụm tốt")
        st.metric("Davies-Bouldin Index", f"{db:.4f}", help="Càng thấp càng tốt")
        st.metric("Thời gian thực thi", f"{cure_time:.4f} giây")
        
        st.markdown("""
        **Ý nghĩa trực quan:**
        * **Các dấu X đen:** Là các điểm đại diện (*representative points*) đã được co về phía trọng tâm theo hệ số $\\alpha$.
        * **Ngôi sao vàng:** Là trọng tâm (*mean*) của cụm.
        * Nhờ các điểm đại diện rải rác trên biên, CURE bao bọc và nhận dạng chính xác hình dạng cụm phức tạp!
        """)

with tab2:
    st.subheader("2. So sánh đối đầu giữa CURE và các thuật toán đã học")
    st.info("Môn Khai phá dữ liệu tại HUIT đã học: **K-Means**, **DBSCAN**, **Hierarchical**. Dưới đây là kết quả kiểm thử trực tiếp trên cùng tập dữ liệu!")
    
    # Chạy K-Means
    t_km0 = time.time()
    kmeans = KMeans(n_clusters=k_clusters, random_state=42, n_init='auto')
    km_labels = kmeans.fit_predict(X_data)
    t_km = time.time() - t_km0
    km_sil = silhouette_score(X_data, km_labels)
    km_db = davies_bouldin_score(X_data, km_labels)
    
    # Chạy DBSCAN
    t_db0 = time.time()
    dbscan = DBSCAN(eps=0.25 if "Moons" in dataset_type or "Circles" in dataset_type else 0.45, min_samples=5)
    db_labels = dbscan.fit_predict(X_data)
    t_db = time.time() - t_db0
    valid_mask = db_labels != -1
    n_db_clusters = len(set(db_labels[valid_mask]))
    if n_db_clusters >= 2:
        db_sil = silhouette_score(X_data[valid_mask], db_labels[valid_mask])
        db_db = davies_bouldin_score(X_data[valid_mask], db_labels[valid_mask])
    else:
        db_sil, db_db = -1.0, 99.0
        
    c_cmp1, c_cmp2, c_cmp3 = st.columns(3)
    
    with c_cmp1:
        st.markdown("##### 🟢 Thuật toán CURE")
        fig_c = px.scatter(x=X_data[:, 0], y=X_data[:, 1], color=[f"Cụm {l+1}" for l in labels], opacity=0.7)
        fig_c.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
        st.plotly_chart(fig_c, use_container_width=True)
        st.write(f"**Silhouette:** `{sil:.4f}` | **Thời gian:** `{cure_time:.3f}s`")
        st.caption("✅ Bắt trọn hình dạng phi cầu, tự nhiên")
        
    with c_cmp2:
        st.markdown("##### 🔵 Thuật toán K-Means")
        fig_k = px.scatter(x=X_data[:, 0], y=X_data[:, 1], color=[f"Cụm {l+1}" for l in km_labels], opacity=0.7)
        fig_k.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
        st.plotly_chart(fig_k, use_container_width=True)
        st.write(f"**Silhouette:** `{km_sil:.4f}` | **Thời gian:** `{t_km:.3f}s`")
        st.caption("❌ Chỉ gom được hình cầu, cắt đôi dải trăng khuyết/vòng tròn")
        
    with c_cmp3:
        st.markdown("##### 🟡 Thuật toán DBSCAN")
        db_colors = [f"Cụm {l+1}" if l != -1 else "Nhiễu" for l in db_labels]
        fig_d = px.scatter(x=X_data[:, 0], y=X_data[:, 1], color=db_colors, opacity=0.7)
        fig_d.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
        st.plotly_chart(fig_d, use_container_width=True)
        st.write(f"**Silhouette:** `{db_sil:.4f}` | **Số cụm:** `{n_db_clusters}`")
        st.caption("⚠️ Rất nhạy cảm với tham số eps và min_samples")
        
    st.markdown("#### 📋 Bảng tổng hợp so sánh chỉ số định lượng")
    df_metrics = pd.DataFrame({
        "Tiêu chí": ["Nhận diện hình dạng phi cầu", "Kháng nhiễu/Ngoại lai", "Độ nhạy tham số", "Silhouette Score", "Thời gian xử lý"],
        "CURE": ["Rất tốt (Đa dạng hình dáng)", "Tốt (Cơ chế co cụm + 2 pha lọc)", "Trung bình (c, alpha, k)", f"{sil:.4f}", f"{cure_time:.3f}s"],
        "K-Means": ["Kém (Chỉ gom hình cầu)", "Kém (Kéo tâm lệch vì outlier)", "Thấp (Chỉ cần k)", f"{km_sil:.4f}", f"{t_km:.3f}s"],
        "DBSCAN": ["Tốt (Dựa trên mật độ)", "Rất tốt (Lọc điểm nhiễu)", "Cao (Rất nhạy eps)", f"{db_sil:.4f}", f"{t_db:.3f}s"]
    })
    st.table(df_metrics)

with tab3:
    st.subheader("3. Bài toán Ví dụ tính tay từng bước (Toy Example)")
    st.markdown("""
    Để giảng viên đánh giá cao bài báo cáo, nhóm đã thiết lập một tập dữ liệu nhỏ gồm **6 điểm 2D**:
    * $P_1(1, 2)$, $P_2(2, 3)$, $P_3(2, 1)$ *(Nhóm bên trái)*
    * $P_4(8, 7)$, $P_5(9, 8)$, $P_6(8, 9)$ *(Nhóm bên phải)*
    
    *Tham số thử nghiệm:* $c = 2$ (2 điểm đại diện), $\\alpha = 0.5$ (hệ số co cụm 50%), mục tiêu gom thành $k = 2$ cụm.
    """)
    
    step_selected = st.radio("Chọn bước sáp nhập để xem chi tiết:", [
        "Bước 0: Khởi tạo 6 cụm ban đầu",
        "Bước 1: Sáp nhập P1 & P2 -> C{1,2}",
        "Bước 2: Sáp nhập P4 & P5 -> C{4,5}",
        "Bước 3: Sáp nhập C{1,2} & P3 -> C{1,2,3}",
        "Bước 4: Sáp nhập C{4,5} & P6 -> C{4,5,6} (Hoàn tất)"
    ], horizontal=True)
    
    step_idx = int(step_selected.split(":")[0].replace("Bước ", ""))
    img_step_path = f"toy_example_steps/step_{step_idx}.png"
    
    import os
    if os.path.exists(img_step_path):
        st.image(img_step_path, caption=f"Minh họa {step_selected}", use_container_width=True)
    else:
        st.warning("Chưa tìm thấy file ảnh minh họa bước này.")

with tab4:
    st.subheader("4. Tài liệu lý thuyết & Thông tin đề tài")
    st.markdown("""
    ### Bối cảnh ra đời của CURE
    * **Tác giả:** Sudipto Guha, Rajeev Rastogi, Kyuseok Shim (Bell Labs, 1998).
    * **Đột phá 1 (Nhiều điểm đại diện):** Thay vì đại diện 1 cụm bằng 1 điểm tâm (Centroid - chỉ gom hình cầu) hoặc tất cả các điểm (All-points - dễ bị chuỗi hóa nối dài do nhiễu), CURE chọn $c$ điểm phân tán tốt nhất trên biên cụm.
    * **Đột phá 2 (Co cụm về mean):** Các điểm đại diện được co về trọng tâm theo tỷ lệ $\\alpha$:
      $$p' = p + \\alpha \\times (\\text{mean} - p)$$
      Điều này kéo các điểm biên vào trong một chút, loại bỏ nguy cơ 2 cụm khác nhau bị dính vào nhau chỉ vì một vài điểm nhiễu ngoại lai!
    
    ---
    **Nhóm thực hiện:** Nhóm sinh viên Khai phá dữ liệu  
    **Trường:** Đại học Công Thương TP. Hồ Chí Minh (HUIT)
    """)
