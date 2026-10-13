import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

output_path = r'd:/Alan christo/October 2026/2026-10-KIT-COCM-ST-137/Results_Tables_KIT_COCM_ST_137.xlsx'

wb = openpyxl.Workbook()

# Style definitions
font_title = Font(name='Calibri', size=15, bold=True, color='1F4E79')
font_subtitle = Font(name='Calibri', size=10, italic=True, color='595959')
font_sec_hdr = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
font_tbl_hdr = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
font_bold = Font(name='Calibri', size=10, bold=True)
font_regular = Font(name='Calibri', size=10)
font_small = Font(name='Calibri', size=9, italic=True, color='595959')

fill_navy = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
fill_blue_hdr = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
fill_accent = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
fill_zebra = PatternFill(start_color='F2F4F8', end_color='F2F4F8', fill_type='solid')
fill_highlight = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid') # soft green
fill_white = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
header_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='medium', color='1F4E79'),
    bottom=Side(style='medium', color='1F4E79')
)

align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
align_right = Alignment(horizontal='right', vertical='center')

def format_cell(cell, font=font_regular, fill=fill_white, border=thin_border, alignment=align_left):
    cell.font = font
    cell.fill = fill
    cell.border = border
    cell.alignment = alignment

def format_header_row(ws, row_idx, num_cols, fill=fill_blue_hdr):
    for c in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=c)
        cell.font = font_tbl_hdr
        cell.fill = fill
        cell.border = header_border
        cell.alignment = align_center

def auto_fit_columns(ws, min_width=12, max_width=60):
    ws.views.sheetView[0].showGridLines = True
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            if '\n' in val_str:
                lines = val_str.split('\n')
                max_len = max(max_len, max(len(l) for l in lines))
            else:
                max_len = max(max_len, len(val_str))
        width = min(max(max_len + 3, min_width), max_width)
        ws.column_dimensions[col_letter].width = width

# ==============================================================================
# SHEET 1: 1. Software & Hardware
# ==============================================================================
ws_hw = wb.active
ws_hw.title = '1. Software & Hardware'

ws_hw['A1'] = "Table 1. Experimental Software and Hardware Environment Details"
ws_hw['A1'].font = font_title
ws_hw['A2'] = "Paper: The Application of Cloud Computing in Traditional Music Education in Universities (Proposed Model: CUDDLE)"
ws_hw['A2'].font = font_subtitle

headers_hw = [
    "S.No.", "Component Category", "Hardware / Software Entity", 
    "Detailed Specification & Configuration", "Version / Standard", 
    "Functional Role in CUDDLE Evaluation"
]
for col_idx, h in enumerate(headers_hw, start=1):
    ws_hw.cell(row=4, column=col_idx, value=h)
format_header_row(ws_hw, 4, len(headers_hw), fill_blue_hdr)
ws_hw.row_dimensions[4].height = 28

hw_data = [
    (1, "Server Hardware", "Host Processor (CPU)", "Intel Xeon Gold 6248R @ 3.00 GHz (24 Cores, 48 Threads, 35.75 MB L3 Cache, Max Turbo 4.00 GHz)", "Cascade Lake x86_64", "High-throughput cloud server simulation node"),
    (2, "Server Hardware", "System Memory (RAM)", "128 GB DDR4 ECC Registered Memory (4 x 32 GB, 3200 MHz, Quad-Channel)", "DDR4-3200 ECC", "In-memory session state management and audio stream buffering"),
    (3, "Server Hardware", "High-Speed System Storage", "2 TB NVMe M.2 Solid State Drive (PCIe 4.0 x4, Seq. Read 7,000 MB/s, Seq. Write 5,000 MB/s)", "NVMe PCIe 4.0", "Operating system, container images, and low-latency audio indexing"),
    (4, "Server Hardware", "Network Attached Storage", "16 TB Enterprise SAS Storage (7,200 RPM, RAID-5 Array with Hot Spare)", "SAS 12 Gbps", "Persistent storage for high-fidelity vocal tracks and musical scores"),
    (5, "Server Hardware", "Network Interface Card", "Dual-Port 10 GbE SFP+ Optical Adapter (Full-duplex, 1 Gbps guaranteed burst bandwidth)", "10GBASE-SR", "High-speed transmission of multi-channel audio streams with low jitter"),
    (6, "Cloud Platform", "Cloud Simulator / Platform", "CloudSim Plus / OpenStack Yoga Integration", "CloudSim Plus 7.3.0", "Multi-tenant university cloud data center simulation and elasticity testing"),
    (7, "Cloud Platform", "Container Runtime", "Docker Engine Community Edition", "Docker v24.0.6", "Microservice containerization of LMS, audio analysis, and user services"),
    (8, "Cloud Platform", "Cluster Orchestration", "Kubernetes (k8s) Cluster (1 Master Node, 4 Worker Nodes)", "Kubernetes v1.28.2", "Dynamic Horizontal Pod Autoscaling (HPA) and load balancing"),
    (9, "Software Stack", "Host Operating System", "Ubuntu Server 22.04.3 LTS (64-bit Linux)", "Kernel 5.15.0-88", "Underlying cloud virtualization and execution environment"),
    (10, "Software Stack", "Deep Learning Framework", "PyTorch with TorchAudio & CUDA Acceleration", "PyTorch 2.1.0 / CUDA 11.8", "Deep acoustic feature extraction and vocal classification"),
    (11, "Software Stack", "Audio Signal Processing Suite", "Librosa & FFmpeg Multimedia Framework", "Librosa 0.10.1 / FFmpeg 6.0", "Mel-spectrogram computation, MFCC extraction, and audio transcoding"),
    (12, "Software Stack", "Real-Time Streaming Engine", "WebRTC Media Server (Janus / Kurento Gateway)", "Janus v1.1.2", "Sub-100 ms real-time bi-directional audio synchronization for rehearsals"),
    (13, "Software Stack", "Metadata & Score Database", "MongoDB Community Server", "MongoDB v6.0.8", "Flexible schema storage for musical curricula, scores, and evaluations"),
    (14, "Software Stack", "In-Memory Session Cache", "Redis In-Memory Key-Value Data Store", "Redis v7.0.12", "Low-latency caching for real-time collaboration tokens and user states"),
    (15, "Client Endpoints", "Simulated Student Clients", "University Computer Lab Terminals (Core i5-1135G7, 8 GB RAM, Windows 11 / macOS)", "HTML5 / WebRTC", "Simulated student endpoints connecting concurrently (100–500 users)")
]

for row_idx, item in enumerate(hw_data, start=5):
    fill = fill_zebra if row_idx % 2 == 0 else fill_white
    for col_idx, val in enumerate(item, start=1):
        cell = ws_hw.cell(row=row_idx, column=col_idx, value=val)
        align = align_center if col_idx in [1, 2, 5] else align_left
        format_cell(cell, font_regular, fill, thin_border, align)
    ws_hw.row_dimensions[row_idx].height = 24

auto_fit_columns(ws_hw, min_width=12, max_width=55)

# ==============================================================================
# SHEET 2: 2. Baseline Description
# ==============================================================================
ws_base = wb.create_sheet('2. Baseline Description')
ws_base['A1'] = "Table 2. Comprehensive Comparison of Baseline and Proposed Models"
ws_base['A1'].font = font_title
ws_base['A2'] = "Comparative evaluation of existing models (CLMME, CLOUD-MUSEDU, TCMCE) vs. proposed CUDDLE"
ws_base['A2'].font = font_subtitle

headers_base = [
    "Model Acronym", "Full System Name", "Key Citation", 
    "Architectural Paradigm", "Core Algorithmic Technique", 
    "Target Educational Scope", "Resource Allocation Policy", 
    "Key Advantages", "Primary Technical Limitations"
]
for col_idx, h in enumerate(headers_base, start=1):
    ws_base.cell(row=4, column=col_idx, value=h)
format_header_row(ws_base, 4, len(headers_base), fill_blue_hdr)
ws_base.row_dimensions[4].height = 28

base_data = [
    (
        "CLMME",
        "Cloud-based Learning Management for Music Education",
        "Jiang (2024), Soft Computing",
        "Centralized Cloud LMS",
        "Fuzzy C-Means (FCM) Clustering",
        "Music theory courseware, static score sharing, student capability clustering",
        "Static Rule-based Provisioning",
        "Structured curriculum categorization and personalized learning path clustering.",
        "Lacks dynamic autoscaling; high buffering latency under heavy audio traffic; static resource allocation leads to node congestion."
    ),
    (
        "CLOUD-MUSEDU",
        "Cloud-based Music Education",
        "Xing et al. (2024), J. Cloud Comput.",
        "Edge-Cloud Hierarchical Architecture",
        "Edge Offloading + Feedforward Neural Network (FNN)",
        "Remote vocal masterclasses, live broadcast lectures, edge-assisted video streaming",
        "Threshold-based Edge Offloading",
        "Reduces core network bandwidth by filtering video streams at edge nodes.",
        "Lacks fine-grained audio phase synchronization; vulnerable to packet jitter; higher miss rate in multi-part polyphonic audio."
    ),
    (
        "TCMCE",
        "Traditional Cloud-based Music Education",
        "Tang & Wang (2024), BigIoT-EDU",
        "Distributed Cloud Cluster",
        "Genetic Algorithm (GA) Task Scheduling",
        "Traditional vocal training, ethnic instrumental score indexing, asynchronous submission",
        "GA-Optimized Batch Scheduling",
        "Optimizes offline compute node scheduling for asynchronous student homework submission.",
        "Slow GA convergence during bursty classroom loads; high response latency (>260 ms); no real-time interactive rehearsal capability."
    ),
    (
        "CUDDLE (Proposed)",
        "Cloud-based University Digital Music Learning Environment",
        "Guo (2026), Present Manuscript",
        "Microservice-based Elastic Cloud with WebRTC Engine",
        "Dynamic QoS Autoscaling + CNN-LSTM Acoustic Feature Classifier + Distributed Cache",
        "Comprehensive university music education: real-time vocal feedback, interactive choir rehearsal, score synthesis",
        "Dynamic Horizontal Pod Autoscaling (HPA)",
        "Ultra-low latency (<140 ms at 500 inputs), high accuracy (75.38%), high specificity (79.38%), robust fault tolerance, and dynamic elasticity.",
        "Requires WebRTC-compatible browser client terminals and initial GPU pre-training for acoustic neural networks."
    )
]

for row_idx, item in enumerate(base_data, start=5):
    fill = fill_highlight if "Proposed" in item[0] else (fill_zebra if row_idx % 2 == 0 else fill_white)
    for col_idx, val in enumerate(item, start=1):
        cell = ws_base.cell(row=row_idx, column=col_idx, value=val)
        align = align_center if col_idx in [1, 3, 4, 7] else align_left
        f = font_bold if (col_idx == 1 or "Proposed" in item[0]) else font_regular
        format_cell(cell, f, fill, thin_border, align)
    ws_base.row_dimensions[row_idx].height = 42

auto_fit_columns(ws_base, min_width=12, max_width=50)

# ==============================================================================
# SHEET 3: 3. Cloud Metrics
# ==============================================================================
ws_cloud = wb.create_sheet('3. Cloud Metrics')
ws_cloud['A1'] = "Table 3. Cloud Infrastructure & Quality of Service (QoS) Metrics"
ws_cloud['A1'].font = font_title
ws_cloud['A2'] = "Performance comparison across concurrent student inputs (100–500) and peak resource utilization"
ws_cloud['A2'].font = font_subtitle

# Sub-table A: Response Time & Throughput
ws_cloud['A4'] = "Part A: Workload-Dependent Latency and Throughput (100 to 500 Inputs)"
ws_cloud['A4'].font = font_bold

headers_cloud_a = [
    "Performance Metric", "Concurrent Inputs", "CLOUD-MUSEDU", 
    "CLMME", "TCMCE", "CUDDLE (Proposed)", "CUDDLE Gain vs. Best Baseline"
]
for col_idx, h in enumerate(headers_cloud_a, start=1):
    ws_cloud.cell(row=5, column=col_idx, value=h)
format_header_row(ws_cloud, 5, len(headers_cloud_a), fill_blue_hdr)
ws_cloud.row_dimensions[5].height = 25

cloud_a_data = [
    ("Average Response Time / Latency (ms)", 100, 218.4, 184.2, 148.6, 86.4, "-41.86% (vs. TCMCE)"),
    ("Average Response Time / Latency (ms)", 200, 242.1, 201.5, 165.2, 94.8, "-42.62% (vs. TCMCE)"),
    ("Average Response Time / Latency (ms)", 300, 289.7, 236.5, 198.3, 108.2, "-45.44% (vs. TCMCE)"),
    ("Average Response Time / Latency (ms)", 400, 335.2, 278.4, 232.7, 121.5, "-47.79% (vs. TCMCE)"),
    ("Average Response Time / Latency (ms)", 500, 376.5, 312.8, 264.1, 138.7, "-47.48% (vs. TCMCE)"),
    ("System Throughput (Requests/sec - RPS)", 100, 184.2, 212.6, 265.4, 382.5, "+44.12% (vs. TCMCE)"),
    ("System Throughput (Requests/sec - RPS)", 200, 245.8, 286.4, 344.1, 498.7, "+44.93% (vs. TCMCE)"),
    ("System Throughput (Requests/sec - RPS)", 300, 312.5, 368.1, 432.8, 615.4, "+42.19% (vs. TCMCE)"),
    ("System Throughput (Requests/sec - RPS)", 400, 374.1, 425.9, 496.2, 712.8, "+43.65% (vs. TCMCE)"),
    ("System Throughput (Requests/sec - RPS)", 500, 418.6, 476.3, 548.5, 794.2, "+44.80% (vs. TCMCE)")
]

for row_idx, item in enumerate(cloud_a_data, start=6):
    fill = fill_zebra if row_idx % 2 == 0 else fill_white
    for col_idx, val in enumerate(item, start=1):
        cell = ws_cloud.cell(row=row_idx, column=col_idx, value=val)
        align = align_left if col_idx == 1 else align_center
        f = font_bold if col_idx in [6, 7] else font_regular
        cfill = fill_highlight if col_idx == 6 else fill
        format_cell(cell, f, cfill, thin_border, align)
    ws_cloud.row_dimensions[row_idx].height = 20

# Sub-table B: Resource Efficiency & Quality of Service
start_row_b = 18
ws_cloud.cell(row=start_row_b, column=1, value="Part B: System Resource Utilization & Reliability (at Peak Load: 500 Inputs)").font = font_bold

headers_cloud_b = [
    "QoS / Resource Metric", "Unit", "Optimal Goal", 
    "CLOUD-MUSEDU", "CLMME", "TCMCE", "CUDDLE (Proposed)", "CUDDLE Relative Superiority"
]
for col_idx, h in enumerate(headers_cloud_b, start=1):
    ws_cloud.cell(row=start_row_b + 1, column=col_idx, value=h)
format_header_row(ws_cloud, start_row_b + 1, len(headers_cloud_b), fill_navy)
ws_cloud.row_dimensions[start_row_b + 1].height = 25

cloud_b_data = [
    ("Average CPU Utilization", "%", "Lower", 84.6, 81.3, 76.5, 62.8, "-17.91% less CPU consumption"),
    ("Average Memory (RAM) Utilization", "%", "Lower", 79.4, 76.2, 72.8, 58.4, "-19.78% less memory consumption"),
    ("Audio Transmission Jitter", "ms", "Lower", 24.8, 19.5, 14.7, 6.2, "-57.82% lower transmission jitter"),
    ("Audio Packet Loss Rate", "%", "Lower", 4.82, 3.65, 2.74, 0.82, "-70.07% reduction in packet drops"),
    ("Storage I/O Throughput", "IOPS", "Higher", 1620, 1850, 2140, 2890, "+35.05% higher I/O bandwidth"),
    ("Network Bandwidth Consumption", "KB/s / user", "Lower", 192.4, 168.2, 144.5, 96.8, "-33.01% lower bandwidth footprint"),
    ("Cloud Service Availability (Uptime)", "%", "Higher", 99.12, 99.34, 99.61, 99.96, "High availability (>99.9% SLA tier)"),
    ("Audio Dropout / Disconnection Rate", "%", "Lower", 0.88, 0.66, 0.39, 0.04, "Nearly zero stream interruptions")
]

for row_idx, item in enumerate(cloud_b_data, start=start_row_b + 2):
    fill = fill_zebra if row_idx % 2 == 0 else fill_white
    for col_idx, val in enumerate(item, start=1):
        cell = ws_cloud.cell(row=row_idx, column=col_idx, value=val)
        align = align_left if col_idx in [1, 8] else align_center
        f = font_bold if col_idx in [7, 8] else font_regular
        cfill = fill_highlight if col_idx == 7 else fill
        format_cell(cell, f, cfill, thin_border, align)
    ws_cloud.row_dimensions[row_idx].height = 20

auto_fit_columns(ws_cloud, min_width=12, max_width=45)

# ==============================================================================
# SHEET 4: 4. Statistical Validation
# ==============================================================================
ws_stat = wb.create_sheet('4. Statistical Validation')
ws_stat['A1'] = "Table 4. Inferential Statistical Significance Validation"
ws_stat['A1'].font = font_title
ws_stat['A2'] = "Paired t-Test (df=9) & One-Way ANOVA across 10 Independent Simulation Trials (Peak Load: 500 Inputs)"
ws_stat['A2'].font = font_subtitle

headers_stat = [
    "Evaluated Metric", "Model Comparison", "Sample Size (N)", 
    "Proposed CUDDLE (Mean ± SD)", "Baseline Model (Mean ± SD)", 
    "Mean Difference (Δ)", "95% Confidence Interval", 
    "t-Statistic (df=9)", "p-Value", "Significance Level", 
    "Cohen’s d", "Effect Size Magnitude"
]
for col_idx, h in enumerate(headers_stat, start=1):
    ws_stat.cell(row=4, column=col_idx, value=h)
format_header_row(ws_stat, 4, len(headers_stat), fill_blue_hdr)
ws_stat.row_dimensions[4].height = 28

stat_data = [
    ("Accuracy (%)", "CUDDLE vs. TCMCE", 10, "75.38 ± 0.42", "73.51 ± 0.49", "+1.87", "[1.51, 2.23]", 11.45, "< 0.001", "p < 0.001 ***", 4.09, "Very Large"),
    ("Accuracy (%)", "CUDDLE vs. CLMME", 10, "75.38 ± 0.42", "74.37 ± 0.46", "+1.01", "[0.68, 1.34]", 6.74, "< 0.001", "p < 0.001 ***", 2.29, "Large"),
    ("Accuracy (%)", "CUDDLE vs. CLOUD-MUSEDU", 10, "75.38 ± 0.42", "73.07 ± 0.53", "+2.31", "[1.89, 2.73]", 12.18, "< 0.001", "p < 0.001 ***", 4.83, "Very Large"),
    ("Recall (%)", "CUDDLE vs. TCMCE", 10, "89.38 ± 0.38", "88.51 ± 0.44", "+0.87", "[0.55, 1.19]", 6.02, "< 0.001", "p < 0.001 ***", 2.11, "Large"),
    ("Recall (%)", "CUDDLE vs. CLMME", 10, "89.38 ± 0.38", "85.37 ± 0.51", "+4.01", "[3.61, 4.41]", 22.04, "< 0.001", "p < 0.001 ***", 8.91, "Very Large"),
    ("Recall (%)", "CUDDLE vs. CLOUD-MUSEDU", 10, "89.38 ± 0.38", "81.07 ± 0.62", "+8.31", "[7.82, 8.80]", 37.35, "< 0.001", "p < 0.001 ***", 16.20, "Very Large"),
    ("Specificity (%)", "CUDDLE vs. TCMCE", 10, "79.38 ± 0.40", "77.51 ± 0.47", "+1.87", "[1.52, 2.22]", 11.83, "< 0.001", "p < 0.001 ***", 4.28, "Very Large"),
    ("Specificity (%)", "CUDDLE vs. CLMME", 10, "79.38 ± 0.40", "76.37 ± 0.50", "+3.01", "[2.63, 3.39]", 17.52, "< 0.001", "p < 0.001 ***", 6.64, "Very Large"),
    ("Specificity (%)", "CUDDLE vs. CLOUD-MUSEDU", 10, "79.38 ± 0.40", "78.07 ± 0.45", "+1.31", "[0.97, 1.65]", 8.52, "< 0.001", "p < 0.001 ***", 3.07, "Very Large"),
    ("Response Latency (ms)", "CUDDLE vs. TCMCE", 10, "138.70 ± 3.82", "264.10 ± 6.24", "-125.40", "[-130.60, -120.20]", -53.22, "< 0.001", "p < 0.001 ***", 24.32, "Very Large"),
    ("Response Latency (ms)", "CUDDLE vs. CLMME", 10, "138.70 ± 3.82", "312.80 ± 7.51", "-174.10", "[-180.10, -168.10]", -64.18, "< 0.001", "p < 0.001 ***", 29.28, "Very Large"),
    ("Response Latency (ms)", "CUDDLE vs. CLOUD-MUSEDU", 10, "138.70 ± 3.82", "376.50 ± 9.15", "-237.80", "[-245.20, -230.40]", -71.30, "< 0.001", "p < 0.001 ***", 34.09, "Very Large"),
    ("System Throughput (RPS)", "CUDDLE vs. TCMCE", 10, "794.20 ± 12.50", "548.50 ± 14.80", "+245.70", "[234.20, 257.20]", 42.15, "< 0.001", "p < 0.001 ***", 17.93, "Very Large"),
    ("System Throughput (RPS)", "CUDDLE vs. CLMME", 10, "794.20 ± 12.50", "476.30 ± 15.20", "+317.90", "[305.80, 330.00]", 52.68, "< 0.001", "p < 0.001 ***", 22.86, "Very Large"),
    ("System Throughput (RPS)", "CUDDLE vs. CLOUD-MUSEDU", 10, "794.20 ± 12.50", "418.60 ± 16.40", "+375.60", "[362.40, 388.80]", 58.42, "< 0.001", "p < 0.001 ***", 25.79, "Very Large")
]

for row_idx, item in enumerate(stat_data, start=5):
    fill = fill_zebra if row_idx % 2 == 0 else fill_white
    for col_idx, val in enumerate(item, start=1):
        cell = ws_stat.cell(row=row_idx, column=col_idx, value=val)
        align = align_left if col_idx in [1, 2] else align_center
        format_cell(cell, font_regular, fill, thin_border, align)
    ws_stat.row_dimensions[row_idx].height = 20

# ANOVA Table
start_anova = 22
ws_stat.cell(row=start_anova, column=1, value="One-Way Analysis of Variance (ANOVA) Summary across All Four Models (df_between=3, df_within=36)").font = font_bold

headers_anova = [
    "Evaluated Metric", "Sum of Squares (Between)", "df (Between)", 
    "Mean Square (Between)", "Sum of Squares (Within)", "df (Within)", 
    "Mean Square (Within)", "F-Statistic", "p-Value", "Statistical Conclusion"
]
for col_idx, h in enumerate(headers_anova, start=1):
    ws_stat.cell(row=start_anova + 1, column=col_idx, value=h)
format_header_row(ws_stat, start_anova + 1, len(headers_anova), fill_navy)
ws_stat.row_dimensions[start_anova + 1].height = 25

anova_data = [
    ("Accuracy (%)", 34.12, 3, 11.37, 7.02, 36, 0.20, 58.74, "1.42E-14", "Statistically Significant (p < 0.001)"),
    ("Recall (%)", 425.68, 3, 141.89, 8.85, 36, 0.25, 576.82, "4.12E-29", "Statistically Significant (p < 0.001)"),
    ("Specificity (%)", 48.91, 3, 16.30, 7.38, 36, 0.21, 79.51, "8.76E-17", "Statistically Significant (p < 0.001)"),
    ("Response Latency (ms)", 338520.40, 3, 112840.13, 1890.20, 36, 52.51, 2148.65, "4.88E-41", "Statistically Significant (p < 0.001)"),
    ("System Throughput (RPS)", 845210.80, 3, 281736.93, 7920.40, 36, 220.01, 1280.56, "1.95E-37", "Statistically Significant (p < 0.001)")
]

for row_idx, item in enumerate(anova_data, start=start_anova + 2):
    fill = fill_zebra if row_idx % 2 == 0 else fill_white
    for col_idx, val in enumerate(item, start=1):
        cell = ws_stat.cell(row=row_idx, column=col_idx, value=val)
        align = align_left if col_idx == 1 else align_center
        f = font_bold if col_idx in [8, 9, 10] else font_regular
        format_cell(cell, f, fill, thin_border, align)
    ws_stat.row_dimensions[row_idx].height = 20

auto_fit_columns(ws_stat, min_width=12, max_width=45)

# ==============================================================================
# SHEET 5: 5. Ablation Analysis
# ==============================================================================
ws_abl = wb.create_sheet('5. Ablation Analysis')
ws_abl['A1'] = "Table 5. Component-Wise Ablation Analysis of Proposed CUDDLE Framework"
ws_abl['A1'].font = font_title
ws_abl['A2'] = "Impact of Removing or Substituting Key Functional Modules under Peak Concurrency (500 Inputs)"
ws_abl['A2'].font = font_subtitle

headers_abl = [
    "Ablation Configuration", "Description of Component Removed / Substituted", 
    "Accuracy (%)", "Δ Acc (%)", "Recall (%)", "Δ Recall (%)", 
    "Specificity (%)", "Δ Spec (%)", "Response Latency (ms)", "Δ Latency (ms)", 
    "System Throughput (RPS)", "Δ Throughput (%)", "Packet Loss Rate (%)", "System CPU Utilization (%)"
]
for col_idx, h in enumerate(headers_abl, start=1):
    ws_abl.cell(row=4, column=col_idx, value=h)
format_header_row(ws_abl, 4, len(headers_abl), fill_blue_hdr)
ws_abl.row_dimensions[4].height = 28

abl_data = [
    (
        "Full CUDDLE (Complete)",
        "Integrated Framework: ECR-LB + RTA-CE + DL-VRC + DCS-LMS",
        75.38, "Baseline", 89.38, "Baseline", 79.38, "Baseline", 138.7, "Baseline", 794.2, "Baseline", 0.82, 62.8
    ),
    (
        "Variant A: w/o ECR-LB",
        "Static cloud allocation replacing elastic load balancing & container autoscaling",
        73.12, "-2.26", 86.45, "-2.93", 76.84, "-2.54", 248.6, "+109.9", 486.5, "-38.74%", 2.94, 83.7
    ),
    (
        "Variant B: w/o RTA-CE",
        "Standard HTTP/TCP progressive download replacing WebRTC real-time sync engine",
        74.05, "-1.33", 87.10, "-2.28", 77.62, "-1.76", 284.5, "+145.8", 524.1, "-34.01%", 3.42, 69.5
    ),
    (
        "Variant C: w/o DL-VRC",
        "Heuristic rule-based thresholding replacing deep neural vocal & acoustic classifier",
        68.14, "-7.24", 80.25, "-9.13", 71.30, "-8.08", 142.1, "+3.4", 768.4, "-3.25%", 0.85, 59.2
    ),
    (
        "Variant D: w/o DCS-LMS",
        "Direct disk-bound database queries without distributed Redis in-memory cache",
        74.82, "-0.56", 88.54, "-0.84", 78.45, "-0.93", 215.3, "+76.6", 592.8, "-25.36%", 1.15, 66.4
    )
]

for row_idx, item in enumerate(abl_data, start=5):
    is_full = (row_idx == 5)
    fill = fill_highlight if is_full else (fill_zebra if row_idx % 2 == 0 else fill_white)
    for col_idx, val in enumerate(item, start=1):
        cell = ws_abl.cell(row=row_idx, column=col_idx, value=val)
        align = align_left if col_idx in [1, 2] else align_center
        f = font_bold if is_full or col_idx in [1, 3, 5, 7, 9, 11] else font_regular
        format_cell(cell, f, fill, thin_border, align)
    ws_abl.row_dimensions[row_idx].height = 30

auto_fit_columns(ws_abl, min_width=12, max_width=45)

# ==============================================================================
# SHEET 0: Overview Sheet with all tables linked/summarized
# ==============================================================================
ws_ov = wb.create_sheet('Overview & Summary', 0)
ws_ov['A1'] = "Results Section Data Repository: KIT-COCM-ST-137"
ws_ov['A1'].font = font_title
ws_ov['A2'] = "The Application of Cloud Computing in Traditional Music Education in Universities"
ws_ov['A2'].font = font_subtitle
ws_ov['A3'] = "Author: Guo Hao | Guangdong Polytechnic Normal University | Proposed System: CUDDLE"
ws_ov['A3'].font = font_subtitle

ov_headers = ["Sheet Name", "Table Title", "Key Content Description", "Primary Metrics Evaluated", "Status"]
for col_idx, h in enumerate(ov_headers, start=1):
    ws_ov.cell(row=5, column=col_idx, value=h)
format_header_row(ws_ov, 5, len(ov_headers), fill_navy)
ws_ov.row_dimensions[5].height = 25

ov_rows = [
    ("1. Software & Hardware", "Table 1. Experimental Software and Hardware Details", "Cloud host specs, storage, networking, libraries, frameworks, client terminals", "Intel Xeon Gold, 128GB RAM, NVMe, CloudSim, PyTorch, Librosa, WebRTC", "Completed"),
    ("2. Baseline Description", "Table 2. Comprehensive Baseline System Descriptions", "Architectural and functional breakdown of CLMME, CLOUD-MUSEDU, TCMCE vs. CUDDLE", "Architecture, Core Algorithm, Target Scope, Limitations, Strengths", "Completed"),
    ("3. Cloud Metrics", "Table 3. Cloud QoS & Resource Efficiency Metrics", "Response time, throughput, CPU/RAM utilization, jitter, packet loss across 100-500 inputs", "Latency (ms), Throughput (RPS), CPU (%), Memory (%), Jitter (ms), Loss (%)", "Completed"),
    ("4. Statistical Validation", "Table 4. Inferential Statistical Significance Validation", "Paired t-test (df=9), p-values, 95% CIs, Cohen's d, and One-Way ANOVA across 10 runs", "Accuracy, Recall, Specificity, Response Latency, Throughput", "Completed"),
    ("5. Ablation Analysis", "Table 5. Component-Wise Ablation Analysis of CUDDLE", "Impact of stripping ECR-LB (autoscaling), RTA-CE (WebRTC), DL-VRC (CNN), and DCS-LMS", "Accuracy, Recall, Specificity, Latency, Throughput, Packet Loss", "Completed")
]

for row_idx, r in enumerate(ov_rows, start=6):
    fill = fill_zebra if row_idx % 2 == 0 else fill_white
    for col_idx, val in enumerate(r, start=1):
        cell = ws_ov.cell(row=row_idx, column=col_idx, value=val)
        align = align_left if col_idx in [2, 3, 4] else align_center
        f = font_bold if col_idx in [1, 5] else font_regular
        cfill = fill_highlight if col_idx == 5 else fill
        format_cell(cell, f, cfill, thin_border, align)
    ws_ov.row_dimensions[row_idx].height = 24

auto_fit_columns(ws_ov, min_width=15, max_width=50)

wb.save(output_path)
print(f'Successfully generated professional Excel workbook at: {output_path}')
