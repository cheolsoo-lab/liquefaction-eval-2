import io
import re
import traceback
from datetime import date
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.patches import FancyBboxPatch, Polygon

# 검토서(Word) 안에 넣을 차트는 이미지(PNG)로 렌더링해야 하므로, 한글이 깨지지 않도록
# 시스템에 설치된 한글 폰트를 찾아 matplotlib 기본 폰트로 지정한다. 실행 환경(주로 Windows)에
# 아래 폰트 중 하나라도 있으면 정상적으로 표시되며, 하나도 없으면 matplotlib 기본값을 쓴다
# (그 경우 차트의 한글 라벨이 깨질 수 있다 - 표/본문 텍스트는 Word 자체 폰트라 영향 없음).
for _cand in ["Malgun Gothic", "AppleGothic", "NanumGothic", "Noto Sans CJK KR", "Noto Sans CJK JP", "Noto Sans KR"]:
    if any(_cand.lower() in f.name.lower() for f in fm.fontManager.ttflist):
        matplotlib.rcParams["font.family"] = _cand
        break
matplotlib.rcParams["axes.unicode_minus"] = False

# ---------------------------------------------------------
# 0. 지반응답해석(SHAKE 대체) 엔진 - 1차원 등가선형해석
# ---------------------------------------------------------
# 1차원 등가선형 지반응답해석 (SHAKE, Schnabel/Lysmer/Seed 1972 방식) 핵심 엔진.
#
# 이론 개요 (Kramer, 1996, Geotechnical Earthquake Engineering, Ch.7):
# - n개 토층(1..n) + 기반암(반무한 탄성 반공간, layer n+1)으로 구성된 수평 성층 지반에
#   연직 상향 전파하는 전단파(SH-wave)를 가정.
# - 층 m 내부 변위: u_m(z,t) = [A_m*exp(i k*_m z) + B_m*exp(-i k*_m z)] * exp(i w t)
#   (z: 층 상단 기준 국소좌표, A: 상향파, B: 하향파 진폭)
# - 복소전단탄성계수 G*_m = G_m(1 + 2i*beta_m), 복소전단파속도 Vs*_m = sqrt(G*_m/rho_m),
#   복소파수 k*_m = w / Vs*_m
# - 지표면(층1 상단) 전단응력=0 경계조건 -> A_1 = B_1
# - 층 경계면에서 변위/전단응력 연속조건으로 아래 점화식 유도(Kramer eq. 7.18~7.19):
#     A_(m+1) = 0.5*A_m*(1+a_m)*exp(i k*_m h_m) + 0.5*B_m*(1-a_m)*exp(-i k*_m h_m)
#     B_(m+1) = 0.5*A_m*(1-a_m)*exp(i k*_m h_m) + 0.5*B_m*(1+a_m)*exp(-i k*_m h_m)
#   여기서 a_m = (rho_m*Vs*_m)/(rho_(m+1)*Vs*_(m+1)) (임피던스비)
# - 입력파는 기반암 노두운동(outcrop motion)으로 가정(암반이 지표에 노출된 경우의 자유면 운동,
#   즉 기반암 내 상향파 진폭의 2배). 최하층(기반암)에서 A_(n+1)을 구하면
#   스케일 팩터 = U_oc(w) / (2*A_(n+1)(w)) 이고, 임의 층 m 상단의 운동은
#   U_m(w) = 스케일팩터 * (A_m+B_m) 로 계산된다.
# - 등가선형 반복계산: 각 층의 변형률 시간이력에서 최대변형률 x 0.65(SHAKE 기본값)를
#   유효변형률로 삼아 G/Gmax-변형률, 감쇠비-변형률 곡선에서 새 물성을 구하고, 물성이
#   수렴할 때까지(통상 5~10% 이내) 반복한다.


def _complex_shear_modulus(g_max, damping_ratio):
    """G* = G(1+2i*beta) (근사식, SHAKE에서 사용하는 표준 근사)."""
    return g_max * (1.0 + 2.0j * damping_ratio)


def site_response_frequency_domain(
    omega,
    thickness,
    rho,
    g_complex,
    bedrock_rho,
    bedrock_g_complex,
):
    """각 각속도 omega(1D array, rad/s)에 대해 지표~기반암 상단까지 n+1개 지점의
    "기반암 노두운동 대비 전달함수"를 계산. g_complex: shape (n_freq, n) 또는 (n,) 허용.
    반환: H (n_freq, n+1) 복소배열, A_top(n_freq, n+1), B_top(n_freq, n+1) (각 층 상단 상향/하향파
    진폭, 스케일 적용 전 raw 값 - strain 계산에 필요)
    """
    n = len(thickness)
    n_freq = len(omega)

    if g_complex.ndim == 1:
        g_complex = np.tile(g_complex, (n_freq, 1))

    vs_c = np.sqrt(g_complex / rho[None, :])  # (n_freq, n)
    vs_c_bed = np.sqrt(bedrock_g_complex / bedrock_rho)  # scalar or (n_freq,)
    if np.isscalar(vs_c_bed) or np.ndim(vs_c_bed) == 0:
        vs_c_bed = np.full(n_freq, vs_c_bed, dtype=complex)

    # 각 층 파수 k*_m = omega / Vs*_m  (omega=0 이면 k=0으로 처리)
    with np.errstate(divide="ignore", invalid="ignore"):
        k_c = omega[:, None] / vs_c  # (n_freq, n)
    k_c = np.where(omega[:, None] == 0, 0.0 + 0.0j, k_c)

    A = np.ones(n_freq, dtype=complex)
    B = np.ones(n_freq, dtype=complex)

    A_top = np.zeros((n_freq, n + 1), dtype=complex)
    B_top = np.zeros((n_freq, n + 1), dtype=complex)
    A_top[:, 0] = A
    B_top[:, 0] = B

    for m in range(n):
        if m < n - 1:
            rho_next = rho[m + 1]
            vs_next = vs_c[:, m + 1]
        else:
            rho_next = bedrock_rho
            vs_next = vs_c_bed

        imped_num = rho[m] * vs_c[:, m]
        imped_den = rho_next * vs_next
        alpha = imped_num / imped_den

        km_hm = k_c[:, m] * thickness[m]
        exp_p = np.exp(1j * km_hm)
        exp_m = np.exp(-1j * km_hm)

        A_next = 0.5 * A * (1 + alpha) * exp_p + 0.5 * B * (1 - alpha) * exp_m
        B_next = 0.5 * A * (1 - alpha) * exp_p + 0.5 * B * (1 + alpha) * exp_m

        A, B = A_next, B_next
        A_top[:, m + 1] = A
        B_top[:, m + 1] = B

    # 기반암 노두운동 대비 전달함수: scale(w) = 1/(2*A_bedrock) 이므로
    # H_m(w) = (A_top[:,m]+B_top[:,m]) / (2*A_top[:,n])
    denom = 2.0 * A_top[:, n]
    denom = np.where(np.abs(denom) < 1e-30, 1e-30 + 0j, denom)
    H = (A_top + B_top) / denom[:, None]

    return H, A_top, B_top, k_c


def _interp_curve(strain_pct, curve_strain_pct, curve_value):
    """변형률(%)에 대해 G/Gmax 또는 감쇠비(%) 곡선을 로그-선형 보간. 범위를 벗어나면
    양 끝단 값으로 고정(외삽 금지, SHAKE의 일반적 처리 방식)."""
    strain_pct = np.clip(strain_pct, curve_strain_pct[0], curve_strain_pct[-1])
    log_strain = np.log10(np.maximum(strain_pct, 1e-6))
    log_curve_strain = np.log10(curve_strain_pct)
    return np.interp(log_strain, log_curve_strain, curve_value)


DEFAULT_CURVES = {
    # Seed & Idriss (1970) 평균 모래 곡선 (변형률 %, G/Gmax, 감쇠비 %)
    "sand": {
        "strain_pct": np.array([0.0001, 0.0003, 0.001, 0.003, 0.01, 0.03, 0.1, 0.3, 1.0, 3.0]),
        "g_gmax": np.array([1.00, 0.99, 0.96, 0.90, 0.79, 0.63, 0.45, 0.28, 0.14, 0.06]),
        "damping_pct": np.array([0.5, 0.6, 1.0, 1.5, 3.0, 6.0, 10.0, 15.5, 21.0, 25.0]),
    },
    # 암반 - 변형률에 거의 무관 (거의 선형탄성)
    "rock": {
        "strain_pct": np.array([0.0001, 0.001, 0.01, 0.1, 1.0, 3.0]),
        "g_gmax": np.array([1.00, 1.00, 0.99, 0.97, 0.93, 0.90]),
        "damping_pct": np.array([0.5, 0.5, 0.6, 0.8, 1.2, 1.5]),
    },
}

# Vucetic & Dobry (1991) 방식 - 점토는 PI(소성지수)에 따라 G/Gmax-변형률·감쇠비-변형률
# 곡선이 달라진다(PI가 클수록 같은 변형률에서 강성 저하가 적고 감쇠도 작다). 문헌의
# 대표 PI값(0/15/30/50/100/200)별 근사 곡선을 아래에 두고, 실제 층의 PI는 이 값들
# 사이를 선형보간해서 사용한다. 프로젝트 자체 실내시험 곡선이 있다면 curves 인자로
# 교체할 수 있다.
CLAY_STRAIN_PCT = np.array([0.0001, 0.0003, 0.001, 0.003, 0.01, 0.03, 0.1, 0.3, 1.0, 3.0])
CLAY_CURVES_BY_PI = {
    0: {
        "strain_pct": CLAY_STRAIN_PCT,
        "g_gmax": np.array([1.00, 0.99, 0.96, 0.90, 0.78, 0.60, 0.38, 0.20, 0.08, 0.03]),
        "damping_pct": np.array([1.0, 1.5, 2.5, 4.0, 7.0, 12.0, 18.0, 24.0, 29.0, 32.0]),
    },
    15: {
        "strain_pct": CLAY_STRAIN_PCT,
        "g_gmax": np.array([1.00, 0.99, 0.97, 0.92, 0.82, 0.68, 0.47, 0.28, 0.13, 0.06]),
        "damping_pct": np.array([1.0, 1.3, 2.1, 3.3, 5.7, 9.5, 14.5, 20.0, 25.0, 28.0]),
    },
    30: {
        "strain_pct": CLAY_STRAIN_PCT,
        "g_gmax": np.array([1.00, 1.00, 0.98, 0.94, 0.86, 0.74, 0.55, 0.36, 0.18, 0.09]),
        "damping_pct": np.array([1.0, 1.1, 1.7, 2.7, 4.5, 7.5, 11.5, 16.0, 20.5, 24.0]),
    },
    50: {
        "strain_pct": CLAY_STRAIN_PCT,
        "g_gmax": np.array([1.00, 1.00, 0.99, 0.96, 0.90, 0.80, 0.64, 0.45, 0.25, 0.14]),
        "damping_pct": np.array([1.0, 1.0, 1.4, 2.1, 3.5, 5.8, 9.0, 13.0, 17.0, 20.0]),
    },
    100: {
        "strain_pct": CLAY_STRAIN_PCT,
        "g_gmax": np.array([1.00, 1.00, 0.99, 0.97, 0.93, 0.86, 0.73, 0.56, 0.35, 0.21]),
        "damping_pct": np.array([1.0, 1.0, 1.1, 1.6, 2.6, 4.3, 6.8, 10.0, 13.5, 16.5]),
    },
    200: {
        "strain_pct": CLAY_STRAIN_PCT,
        "g_gmax": np.array([1.00, 1.00, 1.00, 0.99, 0.96, 0.91, 0.82, 0.68, 0.48, 0.32]),
        "damping_pct": np.array([1.0, 1.0, 1.0, 1.2, 1.8, 3.0, 4.8, 7.2, 10.0, 12.5]),
    },
}
DEFAULT_CLAY_PI = 30.0  # PI 정보가 없는 점토층에 적용할 기본값


def get_clay_curve_for_pi(pi):
    """소성지수(PI)로 Vucetic&Dobry(1991) 곡선을 선형보간해서 반환. PI가 없으면(NaN)
    DEFAULT_CLAY_PI(=30)를 대표값으로 사용한다."""
    if pi is None or (isinstance(pi, float) and np.isnan(pi)):
        pi = DEFAULT_CLAY_PI
    keys = sorted(CLAY_CURVES_BY_PI.keys())
    if pi <= keys[0]:
        return CLAY_CURVES_BY_PI[keys[0]]
    if pi >= keys[-1]:
        return CLAY_CURVES_BY_PI[keys[-1]]
    lo = max(k for k in keys if k <= pi)
    hi = min(k for k in keys if k >= pi)
    if lo == hi:
        return CLAY_CURVES_BY_PI[lo]
    w = (pi - lo) / (hi - lo)
    lo_c, hi_c = CLAY_CURVES_BY_PI[lo], CLAY_CURVES_BY_PI[hi]
    return {
        "strain_pct": lo_c["strain_pct"],
        "g_gmax": lo_c["g_gmax"] * (1 - w) + hi_c["g_gmax"] * w,
        "damping_pct": lo_c["damping_pct"] * (1 - w) + hi_c["damping_pct"] * w,
    }


def run_equivalent_linear(
    thickness,           # (n,) m
    unit_weight,         # (n,) kN/m3
    vs_initial,          # (n,) m/s, 초기(저변형률) 전단파속도
    material_key,        # (n,) 각 층의 곡선 키 (DEFAULT_CURVES의 키)
    bedrock_unit_weight,
    bedrock_vs,
    dt,
    accel_g,             # (npts,) 기반암 노두운동 가속도(단위: g)
    curves=None,
    pi_values=None,      # (n,) 점토층 소성지수(PI)
    strain_ratio=0.65,   # 유효변형률비 (SHAKE 기본 0.65, Idriss&Sun: (M-1)/10)
    curve_shift=1.0,     # 동적곡선 기준변형률 배율 (>1: 선형성 증가(상한), <1: 비선형성 증가(하한))
    max_iter=15,
    tol=0.05,
    initial_damping=0.02,
    g_grav=9.81,
    freq_max=25.0,
):
    """등가선형 1차원 지반응답해석 (SHAKE 방식, 벡터화 버전).
    반복계산 중에는 전 층의 변형률 시간이력을 한 번의 역FFT로 구하고, 전달함수는
    유효 주파수(freq_max 이하) 대역에서만 계산한다. 결과는 기존 구현과 동일하다.
    반환: amax_g (지표~기반암상단 n+1개 지점), accel_time (지표 1개 지점, shape (1, npts)) 등
    """
    if curves is None:
        curves = DEFAULT_CURVES
    n = len(thickness)
    if pi_values is None:
        pi_values = np.full(n, np.nan)
    thickness = np.asarray(thickness, dtype=float)
    vs_initial = np.asarray(vs_initial, dtype=float)
    rho = np.asarray(unit_weight, dtype=float) / g_grav
    bedrock_rho = bedrock_unit_weight / g_grav
    accel_g = np.asarray(accel_g, dtype=float)
    npts = len(accel_g)
    freqs = np.fft.rfftfreq(npts, d=dt)
    omega = 2 * np.pi * freqs
    valid = freqs <= freq_max
    om_v = omega[valid]
    accel_fft = np.fft.rfft(accel_g)
    with np.errstate(divide="ignore", invalid="ignore"):
        # 입력 가속도가 g 단위이므로 변위(m)로 바꿀 때 중력가속도를 곱해야 변형률이 올바른 크기가 된다
        disp_fft = np.where(omega == 0, 0.0, -accel_fft * g_grav / np.where(omega == 0, 1.0, omega ** 2))

    mats = [
        get_clay_curve_for_pi(pi_values[m]) if material_key[m] == "clay"
        else curves.get(material_key[m], curves["sand"])
        for m in range(n)
    ]
    g0 = rho * vs_initial ** 2
    g_max = g0.copy()
    damping = np.full(n, initial_damping)
    bedrock_gc = np.full(len(om_v), bedrock_rho * bedrock_vs ** 2 * (1 + 2j * initial_damping))
    zc = thickness / 2.0

    def _solve(gm, dmp):
        gc = np.tile(gm * (1 + 2j * dmp), (len(om_v), 1))
        return site_response_frequency_domain(om_v, thickness, rho, gc, bedrock_rho, bedrock_gc)

    history = []
    for it in range(max_iter):
        H, A_top, B_top, k_c = _solve(g_max, damping)
        denom = 2.0 * A_top[:, n:n + 1]
        strain_tf = 1j * k_c * (A_top[:, :n] * np.exp(1j * k_c * zc) - B_top[:, :n] * np.exp(-1j * k_c * zc)) / denom
        spec = np.zeros((len(omega), n), dtype=complex)
        spec[valid] = strain_tf * disp_fft[valid][:, None]
        gamma_max = np.max(np.abs(np.fft.irfft(spec, n=npts, axis=0)), axis=0)
        gamma_eff_pct = strain_ratio * gamma_max * 100.0 / curve_shift
        new_g = np.empty(n)
        new_d = np.empty(n)
        for m in range(n):
            new_g[m] = g0[m] * _interp_curve(gamma_eff_pct[m], mats[m]["strain_pct"], mats[m]["g_gmax"])
            new_d[m] = _interp_curve(gamma_eff_pct[m], mats[m]["strain_pct"], mats[m]["damping_pct"]) / 100.0
        max_change = float(max(
            np.max(np.abs(new_g - g_max) / np.maximum(g_max, 1e-9)),
            np.max(np.abs(new_d - damping) / np.maximum(damping, 1e-9)),
        ))
        history.append({"iter": it + 1, "max_change": max_change})
        g_max, damping = new_g, new_d
        if max_change < tol:
            break

    H, A_top, B_top, k_c = _solve(g_max, damping)
    Hfull = np.ones((len(omega), n + 1), dtype=complex)
    Hfull[valid] = H
    acc = np.fft.irfft(Hfull * accel_fft[:, None], n=npts, axis=0)  # (npts, n+1)
    amax_g = np.max(np.abs(acc), axis=0)
    # 층 상단 최대전단응력 τmax (kPa) - CSR = 0.65·(τmax)d,GRA/σv' (KDS 64 17 00:2024, 구조계산서 표기)
    # τ = G*·∂u/∂z, 층 m 상단(국부좌표 z=0)에서 ∂u/∂z = i·k*(A−B), 기반암 상단은 마지막 층 하단(z=h)
    gc_fin = g_max * (1 + 2j * damping)
    denom = 2.0 * A_top[:, n:n + 1]
    tau_tf = np.zeros((len(om_v), n + 1), dtype=complex)
    tau_tf[:, :n] = gc_fin[None, :] * 1j * k_c * (A_top[:, :n] - B_top[:, :n]) / denom
    kh = k_c[:, n - 1] * thickness[n - 1]
    tau_tf[:, n] = (gc_fin[n - 1] * 1j * k_c[:, n - 1]
                    * (A_top[:, n - 1] * np.exp(1j * kh) - B_top[:, n - 1] * np.exp(-1j * kh)) / denom[:, 0])
    tau_spec = np.zeros((len(omega), n + 1), dtype=complex)
    tau_spec[valid] = tau_tf * disp_fft[valid][:, None]
    tau_max = np.max(np.abs(np.fft.irfft(tau_spec, n=npts, axis=0)), axis=0)
    return {
        "amax_g": amax_g,
        "tau_max_kpa": tau_max,
        "accel_time": acc[:, :1].T,  # 지표면 가속도 시간이력만 보관(메모리 절약)
        "g_max_final": g_max,
        "damping_final": damping,
        "n_iter": len(history),
        "converged": history[-1]["max_change"] < tol if history else True,
        "history": history,
    }


def sdof_response_spectrum(accel, dt, periods, damping=0.05):
    """Nigam-Jennings 정확해로 여러 주기(periods)에 대한 의사가속도스펙트럼 Sa(T)를 계산.
    accel: 지반가속도 시간이력 (임의 단위, 보통 g). 반환 Sa도 동일 단위.
    """
    accel = np.asarray(accel, dtype=float)
    npts = len(accel)
    sa = np.zeros(len(periods))

    for ip, T in enumerate(periods):
        wn = 2 * np.pi / T
        xi = damping
        wd = wn * np.sqrt(1 - xi**2)
        ewt = np.exp(-xi * wn * dt)
        swd = np.sin(wd * dt)
        cwd = np.cos(wd * dt)
        sq = np.sqrt(1 - xi**2)

        A = ewt * ((xi / sq) * swd + cwd)
        B = ewt * (swd / wd)
        C = (1.0 / wn**2) * (
            (2 * xi) / (wn * dt)
            + ewt * (((1 - 2 * xi**2) / (wd * dt) - xi / sq) * swd - (1 + (2 * xi) / (wn * dt)) * cwd)
        )
        D = (1.0 / wn**2) * (
            1 - (2 * xi) / (wn * dt)
            + ewt * (((2 * xi**2 - 1) / (wd * dt)) * swd + ((2 * xi) / (wn * dt)) * cwd)
        )
        Ap = -ewt * ((wn / sq) * swd)
        Bp = ewt * (cwd - (xi / sq) * swd)
        Cp = (1.0 / wn**2) * (
            -1.0 / dt + ewt * ((wn / sq + xi / (dt * sq)) * swd + (1.0 / dt) * cwd)
        )
        Dp = (1.0 / (wn**2 * dt)) * (1 - ewt * ((xi / sq) * swd + cwd))

        u = np.zeros(npts)
        v = np.zeros(npts)
        for i in range(npts - 1):
            p_i = -accel[i]
            p_i1 = -accel[i + 1]
            u[i + 1] = A * u[i] + B * v[i] + C * p_i + D * p_i1
            v[i + 1] = Ap * u[i] + Bp * v[i] + Cp * p_i + Dp * p_i1

        sa[ip] = wn**2 * np.max(np.abs(u))

    return sa


# KDS 17 10 00:2024 표 4.2-1(단주기 Fa) / 표 4.2-2(1초주기 Fv) 지반증폭계수.
# 유효지반가속도 s = 0.1 이하 / 0.2 / 0.3 기준값이며, 중간값은 직선보간한다.
KDS_FA_TABLE = {
    "S1": [1.12, 1.12, 1.12],
    "S2": [1.4, 1.4, 1.3],
    "S3": [1.7, 1.5, 1.3],
    "S4": [1.6, 1.4, 1.2],
    "S5": [1.8, 1.3, 1.3],
}
KDS_FV_TABLE = {
    "S1": [0.84, 0.84, 0.84],
    "S2": [1.5, 1.4, 1.3],
    "S3": [1.7, 1.6, 1.5],
    "S4": [2.2, 2.0, 1.8],
    "S5": [3.0, 2.7, 2.4],
}
_KDS_S_POINTS = [0.1, 0.2, 0.3]


def get_kds_site_factors(soil_class, S):
    """지반종류와 유효수평지반가속도 S로 Fa, Fv를 표에서 직선보간해 구한다."""
    fa_row = KDS_FA_TABLE.get(soil_class, KDS_FA_TABLE["S1"])
    fv_row = KDS_FV_TABLE.get(soil_class, KDS_FV_TABLE["S1"])
    s_clipped = min(max(S, _KDS_S_POINTS[0]), _KDS_S_POINTS[-1])
    fa = float(np.interp(s_clipped, _KDS_S_POINTS, fa_row))
    fv = float(np.interp(s_clipped, _KDS_S_POINTS, fv_row))
    return fa, fv


def kds_design_spectrum(periods, S, soil_class="S1", TL=3.0):
    """KDS 17 10 00:2024 5%감쇠 가속도 표준설계응답스펙트럼 Sa(T) [g].

    S: 유효수평지반가속도(g) = 지진구역계수(Z) × 위험도계수(I).
       (S에 이미 위험도계수가 반영되어 있으므로 여기서 I를 다시 곱하면 안 된다.)
    soil_class: 지반종류(S1~S5). 단주기 스펙트럼증폭계수 αA = 2.5×Fa 이며,
       전이주기는 Ts = Fv/(2.5×Fa), T0 = 0.2×Ts 로 산정한다.
       S1(암반)의 경우 Fa=1.12, Fv=0.84 → αA=2.8, Ts=0.30s, T0=0.06s 로
       KDS 암반지반 표준값과 정확히 일치한다.
    """
    periods = np.asarray(periods, dtype=float)
    sa = np.zeros_like(periods)
    fa, fv = get_kds_site_factors(soil_class, S)
    alpha_a = 2.5 * fa
    ts = fv / alpha_a
    t0 = 0.2 * ts
    for i, T in enumerate(periods):
        if T <= t0:
            # 암반(S1): (1+30T)·S, 토사(S2~S5): T=0에서 S·Fa → T0에서 2.5·Fa·S 직선
            sa0 = S if soil_class == "S1" else S * fa
            sa[i] = sa0 + (S * alpha_a - sa0) * (T / t0)
        elif T <= ts:
            sa[i] = S * alpha_a
        elif T <= TL:
            sa[i] = S * alpha_a * ts / T
        else:
            sa[i] = S * alpha_a * ts * TL / (T**2)
    return sa


# ---------------------------------------------------------
# 0-1. KDS 17 10 00:2024 설계지진 조건 및 지진응답해석 결과 검토
# ---------------------------------------------------------
KDS_ZONE_FACTOR = {"지진구역 Ⅰ (Z=0.11g)": 0.11, "지진구역 Ⅱ (Z=0.07g)": 0.07}   # 표 4.2-2
KDS_RISK_FACTOR = {"50년": 0.40, "100년": 0.57, "200년": 0.73, "500년": 1.00,       # 표 4.2-3
                   "1,000년": 1.40, "2,400년": 2.00, "4,800년": 2.60}
# 설계지진 규모. KDS SPT 방법의 MSF = 1 + 0.376(MSFmax−1)의 0.376은 Idriss & Boulanger(2014)의
# 8.64·exp(−M/4) − 1.325 에 M = 6.5를 넣은 값이므로, 응력감소계수 rd도 같은 규모를 쓴다.
DESIGN_M = 6.5


def idriss_rd(z, M=DESIGN_M):
    """응력감소계수 rd (Idriss 1999, Idriss & Boulanger 2008). z: 심도(m).
    KDS 17 10 00:2024 식(4.7-2) 단서의 '지표면 최대지반가속도와 응력감소계수' 방법에 사용."""
    z = np.asarray(z, dtype=float)
    a = -1.012 - 1.126 * np.sin(z / 11.73 + 5.133)
    b = 0.106 + 0.118 * np.sin(z / 11.28 + 5.142)
    return np.where(z <= 34.0, np.exp(a + b * M), 0.12 * np.exp(0.22 * M))


def kds_sfa(S, soil_class):
    """KDS 4.7(6)① 단서: 지표면 최대지반가속도 하한 = S × Fa (기반암 유효수평지반가속도 × 단주기지반증폭계수)."""
    fa, _ = get_kds_site_factors(soil_class, S)
    return S * fa, fa


def surface_pga_table(df_amax, wave_cols, S, soil_class):
    """시추공별 지표면 최대가속도(지진파 평균)와 KDS 하한(S×Fa) 비교표.
    지표면 값은 시추공의 최상단 층(심도 최소) amax를 사용한다."""
    sfa, fa = kds_sfa(S, soil_class)
    rows = []
    wc = [w for w in wave_cols if w in df_amax.columns]
    for bh, g in df_amax.groupby("Borehole", sort=False):
        top = g.sort_values("심도(GL,m)").iloc[0]
        vals = np.array([float(top[w]) for w in wc if pd.notna(top[w])], dtype=float)
        if len(vals) == 0:
            continue
        mean = float(vals.mean())
        rows.append({
            "Borehole": bh, "지표 PGA 평균(g)": mean, "지표 PGA 최소(g)": float(vals.min()),
            "지표 PGA 최대(g)": float(vals.max()), "S(g)": S, "Fa": fa, "하한 S×Fa(g)": sfa,
            "평균/하한": mean / sfa if sfa > 0 else np.nan,
            "하한 미달 지진파 수": int((vals < sfa - 1e-9).sum()), "지진파 수": int(len(vals)),
            "판정": "만족" if mean >= sfa - 1e-9 else "미달",
        })
    return pd.DataFrame(rows)


def _style_map(styler, func, subset=None):
    """pandas 버전과 무관하게 셀 단위 스타일 적용 (2.1 이상 map, 미만 applymap)."""
    f = getattr(styler, "map", None) or getattr(styler, "applymap")
    return f(func, subset=subset)


def kds_rock_design_spectrum(periods, S, **kwargs):
    """하위호환용 래퍼 - 암반지반(S1) 표준설계응답스펙트럼."""
    return kds_design_spectrum(periods, S, soil_class="S1")


def _corr_pair(a, dt_a, b, dt_b):
    """샘플링 간격이 다른 두 시간이력을 공통 시간축으로 리샘플링한 뒤 상관계수를 구한다."""
    dur = min(len(a) * dt_a, len(b) * dt_b)
    dt_common = min(dt_a, dt_b)
    n = max(int(dur / dt_common), 2)
    t_common = np.arange(n) * dt_common
    ta = np.arange(len(a)) * dt_a
    tb = np.arange(len(b)) * dt_b
    a_r = np.interp(t_common, ta, a)
    b_r = np.interp(t_common, tb, b)
    return float(np.corrcoef(a_r, b_r)[0, 1])


def check_ground_motion_suite(
    wave_dict, dt_per_wave, S, I, periods=None, damping=0.05,
    artificial=True, T0=0.06, TL=3.0, corr_limit=0.16, soil_class="S1",
):
    """여러 지진파(wave_dict: {이름: 가속도시간이력(g)}, dt_per_wave: {이름: 시간간격(s)})에
    대해 KDS 17 10 00:2024 적합성 기준을 종합 검토한다. artificial=True면 인공합성파
    기준(90~130%), False면 실지진기록 기준(90% 이상만)을 적용한다. 지진파마다 dt가
    달라도 되며, 상관계수 계산 시에는 공통 시간축으로 리샘플링해서 비교한다.

    soil_class: 비교 대상 표준설계응답스펙트럼의 지반종류.
      KDS 17 10 00:2024 4.2.1.4(9)③·(10)①에 따라 "기반암에 입력하는" 시간이력은
      S1(암반) 표준설계응답스펙트럼에 맞추어 수정/합성하므로 입력파형 검증은 S1을 쓴다.
      반면 같은 조 (10)⑤는 "지반응답해석으로 계산된 S2~S6 지반의" 결과(지표면 운동)를
      해당 지반종류의 표준설계응답스펙트럼과 비교하도록 규정하므로, 지표면 운동 검증에는
      실제 부지의 지반종류(S2~S5)를 넘겨 쓴다.
    반환: (개별 스펙트럼 dict, 평균 스펙트럼, 목표스펙트럼, 판정 dict)
    """
    if periods is None:
        periods = np.concatenate(
            [np.arange(0.02, 0.2, 0.02), np.arange(0.2, 1.0, 0.05), np.arange(1.0, 4.01, 0.1)]
        )

    target = kds_design_spectrum(periods, S, soil_class=soil_class)
    spectra = {}
    for name, accel in wave_dict.items():
        spectra[name] = sdof_response_spectrum(accel, dt_per_wave[name], periods, damping=damping)

    avg_spectrum = np.mean(list(spectra.values()), axis=0)

    mask = (periods >= T0) & (periods <= TL)
    ratio = np.where(target[mask] > 0, avg_spectrum[mask] / target[mask], np.nan)
    min_ratio_pct = float(np.nanmin(ratio) * 100)
    max_ratio_pct = float(np.nanmax(ratio) * 100)

    pgas = {name: float(np.max(np.abs(a))) for name, a in wave_dict.items()}
    avg_pga = float(np.mean(list(pgas.values())))
    pga_ratio_pct = avg_pga / S * 100.0

    names = list(wave_dict.keys())
    corr_matrix = {}
    max_corr = 0.0
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            c = _corr_pair(
                wave_dict[names[i]], dt_per_wave[names[i]],
                wave_dict[names[j]], dt_per_wave[names[j]],
            )
            corr_matrix[(names[i], names[j])] = c
            max_corr = max(max_corr, abs(c))

    lower_ok = min_ratio_pct >= 90.0
    upper_ok = (max_ratio_pct <= 130.0) if artificial else True
    pga_ok = pga_ratio_pct >= 90.0
    corr_ok = max_corr <= corr_limit

    judgement = {
        "평균스펙트럼_최소비율(%)": round(min_ratio_pct, 1),
        "평균스펙트럼_최대비율(%)": round(max_ratio_pct, 1),
        "스펙트럼_하한기준(90%)_만족": lower_ok,
        "스펙트럼_상한기준(130%, 인공합성파만)_만족": upper_ok,
        "평균PGA(g)": round(avg_pga, 4),
        "평균PGA_비율(%)": round(pga_ratio_pct, 1),
        "PGA기준(90%)_만족": pga_ok,
        "최대_상관계수": round(max_corr, 4),
        "상관계수기준(0.16이하)_만족": corr_ok,
        "종합판정": "적합" if (lower_ok and upper_ok and pga_ok and corr_ok) else "부적합(하단 세부기준 확인)",
    }

    return periods, spectra, avg_spectrum, target, judgement, corr_matrix


def compute_scale_factor(judgement, artificial):
    """KDS 17 10 00:2024 4.2.1.4(9)④·(10)⑥⑦에 따른 '스펙트럼보정(스케일링)' 계수 산정.
    선형계로 가정하면 시간이력 진폭을 k배 하면 응답스펙트럼과 PGA도 그대로 k배가 되므로,
    하한기준(평균스펙트럼 ≥ 목표의 90%, 평균PGA ≥ S의 90%)을 만족시키는 데 필요한 최소
    배율을 구한다. 인공합성파는 상한(130%)도 있어서, 하한을 만족시키는 배율이 상한을
    넘기면 상한에 맞춰 절충한다(이 경우 하한을 완전히 만족 못할 수 있어 그 사실을 함께
    반환한다). 상관계수 기준은 파형의 진폭이 아니라 파형 자체의 모양(위상)에 의해 결정되는
    값이라 스케일링으로는 개선할 수 없다 - 그 경우는 파형 조합을 다시 선정해야 한다.
    """
    min_ratio = judgement["평균스펙트럼_최소비율(%)"] / 100.0
    max_ratio = judgement["평균스펙트럼_최대비율(%)"] / 100.0
    pga_ratio = judgement["평균PGA_비율(%)"] / 100.0

    k_lower = (0.90 / min_ratio) if min_ratio > 0 else 1.0
    k_pga = (0.90 / pga_ratio) if pga_ratio > 0 else 1.0
    k = max(1.0, k_lower, k_pga)

    capped = False
    if artificial and max_ratio > 0:
        k_upper_limit = 1.30 / max_ratio
        if k > k_upper_limit:
            k = max(k_upper_limit, 1e-6)
            capped = True

    return round(k, 4), capped


def scale_wave_dict(wave_dict, k):
    return {name: np.asarray(accel, dtype=float) * k for name, accel in wave_dict.items()}


def classify_material_key(soil_name):
    """토층 명칭으로 동적물성곡선(모래/점토/암반) 자동 분류. '재료유형' 컬럼을
    직접 지정하면 그 값이 우선하고, 비어 있을 때만 이 자동판정을 사용한다.
    순서 중요: '실트질 모래'처럼 모래/사질 키워드가 있으면 모래로 우선 분류하고,
    그 외 세립토 계열(점토/실트/CL/CH 등)만 점토로, 암반 계열은 암반으로 분류한다."""
    s = str(soil_name)
    rock_kw = ["암", "riprap", "Rock", "ROCK"]
    sand_kw = ["사질", "모래", "매립", "SM", "SP", "SW", "SC", "SD", "SF"]
    clay_kw = ["점토", "실트", "니토", "CL", "CH", "ML", "MH", "clay", "Clay"]
    if any(kw in s for kw in rock_kw):
        return "rock"
    if any(kw in s for kw in sand_kw):
        return "sand"
    if any(kw in s for kw in clay_kw):
        return "clay"
    return "sand"


def run_site_response_all_boreholes(
    df_raw,
    wave_dict,
    dt_per_wave,
    bedrock_vs_default,
    bedrock_gamma_default,
    gw_col="지하수위",
    bedrock_vs_criterion=760.0,
    strain_ratio=0.65,
    curve_shift=1.0,
    vs_factor=1.0,
    progress_cb=None,
    store_surface=False,
):
    """해석 케이스 파라미터: strain_ratio(유효변형률비), curve_shift(동적곡선 기준변형률 배율),
    vs_factor(토층 전단파속도 배율 - 지반응답해석에만 적용, Vs법 CRR에는 실측 Vs 사용).
    progress_cb(완료 시추공 수, 전체 시추공 수)로 진행률을 알린다.

    df_raw(기존 입력엑셀, 지하수위/토층/심도/층두께/γsat/γt/Vs 컬럼 보유)의 각 시추공에 대해
    등가선형 지반응답해석을 실행하고, wave_dict의 각 지진파 이름으로 amax/g 컬럼을
    df_raw에 채워 반환한다(이후 run_liquefaction_eval은 이 컬럼들을 그대로 사용).
    wave_dict: {지진파이름: 가속도시간이력 배열(g)}, dt_per_wave: {지진파이름: 시간간격(s)}
    (지진파마다 샘플링 간격이 달라도 되도록 지진파별 dt를 따로 받는다).

    기반암 선정: KDS 17 10 00:2024는 "기반암(bed rock): 전단파속도가 760m/s 이상인 지층"
    으로 정의한다. 따라서 시추주상도의 '가장 깊은 층'을 무조건 기반암으로 쓰지 않고,
    지표에서부터 내려가며 Vs(m/s) ≥ bedrock_vs_criterion(기본 760)을 처음 만족하는 층을
    찾아 그 층 상단을 기반암 상단으로 삼는다. 그 층(및 그 아래)의 물성으로 반무한
    탄성 반공간을 정의하고, 그 위의 층들만 등가선형해석의 유한 토층으로 사용한다.
    시추공별 '기반암Vs(m/s)'·'기반암단위중량(kN/m3)' 컬럼을 입력엑셀에 직접 넣으면
    이 자동판정보다 우선 적용된다(예: 시추공이 기반암까지 도달하지 못해 별도로 정의된
    기준암반 조건을 써야 하는 경우). 어느 층도 760m/s에 도달하지 못하면 자동판정을
    포기하고 전역 기본값(bedrock_vs_default/bedrock_gamma_default)을 사용하며, 로그에
    '기반암미도달'로 표시한다(시추가 기반암까지 도달하지 못했다는 뜻이므로 실제로는
    추가 시추 검토가 필요할 수 있다).
    """
    df = df_raw.copy()
    wave_names = list(wave_dict.keys())
    for w in wave_names:
        if w not in df.columns:
            df[w] = np.nan

    site_response_log = []
    surface_motions = {}

    _groups = list(df.groupby("Borehole", sort=False).groups.items())
    for _bi, (bh, idx) in enumerate(_groups):
        if progress_cb is not None:
            progress_cb(_bi, len(_groups))
        # 주의: 여기서 reset_index를 하면 원본 df의 행 번호를 잃어버려, 계산 결과(amax/g)를
        # df에 되돌려 쓸 때 엉뚱한 행(앞선 시추공의 행)에 써지게 된다. 원본 인덱스를 반드시
        # 보존해야 시추공이 여러 개일 때도 각자 자기 행에 결과가 들어간다.
        g_full = df.loc[idx].sort_values(by="심도(GL,m)")

        gw_mode_list = [_parse_gw_flag(v)[0] for v in g_full[gw_col]]
        _gw_parsed = [_parse_gw_flag(v) for v in g_full[gw_col]]
        _mid = (g_full["심도(GL,m)"] + g_full["층두께(m)"] / 2.0).to_numpy(dtype=float)
        is_sat_list = [m == "submerged" or (m == "depth" and _mid[k] > dpt)
                       for k, (m, dpt) in enumerate(_gw_parsed)]
        unit_weight_full = np.array(
            [
                g_full["γsat(kN/m3)"].iloc[i] if is_sat_list[i] else g_full["γt(kN/m3)"].iloc[i]
                for i in range(len(g_full))
            ],
            dtype=float,
        )
        vs_full = g_full["Vs(m/s)"].to_numpy(dtype=float)

        has_override = (
            "기반암Vs(m/s)" in g_full.columns and g_full["기반암Vs(m/s)"].notna().any()
        )
        bedrock_note = ""
        if has_override:
            bedrock_vs = float(g_full["기반암Vs(m/s)"].dropna().iloc[0])
            bedrock_gamma = (
                float(g_full["기반암단위중량(kN/m3)"].dropna().iloc[0])
                if "기반암단위중량(kN/m3)" in g_full.columns
                and g_full["기반암단위중량(kN/m3)"].notna().any()
                else bedrock_gamma_default
            )
            g = g_full  # 전체 층을 유한 토층으로 사용(기반암은 그 아래 별도 조건)
            bedrock_note = "입력엑셀 지정값 사용"
        else:
            bedrock_idx = np.where(vs_full >= bedrock_vs_criterion)[0]
            if len(bedrock_idx) > 0:
                b_idx = int(bedrock_idx[0])
                bedrock_vs = float(vs_full[b_idx])
                bedrock_gamma = float(unit_weight_full[b_idx])
                g = g_full.iloc[:b_idx]  # 기반암 상단 윗층만 유한 토층으로 사용
                bedrock_note = (
                    f"자동판정(Vs≥{bedrock_vs_criterion:.0f}m/s, "
                    f"{g_full['토층'].iloc[b_idx]} 층 상단)"
                )
            else:
                bedrock_vs = bedrock_vs_default
                bedrock_gamma = bedrock_gamma_default
                g = g_full
                bedrock_note = "기반암미도달(전역 기본값 사용, 추가시추 검토 권장)"

        thickness = g["층두께(m)"].to_numpy(dtype=float)
        unit_weight = unit_weight_full[: len(g)]
        vs_initial = vs_full[: len(g)] * vs_factor

        if len(thickness) == 0:
            site_response_log.append(
                {
                    "Borehole": bh, "지진파": "-", "반복횟수": 0, "수렴여부": None,
                    "지표amax/g": None, "기반암상단amax/g": None,
                    "기반암조건": "유효 토층 없음(시추 전체가 기반암) - 해석 생략",
                }
            )
            continue

        if "재료유형" in g.columns:
            material_key = np.array(
                [
                    m if isinstance(m, str) and m in DEFAULT_CURVES
                    else classify_material_key(soil)
                    for m, soil in zip(g["재료유형"], g["토층"])
                ]
            )
        else:
            material_key = np.array([classify_material_key(s) for s in g["토층"]])

        pi_values = (
            g["PI"].to_numpy(dtype=float) if "PI" in g.columns else np.full(len(g), np.nan)
        )

        for w in wave_names:
            accel_g = np.asarray(wave_dict[w], dtype=float)
            res = run_equivalent_linear(
                thickness,
                unit_weight,
                vs_initial,
                material_key,
                bedrock_unit_weight=bedrock_gamma,
                bedrock_vs=bedrock_vs,
                dt=dt_per_wave[w],
                accel_g=accel_g,
                pi_values=pi_values,
                strain_ratio=strain_ratio,
                curve_shift=curve_shift,
            )
            n = len(thickness)
            df.loc[g_full.iloc[:n].index, w] = res["amax_g"][:n]
            df.loc[g_full.iloc[:n].index, f"τmax_{w}"] = res["tau_max_kpa"][:n]
            # 지표면(층1 상단) 가속도 시간이력을 보관해 둔다. KDS 17 10 00:2024
            # 4.2.1.4(10)⑤에 따라 S2~S6 지반은 지반응답해석 결과(지표면 운동)의 평균
            # 응답스펙트럼을 해당 지반종류 표준설계응답스펙트럼과 비교해야 하기 때문이다.
            if store_surface:
                surface_motions.setdefault(bh, {})[w] = res["accel_time"][0]
            site_response_log.append(
                {
                    "Borehole": bh,
                    "지진파": w,
                    "반복횟수": res["n_iter"],
                    "수렴여부": res["converged"],
                    "지표amax/g": round(float(res["amax_g"][0]), 4),
                    "기반암상단amax/g": round(float(res["amax_g"][-1]), 4),
                    "기반암조건": bedrock_note,
                }
            )

    return df, pd.DataFrame(site_response_log), surface_motions


# ---------------------------------------------------------
# 0-2. 지진응답해석 케이스(입력가정 민감도) 일괄해석 · 종합판정 · 액상화 분포도
# ---------------------------------------------------------
# SHAKE 결과가 설계자마다 달라지는 주된 원인은 입력가정(동적물성곡선, 유효변형률비,
# 전단파속도 불확실성, 기반암 가정)이다. 이를 기술적으로 타당한 범위에서 바꾼 케이스를
# 한 번에 해석해 비교하고, 설계자가 적용 케이스를 선택하도록 한다.
SR_CASES = [
    {"id": "C1", "케이스": "C1 기본 해석", "동적곡선": "평균곡선", "curve_shift": 1.0, "변형률비": 0.65,
     "vs_factor": 1.0, "bedrock_vs": 760.0,
     "근거": "SHAKE 표준가정 (Seed & Idriss 평균곡선, 유효변형률비 0.65, 실측 Vs)"},
    {"id": "C2", "케이스": "C2 유효변형률비 0.55", "동적곡선": "평균곡선", "curve_shift": 1.0, "변형률비": 0.55,
     "vs_factor": 1.0, "bedrock_vs": 760.0,
     "근거": "Idriss & Sun(1992) 규모 의존식 (M−1)/10, 설계지진 M = 6.5"},
    {"id": "C3", "케이스": "C3 동적곡선 상한", "동적곡선": "상한(기준변형률×2)", "curve_shift": 2.0, "변형률비": 0.65,
     "vs_factor": 1.0, "bedrock_vs": 760.0,
     "근거": "G/Gmax·감쇠비 곡선의 기준변형률 2배 이동 (비선형성·감쇠 작음)"},
    {"id": "C4", "케이스": "C4 동적곡선 하한", "동적곡선": "하한(기준변형률×0.5)", "curve_shift": 0.5, "변형률비": 0.65,
     "vs_factor": 1.0, "bedrock_vs": 760.0,
     "근거": "G/Gmax·감쇠비 곡선의 기준변형률 1/2 이동 (비선형성·감쇠 큼)"},
    {"id": "C5", "케이스": "C5 전단파속도 +20%", "동적곡선": "평균곡선", "curve_shift": 1.0, "변형률비": 0.65,
     "vs_factor": 1.2, "bedrock_vs": 760.0,
     "근거": "토층 Vs 측정 불확실성 상한 (+20%)"},
    {"id": "C6", "케이스": "C6 전단파속도 −20%", "동적곡선": "평균곡선", "curve_shift": 1.0, "변형률비": 0.65,
     "vs_factor": 0.8, "bedrock_vs": 760.0,
     "근거": "토층 Vs 측정 불확실성 하한 (−20%)"},
    {"id": "C7", "케이스": "C7 기반암 Vs 1,500m/s", "동적곡선": "평균곡선", "curve_shift": 1.0, "변형률비": 0.65,
     "vs_factor": 1.0, "bedrock_vs": 1500.0,
     "근거": "기반암 미도달 시추공의 가정 기반암 Vs 상향 (도달 시추공은 C1과 동일)"},
]
for _c in SR_CASES:
    _c.setdefault("input_scale", 1.0)
SR_DERIVED = ["전체 케이스 평균", "전체 케이스 포락(최대)"]

# 사용자 정의 케이스의 '검토서 적용 가능' 범위. C1~C7 민감도 케이스가 쓰는 문헌 범위와 같으며,
# 입력지진 배율은 KDS 설계지반운동수준(S = Z×I)을 바꾸는 것이므로 1.0만 허용한다.
CUSTOM_CASE_RANGES = [
    ("변형률비", "유효변형률비", 0.50, 0.70, "SHAKE 0.65, Idriss & Sun (M−1)/10 = 0.55 등 문헌 범위"),
    ("curve_shift", "동적곡선 기준변형률 배율", 0.50, 2.00, "동적물성곡선 상·하한 (×0.5 ~ ×2)"),
    ("vs_factor", "토층 Vs 배율", 0.80, 1.20, "전단파속도 측정 불확실성 ±20%"),
    ("bedrock_vs", "기반암 Vs (m/s)", 760.0, 3000.0, "KDS 기반암 정의 Vs ≥ 760 m/s"),
    ("input_scale", "입력지진 배율", 1.00, 1.00, "KDS 설계지반운동수준 S = Z×I (변경 불가)"),
]


def custom_case_violations(case):
    """사용자 정의 케이스 설정값 중 검토서 적용 범위를 벗어난 항목 목록."""
    out = []
    for key, label, lo, hi, why in CUSTOM_CASE_RANGES:
        v = float(case.get(key, 1.0))
        if v < lo - 1e-9 or v > hi + 1e-9:
            rng = f"{lo:g}" if lo == hi else f"{lo:g} ~ {hi:g}"
            out.append(f"{label} {v:g} (허용 {rng}: {why})")
    return out


@st.cache_data(show_spinner=False, max_entries=2000)
def _sr_borehole_cached(df_bh, wave_key, _wave_dict, dt_items, curve_shift, strain_ratio, vs_factor, bedrock_vs):
    """시추공 1개 × 해석 케이스 1개의 지반응답해석 (순수 계산 - 화면 요소 호출 없음).
    Streamlit 캐시는 함수 안에서 호출된 화면 요소를 기록했다가 재생하는데, 바깥에서 만든 요소
    (진행률 표시줄 등)를 갱신하면 캐시 적중 시 CacheReplayClosureError가 난다. 그래서 캐시 함수에는
    계산만 두고, 진행률은 바깥 반복문에서 갱신한다. 지진파 배열은 해시 비용을 줄이기 위해 내용 해시
    (wave_key)로 캐시 키를 만들고 배열 자체(_wave_dict)는 해시에서 제외한다."""
    df_out, log_df, _ = run_site_response_all_boreholes(
        df_bh, _wave_dict, dict(dt_items),
        bedrock_vs_default=bedrock_vs, bedrock_gamma_default=21.0,
        strain_ratio=strain_ratio, curve_shift=curve_shift, vs_factor=vs_factor,
    )
    return df_out, log_df


def _wave_hash(wave_dict, dt_per_wave):
    import hashlib
    h = hashlib.md5()
    for w in sorted(wave_dict):
        h.update(str(w).encode("utf-8"))
        h.update(np.ascontiguousarray(np.asarray(wave_dict[w], dtype=float)).tobytes())
        h.update(repr(float(dt_per_wave[w])).encode())
    return h.hexdigest()


def run_sr_case_cached(df_raw, wave_dict, dt_per_wave, curve_shift, strain_ratio, vs_factor, bedrock_vs,
                       _progress_cb=None):
    """케이스 1개에 대한 전 시추공 지반응답해석. 시추공 단위로 캐시하고(같은 입력은 재계산 없이 사용),
    진행률 콜백은 캐시 함수 바깥에서 호출한다."""
    wave_key = _wave_hash(wave_dict, dt_per_wave)
    dt_items = tuple(sorted((str(k), float(v)) for k, v in dt_per_wave.items()))
    bhs = list(dict.fromkeys(df_raw["Borehole"]))
    outs, logs = [], []
    for i, bh in enumerate(bhs):
        if _progress_cb is not None:
            _progress_cb(i, len(bhs))
        d_bh, lg_bh = _sr_borehole_cached(
            df_raw[df_raw["Borehole"] == bh], wave_key, wave_dict, dt_items,
            float(curve_shift), float(strain_ratio), float(vs_factor), float(bedrock_vs),
        )
        outs.append(d_bh)
        logs.append(lg_bh)
    df_out = pd.concat(outs).reindex(df_raw.index)
    log_df = pd.concat(logs, ignore_index=True) if logs else pd.DataFrame()
    return df_out, log_df


def compare_external_sr(ext_df, prog_df, wave_cols, S, soil_class, n_skip=25.0, clay_zones=None,
                        vs1_mode="original", csr_method="tau"):
    """외부 지진응답해석 결과(amax 입력엑셀)와 프로그램 해석결과 비교.
    반환: dict(accel=시추공별 가속도 비교표, fs_spt/fs_vs=산정방법별 최소안전율표, prof=심도별 비교자료)"""
    w = [c for c in wave_cols if c in ext_df.columns and c in prog_df.columns]
    if not w:
        return None
    sfa, _ = kds_sfa(S, soil_class)
    key = ["Borehole", "심도(GL,m)"]
    m = ext_df[key + w].merge(prog_df[key + w], on=key, suffixes=("_외부", "_프로그램"))
    if m.empty:
        return None
    pe = surface_pga_table(ext_df, w, S, soil_class).set_index("Borehole")
    pp = surface_pga_table(prog_df, w, S, soil_class).set_index("Borehole")
    rows = []
    for bh, g in m.groupby("Borehole", sort=False):
        e_mean = g[[f"{x}_외부" for x in w]].mean(axis=1).to_numpy(float)
        p_mean = g[[f"{x}_프로그램" for x in w]].mean(axis=1).to_numpy(float)
        rows.append({
            "Borehole": bh,
            "지표 PGA 외부(g)": pe.loc[bh, "지표 PGA 평균(g)"] if bh in pe.index else np.nan,
            "지표 PGA 프로그램(g)": pp.loc[bh, "지표 PGA 평균(g)"] if bh in pp.index else np.nan,
            "하한 S×Fa(g)": sfa,
            "심도별 amax 비(외부/프로그램)": float(np.nanmean(e_mean / p_mean)),
            "최하단 amax 외부(g)": float(e_mean[-1]), "최하단 amax 프로그램(g)": float(p_mean[-1]),
        })
    acc = pd.DataFrame(rows)
    acc["지표 PGA 비(외부/프로그램)"] = acc["지표 PGA 외부(g)"] / acc["지표 PGA 프로그램(g)"]
    acc["외부/하한"] = acc["지표 PGA 외부(g)"] / sfa
    acc["프로그램/하한"] = acc["지표 PGA 프로그램(g)"] / sfa

    def _mins(df, method, floor=None):
        sp, vs, _ = run_liquefaction_eval(df, n_skip=n_skip, clay_zones=clay_zones, csr_method=method,
                                          vs1_mode=vs1_mode, pga_floor=floor)
        f = lambda d: d[d["대상층여부"] == "대상층"].groupby("Borehole", sort=False)["FS"].min()
        return f(sp), f(vs)
    scen = [("① 외부결과 · amax 근사", ext_df, "amax", None),
            ("② 프로그램 · amax 근사", prog_df, "amax", None),
            ("③ 프로그램 · τmax (원칙)", prog_df, "tau", None),
            ("④ 외부결과 · KDS PGA×rd", ext_df, "pga_rd", sfa),
            ("⑤ 프로그램 · KDS PGA×rd", prog_df, "pga_rd", sfa)]
    fs_spt, fs_vs = {}, {}
    for lab, d, meth, fl in scen:
        a, b = _mins(d, meth, fl)
        fs_spt[lab], fs_vs[lab] = a, b
    fs_spt = pd.DataFrame(fs_spt).reset_index().rename(columns={"index": "Borehole"})
    fs_vs = pd.DataFrame(fs_vs).reset_index().rename(columns={"index": "Borehole"})
    return {"accel": acc, "fs_spt": fs_spt, "fs_vs": fs_vs, "prof": m, "waves": w, "sfa": sfa}


def backcalc_input_scale(ext_df, df_input, wave_dict, dt_per_wave, scales=(0.4, 0.5, 0.6, 0.7, 0.8, 0.9, 1.0),
                         curve_shift=1.0, strain_ratio=0.65, vs_factor=1.0, bedrock_vs=760.0, _progress_cb=None):
    """외부 결과(심도별 amax)를 가장 잘 재현하는 입력지진 배율 역산(진단용).
    입력 배율만 바꾸고 나머지 조건은 기본 케이스(C1)와 같게 두며, 심도별 amax 비(외부/프로그램)의
    평균과 평균 상대오차를 반환한다."""
    w = [c for c in wave_dict if c in ext_df.columns]
    key = ["Borehole", "심도(GL,m)"]
    rows = []
    for i, k in enumerate(scales):
        if _progress_cb is not None:
            _progress_cb(i, len(scales))
        d, _ = run_sr_case_cached(df_input, scale_wave_dict(wave_dict, k), dt_per_wave,
                                  curve_shift, strain_ratio, vs_factor, bedrock_vs)
        m = ext_df[key + w].merge(d[key + w], on=key, suffixes=("_e", "_p"))
        e = m[[f"{x}_e" for x in w]].to_numpy(float)
        p_ = m[[f"{x}_p" for x in w]].to_numpy(float)
        ok = (p_ > 0) & np.isfinite(e)
        rows.append({"입력배율": k, "amax 비 평균(외부/프로그램)": float(np.mean(e[ok] / p_[ok])),
                     "평균 상대오차": float(np.mean(np.abs(e[ok] - p_[ok]) / np.maximum(e[ok], 1e-9)))})
    return pd.DataFrame(rows)


def external_diagnosis_text(cmp):
    """비교 결과로부터 차이 원인 설명 문장 생성."""
    acc, fs = cmp["accel"], cmp["fs_spt"]
    out = []
    r_pga = float(acc["지표 PGA 비(외부/프로그램)"].mean())
    r_prof = float(acc["심도별 amax 비(외부/프로그램)"].mean())
    out.append(f"가속도 크기 차이(주원인): 외부 해석의 지표 PGA는 프로그램 해석의 평균 {r_pga * 100:.0f}%, "
               f"심도별 amax는 평균 {r_prof * 100:.0f}% 수준입니다. CSR이 가속도에 비례하므로 안전율은 약 "
               f"{1 / max(r_prof, 1e-6):.2f}배 차이가 납니다.")
    c1, c2, c3 = "① 외부결과 · amax 근사", "② 프로그램 · amax 근사", "③ 프로그램 · τmax (원칙)"
    if {c2, c3} <= set(fs.columns):
        k = float(np.nanmean(fs[c3] / fs[c2]))
        out.append(f"CSR 산정방법 차이(반대 방향): 외부 amax 입력은 심도별 amax 근사(0.65·amax·σv/σv′)로, "
                   f"프로그램 해석은 τmax(원칙)로 계산됩니다. 같은 해석결과에서 τmax 방식의 안전율이 amax 근사보다 "
                   f"평균 {k:.2f}배 큽니다(가속도 최대치가 심도별로 동시에 발생한다고 보는 amax 근사가 보수적).")
    low = acc[acc["외부/하한"] < 1.0]
    if len(low):
        out.append(f"KDS 4.7(6)① 단서 검토: 외부 결과의 지표 PGA가 하한 S×Fa({cmp['sfa']:.3f}g)의 "
                   f"{low['외부/하한'].min() * 100:.0f}~{low['외부/하한'].max() * 100:.0f}% 수준으로 미달합니다. "
                   "가속도 기반 CSR(PGA×rd)로 평가하려면 S×Fa를 하한으로 적용해야 합니다(④·⑤ 열).")
    if float(acc["지표 PGA 외부(g)"].mean()) < 0.9 * float(acc["지표 PGA 프로그램(g)"].mean()):
        out.append("외부 해석에서 확인할 항목: ① 입력 지진파 최대가속도와 '노두운동(outcrop)' 입력 설정, "
                   "② 건축물 기준(KDS 41)의 2/3 저감 등 설계지반운동 저감 적용 여부, ③ 토층 Vs·단위중량, "
                   "④ 동적물성곡선(G/Gmax·감쇠비), ⑤ 기반암 깊이·Vs, ⑥ 출력 위치(층 상단/중앙).")
    return out


def derive_case_df(case_dfs, wave_cols, how):
    """여러 케이스의 심도별 amax/g를 평균 또는 최대(포락)로 결합."""
    base = next(iter(case_dfs.values())).copy()
    # amax뿐 아니라 최대전단응력(τmax_지진파) 열도 함께 결합해야 τmax 방식 CSR에 평균·포락이 반영된다
    cols = list(wave_cols) + [f"τmax_{w}" for w in wave_cols if f"τmax_{w}" in base.columns]
    stack = np.stack([d[cols].to_numpy(dtype=float) for d in case_dfs.values()])
    base[cols] = np.nanmean(stack, axis=0) if how == "mean" else np.nanmax(stack, axis=0)
    return base


def borehole_judgement(df_spt, df_vs):
    """시추공별 종합판정 (KDS 17 10 00:2024 절차: 예비평가 → 본평가 FS = CRR/CSR > 1.0).
    두 평가방법(SPT·Vs) 중 작은 안전율로 판정한다(보수적)."""
    rows = []
    for bh in dict.fromkeys(df_spt["Borehole"]):
        g1 = df_spt[(df_spt["Borehole"] == bh) & (df_spt["대상층여부"] == "대상층")]
        g2 = df_vs[(df_vs["Borehole"] == bh) & (df_vs["대상층여부"] == "대상층")]
        t1, t2 = g1["FS"].dropna(), g2["FS"].dropna()
        fs1 = float(t1.min()) if len(t1) else np.nan
        fs2 = float(t2.min()) if len(t2) else np.nan
        fs = float(np.nanmin([fs1, fs2])) if (len(t1) or len(t2)) else np.nan
        ng = pd.concat([g1[g1["FS"] <= 1.0], g2[g2["FS"] <= 1.0]])
        if not (len(t1) or len(t2)):
            j = "평가생략(예비평가)"
        elif fs > 1.0:
            j = "안전 (FS > 1.0)"
        else:
            j = "액상화 발생 (FS ≤ 1.0)"
        ng_rng = (f"{ng['시작심도(m)'].min():.1f} ~ {ng['종료심도(m)'].max():.1f}" if len(ng) else "-")
        rows.append({"Borehole": bh, "최소FS(SPT)": fs1, "최소FS(Vs)": fs2, "최소FS": fs,
                     "NG층수(SPT)": int((g1["FS"] <= 1.0).sum()), "NG층수(Vs)": int((g2["FS"] <= 1.0).sum()),
                     "액상화 발생구간(GL-m)": ng_rng, "종합판정": j})
    return pd.DataFrame(rows)


JUDGE_COLORS = {"안전 (FS > 1.0)": "#1E8449", "액상화 발생 (FS ≤ 1.0)": "#C0392B",
                "평가생략(예비평가)": "#7F8C8D"}


def case_metrics(df_spt, df_vs, df_amax, wave_cols):
    jd = borehole_judgement(df_spt, df_vs)
    st_ = df_spt[df_spt["대상층여부"] == "대상층"]
    vt_ = df_vs[df_vs["대상층여부"] == "대상층"]
    surf = df_amax.sort_values("심도(GL,m)").groupby("Borehole", sort=False).head(1)[wave_cols]
    return {
        "지표 amax 평균(g)": float(np.nanmean(surf.to_numpy(dtype=float))),
        "지표 amax 최대(g)": float(np.nanmax(surf.to_numpy(dtype=float))),
        "최소FS(SPT)": float(st_["FS"].min()) if st_["FS"].notna().any() else np.nan,
        "최소FS(Vs)": float(vt_["FS"].min()) if vt_["FS"].notna().any() else np.nan,
        "NG 시추공": int((jd["최소FS"] <= 1.0).sum()),
        "NG 층수(SPT)": int((st_["FS"] <= 1.0).sum()),
        "NG 층수(Vs)": int((vt_["FS"] <= 1.0).sum()),
        "액상화 미발생": bool((jd["최소FS"] <= 1.0).sum() == 0),
    }


def idw_grid(x, y, v, nx=120, ny=120, power=2.0, pad=0.08):
    """역거리가중(IDW) 보간 격자 (액상화 분포도 등고선용)."""
    x, y, v = (np.asarray(a, dtype=float) for a in (x, y, v))
    dx, dy = (x.max() - x.min()) or 1.0, (y.max() - y.min()) or 1.0
    gx = np.linspace(x.min() - pad * dx, x.max() + pad * dx, nx)
    gy = np.linspace(y.min() - pad * dy, y.max() + pad * dy, ny)
    GX, GY = np.meshgrid(gx, gy)
    d = np.sqrt((GX[..., None] - x) ** 2 + (GY[..., None] - y) ** 2)
    w = 1.0 / np.maximum(d, 1e-6) ** power
    return gx, gy, (w * v).sum(-1) / w.sum(-1)


def make_liq_map_plotly(coords_df, judge_df):
    """액상화 분포도: 시추공별 종합판정(심벌) + 최소 안전율 IDW 분포(FS = 1.0 경계선 강조)."""
    m = coords_df.merge(judge_df, on="Borehole", how="inner")
    fig = go.Figure()
    mv = m.dropna(subset=["최소FS"])
    if len(mv) >= 3:
        gx, gy, gz = idw_grid(mv["X"], mv["Y"], np.clip(mv["최소FS"], 0, 3))
        fig.add_trace(go.Contour(
            x=gx, y=gy, z=gz, zmin=0, zmax=2.0,
            contours=dict(start=0, end=2.0, size=0.25, showlabels=True),
            colorscale=[[0, "#C0392B"], [0.45, "#F5B7B1"], [0.5, "#FDFEFE"], [0.75, "#D5F5E3"], [1, "#1E8449"]],
            colorbar=dict(title="최소 FS"), opacity=0.7, name="최소 안전율 분포",
        ))
        fig.add_trace(go.Contour(
            x=gx, y=gy, z=gz, contours=dict(start=1.0, end=1.0, size=1, coloring="lines", showlabels=True),
            line=dict(color="#922B21", width=3), showscale=False, name="FS = 1.0 경계", hoverinfo="skip",
        ))
    for j, sub in m.groupby("종합판정"):
        fig.add_trace(go.Scatter(
            x=sub["X"], y=sub["Y"], mode="markers+text", text=sub["Borehole"], textposition="top center",
            marker=dict(size=12, color=JUDGE_COLORS.get(j, "#555"), line=dict(width=1, color="black")),
            name=j, customdata=sub[["최소FS"]].to_numpy(),
            hovertemplate="%{text}<br>최소 FS = %{customdata[0]:.2f}<extra></extra>",
        ))
    fig.update_layout(title="<b>액상화 분포도 (시추공별 판정 + 최소 안전율 분포)</b>", template="plotly_white",
                      height=580, xaxis_title="X (m)", yaxis_title="Y (m)",
                      yaxis=dict(scaleanchor="x", scaleratio=1),
                      legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="center", x=0.5))
    return fig


def fig_kds_flow_png():
    """KDS 17 10 00:2024 「내진설계 일반」 4.7 액상화 평가 흐름도.
    예비평가(4.7(3)·(4)) → 본평가(4.7(6), FS ≤ 1.0 액상화 발생) → 액상화 영향 반영 안정성 평가·대책(4.7(7))."""
    from matplotlib.patches import FancyBboxPatch as _FB, Polygon as _PG
    fig, ax = plt.subplots(figsize=(12.6, 8.6))
    ax.set_xlim(0, 16.2)
    ax.set_ylim(-0.2, 11.0)
    ax.axis("off")
    cols = {"s1": "#5D6D7E", "s2": "#1F3A5F", "s3": "#943126"}
    X0 = (0.15, 5.5, 10.85)
    for x0, key, title in ((X0[0], "s1", "STEP 1   예비평가  [4.7(3)·(4)]"),
                           (X0[1], "s2", "STEP 2   본평가  [4.7(6)]"),
                           (X0[2], "s3", "STEP 3   영향 평가·대책  [4.7(7)]")):
        ax.add_patch(_FB((x0, 0.0), 5.2, 10.35, boxstyle="round,pad=0.02,rounding_size=0.15",
                         fc="#FBFCFC", ec=cols[key], lw=1.4, ls="--"))
        ax.text(x0 + 2.6, 10.05, title, ha="center", va="center", fontsize=11, fontweight="bold",
                color="white", bbox=dict(boxstyle="round,pad=0.35", fc=cols[key], ec="none"))

    def box(x, y, t, fc="#EAF1FB", ec="#1F3A5F", w=4.6, h=1.05, bold=False, fs=8.9):
        ax.add_patch(_FB((x - w / 2, y - h / 2), w, h, boxstyle="round,pad=0.02,rounding_size=0.1",
                         fc=fc, ec=ec, lw=1.3))
        ax.text(x, y, t, ha="center", va="center", fontsize=fs, fontweight="bold" if bold else "normal",
                linespacing=1.35)

    def dia(x, y, t, w=4.6, h=1.8, fs=8.7):
        ax.add_patch(_PG([(x, y + h / 2), (x + w / 2, y), (x, y - h / 2), (x - w / 2, y)],
                         fc="#FEF5E7", ec="#B9770E", lw=1.4))
        ax.text(x, y, t, ha="center", va="center", fontsize=fs, linespacing=1.3)

    def arr(p0, p1, lab=None, off=(0.1, 0.0), color="#333"):
        ax.annotate("", xy=p1, xytext=p0, arrowprops=dict(arrowstyle="-|>", color=color, lw=1.3))
        if lab:
            ax.text((p0[0] + p1[0]) / 2 + off[0], (p0[1] + p1[1]) / 2 + off[1], lab, fontsize=8.8,
                    color="#B03A2E", fontweight="bold", ha="left", va="center")

    def route(pts, lab=None, lab_at=None, color="#333", ls="-"):
        """직각으로 꺾이는 연결선."""
        xs = [p[0] for p in pts[:-1]]
        ys = [p[1] for p in pts[:-1]]
        ax.plot(xs, ys, color=color, lw=1.3, ls=ls)
        ax.annotate("", xy=pts[-1], xytext=pts[-2], arrowprops=dict(arrowstyle="-|>", color=color, lw=1.3))
        if lab:
            ax.text(lab_at[0], lab_at[1], lab, fontsize=8.8, color="#B03A2E", fontweight="bold",
                    ha="center", va="center", bbox=dict(fc="white", ec="none", pad=0.5))

    # ---------------- STEP 1 예비평가 ----------------
    x1 = X0[0] + 2.6
    box(x1, 9.05, "지반조사 자료 분석\n주상도 · SPT-N(에너지 보정 전) · 입도(FC)\nPI · LL · wc · 연중 최고지하수위", h=1.25)
    box(x1, 7.55, "단위깊이(지층두께 1.5m 이하)별 평가\nN값은 지층 평균이 아닌 측점별 적용", h=0.95)
    dia(x1, 5.05, "평가 생략조건 해당 여부\n① 연중 최고지하수위 상부 지층\n② 심도 20m보다 깊은 지층\n"
                  "③ 주상도 SPT-N값 25 초과 지층\n④ 고소성 점토거동 (그림 4.7-1 영역 C)", w=5.0, h=3.3, fs=8.4)
    box(x1, 1.75, "액상화 평가 생략", fc="#E8F6EF", ec="#1E8449", bold=True, h=0.75, w=3.6)
    ax.text(x1, 0.72, "※ 영역 C : 반복연화에 의한 강도저하·대변형 별도 검토\n"
                      "※ 내진특등급 실트·점토 : 비배수 반복시험 검토 [4.7(5)]",
            ha="center", va="center", fontsize=7.6, color="#5D6D7E", linespacing=1.4)
    arr((x1, 8.42), (x1, 8.03))
    arr((x1, 7.07), (x1, 6.71))
    arr((x1, 3.39), (x1, 2.13), "해당", off=(0.12, 0))

    # ---------------- STEP 2 본평가 ----------------
    x2 = X0[1] + 2.6
    box(x2, 9.05, "설계지반운동 결정\nS = Z × I  (시설물 성능목표별 재현주기)", h=0.95)
    box(x2, 7.8, "지진응답해석 (1차원 등가선형)\n입력 : 기반암(Vs ≥ 760 m/s) 노두운동", h=0.95)
    box(x2, 6.4, "반복전단응력비  [식 4.7-2]\nCSR = 0.65 · (τmax)d,GRA / σv′\n"
                 "(또는 지표 PGA · rd,  단 PGA ≥ S × Fa)", h=1.3)
    box(x2, 4.75, "반복저항응력비  CRR = CRR7.5 · Kσ · MSF\nSPT [4.7-3~8] · CPT [4.7-9~12] · Vs [4.7-13~19]\n"
                  "(현장상황에 따라 비배수 반복시험 적용)", h=1.3)
    dia(x2, 2.75, "안전율  FS = CRR / CSR  [식 4.7-1]\n(2개 이상 시험 시 최소 FS)", w=4.9, h=1.75, fs=8.6)
    box(x2, 0.85, "액상화 발생 안 함", fc="#E8F6EF", ec="#1E8449", bold=True, h=0.7, w=3.6)
    route([(x1 + 2.5, 5.05), (X0[1] - 0.17, 5.05), (X0[1] - 0.17, 9.05), (x2 - 2.3, 9.05)])
    ax.text(x1 + 2.05, 4.42, "비해당", fontsize=8.8, color="#B03A2E", fontweight="bold", ha="center", va="center")
    arr((x2, 8.57), (x2, 8.28))
    arr((x2, 7.32), (x2, 7.06))
    arr((x2, 5.74), (x2, 5.41))
    arr((x2, 4.09), (x2, 3.63))
    arr((x2, 1.87), (x2, 1.21), "FS > 1.0", off=(0.12, 0))

    # ---------------- STEP 3 영향 평가 · 대책 ----------------
    x3 = X0[2] + 2.6
    box(x3, 9.05, "액상화 발생 판정 (FS ≤ 1.0)\n위험구간 선정 (시추공 · 심도)", fc="#FDEDEC", ec="#943126",
        bold=True, h=0.95)
    box(x3, 7.6, "액상화 영향을 반영한 시설물 안정성 평가\n시설물 유형 · 기초 형식 · 지반 특성\n"
                 "액상화에 의한 지반 강도 저하 고려  [①]", h=1.3)
    dia(x3, 5.55, "시설물 안정성\n확보 여부", w=3.6, h=1.6, fs=9.0)
    box(x3 - 1.27, 3.35, "확보\n기초·지반 설계에\n액상화 영향 반영", fc="#E8F6EF", ec="#1E8449", w=2.35, h=1.25,
        fs=8.4)
    box(x3 + 1.27, 3.35, "미확보\n대책방안 수립·적용  [②]\n(밀도증대·고결·배수·치환)", fc="#FDEDEC", ec="#943126",
        w=2.35, h=1.25, fs=8.0)
    box(x3 + 1.27, 1.45, "대책 적용 후\n액상화 재평가", fc="#FDEDEC", ec="#943126", w=2.35, h=0.9, fs=8.4)
    route([(x2 + 2.45, 2.75), (X0[2] - 0.17, 2.75), (X0[2] - 0.17, 9.05), (x3 - 2.3, 9.05)])
    ax.text(x2 + 1.8, 2.12, "FS ≤ 1.0", fontsize=8.8, color="#B03A2E", fontweight="bold", ha="center", va="center")
    arr((x3, 8.57), (x3, 8.26))
    arr((x3, 6.94), (x3, 6.36))
    ax.plot([x3, x3], [4.75, 4.38], color="#333", lw=1.3)
    ax.plot([x3 - 1.27, x3 + 1.27], [4.38, 4.38], color="#333", lw=1.3)
    arr((x3 - 1.27, 4.38), (x3 - 1.27, 3.99))
    arr((x3 + 1.27, 4.38), (x3 + 1.27, 3.99))
    arr((x3 + 1.27, 2.71), (x3 + 1.27, 1.91))
    route([(x3 + 2.45, 1.45), (X0[2] + 5.05, 1.45), (X0[2] + 5.05, 7.6), (x3 + 2.3, 7.6)], color="#943126", ls="--")
    ax.text(X0[2] + 5.0, 4.6, "대책 후 재검토", fontsize=7.8, color="#943126", rotation=90, ha="right", va="center")
    ax.text(8.1, -0.45, "KDS 17 10 00 : 2024 「내진설계 일반」 4.7 액상화 (식 번호는 KDS 원문 기준)", ha="center",
            fontsize=9, color="#5D6D7E")
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


# ---------------------------------------------------------
# 0-3. 액상화 대책공법 상세 검토: 개량범위(면적·심도) 산정 → 공법 적용성 → 대안별 개략공사비
# ---------------------------------------------------------
# ---------------------------------------------------------
# 공법별 적용 단가 (2026.10 검토) - 직접공사비, 공법 고유 단위
# ---------------------------------------------------------
# SCP·GCP·DCM·JSP·약액주입·동다짐은 국토교통부 표준시장단가(2025년 상반기 'C. 지반개량공사'는 보강토옹벽만 수록)
# 공종이 아니어서 표준품셈 일위대가 또는 견적으로 산정하는 공종이다. 설계자 제시 단가를 기준으로 공개 자료
# (표준시장단가 토공 항목, 관급 골재 단가, SGR 공사비 내역, FHWA NHI-16-027 표 1-6, 일본 우라야스시 액상화
# 대책 검토 단가)와 비교해 적용성을 검토하고 아래 값을 확정하였다. 상세 비교는 '공법별_단가_검토표.xlsx' 참조.
# 단위별 개략공사비(원/m², 평면 1m²·개량심도 D 기준):
#   원/m²            : 단가 (개량심도가 공법 최대심도 이내일 때)
#   원/m³            : 단가 × D                      (치환)
#   원/m³(개량체)    : 단가 × 개량률 × D             (DCM, 격자형 개량)
#   원/m             : 단가 × D / 간격²              (SCP·GCP·JSP·약액주입, 정사각형 배치)
REMED_PRICE_BASIS_DEFAULT = ("공법별 검토 확정단가(2026.10, 직접공사비) - 표준시장단가 토공 항목·관급 골재 단가 기반 "
                             "산출 및 설계자 제시 단가 비교 검토")
REMED_COST_DB = [
    {"공법": "유압해머다짐", "분류": "다짐(밀도증대)", "최대심도(m)": 5.0, "FC상한(%)": 20.0,
     "단가": 22000, "단위": "원/m²", "직경(m)": np.nan, "간격(m)": np.nan, "개량률": np.nan,
     "소음진동": "높음", "협소부지": "불가", "비고": "표층 4~5m 다짐, 시공속도 빠름",
     "단가 근거": "설계자 제시 단가 유지(유압해머 장비 조합, 표준 다짐두께). 공개 국내 단가 미확인 - 동다짐보다 높게 "
                 "제시되어 있어 견적 확인 권장"},
    {"공법": "동다짐", "분류": "다짐(밀도증대)", "최대심도(m)": 8.0, "FC상한(%)": 35.0,
     "단가": 16000, "단위": "원/m²", "직경(m)": np.nan, "간격(m)": np.nan, "개량률": np.nan,
     "소음진동": "매우 높음", "협소부지": "불가",
     "비고": "추 15t급(유효심도 약 7~8m), FC 15% 초과 시 간극수압 소산 대기·시험시공 필수",
     "단가 근거": "설계자 제시 단가 유지. FHWA(2016) 미국 단가 $12~36/m²의 하한 수준으로 국내 적용 타당"},
    {"공법": "치환(양질토)", "분류": "치환", "최대심도(m)": 5.0, "FC상한(%)": 100.0,
     "단가": 45000, "단위": "원/m³", "직경(m)": np.nan, "간격(m)": np.nan, "개량률": np.nan,
     "소음진동": "보통", "협소부지": "가능", "비고": "굴착·사토 후 양질토 채움·다짐, 지하수위 하부는 배수·흙막이 필요",
     "단가 근거": "설계자 제시 단가 유지. 표준시장단가(2025 상반기) 터파기·적재·되메우기 다짐 + 양질토 반입 산출값 "
                 "약 42,000원/m³ 대비 +7%(보수적)"},
    {"공법": "쇄석치환", "분류": "치환", "최대심도(m)": 5.0, "FC상한(%)": 100.0,
     "단가": 60000, "단위": "원/m³", "직경(m)": np.nan, "간격(m)": np.nan, "개량률": np.nan,
     "소음진동": "보통", "협소부지": "가능", "비고": "현장 실무형 상부 치환(쇄석 채움·다짐, 배수성 우수)",
     "단가 근거": "신규 산출: 표준시장단가(2025 상반기) 토공 항목 + 쇄석 자재(관급 골재 단가 15,400~42,275원/m³ 범위, "
                 "28,000원/m³ 적용, 다짐할증 15%) + 배수·가시설 + 2026 상승분(2.98%) → 약 60,000원/m³"},
    {"공법": "SCP (다짐모래기둥)", "분류": "다짐(밀도증대)", "최대심도(m)": 20.0, "FC상한(%)": 25.0,
     "단가": 55000, "단위": "원/m", "직경(m)": 0.7, "간격(m)": 2.0, "개량률": np.nan,
     "소음진동": "높음", "협소부지": "불가", "비고": "케이싱 Φ500(형성 Φ700), 모래 자재 포함, 정사각형 2.0m 배치",
     "단가 근거": "설계자 제시 단가 유지. 일본 2011 시장단가(手間 2,500엔/m + 모래)·FHWA 하한과 비교 시 타당 범위"},
    {"공법": "GCP (다짐쇄석기둥)", "분류": "다짐+배수", "최대심도(m)": 25.0, "FC상한(%)": 35.0,
     "단가": 62000, "단위": "원/m", "직경(m)": 0.7, "간격(m)": 2.0, "개량률": np.nan,
     "소음진동": "보통", "협소부지": "불가", "비고": "케이싱 Φ500(형성 Φ700), 쇄석 자재 포함, 정사각형 2.0m 배치",
     "단가 근거": "설계자 제시 단가 유지. FHWA 쇄석기둥 $49~197/m 하한 이하(국내 저가), SCP 대비 +13%(쇄석 자재 차) 타당"},
    {"공법": "DCM (심층혼합처리)", "분류": "고결", "최대심도(m)": 30.0, "FC상한(%)": 100.0,
     "단가": 85000, "단위": "원/m³(개량체)", "직경(m)": np.nan, "간격(m)": np.nan, "개량률": 0.40,
     "소음진동": "낮음", "협소부지": "불가", "비고": "2축 교반, 시멘트 고화재 포함, 액상화 격자형 개량률 40%",
     "단가 근거": "설계자 제시 단가 유지(개량체 체적 기준). 개량률을 곱해 개량지반 체적당 34,000원/m³로 환산"},
    {"공법": "고압분사주입 (JSP)", "분류": "고결", "최대심도(m)": 25.0, "FC상한(%)": 100.0,
     "단가": 75000, "단위": "원/m", "직경(m)": 0.8, "간격(m)": 1.1, "개량률": np.nan,
     "소음진동": "낮음", "협소부지": "가능", "비고": "Φ800 이중관, 격자 1.1m 배치(개량률 약 42%), 대구경(RJP·SIG)은 별도",
     "단가 근거": "설계자 제시 단가 유지(Φ800 기준). 대구경 고압분사(Φ1.5~2.0m)는 m당 단가가 크게 높아 별도 견적 필요"},
    {"공법": "약액주입", "분류": "고결", "최대심도(m)": 15.0, "FC상한(%)": 20.0,
     "단가": 120000, "단위": "원/m", "직경(m)": np.nan, "간격(m)": 1.0, "개량률": np.nan,
     "소음진동": "없음", "협소부지": "가능", "비고": "액상화 방지용 침투고결(간극 충전), 1.0m 격자, 주입량 약 0.36m³/m",
     "단가 근거": "설계자 제시 48,000원/m는 차수용 주입량(약 0.1m³/m) 수준. SGR 공사비 내역의 천공비:주입비 구성으로 "
                 "액상화용 주입량(간극률 0.4×충전 0.9)에 맞춰 약 2.5배 조정 → 120,000원/m"},
]
# 쇄석치환·양질토 치환 단가 산출 내역 (원/m³, 2025 상반기 표준시장단가 기준 + 2026 상승 2.98%)
REPLACEMENT_BUILDUP = [
    ("기계터파기 보통토사 5m이하(불량)", "표준시장단가 DE130.11001", 2252, 2252),
    ("흙운반 적재(굴삭기+덤프 20톤)", "표준시장단가 DS100.20200", 3372, 3372),
    ("사토 운반(약 10km, 가정)", "덤프 운반 가정값", 5000, 5000),
    ("잔토처리(토사)", "표준시장단가 DF310.00000", 786, 786),
    ("채움재 자재(현장도착, 다짐할증 15%)", "쇄석 28,000×1.15 / 양질토 반입 15,000", 32200, 15000),
    ("포설·다짐(대형장비)", "표준시장단가 DH110.00001", 6735, 6735),
    ("지하수 배수·간이 가시설(가정)", "가정값", 7500, 7500),
]
SHALLOW_LIMIT_M = 5.0


def _num(v, default=np.nan):
    try:
        v = float(v)
        return v if np.isfinite(v) else default
    except (TypeError, ValueError):
        return default


def remed_cost_per_area(m, depth):
    """평면 1m²를 개량심도 depth(m)까지 개량하는 직접공사비(원/m²) - 공법 단가 단위별 환산."""
    price, unit, depth = _num(m.get("단가"), 0.0), str(m.get("단위", "원/m³")), float(depth)
    if unit == "원/m²":
        return price
    if unit == "원/m³(개량체)":
        return price * _num(m.get("개량률"), 1.0) * depth
    if unit == "원/m":
        s_ = _num(m.get("간격(m)"), 1.0)
        return price * depth / max(s_ * s_, 1e-6)
    return price * depth          # 원/m³ (치환)


def remed_equiv_m3(m, depth=10.0):
    """비교용 개량지반 체적당 환산단가(원/m³) - 개량심도 depth 기준."""
    return remed_cost_per_area(m, depth) / max(float(depth), 1e-6)


def remed_quantity_text(m):
    u = str(m.get("단위", ""))
    if u == "원/m":
        d_ = _num(m.get("직경(m)"))
        return (f"Φ{d_ * 1000:.0f}mm, " if np.isfinite(d_) else "") + f"간격 {_num(m.get('간격(m)'), 1.0):.1f}m 정사각형"
    if u == "원/m³(개량체)":
        return f"개량률 {_num(m.get('개량률'), 1.0) * 100:.0f}%"
    if u == "원/m²":
        return f"최대 개량심도 {_num(m.get('최대심도(m)')):.0f}m"
    return "굴착·채움 체적"
# ---------------------------------------------------------
# 현장 실무형 검토: 상부 비액상화층(H1) 확보에 의한 지표 액상화 피해 억제
# ---------------------------------------------------------
# Ishihara(1985)는 수평지반에서 비액상화 표층 두께(H1)와 하부 액상화층 두께(H2), 지표 최대가속도에 따라
# 지표 피해(분사·균열) 발생 경계를 제시했다. 국내 현장에서 상부 쇄석치환·동다짐을 많이 쓰는 근거가 이것으로,
# 상부를 개량해 비액상화층을 충분히 두껍게 만들면 하부 액상화층이 남아 있어도 지표 피해를 억제할 수 있다.
# 아래 필요 H1은 Ishihara(1985) 경계곡선의 상한(평탄부)을 보수적으로 읽은 근사값이며, H2와 무관하게 적용한다.
# 적용 한계: 수평지반의 지표 피해 판단용이며 경사지·호안 인접부 측방유동, 얕은기초 지지력, 말뚝 부마찰력·
# 수평저항 감소, 매설관 부상은 별도 검토해야 한다(KDS 17 10 00:2024 4.7(7) 시설물 안정성 평가).
ISHIHARA_H1_POINTS = [(0.20, 3.0), (0.30, 6.0), (0.45, 9.0)]   # (지표 PGA g, 필요 비액상화층 두께 m)


def ishihara_h1_required(pga):
    xs = [p for p, _ in ISHIHARA_H1_POINTS]
    ys = [h for _, h in ISHIHARA_H1_POINTS]
    return float(np.interp(pga, xs, ys, left=ys[0], right=ys[-1]))


def dc_effective_depth(weight_t, drop_m, fc):
    """동다짐 유효개량심도 D = n·√(W·H) (Menard & Broise 1975, Lukas 1995: 투수성 사질토 n≈0.5,
    비소성 실트질 모래 n≈0.35~0.4). FC 35% 초과는 적용 곤란으로 0 반환."""
    if fc is None or not np.isfinite(fc):
        fc = 0.0
    if fc > 35.0:
        return 0.0
    n = 0.5 if fc <= 15.0 else 0.4
    return float(n * np.sqrt(max(weight_t, 0.0) * max(drop_m, 0.0)))


def crust_assessment(df_spt, df_vs, judge_df, surf_pga, sfa, rep_limit=3.0, dc_weight=15.0, dc_drop=20.0,
                     urban=False, narrow=False):
    """액상화 발생 시추공별 상부 비액상화층 확보 검토.
    surf_pga: {시추공: 지표 PGA 평균(g)}. 평가 PGA = max(지표 PGA, S×Fa) (KDS 4.7(6)① 하한)."""
    rows = []
    for bh in judge_df["Borehole"]:
        ng = pd.concat([df_spt[(df_spt["Borehole"] == bh) & (df_spt["대상층여부"] == "대상층") & (df_spt["FS"] <= 1.0)],
                        df_vs[(df_vs["Borehole"] == bh) & (df_vs["대상층여부"] == "대상층") & (df_vs["FS"] <= 1.0)]])
        if ng.empty:
            continue
        ztop, zbot = float(ng["시작심도(m)"].min()), float(ng["종료심도(m)"].max())
        fc = float(ng["FC(%)"].mean())
        pga = max(float(surf_pga.get(bh, np.nan)) if pd.notna(surf_pga.get(bh, np.nan)) else 0.0, float(sfa or 0.0))
        h1_req = ishihara_h1_required(pga)
        d_req = 0.0 if ztop >= h1_req - 1e-9 else min(h1_req, zbot)
        d_eff = dc_effective_depth(dc_weight, dc_drop, fc)
        opts = []
        if d_req > 0:
            if d_req <= rep_limit + 1e-9:
                opts.append("상부 치환")
            if d_eff >= d_req - 1e-9 and not urban and not narrow:
                opts.append("동다짐")
        if d_req <= 0:
            verdict = "H1 충분 - 지표피해 가능성 낮음"
            note = "전면 지반개량 대신 시설물 안정성 평가(KDS 4.7(7))·기초 설계에 액상화 영향 반영"
        elif opts:
            verdict = "상부 처리로 H1 확보 가능"
            note = f"GL-{d_req:.1f}m까지 " + " 또는 ".join(opts) + " → 하부 액상화층은 기초 설계에 반영"
        else:
            verdict = "상부 처리로 불충분"
            why = []
            if d_req > rep_limit:
                why.append(f"치환 한계 {rep_limit:g}m 초과")
            if d_eff < d_req:
                why.append(f"동다짐 유효심도 {d_eff:.1f}m 부족" if d_eff > 0 else "동다짐 적용 곤란(FC>35%)")
            if urban or narrow:
                why.append("입지 제약(소음·진동/협소)")
            note = "심층개량(SCP·GCP·DCM 등) 또는 구조물 대책 검토 (" + ", ".join(why) + ")"
        rows.append({"Borehole": bh, "액상화층 상단 H1(m)": ztop, "액상화층 하단(m)": zbot, "H2(m)": zbot - ztop,
                     "평가 PGA(g)": pga, "필요 H1(m)": h1_req, "필요 처리심도(m)": d_req, "평균FC(%)": fc,
                     "동다짐 유효심도(m)": d_eff, "판정": verdict, "현장 실무형 대책": note})
    return pd.DataFrame(rows)


# ---------------------------------------------------------
# 시추공 좌표 입력양식 / 좌표 파일 해석 (액상화 분포도)
# ---------------------------------------------------------
def build_coord_template(boreholes=None, n_blank=100):
    """시추공 좌표 입력양식(xlsx). boreholes를 주면 시추공명을 미리 채운다(액상화 입력엑셀과 이름 일치)."""
    out = io.BytesIO()
    boreholes = [str(b) for b in (boreholes or [])]
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb = w.book
        ws = wb.add_worksheet("좌표입력")
        w.sheets["좌표입력"] = ws
        hdr = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#1F4E78",
                             "font_color": "white", "border": 1, "text_wrap": True})
        opt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#8497B0",
                             "font_color": "white", "border": 1, "text_wrap": True})
        bh_f = wb.add_format({"border": 1, "align": "center"})
        inp = wb.add_format({"border": 1, "bg_color": "#FFF2CC", "num_format": "0.000"})
        inp_t = wb.add_format({"border": 1, "bg_color": "#FFF2CC"})
        o_f = wb.add_format({"border": 1, "num_format": "0.00"})
        heads = ["Borehole\n(시추공명)", "동향 E (m)\n[X]", "북향 N (m)\n[Y]", "지반고 EL.(m)\n(선택)", "비고\n(선택)"]
        for j, h in enumerate(heads):
            ws.write(0, j, h, hdr if j < 3 else opt)
        ws.set_row(0, 34)
        n_rows = max(len(boreholes), 0) if boreholes else n_blank
        for i in range(n_rows):
            r = i + 1
            if boreholes:
                ws.write(r, 0, boreholes[i], bh_f)
            else:
                ws.write_blank(r, 0, None, inp_t)
            ws.write_blank(r, 1, None, inp)
            ws.write_blank(r, 2, None, inp)
            ws.write_blank(r, 3, None, o_f)
            ws.write_blank(r, 4, None, bh_f)
        ws.data_validation(1, 1, max(n_rows, 1), 2, {
            "validate": "decimal", "criteria": "between", "minimum": -1e8, "maximum": 1e8,
            "error_title": "숫자 입력", "error_message": "좌표는 m 단위 숫자로 입력하세요."})
        ws.set_column(0, 0, 16)
        ws.set_column(1, 2, 16)
        ws.set_column(3, 3, 13)
        ws.set_column(4, 4, 24)
        ws.freeze_panes(1, 1)
        ws.write_comment(0, 1, "평면직각좌표 동쪽 방향(Easting). 측량성과의 Y값에 해당합니다.")
        ws.write_comment(0, 2, "평면직각좌표 북쪽 방향(Northing). 측량성과의 X값에 해당합니다.")
        # 작성방법
        g = wb.add_worksheet("작성방법")
        t = wb.add_format({"bold": True, "font_size": 13, "font_color": "#1F4E78"})
        b = wb.add_format({"bold": True})
        tx = wb.add_format({"text_wrap": True, "valign": "top"})
        g.set_column(0, 0, 4)
        g.set_column(1, 1, 22)
        g.set_column(2, 2, 92)
        g.write(0, 1, "시추공 좌표 입력양식 작성방법 (액상화 분포도)", t)
        rows = [
            ("입력 시트", "첫 번째 시트 '좌표입력'만 읽습니다. 1행 제목은 그대로 두고 2행부터 입력하세요. 노란 칸이 입력 칸입니다."),
            ("Borehole", "액상화 평가 입력엑셀의 Borehole 이름과 똑같이 입력해야 결과와 연결됩니다(대소문자·하이픈 포함)."),
            ("동향 E / 북향 N", "평면직각좌표(m)를 모든 시추공에 같은 좌표계로 입력합니다(예: TM 중부원점, UTM-K). "
                               "국내 측량성과는 X = 북향, Y = 동향이므로 측량성과의 Y값을 '동향 E'에, X값을 '북향 N'에 넣으세요."),
            ("위경도 입력", "열 이름을 '경도'·'위도'로 바꿔 십진수(예: 126.9781, 37.5665)로 입력하면 프로그램이 평면좌표(m)로 근사 "
                           "변환합니다(부지 규모 수 km 이내 권장)."),
            ("열 이름 인식", "Borehole/시추공/공번, 동향/E/X, 북향/N/Y, 경도/위도 등을 자동 인식합니다. "
                            "X·Y로만 입력한 측량성과(X = 북향)는 사이드바의 '측량좌표(X=북향)' 옵션을 켜면 바꿔 읽습니다."),
            ("분포도 작성 조건", "좌표가 3개소 이상이면 최소 안전율 분포(역거리가중 보간)와 FS = 1.0 경계선을 함께 그립니다. "
                               "2개소 이하는 시추공 위치·판정만 표시합니다."),
            ("대책공법 연계", "좌표가 3개소 이상이면 시추공별 영향면적(티센 분할)을 계산해 개량면적·개략공사비에 반영합니다."),
            ("선택 항목", "지반고·비고는 참고용이며 계산에는 쓰지 않습니다."),
        ]
        for i, (k, v) in enumerate(rows):
            g.write(2 + i, 1, k, b)
            g.write(2 + i, 2, v, tx)
            g.set_row(2 + i, 33)
        g.write(12, 1, "입력 예시", t)
        ex = [("BH-01", 198520.350, 552310.120), ("BH-02", 198585.800, 552298.400), ("BH-03", 198560.150, 552360.900),
              ("BH-04", 198610.420, 552345.700)]
        g.write_row(13, 1, ["Borehole", "동향 E (m)  /  북향 N (m)"], hdr)
        for i, (bh, e, nn) in enumerate(ex):
            g.write(14 + i, 1, bh, bh_f)
            g.write(14 + i, 2, f"{e:,.3f}  /  {nn:,.3f}", bh_f)
        g.write(19, 1, "※ 위 예시는 형식 안내용입니다. 프로그램은 첫 번째 '좌표입력' 시트만 읽습니다.")
    return out.getvalue()


_COORD_BH = ("BOREHOLE", "BH", "시추공", "시추공명", "시추공번호", "공번", "공명", "HOLE", "HOLENO")
_COORD_X = ("X", "X(M)", "X좌표", "동향", "동향E(M)", "동향E", "E", "E(M)", "EASTING", "EAST")
_COORD_Y = ("Y", "Y(M)", "Y좌표", "북향", "북향N(M)", "북향N", "N", "N(M)", "NORTHING", "NORTH")


def _coord_key(c):
    import re as _re2
    k = str(c).upper().replace(" ", "").replace("\n", "")
    return _re2.sub(r"\[.*?\]|\(선택\)", "", k)


def parse_coords_table(raw, swap_xy=False):
    """좌표 엑셀(머리글 위치 자동 탐색) → (coords_df[Borehole, X(동향), Y(북향)], 안내 메시지)."""
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    hr = None
    for i in range(min(len(raw), 20)):
        keys = [_coord_key(v) for v in raw.iloc[i].tolist()]
        has_bh = any(k.startswith(a) for k in keys for a in ("BOREHOLE", "시추공", "공번", "BH", "HOLE"))
        has_xy = any(k in _COORD_X + _COORD_Y or k.startswith(("동향", "북향", "경도", "위도", "LON", "LAT", "EAST", "NORTH"))
                     for k in keys)
        if has_bh and has_xy:
            hr = i
            break
    if hr is None:
        return None, "좌표 파일에서 시추공(Borehole) 열 머리글을 찾지 못했습니다."
    cols = [_coord_key(v) for v in raw.iloc[hr].tolist()]
    body = raw.iloc[hr + 1:].copy()
    body.columns = range(body.shape[1])

    def _find(cands, starts=()):
        for j, k in enumerate(cols):
            if k in cands:
                return j
        for j, k in enumerate(cols):
            if any(k.startswith(a) for a in starts):
                return j
        return None
    jb = _find(_COORD_BH, ("BOREHOLE", "시추공", "공번", "BH"))
    jlon = _find(("경도", "LON", "LONGITUDE", "LNG", "경도(°)"), ("경도", "LON"))
    jlat = _find(("위도", "LAT", "LATITUDE", "위도(°)"), ("위도", "LAT"))
    jx = _find(_COORD_X, ("동향", "EAST", "X"))
    jy = _find(_COORD_Y, ("북향", "NORTH", "Y"))
    msg = ""
    if jb is None:
        return None, "시추공(Borehole) 열을 찾지 못했습니다."
    if jlon is not None and jlat is not None and (jx is None or jy is None):
        lon = pd.to_numeric(body[jlon], errors="coerce")
        lat = pd.to_numeric(body[jlat], errors="coerce")
        X, Y = lon, lat
        geo = True
    elif jx is not None and jy is not None:
        X = pd.to_numeric(body[jx], errors="coerce")
        Y = pd.to_numeric(body[jy], errors="coerce")
        generic = cols[jx] in ("X", "X(M)", "X좌표") and cols[jy] in ("Y", "Y(M)", "Y좌표")
        if swap_xy and generic:
            X, Y = Y, X
            msg = "측량좌표(X = 북향, Y = 동향)로 보고 X·Y를 바꿔 읽었습니다."
        geo = bool(X.between(124, 132).all() and Y.between(33, 39).all() and X.notna().any())
    else:
        return None, "동향(X)·북향(Y) 또는 경도·위도 열을 찾지 못했습니다."
    out = pd.DataFrame({"Borehole": body[jb].astype(str).str.strip(), "X": X, "Y": Y})
    out = out[(out["Borehole"] != "") & (out["Borehole"].str.lower() != "nan")].dropna()
    if geo and len(out):
        lat0 = float(out["Y"].mean())
        lon0 = float(out["X"].mean())
        out["X"] = (out["X"] - lon0) * 111320.0 * np.cos(np.radians(lat0))
        out["Y"] = (out["Y"] - lat0) * 110540.0
        msg = "경도·위도를 부지 중심 기준 평면좌표(m)로 근사 변환했습니다."
    out = out.drop_duplicates("Borehole").reset_index(drop=True)
    return out, msg


def influence_areas(coords_df, boreholes, site_area=None, n=220):
    """시추공별 영향면적(m²): 좌표 3개소 이상이면 볼록다각형 내부를 격자로 나눠 가장 가까운 시추공에
    배분(티센 방식). 부지면적을 입력하면 그 면적으로 비례 환산. 좌표가 없으면 부지면적을 균등 배분."""
    bhs = list(boreholes)
    if coords_df is not None:
        c = coords_df[coords_df["Borehole"].isin(bhs)].drop_duplicates("Borehole")
    else:
        c = None
    if c is not None and len(c) >= 3:
        from matplotlib.path import Path
        from scipy.spatial import ConvexHull
        pts = c[["X", "Y"]].to_numpy(float)
        hull = ConvexHull(pts)
        poly = Path(pts[hull.vertices])
        xs = np.linspace(pts[:, 0].min(), pts[:, 0].max(), n)
        ys = np.linspace(pts[:, 1].min(), pts[:, 1].max(), n)
        GX, GY = np.meshgrid(xs, ys)
        cells = np.c_[GX.ravel(), GY.ravel()]
        inside = poly.contains_points(cells, radius=1e-6)
        cell_a = (xs[1] - xs[0]) * (ys[1] - ys[0])
        nearest = np.argmin(((cells[inside][:, None, :] - pts[None, :, :]) ** 2).sum(-1), axis=1)
        a = np.bincount(nearest, minlength=len(pts)) * cell_a
        hull_area = float(hull.volume)
        if site_area and site_area > 0:
            a = a * (site_area / a.sum())
        area = dict(zip(c["Borehole"], a))
        for bh in bhs:
            area.setdefault(bh, 0.0)
        return area, (site_area if site_area else hull_area), "좌표 기반 영향면적(티센)"
    tot = float(site_area) if site_area else 10000.0
    return {bh: tot / max(len(bhs), 1) for bh in bhs}, tot, "부지면적 균등배분(좌표 미입력)"


def improvement_zones(df_spt, df_vs, judge_df, coords_df=None, site_area=None):
    """액상화 발생(FS≤1) 시추공별 개량범위: 개량심도 = 최하단 NG층 하단, 면적 = 영향면적."""
    bhs = list(judge_df["Borehole"])
    area, tot, how = influence_areas(coords_df, bhs, site_area)
    rows = []
    for bh in bhs:
        ng = pd.concat([df_spt[(df_spt["Borehole"] == bh) & (df_spt["대상층여부"] == "대상층") & (df_spt["FS"] <= 1.0)],
                        df_vs[(df_vs["Borehole"] == bh) & (df_vs["대상층여부"] == "대상층") & (df_vs["FS"] <= 1.0)]])
        if ng.empty:
            continue
        ztop, zbot = float(ng["시작심도(m)"].min()), float(ng["종료심도(m)"].max())
        rows.append({"Borehole": bh, "NG 상단(m)": ztop, "개량심도(m)": zbot, "개량두께(m)": zbot - ztop,
                     "평균FC(%)": float(ng["FC(%)"].mean()), "영향면적(m²)": area.get(bh, 0.0),
                     "개량체적(m³)": area.get(bh, 0.0) * zbot})
    return pd.DataFrame(rows), tot, how


def _feasible(m, zbot, fc, urban, narrow):
    if zbot > m["최대심도(m)"] + 1e-9 or fc > m["FC상한(%)"] + 1e-9:
        return False
    if urban and m["소음진동"] in ("높음", "매우 높음"):
        return False
    if narrow and m["협소부지"] != "가능":
        return False
    return True


def remediation_alternatives(zones, cost_db, urban=False, narrow=False, crust_df=None, overhead=1.0):
    """대안 1: 단일공법 최소공사비 / 대안 2: 심도별 혼합(얕은 구간 ≤5m 경제공법 + 깊은 구간 심층공법)
    대안 3: 저소음·저진동(시공환경 우선) 단일공법. 각 대안의 공사비 = Σ 영향면적 × 개량심도 × 단가."""
    db = [dict(m) for m in cost_db]
    out = {"single": [], "alts": []}
    if zones is None or zones.empty:
        return out
    vol_total = float(zones["개량체적(m³)"].sum())
    _zc = lambda m, area, d: float(area) * remed_cost_per_area(m, d)
    for m in db:
        ok = all(_feasible(m, r["개량심도(m)"], r["평균FC(%)"], urban, narrow) for _, r in zones.iterrows())
        cost = sum(_zc(m, r["영향면적(m²)"], r["개량심도(m)"]) for _, r in zones.iterrows())
        out["single"].append({"공법": m["공법"], "분류": m["분류"], "적용가능": ok,
                              "부적합 사유": "" if ok else _reason(m, zones, urban, narrow),
                              "개량체적(m³)": vol_total, "단가": _num(m.get("단가")), "단위": m.get("단위", ""),
                              "산정조건": remed_quantity_text(m),
                              "환산단가(원/m³)": cost / vol_total if vol_total > 0 else np.nan,
                              "개략공사비(백만원)": cost / 1e6 if ok else np.nan})
    sing = pd.DataFrame(out["single"]).sort_values(["적용가능", "개략공사비(백만원)"], ascending=[False, True])
    out["single"] = sing
    feas = sing[sing["적용가능"]]
    if len(feas):
        b = feas.iloc[0]
        out["alts"].append({"대안": "대안 1  단일공법 최소공사비", "공법 구성": b["공법"],
                            "개량면적(m²)": float(zones["영향면적(m²)"].sum()), "개량체적(m³)": vol_total,
                            "개략공사비(백만원)": b["개략공사비(백만원)"],
                            "특징": "전 구간 동일 공법 - 품질관리·공정 단순"})
    # 대안 2: 심도별 혼합
    mix_cost, parts, ok_all = 0.0, {}, True
    for _, r in zones.iterrows():
        cands = [m for m in db if _feasible(m, r["개량심도(m)"], r["평균FC(%)"], urban, narrow)
                 and (r["개량심도(m)"] <= SHALLOW_LIMIT_M or m["최대심도(m)"] > SHALLOW_LIMIT_M)]
        if not cands:
            ok_all = False
            continue
        m = min(cands, key=lambda x: remed_cost_per_area(x, r["개량심도(m)"]))
        mix_cost += _zc(m, r["영향면적(m²)"], r["개량심도(m)"])
        parts.setdefault(m["공법"], []).append(r["Borehole"])
    if ok_all and parts:
        out["alts"].append({"대안": "대안 2  심도별 혼합공법", "공법 구성": " / ".join(
            f"{k}({len(v)}개소)" for k, v in parts.items()),
            "개량면적(m²)": float(zones["영향면적(m²)"].sum()), "개량체적(m³)": vol_total,
            "개략공사비(백만원)": mix_cost / 1e6, "특징": f"개량심도 {SHALLOW_LIMIT_M:g}m 이하 구간은 경제적 천층공법 적용",
            "_parts": parts})
    # 대안 3: 저소음·저진동
    quiet = sing[sing["적용가능"] & sing["공법"].isin(
        [m["공법"] for m in db if m["소음진동"] in ("없음", "낮음")])]
    if len(quiet):
        q = quiet.iloc[0]
        out["alts"].append({"대안": "대안 3  시공환경 우선(저소음·저진동)", "공법 구성": q["공법"],
                            "개량면적(m²)": float(zones["영향면적(m²)"].sum()), "개량체적(m³)": vol_total,
                            "개략공사비(백만원)": q["개략공사비(백만원)"],
                            "특징": "인접시설·민원 리스크 최소화"})
    # 대안 4: 현장 실무형 - 상부 비액상화층(H1) 확보 (지표 피해 억제 목적, 하부 액상화층은 기초 설계 반영)
    if crust_df is not None and len(crust_df):
        by = {m["공법"]: m for m in db}
        cost4, vol4, parts4, notes4 = 0.0, 0.0, {}, []
        for _, z in zones.iterrows():
            c = crust_df[crust_df["Borehole"] == z["Borehole"]]
            if c.empty:
                continue
            c = c.iloc[0]
            area = float(z["영향면적(m²)"])
            d = float(c["필요 처리심도(m)"])
            if d <= 0:
                parts4.setdefault("개량 없음(기초 설계 반영)", []).append(z["Borehole"])
                continue
            cands = []
            if "상부 치환" in c["현장 실무형 대책"]:
                cands += [by[k] for k in ("쇄석치환", "치환(양질토)") if k in by][:1]
            if "동다짐" in c["현장 실무형 대책"] and "동다짐" in by:
                cands.append(by["동다짐"])
            if not cands:   # 상부 처리 불가 → 해당 시추공은 전 심도 심층개량(최소 단가 공법)
                deep = [m for m in db if _feasible(m, z["개량심도(m)"], z["평균FC(%)"], urban, narrow)]
                if not deep:
                    notes4.append(f"{z['Borehole']} 적용공법 없음")
                    continue
                m = min(deep, key=lambda x: remed_cost_per_area(x, z["개량심도(m)"]))
                dd = float(z["개량심도(m)"])
                parts4.setdefault(f"{m['공법']}(전심도)", []).append(z["Borehole"])
            else:
                m = min(cands, key=lambda x: remed_cost_per_area(x, d))
                dd = d
                parts4.setdefault(f"{m['공법']}(GL-{d:.1f}m)", []).append(z["Borehole"])
            cost4 += _zc(m, area, dd)
            vol4 += area * dd
        if parts4:
            out["alts"].append({"대안": "대안 4  현장 실무형(상부 비액상화층 확보)",
                                "공법 구성": " / ".join(f"{k}({len(v)}개소)" for k, v in parts4.items()),
                                "개량면적(m²)": float(zones["영향면적(m²)"].sum()), "개량체적(m³)": vol4,
                                "개략공사비(백만원)": cost4 / 1e6,
                                "특징": "지표 피해 억제(Ishihara 1985) 목적, 하부 액상화층은 KDS 4.7(7) 시설물 안정성 평가·"
                                        "기초 설계에 반영" + ("; " + ", ".join(notes4) if notes4 else ""),
                                "_parts": parts4})
    if overhead and abs(overhead - 1.0) > 1e-12:
        for a in out["alts"]:
            a["개략공사비(백만원)"] = a["개략공사비(백만원)"] * overhead
        if isinstance(out["single"], pd.DataFrame) and len(out["single"]):
            out["single"]["개략공사비(백만원)"] = out["single"]["개략공사비(백만원)"] * overhead
    return out


def _reason(m, zones, urban, narrow):
    r = []
    if (zones["개량심도(m)"] > m["최대심도(m)"]).any():
        r.append(f"개량심도 초과(최대 {m['최대심도(m)']:g}m)")
    if (zones["평균FC(%)"] > m["FC상한(%)"]).any():
        r.append(f"세립분 초과(FC ≤ {m['FC상한(%)']:g}%)")
    if urban and m["소음진동"] in ("높음", "매우 높음"):
        r.append("소음·진동 제약")
    if narrow and m["협소부지"] != "가능":
        r.append("협소부지 장비 제약")
    return ", ".join(r)


# ---------------------------------------------------------
# 1. 페이지 기본 설정 및 메인 제목 (내진 관련 아이콘으로 변경)
# ---------------------------------------------------------
st.set_page_config(page_title="액상화 평가", page_icon="🏛️", layout="wide")

st.title("🏛️ 내진설계 - 액상화 평가")
st.markdown("---")

# ---------------------------------------------------------
# 2. 사이드바 설정
# ---------------------------------------------------------
st.sidebar.header("📁 액상화 데이터 입력")
uploaded_file = st.sidebar.file_uploader(
    "액상화 평가 엑셀 파일 (.xlsx)", type=["xlsx", "xls"]
)

st.sidebar.markdown("---")
st.sidebar.header("🌐 설계지진 조건 (KDS 17 10 00:2024)")
kds_zone_label = st.sidebar.selectbox("지진구역", list(KDS_ZONE_FACTOR.keys()), index=0, key="kds_zone")
kds_rp = st.sidebar.selectbox("평균재현주기", list(KDS_RISK_FACTOR.keys()), index=4, key="kds_rp",
                              help="내진1등급 붕괴방지수준 1,000년, 내진특등급 2,400년 등 시설물 성능목표에 따릅니다.")
kds_soil_class = st.sidebar.selectbox(
    "지반종류", ["S2", "S3", "S4", "S5"], index=2, key="kds_soil",
    help="KDS 표 4.2-8 단주기지반증폭계수(Fa) 산정용. S6 지반은 부지 고유 해석 대상이므로 가장 유사한 종류를 선택하세요.",
)
KDS_Z = KDS_ZONE_FACTOR[kds_zone_label]
KDS_I = KDS_RISK_FACTOR[kds_rp]
KDS_S = KDS_Z * KDS_I
KDS_SFA, KDS_FA = kds_sfa(KDS_S, kds_soil_class)
st.sidebar.caption(
    f"S = Z×I = {KDS_Z:.2f}×{KDS_I:.2f} = **{KDS_S:.3f}g**, Fa = {KDS_FA:.3f} → "
    f"지표 최대가속도 하한 S×Fa = **{KDS_SFA:.3f}g** (KDS 4.7(6)① 단서)"
)

st.sidebar.markdown("---")
st.sidebar.header("🌍 지반응답해석 (SHAKE 대체, 선택)")
site_response_mode = st.sidebar.radio(
    "amax/g 산정 방식",
    ["① 결과(amax/g) 직접입력", "② 시간이력 → 자체 등가선형해석"],
    label_visibility="collapsed",
)
uploaded_th_file = None
# 기반암 Vs/단위중량은 시추공마다 KDS 17 10 00:2024 기준(Vs≥760m/s)으로 자동 판정되므로
# 사이드바에서 따로 입력받지 않는다. 시추가 기반암까지 도달하지 못한 극히 예외적인 경우에만
# 쓰이는 내부 폴백값이며, 그런 시추공은 결과 화면에 "기반암미도달"로 표시된다.
_BEDROCK_VS_FALLBACK = 760.0
_BEDROCK_GAMMA_FALLBACK = 21.0
target_z, target_i, target_s, soil_class = KDS_Z, KDS_I, KDS_S, kds_soil_class
prematched_waves = True
matching_basis = ""
auto_scale_enabled = False
real_record_waves = []
if site_response_mode == "② 시간이력 → 자체 등가선형해석":
    uploaded_th_file = st.sidebar.file_uploader(
        "지진파 시간이력 엑셀 (.xlsx)", type=["xlsx", "xls"], key="th_uploader"
    )
    st.sidebar.caption(
        "입력 지진파는 설계자가 이미 목표 응답스펙트럼에 맞추어 정합(match)한 값으로 "
        "간주하여, 프로그램은 별도 적합성 검증이나 보정 없이 그대로 등가선형 "
        "지반응답해석의 입력(기반암 노두운동)으로 사용합니다."
    )
    matching_basis = st.sidebar.text_area(
        "정합 근거 (검토서에 기록됨)",
        value="SeismoArtif / SeismoMatch를 이용하여 목표 설계응답스펙트럼에 정합한 "
        "시간이력을 사용함. 실지진기록은 관측소별 원기록을 기반으로 정합하였으며, "
        "정합 결과를 그대로 적용함(프로그램 내 재보정 미적용).",
        height=120,
        help="사용한 정합 도구, 목표스펙트럼 조건(지반종류·재현주기·지표면/기반암노두 기준), "
        "원기록 출처 등을 적어두면 검토서 3장에 그대로 기록됩니다.",
    )

sr_case_choice = [c["케이스"] for c in SR_CASES]
if site_response_mode == "② 시간이력 → 자체 등가선형해석":
    sr_case_choice = st.sidebar.multiselect(
        "해석 케이스 (입력가정 민감도)",
        [c["케이스"] for c in SR_CASES],
        default=[c["케이스"] for c in SR_CASES],
        help="동적물성곡선·유효변형률비·전단파속도·기반암 가정을 타당한 범위에서 바꾼 케이스를 함께 해석하고, "
        "결과를 비교해 적용 케이스를 선택합니다. 케이스가 많을수록 해석시간이 늘어납니다"
        "(같은 입력은 재계산 없이 저장된 결과를 사용).",
    )

custom_case = None
uploaded_ext_file = None
if site_response_mode == "② 시간이력 → 자체 등가선형해석":
    with st.sidebar.expander("🛠️ 사용자 정의 케이스 (설정값 직접 지정)", expanded=False):
        _cu_on = st.checkbox("사용자 정의 케이스 해석", value=False, key="cu_on")
        _cu_name = st.text_input("케이스 이름", value="C8 사용자 정의", key="cu_name")
        _cu_sr = st.number_input("유효변형률비", 0.30, 1.00, 0.65, 0.05, key="cu_sr")
        _cu_cs = st.number_input("동적곡선 기준변형률 배율", 0.10, 5.00, 1.00, 0.05, key="cu_cs",
                                 help="1보다 크면 선형적(감쇠 작음, 상한), 작으면 비선형적(감쇠 큼, 하한) 곡선")
        _cu_vf = st.number_input("토층 Vs 배율", 0.50, 1.50, 1.00, 0.05, key="cu_vf")
        _cu_bv = st.number_input("기반암 Vs (시추 미도달 시, m/s)", 500.0, 3000.0, 760.0, 10.0, key="cu_bv")
        _cu_sc = st.number_input("입력지진 배율", 0.10, 2.00, 1.00, 0.05, key="cu_sc",
                                 help="1.0 = 업로드한 설계 지진파 그대로. 1.0이 아니면 설계지반운동수준이 바뀌므로 참고용입니다.")
        if _cu_on:
            custom_case = {"id": "C8", "케이스": (_cu_name.strip() or "C8 사용자 정의"),
                           "동적곡선": f"기준변형률 ×{_cu_cs:g}", "curve_shift": float(_cu_cs),
                           "변형률비": float(_cu_sr), "vs_factor": float(_cu_vf), "bedrock_vs": float(_cu_bv),
                           "input_scale": float(_cu_sc), "근거": "사용자 지정 설정값", "custom": True}
            custom_case["위반"] = custom_case_violations(custom_case)
            if custom_case["위반"]:
                st.warning("검토서 적용 범위 밖 설정 → **참고용(보기 전용)**\n\n- " + "\n- ".join(custom_case["위반"]))
            else:
                st.success("설정값이 검토서 적용 범위 이내입니다 (적용 케이스로 선택 가능).")
        st.caption("범위 밖 설정의 결과는 화면에서 비교·확인만 가능하며, 검토서·계산서의 적용 결과로는 선택되지 않습니다.")
    uploaded_ext_file = st.sidebar.file_uploader(
        "비교용: 외부 지진응답해석 결과(amax 포함 입력엑셀)", type=["xlsx", "xls"], key="ext_sr_uploader",
        help="외부 SHAKE 등으로 구한 심도별 amax가 들어 있는 액상화 입력엑셀을 올리면, 프로그램 해석결과와 "
        "심도별 가속도·안전율 차이와 원인을 비교합니다(계산에는 사용하지 않음).",
    )

st.sidebar.markdown("---")
st.sidebar.header("🧱 고소성 점토거동 평가 (선택)")
uploaded_clay_file = st.sidebar.file_uploader(
    "고소성 점토거동 평가 엑셀 파일 (.xlsx)", type=["xlsx", "xls"], key="clay_uploader"
)

st.sidebar.markdown("---")
st.sidebar.header("🔎 예비평가 (KDS 17 10 00:2024)")
apply_n25 = st.sidebar.checkbox(
    "주상도 N값 > 25 지층 평가 생략", value=True,
    help="KDS 17 10 00:2024 액상화 평가 생략조건. N값은 에너지 보정 전 현장 낙하횟수이며, "
    "지층 평균이 아닌 측점(층)별로 판정합니다. 지하수위 상부 지층과 고소성 점토거동(영역 C) 지층은 "
    "항상 생략하고, 평가심도는 GL-20m 이내로 합니다.",
)
N_SKIP = 25.0 if apply_n25 else None

st.sidebar.markdown("---")
st.sidebar.header("📐 본평가 산정 방법")
_csr_opt = st.sidebar.radio(
    "반복전단응력비(CSR) 산정",
    ["최대전단응력 τmax (원칙)", "지표 PGA × rd (KDS 단서, PGA ≥ S×Fa)", "심도별 amax 근사 (참고)"],
    help="KDS 17 10 00:2024 식(4.7-2): CSR = 0.65·(τmax)d,GRA / σv′가 원칙입니다. 같은 조항 단서에 따라 "
    "'지표면 최대지반가속도 × 응력감소계수(rd)'로도 산정할 수 있으나, 이때 지표면 최대가속도는 S×Fa 이상이어야 "
    "하므로 프로그램이 지진파별로 max(PGA, S×Fa)를 적용합니다(rd: Idriss 1999, M = 6.5). "
    "'심도별 amax 근사'(0.65·amax(z)·σv/σv′)는 기존 구조계산서 방식으로 KDS에 명시된 방법이 아니어서 참고용입니다. "
    "① 모드(외부 amax 입력)에서는 τmax가 없으므로 τmax를 고르면 심도별 amax 근사로 계산됩니다.",
)
CSR_METHOD = ("tau" if _csr_opt.startswith("최대전단응력")
              else "pga_rd" if _csr_opt.startswith("지표 PGA") else "amax")
PGA_FLOOR = KDS_SFA if CSR_METHOD == "pga_rd" else None
_vs1_opt = st.sidebar.radio(
    "Vs 방법: Vs1 ≥ Vs1* 처리",
    ["원식 적용 (권장)", "Vs1을 Vs1*로 제한 (안전측)"],
    help="원식: Andrus & Stokoe(2000)에서 Vs1 ≥ Vs1*는 액상화 불가 → 방법2를 무한대로 보고 "
    "CRR = 방법1(Kayen) 값 사용(CRR이 연속적으로 증가). 제한: 조밀한 층도 Vs1* 수준의 CRR로 평가하는 "
    "보수적 가정(문헌에 없는 방식).",
)
VS1_MODE = "original" if _vs1_opt.startswith("원식") else "cap"

st.sidebar.markdown("---")
st.sidebar.header("🗺️ 액상화 분포도 (선택)")
uploaded_coord_file = st.sidebar.file_uploader(
    "시추공 좌표 엑셀 (Borehole, 동향 E, 북향 N)", type=["xlsx", "xls", "csv"], key="coord_uploader",
    help="시추공명(Borehole)과 평면좌표(m) 또는 경도·위도 열이 있으면 액상화 분포도를 작성합니다.",
)
coord_swap_xy = st.sidebar.checkbox(
    "측량좌표(X = 북향, Y = 동향)로 입력함", value=False, key="coord_swap",
    help="열 이름이 X·Y뿐이고 국내 측량성과 표기(X가 북쪽)대로 넣은 경우 체크하면 바꿔 읽습니다.",
)
_tpl_bhs = []
if uploaded_file is not None:
    try:
        _tpl_bhs = list(dict.fromkeys(pd.read_excel(uploaded_file)["Borehole"].dropna().astype(str)))
    except Exception:
        _tpl_bhs = []
    finally:
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
st.sidebar.download_button(
    "📥 좌표 입력양식 받기" + (f" (시추공 {len(_tpl_bhs)}개 포함)" if _tpl_bhs else ""),
    data=build_coord_template(_tpl_bhs), file_name="시추공_좌표_입력양식.xlsx", key="coord_tpl_dl",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    help="액상화 평가 엑셀을 먼저 올리면 시추공명이 채워진 양식을 받습니다.",
)

st.sidebar.markdown("---")
st.sidebar.header("🏗️ 보강대책 산정 현장 조건")
site_location = st.sidebar.radio(
    "현장 입지 조건",
    ["도심지/인접구조물 유 (소음·진동 제약)", "해안/매립지/제약없음 (소음·진동 허용)"],
)
space_constraint = st.sidebar.radio(
    "부지 협소 여부",
    ["일반 부지 (대형장비 가능)", "협소 부지 / 층고제한 (소형장비 필요)"],
)
remed_priority = st.sidebar.radio(
    "공법 추천 기준",
    ["경제성 우선 (현장 실무형)", "균형 (성능·경제성·실적)"],
    help="경제성 우선: 공사비 증액 부담이 큰 현장 여건을 반영해 경제성 비중을 높이고, 상부 쇄석치환·동다짐으로 "
    "비액상화층을 확보할 수 있는지(Ishihara 1985 지표피해 기준)를 함께 검토합니다.",
)

#입력 엑셀의 FC(%)는 시추공 전체에 적용되는 단일 설계값으로 간주합니다
#(SPT Δ(N1)60 보정, Vs CASE1 보정, Vs1* 분류 모두 동일한 FC(%) 값을 사용)
#시추공별로 대표 FC(%)를 입력 엑셀의 FC(%) 컬럼에 넣어주세요


# ---------------------------------------------------------
# 3. 보강공법 DB 정의
# ---------------------------------------------------------
# usage_score(실무 활용빈도, 0~100)는 국내외 액상화 보강 시공 사례 조사를 바탕으로 부여한
# "얼마나 자주 1차 대안으로 채택되는가"에 대한 정성적 지표다(절대적 통계치가 아니라
# 문헌·시방서·시공 실적에 근거한 상대 서열). 근거 요약:
#  - SCP(다짐모래기둥): 국내 해안·매립지 액상화 대책의 사실상 표준 공법으로, 인천항/광양항 등
#    대규모 항만·매립지 현장에서 가장 많이 채택되어 온 공법.
#  - GCP(다짐자갈기둥, Vibro Stone Column): 국제적으로도 "가장 널리 쓰이는 액상화 지반개량 공법"
#    으로 꼽히며(Vibro stone column은 배수기능까지 겸해 FC 제약이 SCP보다 완화됨), 국내에서도
#    SCP와 함께 가장 흔히 검토·채택되는 공법.
#  - DCM(심층혼합처리): 근접구조물·고FC 지반에서 채택 빈도가 높아지는 추세이나, SCP/GCP 대비
#    상대적으로 고가여서 대안이 마땅치 않을 때 채택되는 경향.
#  - 동다짐(Dynamic Compaction): 공항·매립지 등 대규모 개활지 다짐에 오랜 기간 널리 사용되어 온
#    가장 경제적인 공법 중 하나이나, 소음·진동 문제로 적용 가능한 현장이 제한적.
#  - 치환공법: 얕은 심도(구조물 기초 하부 등) 액상화 대책으로 가장 오래되고 흔하게 쓰이는 방법.
#  - 배수쇄석기둥: 단독보다는 SCP/GCP의 보조·대안으로 채택되는 경우가 많아 중간 수준.
#  - 고압분사주입/약액주입: 기존구조물 하부 보강, 국소 보강 등 "다른 공법을 못 쓸 때"의 대안으로
#    채택되는 경우가 많아, 1차 대안으로서의 채택 빈도는 상대적으로 낮음(비용도 높음).
METHODS_DB = [
    {
        "공법명": "치환공법 (Soil Replacement)",
        "분류": "굴착/치환",
        "적용심도": "0 ~ 5m",
        "max_depth": 5.0,
        "FC_제약": "제약 없음",
        "base_score": 82,
        "econ_score": 85,
        "usage_score": 75,
        "활용빈도": "높음 (얕은 심도 구조물 기초 대책의 대표 공법)",
        "시공성": "우수 (불량토 완전 제거 및 확실한 개량)",
        "경제성": "우수 (★☆☆, 천층 한정)",
        "소음진동": "중/높음",
        "협소부지": "불가",
        "특징": "액상화 우려 지층을 굴착 제거하고 양질의 토사로 치환. 심도가 얕을 때 가장 확실함.",
    },
    {
        "공법명": "SCP (다짐모래기둥)",
        "분류": "밀시화/다짐",
        "적용심도": "0 ~ 20m",
        "max_depth": 20.0,
        "FC_제약": "FC < 20%",
        "base_score": 92,
        "econ_score": 95,
        "usage_score": 95,
        "활용빈도": "매우 높음 (국내 항만·매립지 액상화 대책의 표준 공법)",
        "시공성": "보통 (진동/소음 발생, 대형장비 필요)",
        "경제성": "매우 우수 (★☆☆)",
        "소음진동": "높음",
        "협소부지": "불가",
        "특징": "모래기둥을 진동 타입하여 지반 밀시화. 세골함유량이 낮고 대면적 시 공사비 대폭 절감.",
    },
    {
        "공법명": "GCP (다짐자갈기둥)",
        "분류": "밀시화/배수",
        "적용심도": "0 ~ 25m",
        "max_depth": 25.0,
        "FC_제약": "FC < 35%",
        "base_score": 90,
        "econ_score": 88,
        "usage_score": 92,
        "활용빈도": "매우 높음 (국내외에서 가장 보편적인 액상화 개량 공법 중 하나)",
        "시공성": "양호 (쇄석 재료 투입, 배수 효과 겸함)",
        "경제성": "우수 (★★☆)",
        "소음진동": "중/높음",
        "협소부지": "불가",
        "특징": "쇄석/자갈을 투입하여 다짐. 과잉간극수압 배수 기능을 겸비하여 SCP 대비 적용 범위 넓음.",
    },
    {
        "공법명": "DCM (심층혼합처리)",
        "분류": "고결/교반",
        "적용심도": "0 ~ 30m",
        "max_depth": 30.0,
        "FC_제약": "제약 없음",
        "base_score": 88,
        "econ_score": 75,
        "usage_score": 80,
        "활용빈도": "높음 (근접구조물·고FC 지반에서 채택 증가 추세)",
        "시공성": "우수 (무진동/무소음, 확실한 고결)",
        "경제성": "보통 (★★★☆)",
        "소음진동": "없음",
        "협소부지": "불가",
        "특징": "시멘트계 슬러리를 지반과 강제 교반. 점토 및 높은 FC 지층에서도 확실한 지반고결.",
    },
    {
        "공법명": "약액주입 (Grouting)",
        "분류": "고결/주입",
        "적용심도": "0 ~ 15m",
        "max_depth": 15.0,
        "FC_제약": "FC < 15%",
        "base_score": 80,
        "econ_score": 65,
        "usage_score": 45,
        "활용빈도": "낮음~보통 (기존구조물 하부·국소보강 등 대안 공법으로 채택)",
        "시공성": "우수 (소형장비 가능, 기존구조물 하부 가능)",
        "경제성": "고가 (★★★★)",
        "소음진동": "없음",
        "협소부지": "가능",
        "특징": "인접구조물 하부나 협소 부지, 국소 불만족 구간 고결에 최적.",
    },
    {
        "공법명": "고압분사주입 (Jet Grouting)",
        "분류": "고결/분사",
        "적용심도": "0 ~ 25m",
        "max_depth": 25.0,
        "FC_제약": "제약 없음",
        "base_score": 84,
        "econ_score": 68,
        "usage_score": 55,
        "활용빈도": "보통 (도심지·협소부지 등 제약조건에서 채택)",
        "시공성": "양호 (협소공간 및 변형지형 시공 가능)",
        "경제성": "고가 (★★★★)",
        "소음진동": "낮음",
        "협소부지": "가능",
        "특징": "초고압 유체로 지반 절삭 및 혼합. 도심지 및 제약부지 강도 증진에 우수.",
    },
    {
        "공법명": "동다짐 (Dynamic Compaction)",
        "분류": "밀시화/타격",
        "적용심도": "0 ~ 10m",
        "max_depth": 10.0,
        "FC_제약": "FC < 35%",
        "base_score": 78,
        "econ_score": 98,
        "usage_score": 78,
        "활용빈도": "높음 (공항·매립지 등 대규모 개활지 다짐의 대표 공법)",
        "시공성": "양호 (대형 부지 다짐 적합)",
        "경제성": "매우 우수 (★☆☆)",
        "소음진동": "매우 높음",
        "협소부지": "불가",
        "특징": "대형 추를 낙하시켜 타격 다짐. 주변 구조물이 없는 광범위 부지에 한정 적용.",
    },
    {
        "공법명": "배수쇄석기둥 (Drain Column)",
        "분류": "배수",
        "적용심도": "0 ~ 15m",
        "max_depth": 15.0,
        "FC_제약": "FC < 25%",
        "base_score": 85,
        "econ_score": 82,
        "usage_score": 65,
        "활용빈도": "보통 (SCP/GCP의 보조·대안 공법으로 채택되는 경우가 많음)",
        "시공성": "양호 (간극수압 신속 소산)",
        "경제성": "우수 (★★☆)",
        "소음진동": "보통",
        "협소부지": "보통",
        "특징": "지진 시 발생하는 과잉간극수압을 신속히 배수시켜 액상화 상쇄.",
    },
]


# ---------------------------------------------------------
# 4. 액상화 평가 연산
# ---------------------------------------------------------
GAMMA_W = 10.0  # KDS 64 17 00:2024 검토서 기준 물의 단위중량(kN/m3). 실무상 9.81을 쓰기도 하나,
                # 검증 대상 엑셀(액상화_검토_NBH-16_-1.xlsx)의 모든 수식이 10을 사용하고 있어 동일하게 맞춤.

# Vs 방법의 MSF(Magnitude Scaling Factor)는 설계지진 규모(Mw)로부터 KDS 64 17 00:2024
# 기준식에 따라 한 번 산정되면 프로젝트 전체에서 고정되는 값이라, 사용자가 매번 바꿀 필요가
# 없다고 판단하여 사이드바 입력 대신 아래 상수로 고정한다. 다른 설계지진(Mw)을 적용하는
# 프로젝트라면 이 두 값만 다시 계산해서 아래 상수를 바꿔주면 된다.
MSF_VS_CASE1 = 1.2165
MSF_VS_CASE2 = 1.2987



def normalize_input_columns(df):
    """액상화 입력엑셀 열 이름 표기 차이 보정 (예: '지하수위(m)' → '지하수위'). 값의 의미는 _parse_gw_flag 참조."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if "지하수위" not in df.columns:
        for c in df.columns:
            if c.replace(" ", "") in ("지하수위(m)", "지하수위(GL,m)", "GWL", "GWL(m)", "지하수"):
                df = df.rename(columns={c: "지하수위"})
                break
    return df


def _parse_gw_flag(val):
    """'지하수위' 열 해석 (층별 플래그).
    1 또는 '수중' : 해당 층이 지하수위 하부(포화)  → 포화단위중량, 간극수압 고려
    0 또는 '수상' : 해당 층이 지하수위 상부(불포화) → 습윤단위중량, KDS 예비평가 생략조건 ①
    그 밖의 숫자   : 지하수위 심도(GL-m)로 보고 층별로 포화 여부를 판정
    (기존에는 0을 '수위 미상 → 전층 수중'으로 처리해 지표 부근 층까지 포화로 계산하는 오류가 있었음)"""
    if isinstance(val, str):
        if "수중" in val:
            return "submerged", None
        if "수상" in val:
            return "dry", None
    try:
        v = float(val)
    except (TypeError, ValueError):
        return "submerged", None
    if v == 1:
        return "submerged", None
    if v == 0:
        return "dry", None
    return "depth", v


def _solve_n1_60cs(n_raw, fc, sigma_v_eff):
    """(N1)60cs 를 KDS 64 17 00:2024 방식(Idriss & Boulanger 2008 세립분 보정)으로 산정.
    CN = min(1.7, (100/σv')^m), m = 0.784 - 0.0768*sqrt((N1)60cs) 이므로 m 과 (N1)60cs 가
    서로를 참조하는 순환식이다. 검토 엑셀도 실제로는 위에서 아래로 값을 그대로 참조하는 구조라
    한 번의 계산으로 끝나지만(엑셀 자체가 순환참조를 피하려고 특정 값으로 고정해 둔 셀도 있음),
    본 프로그램에서는 합의된 방식대로 수렴할 때까지 반복 계산한다.

    N값(N60)이나 FC가 결측(NaN)이면 계산을 진행할 수 없으므로, 예외를 던지는 대신
    전부 NaN을 반환해서 그 층만 "판정불가"로 처리되고 나머지 층/시추공 계산은
    영향받지 않도록 한다.
    """
    if pd.isna(n_raw) or pd.isna(fc) or pd.isna(sigma_v_eff):
        return np.nan, np.nan, np.nan, np.nan

    sigma_v_eff = max(sigma_v_eff, 0.0001)
    delta_n160 = round(
        float(np.exp(1.63 + 9.7 / (fc + 0.01) - (15.7 / (fc + 0.01)) ** 2)), 3
    )
    n1_60cs = 5.0
    cn = 1.0
    n1_60 = 1
    for _ in range(50):
        m = 0.784 - 0.0768 * np.sqrt(min(max(n1_60cs, 0.0), 46.0))  # KDS 식(4.7-7) (반올림 없음)
        cn = min(1.7, (100.0 / sigma_v_eff) ** m)
        # KDS 식(4.7-5) (N1)60 = CN × N60 - 기존 구조계산서의 ROUNDDOWN(정수 내림)은 기준에 없는 보수적 처리로
        # N이 작은 층에서 (N1)60을 최대 40% 이상 과소평가하므로 연속값을 사용한다.
        n1_60 = float(n_raw * cn)
        new_n1_60cs = n1_60 + delta_n160
        if abs(new_n1_60cs - n1_60cs) < 1e-6:
            n1_60cs = new_n1_60cs
            break
        n1_60cs = new_n1_60cs
    return n1_60, delta_n160, n1_60cs, cn


def _spt_crr(n1_60cs, sigma_v_eff):
    """CRR7.5 (Idriss & Boulanger 2008) × Kσ × MSF (검토 엑셀 액상화(SPT) 시트 수식 그대로).
    최종 CRR과 함께, 엑셀 다운로드용 상세계산 시트에 쓸 중간값(Cσ, Kσ, MSFmax, MSF, CRR7.5)도
    딕셔너리로 반환한다."""
    if pd.isna(n1_60cs) or pd.isna(sigma_v_eff):
        return np.nan, {
            "Csigma": np.nan, "Ksigma": np.nan, "MSFmax": np.nan,
            "MSF": np.nan, "CRR7.5": np.nan,
        }
    sigma_v_eff = max(sigma_v_eff, 0.0001)
    denom = 18.9 - 2.55 * np.sqrt(max(n1_60cs, 0.0))
    c_sigma = min(0.3, 1.0 / denom) if denom > 0 else 0.3
    # Kσ = 1 - Cσ·ln(σv'/Pa) ≤ 1.1, Pa = 대기압 100 kPa (σv'와 같은 단위). 기존 /10은 단위 오류였음.
    k_sigma = min(1.1, 1.0 - c_sigma * np.log(sigma_v_eff / 100.0))
    msf_max = min(2.2, round(1.09 + (n1_60cs / 31.5) ** 2, 3))
    msf = 1.0 + 0.376 * (msf_max - 1.0)
    crr75 = float(
        np.exp(
            (n1_60cs / 14.1)
            + (n1_60cs / 126.0) ** 2
            - (n1_60cs / 23.6) ** 3
            + (n1_60cs / 25.4) ** 4
            - 2.8
        )
    )
    crr75 = min(crr75, 10.0)
    crr = crr75 * k_sigma * msf
    detail = {
        "Csigma": round(c_sigma, 4),
        "Ksigma": round(k_sigma, 4),
        "MSFmax": round(msf_max, 4),
        "MSF": round(msf, 4),
        "CRR7.5": round(crr75, 4),
    }
    return crr, detail


def _vs_crr(vs_val, fc, sigma_v_eff, vs1_mode="original"):
    """Kayen et al.(2013) CASE1/CASE2 방식 (KDS 64 17 00:2024 Vs 방법).
    두 CASE 중 더 작은(더 보수적인) CRR을 최종값으로 사용한다.
    FC(%)는 시추공 전체에 적용하는 단일 설계값으로, Vs1* 분류와 CASE1 보정항에
    동일하게 사용한다.

    - Vs1(=Vs×CV)의 상한은 층마다 달라지는 Vs1*(FC로 결정되는 임계 전단파속도) 값으로 clip한다.
      Vs1이 Vs1*를 넘어서면 그 층은 Vs 기준으로 사실상 비액상화성으로 보는 것과 같은 취지이며,
      CASE1 회귀식이 보정범위를 벗어나 무의미하게 발산하는 것도 함께 방지된다.
    - CV(응력보정계수)는 CASE1과 CASE2에서 상한이 다르다: CASE1은 1.5 이하, CASE2는 1.4 이하.

    최종 CRR과 함께, 엑셀 다운로드용 상세계산 시트에 쓸 중간값(CV, Vs1, Vs1*, CRR7.5,
    CASE1/CASE2 CRR)도 딕셔너리로 반환한다.
    """
    sigma_v_eff = max(sigma_v_eff, 0.0001)

    if pd.isna(vs_val) or pd.isna(fc):
        detail = {
            "CV_CASE1": np.nan, "CV_CASE2": np.nan, "Vs1_CASE1": np.nan,
            "Vs1_CASE2": np.nan, "Vs1_star": np.nan, "CRR7.5_CASE1": np.nan,
            "CRR_CASE1": np.nan, "CRR7.5_CASE2": np.nan, "CRR_CASE2": np.nan,
        }
        return np.nan, np.nan, np.nan, detail

    if fc <= 5:
        vs1_star = 215.0
    elif fc < 35:
        vs1_star = 215.0 - 0.5 * (fc - 5)
    else:
        vs1_star = 200.0

    cv_case1 = min(1.5, (100.0 / sigma_v_eff) ** 0.25)  # KDS 식(4.7-15)
    cv_case2 = min(1.4, (100.0 / sigma_v_eff) ** 0.25)  # KDS 식(4.7-18)

    # Vs1 ≥ Vs1* 처리
    #  original(권장): Andrus & Stokoe(2000) 원식 - Vs1 ≥ Vs1*이면 방법2는 액상화 불가(CRR=∞)로 보고
    #                 CRR = 방법1(Kayen) 값을 사용 → CRR이 Vs1에 대해 연속적으로 증가
    #  cap(안전측):    Vs1을 Vs1*로 제한 - 조밀한 층(Vs1 > Vs1*)도 Vs1* 수준의 CRR로 평가(문헌에 없는 보수적 가정)
    if vs1_mode == "cap":
        vs1_case1 = min(vs_val * cv_case1, vs1_star)
        vs1_case2 = min(vs_val * cv_case2, vs1_star)
    else:
        vs1_case1 = vs_val * cv_case1
        vs1_case2 = vs_val * cv_case2

    crr75_case1 = float(
        np.exp(
            (
                (0.0073 * vs1_case1) ** 2.8011
                - 4.8981  # KDS 17 10 00:2024 식(4.7-13) 원문 상수 (최종 CRR = CRR7.5 × MSF 1.2165)
                - 0.0099 * np.log(sigma_v_eff)
                + 0.0028 * fc
                - 0.4984
            )
            / 1.946
        )
    )
    crr_case1 = crr75_case1 * MSF_VS_CASE1

    if vs1_case2 < vs1_star:
        # KDS 식(4.7-16) 원문은 0.022×{(Vs1/100)² + 2.8×(…)} 로 중괄호 표기되어 있으나, 출처식(Andrus & Stokoe, 2000)과
        # 구조계산서는 0.022를 첫 항에만 곱한다. 중괄호대로면 CRR이 1/4 수준으로 비현실적이므로 출처식을 적용한다.
        crr75_case2 = 0.022 * (vs1_case2 / 100.0) ** 2 + 2.8 * (
            (1.0 / (vs1_star - vs1_case2)) - (1.0 / vs1_star)
        )
    else:
        # 원식: 한계 전단파속도 이상 → 액상화 불가(무한대). 상한 방식: 기존과 동일하게 Vs1* 지점 처리
        crr75_case2 = np.inf if vs1_mode != "cap" else 2.0
    crr_case2 = crr75_case2 * MSF_VS_CASE2

    crr_final = min(crr_case1, crr_case2)
    detail = {
        "CV_CASE1": cv_case1,
        "CV_CASE2": cv_case2,
        "Vs1_CASE1": round(vs1_case1, 2),
        "Vs1_CASE2": round(vs1_case2, 2),
        "Vs1_star": round(vs1_star, 2),
        "CRR7.5_CASE1": round(crr75_case1, 4),
        "CRR_CASE1": round(crr_case1, 4),
        "CRR7.5_CASE2": round(crr75_case2, 4),
        "CRR_CASE2": round(crr_case2, 4),
    }
    return crr_final, vs1_case1, vs1_star, detail


# ---------------------------------------------------------
# 4-1. 고소성 점토거동 유형 지반 판별 (KDS 17 10 00:2024, 4.7 액상화 (4)항)
# ---------------------------------------------------------
# KDS 17 10 00:2024 "④ 고소성의 점토거동 유형 지반: 소성지수(PI), 액성한계(LL),
# 현장함수비(wc)가 그림 4.7-1에서의 영역 C에 해당하는 지반(액상화 발생 가능성은
# 낮지만, 반복연화가 발생하여 급격한 강도저하 및 대변형이 예상되는 지반)"은
# 액상화 평가를 생략할 수 있는 조건 중 하나로 규정되어 있다. 그림 4.7-1은 세립토의
# 액상화 취약성을 PI와 wc/LL로 구분하는 Bray & Sancio(2006) 판별 기준과 동일한
# 형태이며, 영역(Zone) A/B/C 경계는 다음과 같다.
#   Zone A (사질토 거동, 액상화 검토 대상)      : PI < 12  이고 wc/LL > 0.85
#   Zone B (중간영역/실트질, 액상화 가능성 있음) : PI < 18  이고 wc/LL > 0.80  (A 제외)
#   Zone C (고소성 점토거동, 액상화 평가 생략 가능): 그 외 (PI ≥ 18 또는 wc/LL ≤ 0.80)
def classify_clay_behavior(pi, wc, ll):
    """PI, wc, LL로 Bray & Sancio(2006)/KDS 17 10 00:2024 그림4.7-1 영역(A/B/C) 판별."""
    if pd.isna(pi) or pd.isna(wc) or pd.isna(ll) or ll <= 0:
        return None, np.nan, "PI/LL/wc 자료 없음 - 판정 불가"

    ratio = wc / ll

    if pi < 12 and ratio > 0.85:
        zone = "A"
        note = "사질토(Sand-like) 거동 - 액상화 평가 대상"
    elif pi < 18 and ratio > 0.80:
        zone = "B"
        note = "중간영역(Silt-like) 거동 - 액상화 가능성 있음, 실내시험(비배수 반복삼축/반복단순전단시험) 권장"
    else:
        zone = "C"
        note = (
            "고소성 점토거동(Clay-like) - 액상화 평가 생략 가능 "
            "(단, 반복연화에 의한 급격한 강도저하·대변형 가능성은 별도 검토 필요)"
        )

    return zone, round(ratio, 3), note


def normalize_clay_columns(df_clay):
    """점토거동 입력 엑셀의 열 이름 표기 차이(공백, 단위 유무, 영문 등)를 표준 이름으로 맞춘다."""
    alias = {
        "borehole": "Borehole", "시추공": "Borehole", "시추공명": "Borehole", "공번": "Borehole",
        "토층": "토층", "토질": "토층", "지층": "토층", "soil": "토층",
        "시작심도(m)": "시작심도(m)", "시작심도": "시작심도(m)", "상단심도(m)": "시작심도(m)", "심도(m)": "시작심도(m)",
        "from(m)": "시작심도(m)", "종료심도(m)": "종료심도(m)", "종료심도": "종료심도(m)", "하단심도(m)": "종료심도(m)",
        "to(m)": "종료심도(m)", "ll(%)": "LL(%)", "ll": "LL(%)", "액성한계(%)": "LL(%)", "액성한계": "LL(%)",
        "pi": "PI", "pi(%)": "PI", "소성지수": "PI", "소성지수(%)": "PI",
        "wc(%)": "wc(%)", "wc": "wc(%)", "w(%)": "wc(%)", "함수비(%)": "wc(%)", "자연함수비(%)": "wc(%)",
        "현장함수비(%)": "wc(%)", "함수비": "wc(%)",
    }
    ren = {}
    for c in df_clay.columns:
        k = str(c).strip().replace(" ", "").lower()
        if k in alias and alias[k] not in ren.values():
            ren[c] = alias[k]
    out = df_clay.rename(columns=ren).copy()
    out = out.dropna(how="all")
    if "Borehole" in out.columns:
        out = out[out["Borehole"].notna()]
    for c in ("시작심도(m)", "종료심도(m)", "LL(%)", "PI", "wc(%)"):
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce")
    if "종료심도(m)" not in out.columns and "시작심도(m)" in out.columns:
        out["종료심도(m)"] = out["시작심도(m)"]
    if "토층" not in out.columns:
        out["토층"] = "-"
    return out.reset_index(drop=True)


def clay_sample_label(row):
    return (f"{row.get('Borehole', '-')} | {row.get('토층', '-')} | "
            f"GL-{float(row.get('시작심도(m)', np.nan)):.1f}~{float(row.get('종료심도(m)', np.nan)):.1f} m")


def clay_criteria_rows(row):
    """시료 1개에 대한 KDS 그림 4.7-1 판정 근거(조건별 만족 여부)."""
    pi, ll, wc = row.get("PI"), row.get("LL(%)"), row.get("wc(%)")
    ratio = row.get("wc/LL")
    ok = lambda b: "만족" if b else "불만족"
    if pd.isna(pi) or pd.isna(ratio):
        return pd.DataFrame([{"구분": "자료", "조건": "PI, LL, wc 필요", "시료값": "-", "결과": "판정 불가"}])
    return pd.DataFrame([
        {"구분": "영역 A", "조건": "PI < 12", "시료값": f"PI = {pi:.1f}", "결과": ok(pi < 12)},
        {"구분": "영역 A", "조건": "wc/LL > 0.85", "시료값": f"{wc:.1f}/{ll:.1f} = {ratio:.3f}", "결과": ok(ratio > 0.85)},
        {"구분": "영역 B", "조건": "PI < 18", "시료값": f"PI = {pi:.1f}", "결과": ok(pi < 18)},
        {"구분": "영역 B", "조건": "wc/LL > 0.80", "시료값": f"{ratio:.3f}", "결과": ok(ratio > 0.80)},
    ])


def run_clay_behavior_eval(df_clay):
    """점토 거동 판별용 입력(Borehole, 토층, 시작심도, 종료심도, LL, PI, wc)을 받아
    층별 Zone(A/B/C) 판정 결과 데이터프레임을 반환한다."""
    df = normalize_clay_columns(df_clay)

    zones, ratios, notes = [], [], []
    for _, row in df.iterrows():
        zone, ratio, note = classify_clay_behavior(
            row.get("PI"), row.get("wc(%)"), row.get("LL(%)")
        )
        zones.append(zone)
        ratios.append(ratio)
        notes.append(note)

    df["wc/LL"] = ratios
    df["판정영역"] = zones
    df["검토 의견"] = notes
    df["액상화평가"] = df["판정영역"].map(
        {
            "A": "평가 필요(SPT/Vs 방법 적용)",
            "B": "평가 필요(실내시험 병행 권장)",
            "C": "생략 가능(반복연화 별도 검토)",
        }
    )
    return df


CLAY_INTRO_MD = """
### 📝 고소성 점토거동 유형 지반 판별
**내진설계 일반(KDS 17 10 00 : 2024, 4.7 액상화 (4)④)** 에 따라 세립토(점토/실트)층의
소성지수(PI)·액성한계(LL)·현장함수비(wc)로 **사질토 거동(영역 A) / 중간영역(영역 B) / 고소성 점토거동(영역 C)** 을
판별합니다. 영역 C는 액상화 평가를 생략할 수 있으나, 반복연화에 의한 급격한 강도저하·대변형은 별도로 검토해야 합니다.
"""


# Bray & Sancio(2006)/KDS 그림4.7-1 형태의 PI - wc/LL 산점도 (Zone A/B/C 경계 포함)
def make_clay_behavior_chart(df_clay_result, highlight=None, title=None):
    """highlight: 강조할 행 인덱스(개별 시료 평가). 지정하면 나머지 시료는 흐리게 표시한다."""
    fig = go.Figure()

    # KDS 17 10 00:2024 그림4.7-1 영역 경계선 (사용자 제공 좌표 그대로 사용)
    # B영역 경계: (0.8, 0) → (0.8, 18) → (1.4, 18)
    fig.add_trace(
        go.Scatter(
            x=[0.8, 0.8, 1.4],
            y=[0, 18, 18],
            mode="lines",
            line=dict(color="#8E44AD", dash="dash", width=2),
            name="B영역 경계",
            showlegend=False,
            hoverinfo="skip",
        )
    )
    # A영역 경계: (0.85, 0) → (0.85, 12) → (1.4, 12)
    fig.add_trace(
        go.Scatter(
            x=[0.85, 0.85, 1.4],
            y=[0, 12, 12],
            mode="lines",
            line=dict(color="#E67E22", dash="dash", width=2),
            name="A영역 경계",
            showlegend=False,
            hoverinfo="skip",
        )
    )

    fig.add_annotation(x=1.15, y=6, text="Zone A (사질토 거동)", showarrow=False, font=dict(color="#2ECC71"))
    fig.add_annotation(x=1.15, y=15, text="Zone B (중간영역)", showarrow=False, font=dict(color="#F39C12"))
    fig.add_annotation(x=0.4, y=30, text="Zone C (점토거동)", showarrow=False, font=dict(color="#9B59B6"))

    color_map = {"A": "#2ECC71", "B": "#F39C12", "C": "#9B59B6", None: "#7F8C8D"}
    valid = df_clay_result.dropna(subset=["wc/LL", "PI"])
    for zone, sub in valid.groupby("판정영역"):
        fig.add_trace(
            go.Scatter(
                x=sub["wc/LL"],
                y=sub["PI"],
                mode="markers",
                name=f"Zone {zone}",
                marker=dict(size=11 if highlight is None else 8, color=color_map.get(zone, "#7F8C8D"),
                            opacity=1.0 if highlight is None else 0.25),
                text=sub["Borehole"].astype(str) + " / " + sub["시작심도(m)"].astype(str) + "m",
                hovertemplate="%{text}<br>wc/LL=%{x}<br>PI=%{y}<extra></extra>",
            )
        )
    if highlight is not None and highlight in df_clay_result.index:
        r = df_clay_result.loc[highlight]
        if pd.notna(r.get("wc/LL")) and pd.notna(r.get("PI")):
            fig.add_trace(go.Scatter(
                x=[r["wc/LL"]], y=[r["PI"]], mode="markers+text", name="선택 시료",
                marker=dict(size=20, symbol="star", color="#C0392B", line=dict(color="black", width=1)),
                text=[f"Zone {r['판정영역']}"], textposition="top center", textfont=dict(size=14, color="#C0392B"),
                hovertemplate=f"{clay_sample_label(r)}<br>wc/LL=%{{x}}<br>PI=%{{y}}<extra></extra>",
            ))
            fig.add_shape(type="line", x0=r["wc/LL"], x1=r["wc/LL"], y0=0, y1=r["PI"],
                          line=dict(color="#C0392B", dash="dot", width=1))
            fig.add_shape(type="line", x0=0, x1=r["wc/LL"], y0=r["PI"], y1=r["PI"],
                          line=dict(color="#C0392B", dash="dot", width=1))
    fig.update_layout(
        title=title or "<b>소성지수를 바탕으로한 액상화 가능 범위(KDS 17 10 00:2024 그림4.7-1)</b>",
        xaxis=dict(title="현장함수비 / 액성한계 (wc/LL)", range=[0, 1.4]),
        yaxis=dict(title="소성지수 PI (%)", range=[0, 40]),
        template="plotly_white",
        height=520,
        legend=dict(orientation="v", yanchor="top", y=1.0, xanchor="left", x=1.02),
        margin=dict(r=160),
    )
    return fig


def build_clay_excel(df_clay_result):
    """고소성 점토거동 판별 결과 단독 엑셀 (판정표 + 시료별 판정근거 + PI–wc/LL 판별도)."""
    out = io.BytesIO()
    cols = [c for c in ["Borehole", "토층", "시작심도(m)", "종료심도(m)", "LL(%)", "PI", "wc(%)", "wc/LL",
                        "판정영역", "액상화평가", "검토 의견"] if c in df_clay_result.columns]
    with pd.ExcelWriter(out, engine="xlsxwriter", engine_kwargs={"options": {"nan_inf_to_errors": True}}) as w:
        wb = w.book
        ws = wb.add_worksheet("점토거동판별")
        tf = wb.add_format({"bold": True, "font_size": 12})
        hf = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#1F4E78",
                            "font_color": "white", "border": 1, "text_wrap": True})
        cf = wb.add_format({"align": "center", "border": 1})
        nf = wb.add_format({"align": "center", "border": 1, "num_format": "0.000"})
        zf = {z: wb.add_format({"align": "center", "border": 1, "bold": True, "bg_color": c})
              for z, c in (("A", "#D5F5E3"), ("B", "#FCF3CF"), ("C", "#EBDEF0"))}
        ws.write(0, 0, "고소성 점토거동 유형 지반 판별 (KDS 17 10 00:2024 4.7(4)④, 그림 4.7-1)", tf)
        ws.write(1, 0, "판정기준: 영역 A = PI < 12 이고 wc/LL > 0.85 / 영역 B = PI < 18 이고 wc/LL > 0.80 (A 제외) / "
                       "영역 C = 그 외 (액상화 평가 생략 가능, 반복연화 별도 검토)")
        for j, c in enumerate(cols):
            ws.write(3, j, c, hf)
        for i, (_, r) in enumerate(df_clay_result.iterrows(), start=4):
            for j, c in enumerate(cols):
                v = r[c]
                if c == "판정영역":
                    ws.write(i, j, v if isinstance(v, str) else "-", zf.get(v, cf))
                elif isinstance(v, (int, float, np.integer, np.floating)) and pd.notna(v):
                    ws.write_number(i, j, float(v), nf if c == "wc/LL" else cf)
                else:
                    ws.write(i, j, "" if (v is None or (isinstance(v, float) and pd.isna(v))) else str(v), cf)
        ws.set_column(0, 0, 12)
        ws.set_column(1, 1, 18)
        ws.set_column(2, len(cols) - 1, 12)
        ws.set_column(cols.index("검토 의견") if "검토 의견" in cols else len(cols), len(cols), 60)
        n = len(df_clay_result)
        # 판별도: 경계선(별도 데이터 영역) + 영역별 시료
        bx = len(cols) + 2
        ws.write_row(3, bx, ["A경계 x", "A경계 y", "B경계 x", "B경계 y"], hf)
        for k, (ax, ay, bxv, byv) in enumerate([(0.85, 0, 0.8, 0), (0.85, 12, 0.8, 18), (1.4, 12, 1.4, 18)]):
            ws.write_row(4 + k, bx, [ax, ay, bxv, byv])
        ch = wb.add_chart({"type": "scatter", "subtype": "straight"})
        ch.add_series({"name": "영역 A 경계", "categories": ["점토거동판별", 4, bx, 6, bx],
                       "values": ["점토거동판별", 4, bx + 1, 6, bx + 1],
                       "line": {"color": "#E67E22", "dash_type": "dash", "width": 1.75}, "marker": {"type": "none"}})
        ch.add_series({"name": "영역 B 경계", "categories": ["점토거동판별", 4, bx + 2, 6, bx + 2],
                       "values": ["점토거동판별", 4, bx + 3, 6, bx + 3],
                       "line": {"color": "#8E44AD", "dash_type": "dash", "width": 1.75}, "marker": {"type": "none"}})
        if n and "wc/LL" in cols and "PI" in cols:
            ch.add_series({"name": "시료", "categories": ["점토거동판별", 4, cols.index("wc/LL"), 3 + n, cols.index("wc/LL")],
                           "values": ["점토거동판별", 4, cols.index("PI"), 3 + n, cols.index("PI")],
                           "line": {"none": True},
                           "marker": {"type": "circle", "size": 8, "fill": {"color": "#C0392B"}, "border": {"color": "black"}}})
        ch.set_title({"name": "PI – wc/LL 판별도 (KDS 그림 4.7-1)"})
        ch.set_x_axis({"name": "wc/LL", "min": 0, "max": 1.4, "major_gridlines": {"visible": True}})
        ch.set_y_axis({"name": "PI (%)", "min": 0, "max": 40, "major_gridlines": {"visible": True}})
        ch.set_size({"width": 620, "height": 440})
        ws.insert_chart(5 + n, 0, ch)
        # 시추공별 요약
        if "Borehole" in df_clay_result.columns:
            sm = (df_clay_result.assign(판정영역=df_clay_result["판정영역"].fillna("판정불가"))
                  .groupby(["Borehole", "판정영역"]).size().unstack(fill_value=0).reset_index())
            sm.to_excel(w, sheet_name="시추공별요약", index=False, startrow=1)
            w.sheets["시추공별요약"].write(0, 0, "시추공별 판정영역 시료 수", tf)
    return out.getvalue()


def render_clay_section(df_clay_result, key="main"):
    """고소성 점토거동 판별 결과 화면 (전체 판정표·판별도·시추공별 요약·개별 시료 평가)."""
    if df_clay_result is None or df_clay_result.empty:
        st.info("판별할 시료가 없습니다. 입력 열(Borehole, 토층, 시작심도(m), 종료심도(m), LL(%), PI, wc(%))을 확인해주세요.")
        return
    need = [c for c in ("LL(%)", "PI", "wc(%)") if c not in df_clay_result.columns]
    if need:
        st.error(f"필수 열이 없습니다: {', '.join(need)}")
        return
    zone_counts = df_clay_result["판정영역"].value_counts()
    zc0, zc1, zc2, zc3 = st.columns(4)
    zc0.metric("판별 시료", f"{len(df_clay_result)}개")
    zc1.metric("Zone A (사질토 거동)", f"{zone_counts.get('A', 0)}개")
    zc2.metric("Zone B (중간영역)", f"{zone_counts.get('B', 0)}개")
    zc3.metric("Zone C (고소성 점토)", f"{zone_counts.get('C', 0)}개")

    t_all, t_one, t_bh = st.tabs(["📋 전체 판정", "🔬 개별 시료 평가", "🧭 시추공별 요약"])
    with t_all:
        show_cols = [c for c in ["Borehole", "토층", "시작심도(m)", "종료심도(m)", "LL(%)", "PI", "wc(%)", "wc/LL",
                                 "판정영역", "액상화평가"] if c in df_clay_result.columns]
        st.dataframe(df_clay_result[show_cols], use_container_width=True, hide_index=True)
        st.plotly_chart(make_clay_behavior_chart(df_clay_result), use_container_width=True)
    with t_one:
        idx_list = list(df_clay_result.index)
        sel = st.selectbox("평가할 시료 선택", idx_list, format_func=lambda i: clay_sample_label(df_clay_result.loc[i]),
                           key=f"clay_one_{key}")
        r = df_clay_result.loc[sel]
        c_l, c_r = st.columns([3, 2])
        with c_l:
            st.plotly_chart(make_clay_behavior_chart(df_clay_result, highlight=sel,
                                                     title=f"<b>{clay_sample_label(r)}</b>"),
                            use_container_width=True)
        with c_r:
            zone = r.get("판정영역")
            st.markdown(f"#### 판정 결과: **Zone {zone if zone else '-'}**")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("LL (%)", "-" if pd.isna(r.get("LL(%)")) else f"{r['LL(%)']:.1f}")
            m2.metric("PI", "-" if pd.isna(r.get("PI")) else f"{r['PI']:.1f}")
            m3.metric("wc (%)", "-" if pd.isna(r.get("wc(%)")) else f"{r['wc(%)']:.1f}")
            m4.metric("wc/LL", "-" if pd.isna(r.get("wc/LL")) else f"{r['wc/LL']:.3f}")
            st.dataframe(clay_criteria_rows(r), use_container_width=True, hide_index=True)
            msg = {"A": ("error", "사질토 거동 → 액상화 본평가 대상 (SPT·Vs 방법 적용)"),
                   "B": ("warning", "중간 영역 → 액상화 가능성 있음. 본평가와 함께 비배수 반복삼축·반복단순전단 실내시험 권장"),
                   "C": ("success", "고소성 점토거동 → KDS 4.7(4)④에 따라 액상화 평가 생략 가능. 단, 반복연화에 의한 "
                                    "강도저하·대변형은 별도 검토 (내진특등급은 실내시험 필수, KDS 4.7(5))")}.get(zone)
            if msg:
                getattr(st, msg[0])(msg[1])
            else:
                st.info(str(r.get("검토 의견", "판정 불가")))
    with t_bh:
        bhs = list(dict.fromkeys(df_clay_result["Borehole"])) if "Borehole" in df_clay_result.columns else []
        sm = (df_clay_result.assign(판정영역=df_clay_result["판정영역"].fillna("판정불가"))
              .groupby(["Borehole", "판정영역"]).size().unstack(fill_value=0).reindex(bhs).reset_index())
        st.dataframe(sm, use_container_width=True, hide_index=True)
        if bhs:
            bsel = st.selectbox("시추공 선택", bhs, key=f"clay_bh_{key}")
            sub = df_clay_result[df_clay_result["Borehole"] == bsel]
            st.plotly_chart(make_clay_behavior_chart(sub, title=f"<b>{bsel} 점토거동 판별도</b>"),
                            use_container_width=True)
            st.dataframe(sub[[c for c in ["토층", "시작심도(m)", "종료심도(m)", "LL(%)", "PI", "wc(%)", "wc/LL",
                                         "판정영역", "액상화평가"] if c in sub.columns]],
                         use_container_width=True, hide_index=True)


def run_liquefaction_eval(df, n_skip=25.0, clay_zones=None, csr_method="tau", vs1_mode="original",
                          pga_floor=None, rd_M=DESIGN_M):
    """csr_method: "tau"   CSR = 0.65·(τmax)d,GRA/σv′ (KDS 식 4.7-2, 원칙)
                   "amax"  CSR = 0.65·amax(z)·σv/σv′ (심도별 최대가속도 근사, rd = 1)
                   "pga_rd" CSR = 0.65·PGA·σv/σv′·rd (KDS 4.7-2 단서: 지표 PGA + 응력감소계수,
                            PGA는 pga_floor(=S×Fa) 이상이어야 하므로 지진파별로 max(PGA, S×Fa) 적용)"""

    df_proc = df.copy()

    df_proc["시작심도(m)"] = df_proc["심도(GL,m)"]
    df_proc["종료심도(m)"] = df_proc["시작심도(m)"] + df_proc["층두께(m)"]
    df_proc["평균심도(m)"] = (
        df_proc["시작심도(m)"] + df_proc["종료심도(m)"]
    ) / 2.0

    vs_cols = [c for c in df_proc.columns if "Vs" in c]
    vs_col = vs_cols[0] if vs_cols else "Vs(m/s)"

    base_cols = [
        "Borehole",
        "지하수위",
        "토층",
        "심도(GL,m)",
        "층두께(m)",
        "γsat(kN/m3)",
        "γt(kN/m3)",
        "SPT-N",
        "FC(%)",
        "시작심도(m)",
        "종료심도(m)",
        "평균심도(m)",
        # 지반응답해석(SHAKE 대체) 모드에서만 쓰이는 부가 컬럼 - amax/g 지진파 컬럼이
        # 아니므로 wave_cols에서 반드시 제외해야 한다.
        "PI",
        "재료유형",
        "기반암Vs(m/s)",
        "기반암단위중량(kN/m3)",
    ] + vs_cols
    # 입력 엑셀의 지진파 컬럼(Hachinohe, Ofunato 등)은 CSR 그 자체가 아니라
    # 지진응답해석(SHAKE)에서 산출된 amax/g(최대수평지반가속도비) 값이다.
    # CSR = 0.65 × (amax/g) × (σv/σv') 로 별도 계산해야 하며, 이 곱셈이 누락되어 있던 것이
    # 이번 SPT/Vs 결과 오차의 가장 큰 원인이었다.
    wave_cols = [c for c in df_proc.columns if c not in base_cols and not str(c).startswith("τmax_")]

    sand_keywords = [
        "모래",
        "사질",
        "매립",
        "풍화사",
        "SM",
        "SP",
        "SW",
        "SC",
        "SD",
        "SF",
    ]

    spt_list, vs_list = [], []

    for bh, group in df_proc.groupby("Borehole", sort=False):
        group = group.sort_values(by="시작심도(m)").reset_index(drop=True)
        # 지표면(최상단 층) 최대가속도 - "pga_rd" 방식에서 사용
        _surf_pga = {w: (float(group[w].iloc[0]) if pd.notna(group[w].iloc[0]) else np.nan) for w in wave_cols}

        # 층 '상단(시작심도)'에서의 누적 응력 - 검토 엑셀은 CSR/CRR을 각 층의
        # 시작심도(=이전 층까지의 누적치)에서 평가하므로, 중간심도가 아닌
        # 시작심도 기준 누적값을 사용해야 원본과 일치한다.
        cum_sig_v_top = 0.0
        cum_sig_v_eff_top = 0.0

        for idx, row in group.iterrows():
            gw_raw = row["지하수위"]
            h = row["층두께(m)"]
            start_d = row["시작심도(m)"]
            end_d = row["종료심도(m)"]
            avg_d = row["평균심도(m)"]
            soil = str(row["토층"])

            is_sand = any(kw in soil for kw in sand_keywords)

            gt = row["γt(kN/m3)"]
            gsat = row["γsat(kN/m3)"]

            gw_mode, gw_depth = _parse_gw_flag(gw_raw)

            if gw_mode == "submerged":
                # 지하수위 하부 층: 포화단위중량, 유효응력은 수중단위중량(γsat − γw)으로 누적
                is_sat = True
                layer_sig_v_full = h * gsat
                layer_sig_v_eff_full = h * (gsat - GAMMA_W)
            elif gw_mode == "dry":
                # 지하수위 상부 층: 습윤단위중량, 간극수압 없음 (예비평가 생략조건 ①)
                is_sat = False
                layer_sig_v_full = h * gt
                layer_sig_v_eff_full = h * gt
            else:
                gw = gw_depth if gw_depth is not None else 0.0
                is_sat = end_d > gw
                if end_d <= gw:
                    layer_sig_v_full = h * gt
                    layer_sig_v_eff_full = h * gt
                elif start_d >= gw:
                    layer_sig_v_full = h * gsat
                    layer_sig_v_eff_full = h * (gsat - GAMMA_W)
                else:
                    h1 = gw - start_d
                    h2 = end_d - gw
                    layer_sig_v_full = h1 * gt + h2 * gsat
                    layer_sig_v_eff_full = h1 * gt + h2 * (gsat - GAMMA_W)

            # CSR 평가 지점: 해당 층 "시작심도"까지의 누적응력 (검토 엑셀과 동일)
            sig_v = cum_sig_v_top
            sig_v_eff = max(cum_sig_v_eff_top, 0.0001)

            cum_sig_v_top += layer_sig_v_full
            cum_sig_v_eff_top += layer_sig_v_eff_full

            # 예비평가 (KDS 17 10 00:2024 액상화 평가 생략조건) - 층별 판정 사유 기록
            _n_field = row.get("SPT-N")
            # 고소성 점토거동 판별(PI·LL·wc) 결과가 있으면 해당 심도의 영역(A/B/C)을 연동한다.
            # 영역 C → 평가 생략, 영역 A·B → 토층명이 세립토라도 본평가 대상(액상화 가능성 있음)
            _zone = None
            if clay_zones is not None and len(clay_zones):
                _cz = clay_zones[(clay_zones["Borehole"] == row["Borehole"])
                                 & (clay_zones["시작심도(m)"] <= start_d + 1e-6)
                                 & (clay_zones["종료심도(m)"] > start_d + 1e-6)]
                if len(_cz):
                    _zone = _cz["판정영역"].iloc[0]
            if not is_sat:
                prelim = "생략: 지하수위 상부"
            elif start_d > 20.0:
                prelim = "생략: 심도 20m 초과"
            elif _zone == "C":
                prelim = "생략: 고소성 점토거동(영역 C)"
            elif not is_sand and _zone not in ("A", "B"):
                prelim = "비대상: 비사질토(점성토·암반 등)"
            elif n_skip is not None and pd.notna(_n_field) and float(_n_field) > n_skip:
                prelim = f"생략: N > {n_skip:g}"
            else:
                prelim = "본평가 대상"
            is_target = prelim == "본평가 대상"

            # 대상층인데 계산에 필요한 값이 비어 있으면(결측) 이후 CRR/FS가 NaN이 되어
            # "판정불가"로 표시된다. 그 원인을 바로 알 수 있도록 결측 항목을 기록해둔다.
            missing_reasons = []
            if is_target:
                if pd.isna(row.get("SPT-N")):
                    missing_reasons.append("SPT-N")
                if pd.isna(row.get(vs_col)):
                    missing_reasons.append("Vs")
                if pd.isna(row.get("FC(%)")):
                    missing_reasons.append("FC(%)")
                if pd.isna(sig_v_eff):
                    missing_reasons.append("유효응력(γ/지하수위 확인)")
                if wave_cols and all(pd.isna(row.get(w)) for w in wave_cols):
                    missing_reasons.append("지진파(amax/g) 전체 결측")

            wave_csr_dict = {}
            wave_fs_spt_dict = {}
            wave_fs_vs_dict = {}

            stress_ratio = sig_v / sig_v_eff if sig_v_eff > 0 else 0.0
            rd_val = float(idriss_rd(start_d, rd_M)) if csr_method == "pga_rd" else None

            for wcol in wave_cols:
                amax_g = row[wcol]
                _tau = row.get(f"τmax_{wcol}") if csr_method == "tau" else None
                if csr_method == "pga_rd":
                    # CSR = 0.65·PGA·(σv/σv′)·rd, PGA ≥ S×Fa (KDS 17 10 00:2024 식 4.7-2 단서)
                    _pga = _surf_pga.get(wcol, np.nan)
                    if pga_floor is not None:
                        _pga = max(_pga if pd.notna(_pga) else 0.0, float(pga_floor))
                    csr_w = 0.65 * _pga * stress_ratio * rd_val
                elif _tau is not None and pd.notna(_tau) and sig_v_eff > 0.001:
                    # CSR = 0.65·(τmax)d,GRA / σv' - 지반응답해석의 최대전단응력 직접 사용(원칙)
                    csr_w = 0.65 * float(_tau) / sig_v_eff
                else:
                    # CSR = 0.65·(amax/g)·(σv/σv') - 심도별 최대가속도 근사(수정 Seed & Idriss, rd=1)
                    csr_w = 0.65 * amax_g * stress_ratio
                wave_csr_dict[wcol] = csr_w

            max_csr = max(wave_csr_dict.values()) if wave_csr_dict else 0.15
            # 최대 CSR을 유발하는 지진파 = 해당 층에서 안전율(FS)이 최소가 되는 지진파
            # (CRR은 층마다 지진파와 무관하게 동일하므로, CSR이 최대인 파형이 곧 FS 최소 파형)
            governing_wave = (
                max(wave_csr_dict, key=wave_csr_dict.get) if wave_csr_dict else None
            )

            # (A) SPT 방법
            if is_target:
                # 입력의 SPT-N은 현장측정 N치이지만, 별도의 로드길이/샘플러/해머에너지
                # 보정 없이 이미 에너지 보정이 끝난 N60으로 간주하여 그대로 사용한다.
                n60_val = row["SPT-N"]
                fc_val = row["FC(%)"]

                n1_60, delta_n160, n1_60cs, cn_spt = _solve_n1_60cs(
                    n60_val, fc_val, sig_v_eff
                )
                crr_spt, spt_detail = _spt_crr(n1_60cs, sig_v_eff)
                fs_spt = crr_spt / max_csr if max_csr > 0 else np.nan

                for wcol, csr_w in wave_csr_dict.items():
                    wave_fs_spt_dict[wcol] = (
                        round(crr_spt / csr_w, 3) if csr_w > 0 else 99.0
                    )

                status_spt = (
                    "안전(만족)" if round(fs_spt, 3) > 1.0
                    else ("NG(불만족)" if pd.notna(fs_spt) else "판정불가")
                )
            else:
                crr_spt, fs_spt = np.nan, np.nan
                status_spt = "비대상층"
                n1_60, delta_n160, n1_60cs, cn_spt = None, None, None, None
                spt_detail = {}

            # (B) Vs 방법
            if is_target:
                vs_val = row[vs_col]
                fc_val = row["FC(%)"]

                crr_vs, vs1, vs1_star, vs_detail = _vs_crr(vs_val, fc_val, sig_v_eff, vs1_mode=vs1_mode)
                fs_vs = crr_vs / max_csr if max_csr > 0 else np.nan

                for wcol, csr_w in wave_csr_dict.items():
                    wave_fs_vs_dict[wcol] = (
                        round(crr_vs / csr_w, 3) if csr_w > 0 else 99.0
                    )

                status_vs = (
                    "안전(만족)" if round(fs_vs, 3) > 1.0
                    else ("NG(불만족)" if pd.notna(fs_vs) else "판정불가")
                )
            else:
                crr_vs, fs_vs = np.nan, np.nan
                status_vs = "비대상층"
                vs_detail = {}

            base_res = {
                "Borehole": row["Borehole"],
                "토층": row["토층"],
                "시작심도(m)": round(row["시작심도(m)"], 2),
                "종료심도(m)": round(row["종료심도(m)"], 2),
                "평균심도(m)": round(avg_d, 2),
                "FC(%)": row["FC(%)"],
                "σv(kPa)": round(sig_v, 3),
                "σv'(kPa)": round(sig_v_eff, 3),
                "CSR": round(max_csr, 3),
                **({"rd": round(rd_val, 3)} if rd_val is not None else {}),
                "대상층여부": "대상층" if is_target else "비대상층",
                "예비평가": prelim,
                "지배지진파": governing_wave if is_target else None,
                "결측항목": ", ".join(missing_reasons) if missing_reasons else None,
            }
            wave_csr_display = {
                f"CSR_{wcol}": round(v, 4) for wcol, v in wave_csr_dict.items()
            }

            res_spt_row = base_res.copy()
            res_spt_row.update(wave_csr_display)
            res_spt_row.update(
                {
                    "N치(N60)": row["SPT-N"],
                    "CN": cn_spt,
                    "(N1)60": n1_60,
                    "Δ(N1)60": delta_n160,
                    "(N1)60cs": n1_60cs,
                    **spt_detail,
                    "CRR": round(crr_spt, 3) if pd.notna(crr_spt) else None,
                    "FS": round(fs_spt, 3) if pd.notna(fs_spt) else None,
                    "판정": status_spt,
                    **{f"FS_{w}": v for w, v in wave_fs_spt_dict.items()},
                    "Wave_FS_dict": wave_fs_spt_dict,
                }
            )
            spt_list.append(res_spt_row)

            res_vs_row = base_res.copy()
            res_vs_row.update(wave_csr_display)
            res_vs_row.update(
                {
                    "Vs(m/s)": row[vs_col],
                    **vs_detail,
                    "CRR": round(crr_vs, 3) if pd.notna(crr_vs) else None,
                    "FS": round(fs_vs, 3) if pd.notna(fs_vs) else None,
                    "판정": status_vs,
                    **{f"FS_{w}": v for w, v in wave_fs_vs_dict.items()},
                    "Wave_FS_dict": wave_fs_vs_dict,
                }
            )
            vs_list.append(res_vs_row)

    return pd.DataFrame(spt_list), pd.DataFrame(vs_list), wave_cols


# 심도별 지진파 안전율 그래프 생성
def make_depth_fs_profile_chart(target_df, wave_cols, title_method):
    fig = go.Figure()

    symbols = [
        "circle",
        "square",
        "diamond",
        "cross",
        "x",
        "triangle-up",
        "pentagon",
        "hexagram",
    ]
    colors = [
        "#1F77B4",
        "#FF7F0E",
        "#2CA02C",
        "#D62728",
        "#9467BD",
        "#8C564B",
        "#E377C2",
        "#7F7F7F",
    ]

    target_df = target_df.sort_values(by="시작심도(m)").reset_index(drop=True)

    # 연속된 행이 같은 토층이면 하나의 구간으로 묶는다(내부 경계선은 생략하고
    # 그 구간 전체의 시작~끝에서만 경계선을 그린다).
    group_id = (target_df["토층"] != target_df["토층"].shift()).cumsum()
    layer_groups = target_df.groupby(group_id).agg(
        토층=("토층", "first"),
        시작심도=("시작심도(m)", "min"),
        종료심도=("종료심도(m)", "max"),
    )

    depth_boundaries = set(layer_groups["시작심도"]).union(set(layer_groups["종료심도"]))
    for d in depth_boundaries:
        fig.add_hline(
            y=d, line_dash="dot", line_color="lightgray", line_width=1
        )

    # 토층명 라벨 - 묶인 구간의 중앙에, 그래프 오른쪽 바깥에 표기
    for _, grp in layer_groups.iterrows():
        mid = (grp["시작심도"] + grp["종료심도"]) / 2.0
        fig.add_annotation(
            x=1, xref="paper", y=mid, yref="y",
            xanchor="left", yanchor="middle",
            text=str(grp["토층"]), showarrow=False,
            font=dict(size=10, color="#555555"),
            xshift=8,
        )

    for idx, wcol in enumerate(wave_cols):
        fs_vals = [
            row["Wave_FS_dict"].get(wcol, np.nan)
            for _, row in target_df.iterrows()
        ]
        # 안전율은 각 층의 "시작심도(층 상단)" 기준으로 산정하므로, 그래프의 점(symbol)도
        # 층 중간심도가 아니라 시작심도에 그려야 실제 계산 위치와 일치한다.
        depths = target_df["시작심도(m)"].tolist()

        fig.add_trace(
            go.Scatter(
                x=fs_vals,
                y=depths,
                mode="markers",
                name=str(wcol),
                marker=dict(
                    symbol=symbols[idx % len(symbols)],
                    size=10,
                    color=colors[idx % len(colors)],
                ),
            )
        )

    # 안전율 수직 기준선 (FS = 1.0)
    fig.add_vline(
        x=1.0,
        line_dash="dash",
        line_color="red",
        line_width=3,
        annotation_text="기준 안전율 (FS = 1.0)",
        annotation_position="bottom right",
    )

    max_fs_val = 3.0
    for _, row in target_df.iterrows():
        for val in row["Wave_FS_dict"].values():
            if val < 50 and val > max_fs_val:
                max_fs_val = val

    fig.update_layout(
        title=f"<b>심도별 지진파 안전율 ({title_method})</b>",
        xaxis=dict(
            title="안전율 (FS)",
            side="top",
            range=[0, min(max_fs_val + 0.5, 5.0)],
        ),
        yaxis=dict(
            title="심도 (m)", autorange="reversed", gridcolor="whitesmoke"
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.15,
            xanchor="center",
            x=0.5,
        ),
        template="plotly_white",
        height=550,
        margin=dict(r=110),
    )
    return fig


# 요약 바 차트 생성 (안전율 FS 표기)
# 시추공별 최소 안전율 + 그 안전율을 유발한 지진파를 함께 요약
def build_min_fs_summary(df_target):
    all_boreholes = list(dict.fromkeys(df_target["Borehole"]))
    valid = df_target.dropna(subset=["FS"])
    count_srs = df_target.groupby("Borehole")["FS"].count()  # NaN은 자동 제외되고 카운트됨

    rows = []
    if not valid.empty:
        idxmin = valid.groupby("Borehole")["FS"].idxmin()
        min_df = valid.loc[idxmin, ["Borehole", "FS", "지배지진파"]].reset_index(drop=True)
        min_df = min_df.rename(columns={"지배지진파": "최소안전율 지진파"})
        min_df["검토_구간수"] = min_df["Borehole"].map(count_srs).values
        min_df["최종판정"] = min_df["FS"].apply(
            lambda x: "안전(만족)" if x > 1.0 else "NG(불만족)"
        )
        rows.append(min_df)

    # 대상층은 있지만 N값/Vs값 결측 등으로 FS가 전부 NaN인 시추공은 조용히 빠뜨리지 않고
    # "판정불가"로 표시한다(원인 파악을 위해 결측된 항목까지 함께 남긴다).
    valid_boreholes = set(valid["Borehole"].unique()) if not valid.empty else set()
    missing = [bh for bh in all_boreholes if bh not in valid_boreholes]
    if missing:
        reason_map = {}
        if "결측항목" in df_target.columns:
            for bh in missing:
                reasons = df_target.loc[
                    df_target["Borehole"] == bh, "결측항목"
                ].dropna().unique()
                reason_map[bh] = ", ".join(reasons) if len(reasons) else "원인불명"
        rows.append(
            pd.DataFrame(
                {
                    "Borehole": missing,
                    "FS": [np.nan] * len(missing),
                    "최소안전율 지진파": [None] * len(missing),
                    "검토_구간수": [count_srs.get(bh, 0) for bh in missing],
                    "최종판정": [
                        f"판정불가({reason_map.get(bh, '원인불명')})" for bh in missing
                    ],
                }
            )
        )

    min_df = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame(
        columns=["Borehole", "검토_구간수", "FS", "최종판정", "최소안전율 지진파"]
    )
    # 표시 순서: 시추공 / 검토구간 / FS / 최종판정 / 지진파
    return min_df[
        ["Borehole", "검토_구간수", "FS", "최종판정", "최소안전율 지진파"]
    ]


# st.dataframe 표시 전용: 헤더가 표 폭에 맞게 2줄로 보이도록 컬럼명을 다듬는다
# (내부 계산/차트에서는 원래 컬럼명을 그대로 쓰고, 화면에 찍기 직전에만 이 함수를 거친다)
def format_min_fs_for_display(min_df):
    return min_df.rename(
        columns={
            "Borehole": "시추공",
            "검토_구간수": "검토\n구간수",
            "FS": "안전율\nFS",
            "최종판정": "최종\n판정",
            "최소안전율 지진파": "지진파\n(최소 FS)",
        }
    )


def _split_borehole_name(name):
    """'TBB-1' -> ('TBB', 1), 'NBH-16' -> ('NBH', 16) 처럼 시추공명을
    엑셀 다운로드용 '시추/넘버' 두 컬럼으로 분리한다 (샘플 출력양식과 동일한 구조)."""
    m = re.match(r"^(.*?)[-_]?(\d+)\s*$", str(name).strip())
    if m and m.group(1):
        return m.group(1), int(m.group(2))
    return str(name), None


# 엑셀 다운로드용 "결과정리(SPT/Vs)" 시트: 시추공별로 안전율이 최소가 되는 층의
# 지진파 8개 안전율을 모두 펼쳐서 보여준다 (업로드해주신 샘플 출력양식과 동일한 구조:
# 시추 / 넘버 / <지진파 8개> / 시추공 / 최소안전율 / 기준안전율 / 평가).
def build_wave_result_summary(df_target, wave_cols):
    all_boreholes = list(dict.fromkeys(df_target["Borehole"]))
    valid = df_target.dropna(subset=["FS"])
    idxmin = valid.groupby("Borehole")["FS"].idxmin() if not valid.empty else pd.Series(dtype=object)
    rows = []
    for bh, idx in idxmin.items():
        row = df_target.loc[idx]
        prefix, num = _split_borehole_name(bh)
        rec = {"시추": prefix, "넘버": num}
        for w in wave_cols:
            rec[w] = row["Wave_FS_dict"].get(w, np.nan)
        rec["시추공"] = bh
        rec["최소안전율"] = row["FS"]
        rec["기준안전율"] = 1.0
        rec["평가"] = "OK" if row["FS"] > 1.0 else "NG"
        rows.append(rec)

    # 대상층은 있지만 N값/Vs값 결측 등으로 FS가 전부 NaN인 시추공도 조용히 빠뜨리지
    # 않고 "판정불가" 행으로 남긴다(원인 파악용).
    valid_boreholes = set(idxmin.index)
    for bh in all_boreholes:
        if bh in valid_boreholes:
            continue
        prefix, num = _split_borehole_name(bh)
        rec = {"시추": prefix, "넘버": num}
        for w in wave_cols:
            rec[w] = np.nan
        rec["시추공"] = bh
        rec["최소안전율"] = np.nan
        rec["기준안전율"] = 1.0
        if "결측항목" in df_target.columns:
            reasons = df_target.loc[
                df_target["Borehole"] == bh, "결측항목"
            ].dropna().unique()
            rec["평가"] = f"판정불가({', '.join(reasons)})" if len(reasons) else "판정불가"
        else:
            rec["평가"] = "판정불가"
        rows.append(rec)

    cols = ["시추", "넘버"] + list(wave_cols) + [
        "시추공",
        "최소안전율",
        "기준안전율",
        "평가",
    ]
    return pd.DataFrame(rows)[cols] if rows else pd.DataFrame(columns=cols)


def make_fs_chart(fs_df, title_method):
    fig = go.Figure()

    max_val = (
        fs_df["FS"].max() if not fs_df["FS"].dropna().empty else 2.0
    )
    y_upper = max(2.5, max_val + 0.5)

    fig.add_trace(
        go.Bar(
            x=fs_df["Borehole"],
            y=fs_df["FS"],
            name="최소 안전율 (FS)",
            text=fs_df["FS"],
            textposition="auto",
            marker_color=[
                "#2ECC71" if x > 1.0 else "#E74C3C" for x in fs_df["FS"]
            ],
            width=0.4,
        )
    )

    # 안전율 수평 기준선 (y=1.0)
    fig.add_hline(
        y=1.0,
        line_dash="solid",
        line_color="red",
        line_width=3,
        annotation_text="기준 안전율 (FS = 1.0)",
        annotation_position="top right",
    )

    fig.update_layout(
        title=f"<b>시추공별 안전율 검토 ({title_method})</b>",
        xaxis_title="시추공 (Borehole)",
        yaxis_title="안전율 (FS)",
        yaxis=dict(range=[0, y_upper]),
        template="plotly_white",
        height=400,
    )
    return fig


# 종합평가용: 시추공별 SPT/Vs 최소 안전율을 한 차트에 비교
def make_summary_comparison_chart(summary_df):
    fig = go.Figure()

    max_val = max(
        summary_df["FS (SPT)"].max() if not summary_df["FS (SPT)"].dropna().empty else 2.0,
        summary_df["FS (Vs)"].max() if not summary_df["FS (Vs)"].dropna().empty else 2.0,
    )
    y_upper = max(2.5, max_val + 0.5)

    fig.add_trace(
        go.Bar(
            x=summary_df["Borehole"],
            y=summary_df["FS (SPT)"],
            name="SPT 방법",
            text=summary_df["FS (SPT)"],
            textposition="auto",
            marker_color="#3498DB",
        )
    )
    fig.add_trace(
        go.Bar(
            x=summary_df["Borehole"],
            y=summary_df["FS (Vs)"],
            name="Vs 방법",
            text=summary_df["FS (Vs)"],
            textposition="auto",
            marker_color="#9B59B6",
        )
    )

    fig.add_hline(
        y=1.0,
        line_dash="solid",
        line_color="red",
        line_width=3,
        annotation_text="기준 안전율 (FS = 1.0)",
        annotation_position="top right",
    )

    fig.update_layout(
        title="<b>시추공별 SPT / Vs 방법 안전율</b>",
        xaxis=dict(title=dict(text="시추공 (Borehole)", standoff=15)),
        yaxis_title="안전율 (FS)",
        yaxis=dict(range=[0, y_upper]),
        barmode="group",
        template="plotly_white",
        height=480,
        margin=dict(b=70, r=170, t=70),
        legend=dict(
            orientation="v",
            yanchor="top",
            y=1.0,
            xanchor="left",
            x=1.02,
        ),
    )
    return fig


# ---------------------------------------------------------
# 4-2. 엑셀 결과보고서 생성 ('검토결과 출력양식 샘플'과 동등한 수준의 다운로드 파일)
# ---------------------------------------------------------
def build_excel_report(
    df_spt, df_vs, wave_cols, summary_df, df_raw, df_clay_result,
    site_response_log_df=None,
    sr_case_table=None, judge_df=None, csr_method="amax", extra_tables=None,
):
    output = io.BytesIO()
    with pd.ExcelWriter(
        output, engine="xlsxwriter",
        engine_kwargs={"options": {"nan_inf_to_errors": True}},
    ) as writer:
        wb = writer.book

        title_fmt = wb.add_format(
            {"bold": True, "font_size": 12, "align": "left", "valign": "vcenter"}
        )
        header_fmt = wb.add_format(
            {
                "bold": True,
                "align": "center",
                "valign": "vcenter",
                "bg_color": "#1F4E78",
                "font_color": "white",
                "border": 1,
                "text_wrap": True,
            }
        )
        ok_fmt = wb.add_format({"font_color": "#1E8449", "bold": True, "align": "center"})
        ng_fmt = wb.add_format({"font_color": "#C0392B", "bold": True, "align": "center"})
        na_fmt = wb.add_format({"font_color": "#7F8C8D", "italic": True, "align": "center"})
        num_fmt = wb.add_format({"num_format": "0.000", "align": "center"})
        center_fmt = wb.add_format({"align": "center"})

        # (0) 점토거동판별 - 맨 앞에 배치, 업로드했을 경우에만 추가.
        #     KDS 17 10 00:2024 그림4.7-1 기준 Zone A/B/C 판정표 + PI-wc/LL 산점도 차트.
        if df_clay_result is not None and not df_clay_result.empty:
            ws_clay = wb.add_worksheet("점토거동판별")
            writer.sheets["점토거동판별"] = ws_clay
            clay_cols = list(df_clay_result.columns)
            ws_clay.merge_range(
                0, 0, 0, len(clay_cols) - 1,
                "0. 고소성 점토거동 유형 지반 판별 (KDS 17 10 00:2024 그림4.7-1)",
                title_fmt,
            )
            for c, h in enumerate(clay_cols):
                ws_clay.write(1, c, h, header_fmt)
            zone_fmt = {
                "A": wb.add_format({"bg_color": "#D5F5E3", "align": "center"}),
                "B": wb.add_format({"bg_color": "#FCF3CF", "align": "center"}),
                "C": wb.add_format({"bg_color": "#EBDEF0", "align": "center"}),
            }
            for r, (_, row) in enumerate(df_clay_result.iterrows(), start=2):
                for c, h in enumerate(clay_cols):
                    val = row[h]
                    fmt = (
                        zone_fmt.get(row.get("판정영역"), center_fmt)
                        if h == "판정영역"
                        else center_fmt
                    )
                    ws_clay.write(r, c, val if not pd.isna(val) else "", fmt)
            n_clay = len(df_clay_result)
            ws_clay.freeze_panes(2, 1)
            ws_clay.set_column(0, len(clay_cols) - 1, 14)

            if "wc/LL" in clay_cols and "PI" in clay_cols:
                ratio_col = clay_cols.index("wc/LL")
                pi_col = clay_cols.index("PI")
                chart_clay = wb.add_chart({"type": "scatter"})
                chart_clay.add_series(
                    {
                        "name": "층별 판정 결과",
                        "categories": [
                            "점토거동판별", 2, ratio_col, 1 + n_clay, ratio_col,
                        ],
                        "values": ["점토거동판별", 2, pi_col, 1 + n_clay, pi_col],
                        "marker": {"type": "circle", "size": 8, "fill": {"color": "#8E44AD"}},
                    }
                )
                chart_clay.set_title({"name": "PI - wc/LL 판별도 (Zone A/B/C 경계 포함)"})
                chart_clay.set_x_axis({"name": "wc/LL", "min": 0, "max": 1.6})
                chart_clay.set_y_axis({"name": "PI (%)", "min": 0, "max": 40})
                chart_clay.set_size({"width": 560, "height": 420})
                ws_clay.insert_chart(1, len(clay_cols) + 1, chart_clay)

        def write_wave_summary_sheet(sheet_name, title, df_target):
            """'결과정리(SPT/Vs)' 시트: 표(시추/넘버/8개파형/시추공/최소안전율/기준안전율/평가)
            + 시추공별 최소안전율 막대그래프(기준안전율 FS=1.0 보조선 포함)를 함께 기록한다."""
            ws = wb.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws
            ncols = 2 + len(wave_cols) + 4
            ws.merge_range(0, 0, 0, ncols - 1, title, title_fmt)

            if df_target.empty:
                ws.write(1, 0, "평가 대상층이 없습니다.")
                return

            res_df = build_wave_result_summary(df_target, wave_cols)
            headers = list(res_df.columns)
            for c, h in enumerate(headers):
                ws.write(1, c, h, header_fmt)

            n_rows = len(res_df)
            for r, (_, row) in enumerate(res_df.iterrows(), start=2):
                for c, h in enumerate(headers):
                    val = row[h]
                    if h == "평가":
                        ws.write(
                            r, c, val,
                            ok_fmt if val == "OK"
                            else (na_fmt if isinstance(val, str) and "판정불가" in val else ng_fmt),
                        )
                    elif h in ("최소안전율",) or h in wave_cols:
                        ws.write(r, c, val if pd.notna(val) else "", num_fmt)
                    else:
                        ws.write(r, c, val, center_fmt)

            ws.freeze_panes(2, 2)
            ws.set_column(0, 1, 8)
            ws.set_column(2, 1 + len(wave_cols), 11)
            ws.set_column(2 + len(wave_cols), ncols - 1, 12)

            # --- 시추공별 최소 안전율 막대그래프 (+ 기준안전율=1.0 보조선) ---
            bh_col = 2 + len(wave_cols)  # "시추공" 컬럼 위치
            fs_col = bh_col + 1  # "최소안전율" 컬럼 위치
            std_col = bh_col + 2  # "기준안전율" 컬럼 위치
            first_row, last_row = 2, 2 + n_rows - 1

            bar_chart = wb.add_chart({"type": "column"})
            bar_chart.add_series(
                {
                    "name": "최소 안전율(FS)",
                    "categories": [sheet_name, first_row, bh_col, last_row, bh_col],
                    "values": [sheet_name, first_row, fs_col, last_row, fs_col],
                    "data_labels": {"value": True, "num_format": "0.00"},
                    "points": [
                        {
                            "fill": {
                                "color": "#95A5A6" if pd.isna(v)
                                else ("#2ECC71" if v > 1.0 else "#E74C3C")
                            }
                        }
                        for v in res_df["최소안전율"]
                    ],
                }
            )
            line_chart = wb.add_chart({"type": "line"})
            line_chart.add_series(
                {
                    "name": "기준안전율 (FS=1.0)",
                    "categories": [sheet_name, first_row, bh_col, last_row, bh_col],
                    "values": [sheet_name, first_row, std_col, last_row, std_col],
                    "line": {"color": "red", "width": 2.25, "dash_type": "dash"},
                    "marker": {"type": "none"},
                }
            )
            bar_chart.combine(line_chart)
            bar_chart.set_title({"name": f"{title} - 시추공별 최소 안전율"})
            bar_chart.set_x_axis({"name": "시추공 (Borehole)"})
            bar_chart.set_y_axis({"name": "안전율 (FS)", "min": 0})
            bar_chart.set_size({"width": 720, "height": 380})
            bar_chart.set_legend({"position": "bottom"})
            ws.insert_chart(last_row + 3, 0, bar_chart)

        # (1) 결과정리(SPT) / 결과정리(Vs) - 시추공별 최소안전율 + 8개 지진파 안전율 일람
        write_wave_summary_sheet(
            "결과정리(SPT)",
            "1. 표준관입시험(SPT)방법의 액상화 평가",
            df_spt[df_spt["대상층여부"] == "대상층"],
        )
        write_wave_summary_sheet(
            "결과정리(Vs)",
            "2. 전단파속도(Vs)방법의 액상화 평가",
            df_vs[df_vs["대상층여부"] == "대상층"],
        )

        # (2) 종합요약 (탭3 종합평가에 표시되는 표와 동일 내용을 보기 좋게 서식화해서 기록)
        def write_summary_sheet(sheet_name, title, sdf):
            ws = wb.add_worksheet(sheet_name)
            writer.sheets[sheet_name] = ws
            headers = list(sdf.columns)
            ncols = len(headers)
            ws.merge_range(0, 0, 0, ncols - 1, title, title_fmt)

            for c, h in enumerate(headers):
                ws.write(1, c, h, header_fmt)

            for r, (_, row) in enumerate(sdf.iterrows(), start=2):
                for c, h in enumerate(headers):
                    val = row[h]
                    if h.startswith("최종판정"):
                        ws.write(
                            r, c, val,
                            ok_fmt if val == "안전(만족)"
                            else (na_fmt if isinstance(val, str) and "판정불가" in val else ng_fmt),
                        )
                    elif h.startswith("FS"):
                        ws.write(r, c, val if pd.notna(val) else "", num_fmt)
                    else:
                        ws.write(r, c, val, center_fmt)

            ws.freeze_panes(2, 1)
            ws.set_column(0, 0, 14)
            ws.set_column(1, ncols - 1, 16)

        write_summary_sheet(
            "종합요약", "3. 액상화 평가 종합 요약 (SPT / Vs 방법 비교)", summary_df
        )

        # (3) SPT_상세계산 / Vs_상세계산 - 층별 전체 중간계산값 (CN, N1_60, Kσ, MSF,
        #     Vs1, CRR7.5 등 모든 파형별 CSR·FS 포함) - 시추공 컬럼으로 필터링해 확인 가능
        spt_detail_cols = [
            c for c in df_spt.columns if c != "Wave_FS_dict"
        ]
        df_spt[spt_detail_cols].to_excel(
            writer, sheet_name="SPT_상세계산", index=False
        )
        vs_detail_cols = [c for c in df_vs.columns if c != "Wave_FS_dict"]
        df_vs[vs_detail_cols].to_excel(
            writer, sheet_name="Vs_상세계산", index=False
        )
        for sn in ["SPT_상세계산", "Vs_상세계산"]:
            ws = writer.sheets[sn]
            ws.freeze_panes(1, 2)

        # (4) 시추공별 심도-지진파 안전율 그래프 (대상층만, 심도순 정렬)
        #     프로그램 화면(탭1/탭2)의 "심도별 지진파 안전율" 그래프와 동일한 내용을
        #     엑셀 네이티브 차트로도 확인할 수 있도록 시추공마다 별도 시트에 기록한다.
        # 시추공별 심도-안전율 차트 시트는 시추공 1개당 시트 1개가 생성된다.
        # 실무에서는 한 프로젝트에 시추공이 수십~100개 이상 들어오는 경우도 있는데,
        # 그 개수만큼 그대로 시트를 만들면 파일이 지나치게 무거워지고 열람 속도가
        # 느려진다. 계산(run_liquefaction_eval)과 요약표(결과정리/종합요약)는
        # 시추공 수와 무관하게 전부 정상 처리되지만, "차트 시트"만 아래 개수로 제한한다.
        MAX_DEPTH_CHART_SHEETS = 100

        def write_depth_profile_sheets(df_target, method_label, sheet_prefix):
            if df_target.empty:
                return
            boreholes = list(dict.fromkeys(df_target["Borehole"]))
            if len(boreholes) > MAX_DEPTH_CHART_SHEETS:
                note_ws = wb.add_worksheet(f"{sheet_prefix}_안내")
                writer.sheets[f"{sheet_prefix}_안내"] = note_ws
                note_ws.set_column(0, 0, 90)
                note_ws.write(
                    0, 0,
                    f"시추공이 {len(boreholes)}개로 많아, 시추공별 심도-안전율 차트 시트는 "
                    f"파일 용량·로딩 속도 보호를 위해 생성하지 않았습니다 "
                    f"(개별 차트 시트는 시추공 {MAX_DEPTH_CHART_SHEETS}개 이하일 때만 생성됩니다).",
                )
                note_ws.write(
                    1, 0,
                    "모든 시추공의 안전율 계산 결과는 '결과정리(SPT)/(Vs)', '종합요약', "
                    "'SPT_상세계산'/'Vs_상세계산' 시트에서 빠짐없이 확인하실 수 있으며, "
                    "웹앱 화면에서는 시추공을 선택해 개별 심도-안전율 그래프를 계속 조회할 수 있습니다.",
                )
                return
            used_names = set(writer.sheets.keys())
            for bh, g in df_target.groupby("Borehole", sort=False):
                # 안전율은 각 층의 시작심도(층 상단) 기준으로 산정되므로, 차트의 점도
                # 평균심도가 아니라 시작심도로 표시해야 실제 계산 위치와 일치한다.
                g = g.sort_values(by="시작심도(m)").reset_index(drop=True)

                base_name = f"{sheet_prefix}_{bh}"
                sheet_name = base_name[:31]
                suffix = 1
                while sheet_name in used_names:
                    sheet_name = f"{base_name[: 31 - len(str(suffix)) - 1]}_{suffix}"[:31]
                    suffix += 1
                used_names.add(sheet_name)

                ws = wb.add_worksheet(sheet_name)
                writer.sheets[sheet_name] = ws
                ws.merge_range(
                    0, 0, 0, len(wave_cols) + 2,
                    f"{bh} 시추공 심도별 지진파 안전율 ({method_label})",
                    title_fmt,
                )
                depth_headers = ["시작심도(m)", "토층"] + [str(w) for w in wave_cols]
                for c, h in enumerate(depth_headers):
                    ws.write(1, c, h, header_fmt)
                for r, (_, row) in enumerate(g.iterrows(), start=2):
                    ws.write(r, 0, row["시작심도(m)"], num_fmt)
                    ws.write(r, 1, row["토층"], center_fmt)
                    for c, w in enumerate(wave_cols, start=2):
                        v = row["Wave_FS_dict"].get(w, None)
                        ws.write(r, c, v if (v is not None and pd.notna(v)) else "", num_fmt)

                n_rows = len(g)
                first_row, last_row = 2, 2 + n_rows - 1
                ws.set_column(0, 1, 13)
                ws.set_column(2, len(wave_cols) + 1, 12)

                scatter = wb.add_chart({"type": "scatter", "subtype": "marker_only"})
                marker_types = [
                    "circle", "square", "diamond", "triangle",
                    "x", "star", "short_dash", "plus",
                ]
                for c, w in enumerate(wave_cols, start=2):
                    scatter.add_series(
                        {
                            "name": str(w),
                            "categories": [sheet_name, first_row, c, last_row, c],
                            "values": [sheet_name, first_row, 0, last_row, 0],
                            "marker": {
                                "type": marker_types[(c - 2) % len(marker_types)],
                                "size": 7,
                            },
                        }
                    )
                scatter.set_title({"name": f"{bh} - 심도별 지진파 안전율 ({method_label})"})
                scatter.set_x_axis({"name": "안전율 (FS)", "min": 0})
                scatter.set_y_axis({"name": "심도 (m, 층 상단 기준)", "reverse": True})
                scatter.set_size({"width": 560, "height": 480})
                scatter.set_legend({"position": "bottom"})
                ws.insert_chart(1, len(wave_cols) + 3, scatter)

        # (4) 시추공별 계산서 시트 - 구조계산서·출력양식과 같은 구성
        #     가) 평가 방법(수식) → 나) 간이평가 결과: (가) 지진파별 심도별 CSR, (나) 심도별 CRR, (다) 안전율 + 그래프
        sec_fmt = wb.add_format({"bold": True, "font_size": 10.5})
        txt_fmt = wb.add_format({"font_size": 9.5})
        h2_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#D9E2F3",
                                "border": 1, "text_wrap": True, "font_size": 9})
        c_fmt = wb.add_format({"align": "center", "border": 1, "font_size": 9})
        n1_fmt = wb.add_format({"align": "center", "border": 1, "font_size": 9, "num_format": "0.0"})
        n2_fmt = wb.add_format({"align": "center", "border": 1, "font_size": 9, "num_format": "0.00"})
        n3_fmt = wb.add_format({"align": "center", "border": 1, "font_size": 9, "num_format": "0.000"})
        ok_c = wb.add_format({"align": "center", "border": 1, "font_size": 9, "bold": True, "font_color": "#1E8449"})
        ng_c = wb.add_format({"align": "center", "border": 1, "font_size": 9, "bold": True, "font_color": "#C0392B",
                              "bg_color": "#FADBD8"})
        na_c = wb.add_format({"align": "center", "border": 1, "font_size": 8, "font_color": "#7F8C8D"})

        def _w(ws, r, c, v, fmt):
            if v is None or (isinstance(v, (float, np.floating)) and not np.isfinite(v)):
                ws.write(r, c, "-", c_fmt)
            else:
                ws.write(r, c, v, fmt)

        def _sheet_name(bh, kind):
            nm = re.sub(r"[\[\]:*?/\\\\]", "", f"({bh})_{kind}")[:31]
            k = 1
            while nm in writer.sheets:
                nm = (re.sub(r"[\[\]:*?/\\\\]", "", f"({bh})_{kind}")[:28] + f"_{k}")[:31]
                k += 1
            return nm

        tau_on = csr_method == "tau" and any(str(c).startswith("τmax_") for c in df_raw.columns)
        raw_idx = df_raw.assign(_z=df_raw["심도(GL,m)"].astype(float).round(3)).set_index(["Borehole", "_z"])

        def write_calc_sheet(bh, kind):
            dfm = df_spt if kind == "SPT" else df_vs
            g = dfm[dfm["Borehole"] == bh].sort_values("시작심도(m)")
            if g.empty:
                return
            sn = _sheet_name(bh, kind)
            ws = wb.add_worksheet(sn)
            writer.sheets[sn] = ws
            ws.set_column(0, 0, 13)
            ws.set_column(1, 30, 9.5)
            r = 0
            ws.write(r, 0, ("1) SPT-N값을 이용한 액상화 평가" if kind == "SPT" else "2) 전단파속도를 이용한 액상화 간이평가")
                     + f"  -  {bh}", title_fmt); r += 2
            ws.write(r, 0, "가) 평가 방법", sec_fmt); r += 1
            lines = ["(1) 반복전단응력비(CSR) 산정 (내진설계 일반, KDS 64 17 00 : 2024)",
                     "     CSR = 0.65 × (τmax)d,GRA / σv'     (τmax : 지반응답해석 최대전단응력, σv' : 평가 깊이의 연직유효응력)",
                     ("     ※ 본 계산은 지반응답해석에서 직접 산정한 τmax를 적용" if tau_on else
                      "     ※ 본 계산은 τmax ≈ σv × amax(z)/g 근사(수정 Seed & Idriss, rd = 1)를 적용"),
                     "(2) 반복저항응력비(CRR) 산정"]
            if kind == "SPT":
                lines += ["     CRR7.5 = exp[(N1)60cs/14.1 + ((N1)60cs/126)² − ((N1)60cs/23.6)³ + ((N1)60cs/25.4)⁴ − 2.8]",
                          "     (N1)60cs = (N1)60 + Δ(N1)60,   (N1)60 = CN × N60,   Δ(N1)60 = exp[1.63 + 9.7/(FC+0.01) − (15.7/(FC+0.01))²]",
                          "     CN = (Pa/σv')^m ≤ 1.7,   m = 0.784 − 0.0768√(N1)60cs,   Pa = 100 kPa",
                          "     CRR = CRR7.5 × Kσ × MSF,   MSF = 1 + 0.376 × (MSFmax − 1),   MSFmax = 1.09 + ((N1)60cs/31.5)² ≤ 2.2",
                          "     Kσ = 1 − Cσ × ln(σv'/Pa) ≤ 1.1,   Cσ = 1 / (18.9 − 2.55√(N1)60cs) ≤ 0.3"]
            else:
                lines += ["     CASE1 : CRR7.5 = exp{[(0.0073 × Vs1)^2.8011 − 4.8981 − 0.0099 × ln(σv') + 0.0028 × FC − 0.4984] / 1.946}   (KDS 17 10 00 식 4.7-13)",
                          "             Vs1 = CV × Vs,   CV = (Pa/σv')^0.25 ≤ 1.5",
                          "     CASE2 : CRR7.5 = 0.022 × (Vs1/100)² + 2.8 × [1/(V*s1 − Vs1) − 1/V*s1],   CV ≤ 1.4   (Andrus & Stokoe, 2000)",
                          "             V*s1 = 215 (FC ≤ 5%),  215 − 0.5(FC − 5) (5% < FC < 35%),  200 (FC ≥ 35%) m/s",
                          f"     CRR = min(CRR7.5,CASE1 × {MSF_VS_CASE1:.4f},  CRR7.5,CASE2 × {MSF_VS_CASE2:.4f})   (MSF : 설계지진 규모 M = 6.5)"]
            lines += ["(3) 안전율 :  F.S = CRR / CSR > 1.0  →  액상화에 대하여 안전"]
            for t in lines:
                ws.write(r, 0, t, txt_fmt); r += 1
            r += 1
            ws.write(r, 0, f"나) 액상화 간이평가 결과  -  {bh}", sec_fmt); r += 1
            # (가) 지진파별 심도별 CSR
            ws.write(r, 0, "(가) 심도별 반복전단응력비(CSR)", sec_fmt); r += 1
            for w in wave_cols:
                ws.write(r, 0, f"■ {w}", txt_fmt); r += 1
                hdr = ["토층", "심도\n(GL,m)", "두께\n(m)", "γsat\n(kN/m³)", "γt\n(kN/m³)", "σv\n(kN/m²)", "σv'\n(kN/m²)",
                       "amax / g", "τmax\n(kN/m²)", "τd / σv'\n(CSR)", "비고"]
                for c, h in enumerate(hdr):
                    ws.write(r, c, h, h2_fmt)
                ws.set_row(r, 26); r += 1
                for _, x in g.iterrows():
                    key = (bh, round(float(x["시작심도(m)"]), 3))
                    rw = raw_idx.loc[key] if key in raw_idx.index else None
                    if isinstance(rw, pd.DataFrame):
                        rw = rw.iloc[0]
                    tau = rw.get(f"τmax_{w}") if (rw is not None and tau_on) else None
                    _w(ws, r, 0, x["토층"], c_fmt)
                    _w(ws, r, 1, float(x["시작심도(m)"]), n1_fmt)
                    _w(ws, r, 2, float(rw["층두께(m)"]) if rw is not None else None, n1_fmt)
                    _w(ws, r, 3, float(rw["γsat(kN/m3)"]) if rw is not None else None, n1_fmt)
                    _w(ws, r, 4, float(rw["γt(kN/m3)"]) if rw is not None else None, n1_fmt)
                    _w(ws, r, 5, x["σv(kPa)"], n2_fmt)
                    _w(ws, r, 6, x["σv'(kPa)"], n2_fmt)
                    _w(ws, r, 7, float(rw[w]) if (rw is not None and w in rw.index and pd.notna(rw[w])) else None, n3_fmt)
                    _w(ws, r, 8, float(tau) if (tau is not None and pd.notna(tau)) else None, n2_fmt)
                    csr = x.get(f"CSR_{w}")
                    _w(ws, r, 9, float(csr) if (csr is not None and pd.notna(csr) and x["대상층여부"] == "대상층") else None, n3_fmt)
                    ws.write(r, 10, x.get("예비평가", "") or "", na_c)
                    r += 1
                r += 1
            # (나) 심도별 CRR
            ws.write(r, 0, "(나) 심도별 반복저항응력비(CRR)", sec_fmt); r += 1
            if kind == "SPT":
                cols = [("토층", "토층", c_fmt), ("심도\n(GL,m)", "시작심도(m)", n1_fmt), ("σv'\n(kN/m²)", "σv'(kPa)", n2_fmt),
                        ("FC\n(%)", "FC(%)", n1_fmt), ("N값\n(N60)", "N치(N60)", n1_fmt), ("CN", "CN", n2_fmt),
                        ("(N1)60", "(N1)60", n1_fmt), ("Δ(N1)60", "Δ(N1)60", n3_fmt), ("(N1)60cs", "(N1)60cs", n3_fmt),
                        ("CRR7.5", "CRR7.5", n3_fmt), ("Cσ", "Csigma", n3_fmt), ("Kσ", "Ksigma", n3_fmt),
                        ("MSFmax", "MSFmax", n3_fmt), ("MSF", "MSF", n3_fmt), ("CRR", "CRR", n3_fmt)]
            else:
                cols = [("토층", "토층", c_fmt), ("심도\n(GL,m)", "시작심도(m)", n1_fmt), ("σv'\n(kN/m²)", "σv'(kPa)", n2_fmt),
                        ("FC\n(%)", "FC(%)", n1_fmt), ("Vs\n(m/s)", "Vs(m/s)", n1_fmt), ("CV\n(CASE1)", "CV_CASE1", n2_fmt),
                        ("Vs1\n(CASE1)", "Vs1_CASE1", n1_fmt), ("CRR7.5\n(CASE1)", "CRR7.5_CASE1", n3_fmt),
                        ("CRR\n(CASE1)", "CRR_CASE1", n3_fmt), ("CV\n(CASE2)", "CV_CASE2", n2_fmt),
                        ("Vs1\n(CASE2)", "Vs1_CASE2", n1_fmt), ("V*s1", "Vs1_star", n1_fmt),
                        ("CRR7.5\n(CASE2)", "CRR7.5_CASE2", n3_fmt), ("CRR\n(CASE2)", "CRR_CASE2", n3_fmt), ("CRR", "CRR", n3_fmt)]
            for c, (h, _, _f) in enumerate(cols):
                ws.write(r, c, h, h2_fmt)
            ws.write(r, len(cols), "비고", h2_fmt)
            ws.set_row(r, 26); r += 1
            for _, x in g.iterrows():
                tgt = x["대상층여부"] == "대상층"
                for c, (_, k, f) in enumerate(cols):
                    v = x.get(k)
                    if c >= 2 and not tgt and k not in ("σv'(kPa)", "FC(%)", "Vs(m/s)", "N치(N60)"):
                        v = None
                    _w(ws, r, c, v if (isinstance(v, str) or v is None) else float(v), f)
                ws.write(r, len(cols), x.get("예비평가", "") or "", na_c)
                r += 1
            r += 1
            # (다) 안전율
            ws.write(r, 0, "(다) 심도별 안전율(F.S = CRR / CSR)", sec_fmt); r += 1
            fs_top = r
            hdr = ["토층", "심도\n(GL,m)"] + [str(w) for w in wave_cols] + ["최소 F.S", "지배 지진파", "판정"]
            for c, h in enumerate(hdr):
                ws.write(r, c, h, h2_fmt)
            ws.set_row(r, 26); r += 1
            first = r
            for _, x in g.iterrows():
                tgt = x["대상층여부"] == "대상층"
                ws.write(r, 0, x["토층"], c_fmt)
                ws.write(r, 1, float(x["시작심도(m)"]), n1_fmt)
                for c, w in enumerate(wave_cols, start=2):
                    v = (x.get("Wave_FS_dict") or {}).get(w) if tgt else None
                    _w(ws, r, c, v, n3_fmt)
                fs = x["FS"] if tgt else None
                _w(ws, r, 2 + len(wave_cols), fs, n3_fmt)
                ws.write(r, 3 + len(wave_cols), (x.get("지배지진파") or "-") if tgt else "-", c_fmt)
                if not tgt or fs is None or not np.isfinite(fs):
                    ws.write(r, 4 + len(wave_cols), "비대상" if not tgt else "-", na_c)
                else:
                    ws.write(r, 4 + len(wave_cols), "O.K" if fs > 1.0 else "N.G", ok_c if fs > 1.0 else ng_c)
                r += 1
            last = r - 1
            ch = wb.add_chart({"type": "scatter", "subtype": "marker_only"})
            mk = ["circle", "square", "diamond", "triangle", "x", "star", "short_dash", "plus"]
            for c, w in enumerate(wave_cols, start=2):
                ch.add_series({"name": str(w), "categories": [sn, first, c, last, c], "values": [sn, first, 1, last, 1],
                               "marker": {"type": mk[(c - 2) % len(mk)], "size": 6}})
            ch.set_title({"name": f"{bh} 심도별 안전율 ({kind})", "name_font": {"size": 10}})
            ch.set_x_axis({"name": "안전율 (F.S)", "min": 0, "max": 5})
            ch.set_y_axis({"name": "심도 (GL, m)", "reverse": True, "min": 0})
            ch.set_legend({"position": "bottom", "font": {"size": 8}})
            ch.set_size({"width": 480, "height": 430})
            ws.insert_chart(fs_top, 6 + len(wave_cols), ch)

        _bh_list = list(dict.fromkeys(df_spt["Borehole"]))
        for _bh in _bh_list:
            write_calc_sheet(_bh, "SPT")
        for _bh in _bh_list:
            write_calc_sheet(_bh, "Vs")


        # (4-1) 시추공별 종합판정 / 지진응답해석 케이스 비교
        _tbls = [("종합판정", "시추공별 종합판정 (KDS 17 10 00:2024, FS = CRR/CSR > 1.0)", judge_df),
                 ("지진응답해석_케이스비교", "지진응답해석 케이스(입력가정)별 결과 비교", sr_case_table)]
        _tbls += list(extra_tables or [])
        for _sn, _title, _df in _tbls:
            if _df is None or len(_df) == 0:
                continue
            _ws = wb.add_worksheet(_sn)
            writer.sheets[_sn] = _ws
            _cols = [c for c in _df.columns if c != "id"]
            _ws.merge_range(0, 0, 0, max(len(_cols) - 1, 1), _title, title_fmt)
            for c, h in enumerate(_cols):
                _ws.write(1, c, h, header_fmt)
            for r, (_, row) in enumerate(_df.iterrows(), start=2):
                for c, h in enumerate(_cols):
                    v = row[h]
                    if isinstance(v, (bool, np.bool_)):
                        _ws.write(r, c, "미발생" if v else "발생", ok_fmt if v else ng_fmt)
                    elif isinstance(v, (int, float, np.integer, np.floating)):
                        _ws.write(r, c, float(v) if pd.notna(v) else "", num_fmt)
                    else:
                        _ws.write(r, c, str(v), ng_fmt if "발생 (FS" in str(v) else center_fmt)
            _ws.freeze_panes(2, 1)
            _ws.set_column(0, len(_cols) - 1, 15)

        # (5) 지반응답해석(SHAKE 대체) - 시간이력 모드로 실행했을 때만 추가
        if site_response_log_df is not None and not site_response_log_df.empty:
            ws_sr = wb.add_worksheet("지반응답해석")
            writer.sheets["지반응답해석"] = ws_sr
            sr_cols = list(site_response_log_df.columns)
            ws_sr.merge_range(
                0, 0, 0, len(sr_cols) - 1,
                "5. 등가선형 지반응답해석(SHAKE 대체) 수행 결과 (KDS 17 10 00:2024)",
                title_fmt,
            )
            for c, h in enumerate(sr_cols):
                ws_sr.write(1, c, h, header_fmt)
            for r, (_, row) in enumerate(site_response_log_df.iterrows(), start=2):
                for c, h in enumerate(sr_cols):
                    val = row[h]
                    if h == "수렴여부":
                        ws_sr.write(r, c, bool(val) if pd.notna(val) else "", ok_fmt if val else ng_fmt)
                    else:
                        ws_sr.write(r, c, val if not pd.isna(val) else "", center_fmt)
            ws_sr.freeze_panes(2, 1)
            ws_sr.set_column(0, len(sr_cols) - 1, 16)

        # (6) 입력데이터 - 업로드한 원본 입력 엑셀 (보기 좋게 서식화)
        ws_in = wb.add_worksheet("입력데이터")
        writer.sheets["입력데이터"] = ws_in
        in_cols = list(df_raw.columns)
        ws_in.merge_range(
            0, 0, 0, len(in_cols) - 1, "6. 업로드 입력 데이터 (원본)", title_fmt
        )
        for c, h in enumerate(in_cols):
            ws_in.write(1, c, h, header_fmt)
        for r, (_, row) in enumerate(df_raw.iterrows(), start=2):
            for c, h in enumerate(in_cols):
                val = row[h]
                ws_in.write(
                    r, c, val if not (isinstance(val, float) and pd.isna(val)) else "",
                    num_fmt if isinstance(val, (int, float)) else center_fmt,
                )
        ws_in.freeze_panes(2, 2)
        ws_in.set_column(0, len(in_cols) - 1, 13)

    return output.getvalue()


# ---------------------------------------------------------
# 4-3. 액상화 평가 검토서(Word) 생성 - 설계보고서 「4.11 액상화 평가」 수준 구성
#   1.검토개요 2.내진설계기준 3.지진응답해석(케이스 비교) 4.액상화평가(예비/본평가/종합판정/분포도)
#   5.대책공법 6.결론 부록(상세계산표·입력자료)
# ---------------------------------------------------------
_RP_FONT = "맑은 고딕"
_RP_NAVY = RGBColor(0x1F, 0x3A, 0x5F)
_RP_HDR_FILL = "D9E2F3"
_RP_SUB_FILL = "F2F2F2"
_RP_NG_FILL = "FADBD8"
_RP_WAVE_MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*", "<", ">", "h", "p"]


# ---------------------------------------------------------
# 서식 도우미
# ---------------------------------------------------------
def _rp_font(run, size=10, bold=False, color=None, italic=False):
    run.font.name = _RP_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color if isinstance(color, RGBColor) else RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(attr), _RP_FONT)


def _rp_para(doc, text="", size=10, bold=False, align=None, color=None,
             space_before=0, space_after=3, indent_cm=0.0, line_spacing=1.25):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line_spacing
    if indent_cm:
        pf.left_indent = Cm(indent_cm)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    elif align == "justify":
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if text:
        _rp_font(p.add_run(text), size=size, bold=bold, color=color)
    return p


def _rp_bullet(doc, text, level=0, size=10):
    """설계보고서 관용 기호(◦ / ∙ / -)로 들여쓴 항목."""
    sym = ["◦", "∙", "-"][min(level, 2)]
    p = _rp_para(doc, "", size=size, indent_cm=0.4 + 0.5 * level, space_after=2, align="justify")
    p.paragraph_format.first_line_indent = Cm(-0.4)
    _rp_font(p.add_run(f"{sym} {text}"), size=size)
    return p


def _rp_cell_fill(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    for old in tcpr.findall(qn("w:shd")):
        tcpr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _rp_table_borders(table, outer=12, inner=4):
    tbl = table._tbl
    tblpr = tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge, sz in (("top", outer), ("bottom", outer), ("left", 0), ("right", 0),
                     ("insideH", inner), ("insideV", inner)):
        el = OxmlElement(f"w:{edge}")
        if sz == 0:
            el.set(qn("w:val"), "nil")
        else:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(sz))
            el.set(qn("w:color"), "404040")
        borders.append(el)
    tblpr.append(borders)


def _rp_cell_text(cell, text, size=9, bold=False, align="center", color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.line_spacing = 1.0
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    for k, line in enumerate(str(text).split("\n")):
        if k:
            p.add_run().add_break()
        _rp_font(p.add_run(line), size=size, bold=bold, color=color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def _rp_table(doc, header_rows, body, widths_cm, body_align=None, row_fills=None,
              merges=None, font_size=9, first_col_fill=False, ng_col_idx=None):
    """header_rows: 머리행 목록(병합할 칸은 None), merges: [(r0,c0,r1,c1)] 머리행 병합 범위."""
    ncol = len(widths_cm)
    nh = len(header_rows)
    table = doc.add_table(rows=nh + len(body), cols=ncol)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _rp_table_borders(table)
    for r, row in enumerate(header_rows):
        for c in range(ncol):
            cell = table.cell(r, c)
            _rp_cell_fill(cell, _RP_HDR_FILL)
            if row[c] is not None:
                _rp_cell_text(cell, row[c], size=font_size, bold=True)
    for (r0, c0, r1, c1) in (merges or []):
        a, b = table.cell(r0, c0), table.cell(r1, c1)
        txt = header_rows[r0][c0]
        m = a.merge(b)
        _rp_cell_text(m, txt, size=font_size, bold=True)
        _rp_cell_fill(m, _RP_HDR_FILL)
    from docx.table import _Cell
    trs = table._tbl.findall(qn("w:tr"))
    for i, row in enumerate(body):
        fill = row_fills[i] if row_fills else None
        tcs = trs[nh + i].findall(qn("w:tc"))
        for c in range(ncol):
            # table.cell()은 호출마다 표 전체를 다시 훑어 행이 많으면 급격히 느려지므로 셀 요소에 직접 접근
            cell = _Cell(tcs[c], table)
            val = row[c] if c < len(row) else ""
            al = body_align[c] if body_align else "center"
            is_ng = ng_col_idx is not None and c in ng_col_idx and str(val).startswith("N")
            _rp_cell_text(cell, val, size=font_size, align=al,
                          bold=is_ng, color=(0xC0, 0x39, 0x2B) if is_ng else None)
            if fill:
                _rp_cell_fill(cell, fill)
            elif first_col_fill and c == 0:
                _rp_cell_fill(cell, _RP_SUB_FILL)
    _rp_fix_widths(table, widths_cm)
    # 머리행 반복 (여러 페이지에 걸칠 때)
    for r in range(nh):
        trpr = table.rows[r]._tr.get_or_add_trPr()
        th = OxmlElement("w:tblHeader")
        th.set(qn("w:val"), "true")
        trpr.append(th)
    return table


def _rp_fix_widths(table, widths_cm):
    """한글·LibreOffice는 셀 너비보다 표 격자(tblGrid)를 우선 읽으므로 격자·표폭·셀폭을 모두 지정하고
    고정 레이아웃으로 둔다."""
    tw = [int(round(w * 567)) for w in widths_cm]
    tbl = table._tbl
    tblpr = tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        for old in tblpr.findall(qn(tag)):
            tblpr.remove(old)
    tblw = OxmlElement("w:tblW")
    tblw.set(qn("w:w"), str(sum(tw)))
    tblw.set(qn("w:type"), "dxa")
    tblpr.append(tblw)
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblpr.append(lay)
    grid = tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), tw):
        gc.set(qn("w:w"), str(w))
    for tr in tbl.findall(qn("w:tr")):
        gi = 0
        for tc in tr.findall(qn("w:tc")):
            tcpr = tc.get_or_add_tcPr()
            gs = tcpr.find(qn("w:gridSpan"))
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            for old in tcpr.findall(qn("w:tcW")):
                tcpr.remove(old)
            tcw = OxmlElement("w:tcW")
            tcw.set(qn("w:w"), str(sum(tw[gi:gi + span])))
            tcw.set(qn("w:type"), "dxa")
            tcpr.insert(0, tcw)
            gi += span


def _rp_png(fig, dpi=170):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def _rp_page_field(paragraph):
    run = paragraph.add_run()
    for kind, text in (("begin", None), (None, "PAGE"), ("end", None)):
        if kind:
            fc = OxmlElement("w:fldChar")
            fc.set(qn("w:fldCharType"), kind)
            run._r.append(fc)
        else:
            it = OxmlElement("w:instrText")
            it.set(qn("xml:space"), "preserve")
            it.text = text
            run._r.append(it)
    _rp_font(run, size=9)


class _Numbering:
    """〈표 3.2-1〉, 〈그림 4.4-2〉 형식의 절별 일련번호."""

    def __init__(self):
        self.sec = "1"
        self.tab = 0
        self.fig = 0

    def set(self, sec):
        self.sec, self.tab, self.fig = sec, 0, 0

    def t(self):
        self.tab += 1
        return f"〈표 {self.sec}-{self.tab}〉"

    def f(self):
        self.fig += 1
        return f"〈그림 {self.sec}-{self.fig}〉"


# ---------------------------------------------------------
# 그림 생성 (matplotlib)
# ---------------------------------------------------------
def _rp_fig_input_waves(wd, dtp):
    names = list(wd.keys())
    n = len(names)
    ncol = 2 if n > 3 else 1
    nrow = int(np.ceil(n / ncol))
    fig, axes = plt.subplots(nrow, ncol, figsize=(7.6, max(1.05 * nrow, 2.2)), squeeze=False)
    for k in range(n, axes.size):
        axes.flat[k].axis("off")
    axes = axes.flatten()[:n]
    tmax = max(len(wd[w]) * dtp[w] for w in names)
    for ax, w in zip(axes, names):
        a = np.asarray(wd[w], dtype=float)
        t = np.arange(len(a)) * dtp[w]
        pga = float(np.max(np.abs(a)))
        ax.plot(t, a, lw=0.55, color="#1F3A5F")
        ax.axhline(0, color="#999999", lw=0.4)
        ax.set_xlim(0, tmax)
        lim = max(pga * 1.25, 1e-3)
        ax.set_ylim(-lim, lim)
        ax.set_ylabel("g", fontsize=7, rotation=0, labelpad=8)
        ax.tick_params(labelsize=6.5)
        ax.text(0.005, 0.83, str(w), transform=ax.transAxes, fontsize=8, fontweight="bold")
        ax.text(0.995, 0.83, f"PGA = {pga:.3f} g", transform=ax.transAxes, fontsize=7.5,
                ha="right", color="#943126")
        ax.grid(alpha=0.25, lw=0.4)
    for ax in axes[-ncol:]:
        ax.set_xlabel("시간 (s)", fontsize=8)
    fig.tight_layout(h_pad=0.25, w_pad=1.0)
    return _rp_png(fig)


def _rp_fig_dyn_curves(curves_sand, clay_curve_fn, curves_rock):
    fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.6, 3.1))
    sets = [("사질토 (Seed & Idriss, 1970)", curves_sand, "#1F77B4", "-")]
    for pi, ls in ((15, "--"), (30, "-."), (50, ":")):
        sets.append((f"점성토 PI={pi} (Vucetic & Dobry, 1991)", clay_curve_fn(pi), "#C0392B", ls))
    sets.append(("암반", curves_rock, "#7F8C8D", "-"))
    for name, c, col, ls in sets:
        a1.semilogx(c["strain_pct"], c["g_gmax"], ls, color=col, lw=1.4, label=name)
        a2.semilogx(c["strain_pct"], c["damping_pct"], ls, color=col, lw=1.4, label=name)
    a1.set_ylabel("G/Gmax")
    a2.set_ylabel("감쇠비 D (%)")
    for a in (a1, a2):
        a.set_xlabel("전단변형률 γ (%)")
        a.grid(which="both", alpha=0.3, lw=0.5)
        a.tick_params(labelsize=7.5)
    a1.set_ylim(0, 1.05)
    a1.set_title("전단탄성계수 감소곡선", fontsize=9.5)
    a2.set_title("감쇠비 곡선", fontsize=9.5)
    h, l = a1.get_legend_handles_labels()
    fig.legend(h, l, loc="lower center", ncol=3, fontsize=7.2, frameon=False,
               bbox_to_anchor=(0.5, -0.06))
    fig.tight_layout(rect=(0, 0.1, 1, 1))
    return _rp_png(fig)


def _rp_layer_bands(g):
    """같은 토층이 연속되면 하나로 묶어 (토층명, 시작, 끝) 목록으로 반환."""
    g = g.sort_values("시작심도(m)")
    bands = []
    for _, r in g.iterrows():
        nm, s, e = str(r["토층"]), float(r["시작심도(m)"]), float(r["종료심도(m)"])
        if bands and bands[-1][0] == nm and abs(bands[-1][2] - s) < 1e-6:
            bands[-1][2] = e
        else:
            bands.append([nm, s, e])
    return bands


def _rp_draw_bands(ax, bands, ymax, xmax):
    for i, (nm, s, e) in enumerate(bands):
        if s > ymax:
            break
        e2 = min(e, ymax)
        if i % 2 == 0:
            ax.axhspan(s, e2, color="#F4F6F7", zorder=0)
        ax.axhline(s, color="#7F8C8D", lw=0.6, ls=":", zorder=1)
        ax.text(xmax * 0.985, (s + e2) / 2, nm, ha="right", va="center", fontsize=6.3,
                color="#5D6D7E", zorder=5,
                bbox=dict(boxstyle="round,pad=0.12", fc="white", ec="none", alpha=0.75))
    if bands:
        ax.axhline(min(bands[-1][2], ymax), color="#7F8C8D", lw=0.6, ls=":", zorder=1)


def _rp_grid_axes(n, ncols=3, cell=(2.55, 4.0)):
    ncols = max(1, min(ncols, n))
    nrows = int(np.ceil(n / ncols))
    fig, axes = plt.subplots(nrows, ncols, figsize=(cell[0] * ncols, cell[1] * nrows), squeeze=False)
    return fig, axes


def _rp_fig_amax_profiles(df_raw, wave_cols, boreholes):
    fig, axes = _rp_grid_axes(len(boreholes))
    handles = None
    for k, bh in enumerate(boreholes):
        ax = axes.flat[k]
        g = df_raw[df_raw["Borehole"] == bh].sort_values("심도(GL,m)")
        depth = g["심도(GL,m)"].to_numpy(float)
        xmax = 0.05
        for j, w in enumerate(wave_cols):
            if w not in g.columns:
                continue
            v = pd.to_numeric(g[w], errors="coerce").to_numpy(float)
            if np.all(np.isnan(v)):
                continue
            xmax = max(xmax, np.nanmax(v))
            ax.plot(v, depth, marker=_RP_WAVE_MARKERS[j % len(_RP_WAVE_MARKERS)], ms=2.6,
                    lw=0.8, label=str(w))
        ymax = float(np.nanmax(depth)) if len(depth) else 20
        ax.set_ylim(ymax + 0.5, 0)
        ax.set_xlim(0, xmax * 1.15)
        ax.set_title(str(bh), fontsize=8.5, fontweight="bold")
        ax.xaxis.set_label_position("top")
        ax.xaxis.tick_top()
        ax.set_xlabel("최대가속도 (g)", fontsize=7)
        ax.set_ylabel("심도 (m)", fontsize=7)
        ax.tick_params(labelsize=6.5)
        ax.grid(alpha=0.3, lw=0.4)
        if handles is None:
            handles = ax.get_legend_handles_labels()
    for k in range(len(boreholes), axes.size):
        axes.flat[k].axis("off")
    if handles:
        fig.legend(*handles, loc="lower center", ncol=min(len(wave_cols), 8), fontsize=7,
                   frameon=False, bbox_to_anchor=(0.5, -0.01))
    fig.tight_layout(rect=(0, 0.04, 1, 1))
    return _rp_png(fig)


def _rp_fig_fs_profiles(df_all, wave_cols, boreholes, method):
    fig, axes = _rp_grid_axes(len(boreholes))
    handles = None
    for k, bh in enumerate(boreholes):
        ax = axes.flat[k]
        g_all = df_all[df_all["Borehole"] == bh]
        g = g_all[g_all["대상층여부"] == "대상층"].sort_values("시작심도(m)")
        fs_max = 3.0
        for _, r in g.iterrows():
            for v in (r["Wave_FS_dict"] or {}).values():
                if v is not None and np.isfinite(v) and v < 50:
                    fs_max = max(fs_max, v)
        xmax = min(fs_max + 0.5, 5.0)
        ymax = float(g["종료심도(m)"].max()) if not g.empty else 20.0
        ymax = max(ymax, 5.0)
        ax.axvspan(0, 1.0, color="#FDEDEC", zorder=0)
        _rp_draw_bands(ax, _rp_layer_bands(g_all[g_all["시작심도(m)"] < ymax]), ymax, xmax)
        for j, w in enumerate(wave_cols):
            xs, ys = [], []
            for _, r in g.iterrows():
                v = (r["Wave_FS_dict"] or {}).get(w)
                if v is not None and np.isfinite(v):
                    xs.append(min(v, xmax * 0.97))
                    ys.append(float(r["시작심도(m)"]))
            ax.scatter(xs, ys, s=11, marker=_RP_WAVE_MARKERS[j % len(_RP_WAVE_MARKERS)],
                       label=str(w), zorder=4, linewidths=0.4, edgecolors="#333333")
        ax.axvline(1.0, color="#C0392B", lw=1.3, ls="--", zorder=3)
        ax.set_xlim(0, xmax)
        ax.set_ylim(ymax, 0)
        ax.set_title(str(bh), fontsize=8.5, fontweight="bold")
        ax.xaxis.set_label_position("top")
        ax.xaxis.tick_top()
        ax.set_xlabel("안전율 (FS)", fontsize=7)
        ax.set_ylabel("심도 (m)", fontsize=7)
        ax.tick_params(labelsize=6.5)
        ax.grid(alpha=0.25, lw=0.4, zorder=0)
        if g.empty:
            ax.text(xmax / 2, ymax / 2, "평가 대상층 없음", ha="center", fontsize=8, color="#7F8C8D")
        if handles is None and not g.empty:
            handles = ax.get_legend_handles_labels()
    for k in range(len(boreholes), axes.size):
        axes.flat[k].axis("off")
    if handles:
        h, l = handles
        from matplotlib.lines import Line2D
        h = list(h) + [Line2D([0], [0], color="#C0392B", ls="--", lw=1.3)]
        l = list(l) + ["기준안전율 FS=1.0"]
        fig.legend(h, l, loc="lower center", ncol=min(len(l), 9), fontsize=6.8,
                   frameon=False, bbox_to_anchor=(0.5, -0.01))
    fig.tight_layout(rect=(0, 0.045, 1, 1))
    return _rp_png(fig)


def _rp_fig_bar(min_df, title):
    d = min_df.dropna(subset=["FS"])
    n = max(len(d), 1)
    fig, ax = plt.subplots(figsize=(7.4, 3.0))
    colors = ["#2E86C1" if v > 1.0 else "#C0392B" for v in d["FS"]]
    x = np.arange(len(d))
    cap = 3.0  # 큰 안전율 하나 때문에 나머지 막대가 납작해지지 않도록 표시 상한을 둔다
    ax.bar(x, np.minimum(d["FS"], cap), color=colors, width=0.6 if n > 3 else 0.35, zorder=3)
    if n <= 3:
        ax.set_xlim(-0.9, n - 0.1)
    ax.axhline(1.0, color="#C0392B", ls="--", lw=1.4, zorder=4, label="기준안전율 FS = 1.0")
    if n <= 30:
        for xi, v in zip(x, d["FS"]):
            ax.text(xi, min(v, cap) + 0.03, f"{v:.2f}" + ("▲" if v > cap else ""), ha="center",
                    fontsize=7 if n > 12 else 8)
    ax.set_xticks(x)
    ax.set_xticklabels(d["Borehole"], rotation=0 if n <= 8 else 90, fontsize=7.5 if n <= 30 else 5.5)
    ax.set_ylabel("최소 안전율 (FS)", fontsize=8.5)
    ax.set_ylim(0, min(max(2.0, float(d["FS"].max()) * 1.2), cap * 1.12) if len(d) else 2.0)
    ax.set_title(title, fontsize=9.5, fontweight="bold")
    ax.grid(axis="y", alpha=0.3, zorder=0)
    ax.legend(fontsize=7.5, loc="upper right")
    fig.tight_layout()
    return _rp_png(fig)


def _rp_fig_clay(df_clay):
    fig, ax = plt.subplots(figsize=(6.0, 3.8))
    ax.add_patch(Polygon([(0.85, 0), (1.4, 0), (1.4, 12), (0.85, 12)], fc="#E8F6EF", ec="none"))
    ax.add_patch(Polygon([(0.8, 0), (0.85, 0), (0.85, 12), (1.4, 12), (1.4, 18), (0.8, 18)],
                         fc="#FEF5E7", ec="none"))
    ax.plot([0.8, 0.8, 1.4], [0, 18, 18], color="#B9770E", lw=1.5)
    ax.plot([0.85, 0.85, 1.4], [0, 12, 12], color="#1E8449", lw=1.5)
    ax.text(1.12, 5.5, "영역 A\n(액상화 가능성 높음)", ha="center", fontsize=8, color="#1E8449")
    ax.text(1.12, 14.8, "영역 B (가능성 있음)", ha="center", fontsize=8, color="#B9770E")
    ax.text(0.4, 26, "영역 C\n(액상화 가능성 낮음, 반복연화 검토)", ha="center", fontsize=8, color="#6C3483")
    col = {"A": "#1E8449", "B": "#B9770E", "C": "#6C3483"}
    if df_clay is not None and len(df_clay) and {"wc/LL", "PI", "판정영역"} <= set(df_clay.columns):
        v = df_clay.dropna(subset=["wc/LL", "PI"])
    else:
        v = pd.DataFrame(columns=["wc/LL", "PI", "판정영역"])
    ax.text(0.02, 38.5, "KDS 17 10 00 : 2024 그림 4.7-1 (Bray & Sancio, 2006)", fontsize=7, color="#5D6D7E", va="top")
    for z, sub in v.groupby("판정영역"):
        ax.scatter(sub["wc/LL"], sub["PI"], s=26, color=col.get(z, "#555555"),
                   edgecolors="black", linewidths=0.4, label=f"영역 {z}", zorder=5)
    ax.set_xlim(0, 1.4)
    ax.set_ylim(0, 40)
    ax.set_xlabel("현장함수비 / 액성한계 (wc/LL)")
    ax.set_ylabel("소성지수 PI (%)")
    ax.grid(alpha=0.3)
    if len(v):
        ax.legend(fontsize=7.5, loc="upper right")
    fig.tight_layout()
    return _rp_png(fig)


def _rp_fig_cases(case_tbl, applied):
    d = case_tbl.copy()
    n = len(d)
    fig, ax = plt.subplots(figsize=(7.6, 3.2))
    x = np.arange(n)
    w = 0.38
    b1 = ax.bar(x - w / 2, d["최소FS(SPT)"], width=w, color="#2E86C1", label="최소 FS (SPT-N 방법)", zorder=3)
    b2 = ax.bar(x + w / 2, d["최소FS(Vs)"], width=w, color="#8E44AD", label="최소 FS (Vs 방법)", zorder=3)
    for i, lab in enumerate(d["케이스"]):
        if lab == applied:
            for bb in (b1[i], b2[i]):
                bb.set_edgecolor("#F39C12")
                bb.set_linewidth(2.4)
        ax.text(i, max(d["최소FS(SPT)"].iloc[i], d["최소FS(Vs)"].iloc[i]) + 0.04,
                f"NG {d['NG 시추공'].iloc[i]}공", ha="center", fontsize=7, color="#943126")
    ax.axhline(1.0, color="#C0392B", ls="--", lw=1.3, zorder=4, label="기준안전율 1.0")
    ax.set_xticks(x)
    ax.set_xticklabels([str(c).replace(" ", "\n", 1) for c in d["케이스"]], fontsize=7)
    ax.set_ylabel("최소 안전율 (FS)", fontsize=8.5)
    ax.set_ylim(0, max(1.6, float(np.nanmax(d[["최소FS(SPT)", "최소FS(Vs)"]].to_numpy())) * 1.25))
    ax.grid(axis="y", alpha=0.3, zorder=0)
    from matplotlib.patches import Patch
    from matplotlib.lines import Line2D
    ax.legend(handles=[Patch(color="#2E86C1", label="최소 FS (SPT-N 방법)"),
                       Patch(color="#8E44AD", label="최소 FS (Vs 방법)"),
                       Line2D([0], [0], color="#C0392B", ls="--", label="기준안전율 1.0")],
              fontsize=7, loc="upper left", ncol=3)
    ax.set_title(f"해석 케이스별 최소 안전율 (주황 테두리: 적용 케이스 - {applied})", fontsize=9)
    fig.tight_layout()
    return _rp_png(fig)


def _rp_fig_case_profiles(case_dfs, wave_cols, bh, applied):
    fig, ax = plt.subplots(figsize=(4.6, 4.6))
    for k, (lab, d) in enumerate(case_dfs.items()):
        g = d[d["Borehole"] == bh].sort_values("심도(GL,m)")
        v = g[wave_cols].apply(pd.to_numeric, errors="coerce").mean(axis=1)
        ax.plot(v, g["심도(GL,m)"], marker=_RP_WAVE_MARKERS[k % len(_RP_WAVE_MARKERS)], ms=3,
                lw=2.4 if lab == applied else 1.0, label=str(lab), zorder=5 if lab == applied else 3)
    ax.invert_yaxis()
    ax.xaxis.tick_top()
    ax.xaxis.set_label_position("top")
    ax.set_xlabel("최대가속도 (g, 지진파 평균)", fontsize=8)
    ax.set_ylabel("심도 (m)", fontsize=8)
    ax.set_title(f"{bh} 케이스별 심도별 최대가속도", fontsize=9, pad=24)
    ax.grid(alpha=0.3)
    ax.tick_params(labelsize=7)
    ax.legend(fontsize=6.5, loc="lower right")
    fig.tight_layout()
    return _rp_png(fig)


def _rp_fig_map(coords, judge):
    m = coords.merge(judge, on="Borehole", how="inner")
    fig, ax = plt.subplots(figsize=(7.4, 5.6))
    mv = m.dropna(subset=["최소FS"])
    if len(mv) >= 3:
        gx, gy, gz = idw_grid(mv["X"], mv["Y"], np.clip(mv["최소FS"], 0, 2.0))
        lv = [0, 0.5, 0.75, 1.0, 1.25, 1.5, 2.0]
        cf = ax.contourf(gx, gy, gz, levels=lv, colors=["#C0392B", "#E6B0AA", "#FADBD8", "#E8F6EF",
                                                       "#ABEBC6", "#58D68D"], alpha=0.75)
        cs = ax.contour(gx, gy, gz, levels=[1.0], colors="#7B241C", linewidths=2.2)
        ax.clabel(cs, fmt="FS=1.0", fontsize=7.5)
        cb = fig.colorbar(cf, ax=ax, shrink=0.85)
        cb.set_label("최소 안전율 (FS)", fontsize=8)
    for j, sub in m.groupby("종합판정"):
        ax.scatter(sub["X"], sub["Y"], s=55, color=JUDGE_COLORS.get(j, "#555"), edgecolors="black",
                   linewidths=0.6, label=j, zorder=5)
    for _, r in m.iterrows():
        ax.annotate(str(r["Borehole"]), (r["X"], r["Y"]), xytext=(4, 5), textcoords="offset points",
                    fontsize=6.8, zorder=6)
    ax.set_aspect("equal", adjustable="datalim")
    ax.set_xlabel("X (m)", fontsize=8)
    ax.set_ylabel("Y (m)", fontsize=8)
    ax.tick_params(labelsize=7)
    ax.grid(alpha=0.25)
    ax.legend(fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3, frameon=False)
    fig.tight_layout()
    return _rp_png(fig)


def _rp_fig_alt_cost(alts, best_i):
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    names = [a["대안"].replace("  ", "\n", 1) for a in alts]
    vals = [a["개략공사비(백만원)"] for a in alts]
    cols = ["#E67E22" if i == best_i else "#1F4E78" for i in range(len(alts))]
    ax.barh(names, vals, color=cols, zorder=3)
    for i, v in enumerate(vals):
        ax.text(v, i, f"  {v:,.0f}", va="center", fontsize=8)
    ax.invert_yaxis()
    ax.set_xlabel("개략공사비 (백만원)", fontsize=8.5)
    ax.set_xlim(0, max(vals) * 1.22)
    ax.grid(axis="x", alpha=0.3, zorder=0)
    ax.tick_params(labelsize=7.5)
    fig.tight_layout()
    return _rp_png(fig)


def _rp_fig_improve_map(coords, judge, zones, n=220):
    from matplotlib.path import Path
    from scipy.spatial import ConvexHull
    c = coords.merge(judge[["Borehole", "종합판정"]], on="Borehole").drop_duplicates("Borehole")
    pts = c[["X", "Y"]].to_numpy(float)
    hull = ConvexHull(pts)
    poly = Path(pts[hull.vertices])
    xs = np.linspace(pts[:, 0].min(), pts[:, 0].max(), n)
    ys = np.linspace(pts[:, 1].min(), pts[:, 1].max(), n)
    GX, GY = np.meshgrid(xs, ys)
    cells = np.c_[GX.ravel(), GY.ravel()]
    near = np.argmin(((cells[:, None, :] - pts[None, :, :]) ** 2).sum(-1), axis=1)
    depth = dict(zip(zones["Borehole"], zones["개량심도(m)"]))
    zval = np.array([depth.get(b, np.nan) for b in c["Borehole"]])[near]
    zval[~poly.contains_points(cells, radius=1e-6)] = np.nan
    fig, ax = plt.subplots(figsize=(7.4, 5.4))
    im = ax.pcolormesh(GX, GY, zval.reshape(GX.shape), cmap="OrRd", shading="auto", alpha=0.85)
    cb = fig.colorbar(im, ax=ax, shrink=0.85)
    cb.set_label("개량심도 (GL-m)", fontsize=8)
    hv = np.r_[hull.vertices, hull.vertices[0]]
    ax.plot(pts[hv, 0], pts[hv, 1], color="#5D6D7E", lw=1.0, ls="--")
    for j, sub in c.groupby("종합판정"):
        ax.scatter(sub["X"], sub["Y"], s=40, color=JUDGE_COLORS.get(j, "#555"), edgecolors="black",
                   linewidths=0.5, label=j, zorder=5)
    for _, r in c.iterrows():
        ax.annotate(str(r["Borehole"]), (r["X"], r["Y"]), xytext=(3, 4), textcoords="offset points", fontsize=6.5)
    ax.set_aspect("equal", adjustable="datalim")
    ax.set_xlabel("X (m)", fontsize=8)
    ax.set_ylabel("Y (m)", fontsize=8)
    ax.tick_params(labelsize=7)
    ax.legend(fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=3, frameon=False)
    fig.tight_layout()
    return _rp_png(fig)


def _rp_eq_png(lines, fontsize=12.5):
    """수식 이미지(matplotlib mathtext). lines: [(라벨, LaTeX식)]"""
    h = 0.42 * len(lines) + 0.15
    fig = plt.figure(figsize=(7.2, h))
    for i, (lab, eq) in enumerate(lines):
        y = 1 - (i + 0.6) / len(lines)
        if lab:
            fig.text(0.01, y, lab, fontsize=9.5, va="center", color="#1F3A5F", fontweight="bold")
        fig.text(0.2 if lab else 0.04, y, eq, fontsize=fontsize, va="center")
    return _rp_png(fig, dpi=200)


# ---------------------------------------------------------
# 계산 도우미
# ---------------------------------------------------------
def _fmt(v, nd=3, dash="-"):
    try:
        if v is None or (isinstance(v, float) and not np.isfinite(v)):
            return dash
        return f"{float(v):.{nd}f}"
    except Exception:
        return str(v) if v is not None else dash


def _grid_w(chunk):
    return {1: 9.0, 2: 13.0}.get(len(chunk), 16.0)


def _chunks(seq, n):
    seq = list(seq)
    return [seq[i:i + n] for i in range(0, len(seq), n)]


# ---------------------------------------------------------
# 검토서 본체
# ---------------------------------------------------------
def build_review_report_docx_v2(ctx):
    """ctx(dict) 주요 키:
    df_raw(심도별 amax/g 포함), df_spt, df_vs, wave_cols, spt_min_df, vs_min_df,
    site_response_mode, site_response_log_df, wd(지진파 dict), dtp(dt dict), matching_basis,
    df_clay_result, remed_df, remed_msg, site_location, space_constraint,
    info{과업명, 검토기관, 작성자, 내진등급, 내진성능수준, 재현주기, Z, I},
    curves{sand, rock, clay_fn}, msf{spt_M, vs_case1, vs_case2},
    detail_row_limit(부록 상세계산표 최대 행수)"""
    info = ctx.get("info", {})
    proj = info.get("과업명") or "○○ 조성공사"
    org = info.get("검토기관") or ""
    author = info.get("작성자") or ""
    grade = info.get("내진등급", "내진1등급")
    perf = info.get("내진성능수준", "붕괴방지수준")
    rp = info.get("재현주기", "1,000년")
    Z = float(info.get("Z", 0.11))
    I = float(info.get("I", 1.4))
    S = Z * I

    df_raw = ctx["df_raw"]
    df_spt, df_vs = ctx["df_spt"], ctx["df_vs"]
    wave_cols = list(ctx["wave_cols"])
    spt_min, vs_min = ctx["spt_min_df"], ctx["vs_min_df"]
    mode2 = ctx.get("site_response_mode", "").startswith("②")
    wd, dtp = ctx.get("wd") or {}, ctx.get("dtp") or {}
    sr_log = ctx.get("site_response_log_df")
    clay = ctx.get("df_clay_result")
    remed_df = ctx.get("remed_df")
    boreholes = list(dict.fromkeys(df_raw["Borehole"]))
    spt_t = df_spt[df_spt["대상층여부"] == "대상층"]
    vs_t = df_vs[df_vs["대상층여부"] == "대상층"]

    num = _Numbering()
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin, sec.bottom_margin = Cm(2.2), Cm(2.0)
    sec.header_distance, sec.footer_distance = Cm(1.0), Cm(1.0)
    sec.different_first_page_header_footer = True
    st_normal = doc.styles["Normal"]
    st_normal.font.name = _RP_FONT
    st_normal.font.size = Pt(10)
    st_normal.element.rPr.rFonts.set(qn("w:eastAsia"), _RP_FONT)

    # 머리말 / 꼬리말
    hp = sec.header.paragraphs[0]
    hp.text = ""
    hp.paragraph_format.tab_stops.add_tab_stop(Cm(17.0), alignment=2)
    _rp_font(hp.add_run(proj), size=8.5, color=(0x5D, 0x6D, 0x7E))
    _rp_font(hp.add_run("\t액상화 평가 검토서"), size=8.5, bold=True, color=_RP_NAVY)
    pbdr = OxmlElement("w:pBdr")
    bt = OxmlElement("w:bottom")
    for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:color", "1F3A5F"), ("w:space", "1")):
        bt.set(qn(k), v)
    pbdr.append(bt)
    hp._p.get_or_add_pPr().append(pbdr)
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _rp_font(fp.add_run("- "), size=9)
    _rp_page_field(fp)
    _rp_font(fp.add_run(" -"), size=9)

    def H1(t):
        p = _rp_para(doc, "", space_before=6, space_after=6)
        p.paragraph_format.keep_with_next = True
        _rp_font(p.add_run(t), size=15, bold=True, color=_RP_NAVY)
        pb = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        for k, v in (("w:val", "single"), ("w:sz", "12"), ("w:color", "1F3A5F"), ("w:space", "2")):
            b.set(qn(k), v)
        pb.append(b)
        p._p.get_or_add_pPr().append(pb)

    def H2(t, secno):
        num.set(secno)
        p = _rp_para(doc, "", space_before=10, space_after=4)
        p.paragraph_format.keep_with_next = True
        _rp_font(p.add_run(t), size=12, bold=True, color=_RP_NAVY)

    def H3(t):
        _rp_para(doc, t, size=10.5, bold=True, space_before=6, space_after=3).paragraph_format.keep_with_next = True

    def cap_table(t):
        _rp_para(doc, f"{num.t()} {t}", size=9, bold=True, align="center", space_before=6,
                 space_after=2).paragraph_format.keep_with_next = True

    def figure(buf, width_cm, caption):
        p = _rp_para(doc, "", align="center", space_before=4, space_after=0)
        p.paragraph_format.keep_with_next = True
        p.add_run().add_picture(buf, width=Cm(width_cm))
        _rp_para(doc, f"{num.f()} {caption}", size=9, bold=True, align="center", space_after=8)

    def page_break():
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

    # ===================== 표지 =====================
    for _ in range(7):
        _rp_para(doc, "")
    _rp_para(doc, proj, size=16, bold=True, align="center", color=(0x5D, 0x6D, 0x7E), space_after=10)
    t = _rp_para(doc, "액상화 평가 검토서", size=30, bold=True, align="center", color=_RP_NAVY, space_after=8)
    pb = OxmlElement("w:pBdr")
    for edge in ("top", "bottom"):
        b = OxmlElement(f"w:{edge}")
        for k, v in (("w:val", "single"), ("w:sz", "18"), ("w:color", "1F3A5F"), ("w:space", "8")):
            b.set(qn(k), v)
        pb.append(b)
    t._p.get_or_add_pPr().append(pb)
    _rp_para(doc, "(KDS 17 10 00 : 2024 내진설계 일반 기준 적용)", size=12, align="center",
             color=(0x5D, 0x6D, 0x7E), space_after=40)
    cover = [
        ("검토 대상", f"시추공 {len(boreholes)}개소 ({', '.join(map(str, boreholes[:6]))}{' 외' if len(boreholes) > 6 else ''})"),
        ("적용 지진파", f"{len(wave_cols)}종 ({', '.join(map(str, wave_cols))})"),
        ("평가 방법", "표준관입시험(SPT-N) 방법, 전단파속도(Vs) 방법"),
        ("설계지반운동", f"{grade} · {perf} · 재현주기 {rp} (S = {S:.3f}g)"),
    ]
    _rp_table(doc, [["구 분", "내 용"]], cover, [3.5, 11.5], body_align=["center", "left"],
              first_col_fill=True, font_size=9.5)
    for _ in range(6):
        _rp_para(doc, "")
    _rp_para(doc, date.today().strftime("%Y. %m."), size=13, bold=True, align="center")
    if org or author:
        _rp_para(doc, " / ".join(x for x in (org, author) if x), size=12, align="center")
    page_break()

    # ===================== 목차 =====================
    H1("목  차")
    toc = [
        ("1. 검토 개요", ["1.1 검토 목적 및 범위", "1.2 검토 대상 및 지반조건"]),
        ("2. 내진설계 기준", ["2.1 내진설계 적용기준", "2.2 지진구역 및 위험도계수", "2.3 내진성능수준 결정",
                          "2.4 액상화 판정기준"]),
        ("3. 지진응답해석", ["3.1 개요", "3.2 설계 입력지진파", "3.3 해석방법 및 동적물성치",
                         "3.4 지반가속도 산정결과", "3.5 적용 해석조건"]),
        ("4. 액상화 평가", ["4.1 검토 개요", "4.2 예비평가", "4.3 본평가 방법",
                        "4.4 표준관입시험(SPT) 이용 결과", "4.5 전단파속도(Vs) 이용 결과",
                        "4.6 액상화 본평가 결과 정리", "4.7 시추공별 종합판정",
                        "4.8 액상화 위험구간", "4.9 액상화 분포도"]),
        ("5. 액상화 대책공법 검토", []),
        ("6. 결론", []),
        ("부록", ["A. 시추공별 상세 계산표", "B. 입력 지반자료"]),
    ]
    for ch, subs in toc:
        _rp_para(doc, ch, size=11, bold=True, space_before=6, space_after=2)
        for s_ in subs:
            _rp_para(doc, s_, size=10, indent_cm=0.8, space_after=1)
    page_break()

    # ===================== 1. 검토 개요 =====================
    H1("1. 검토 개요")
    H2("1.1 검토 목적 및 범위", "1.1")
    _rp_bullet(doc, "본 검토는 「내진설계 일반(KDS 17 10 00 : 2024)」의 액상화 평가 절차에 따라, 사업대상지 "
                    "시추조사 결과를 이용하여 설계지진 시 포화 사질토 지반의 액상화 발생 가능성을 평가하고, "
                    "액상화 발생 우려 구간에 대한 대책공법을 검토하기 위하여 수행하였다.")
    _rp_bullet(doc, f"검토 대상은 시추공 {len(boreholes)}개소이며, {len(wave_cols)}종의 지진파에 대한 지진응답해석 "
                    "결과를 토대로 표준관입시험(SPT-N) 방법과 전단파속도(Vs) 방법을 병행하여 평가하였다.")
    _rp_bullet(doc, "평가는 KDS 17 10 00 : 2024 절차에 따라 설계기준에 의한 검토 → 예비평가(평가 생략조건 검토) "
                    "→ 본평가(FS = CRR / CSR > 1.0) → 판정 및 대책공법 검토 순으로 단계적으로 수행하였다.")

    H2("1.2 검토 대상 및 지반조건", "1.2")
    cap_table("검토 대상 시추공 현황")
    rows = []
    for bh in boreholes:
        g = df_raw[df_raw["Borehole"] == bh].sort_values("심도(GL,m)")
        tot = float((g["심도(GL,m)"] + g["층두께(m)"]).max())
        gw = g["지하수위"].iloc[0] if "지하수위" in g.columns else "-"
        gw_txt = "전 지층 수중" if str(gw) in ("1", "1.0", "수중") else ("수상" if str(gw) in ("0", "0.0", "수상") else f"GL-{gw}m")
        tg = spt_t[spt_t["Borehole"] == bh]
        rng = f"{tg['시작심도(m)'].min():.1f} ~ {tg['종료심도(m)'].max():.1f}" if len(tg) else "-"
        layers = " / ".join(b[0] for b in _rp_layer_bands(
            g.rename(columns={"심도(GL,m)": "시작심도(m)"}).assign(**{"종료심도(m)": g["심도(GL,m)"] + g["층두께(m)"]})))
        rows.append([bh, f"{tot:.1f}", gw_txt, rng, str(len(tg)), layers])
    _rp_table(doc, [["시추공", "조사심도\n(m)", "지하수위", "평가대상 심도\n(GL-m)", "대상\n층수", "지층 구성"]],
              rows, [2.2, 1.7, 2.1, 2.6, 1.3, 7.1],
              body_align=["center"] * 5 + ["left"], font_size=8.5)
    page_break()

    # ===================== 2. 내진설계 기준 =====================
    H1("2. 내진설계 기준")
    H2("2.1 내진설계 적용기준", "2.1")
    _rp_bullet(doc, "상위개념 : 내진설계기준연구(Ⅱ) (건설교통부, 1997)")
    _rp_bullet(doc, "하위개념 : 국가건설기준 「내진설계 일반(KDS 17 10 00 : 2024)」")
    cap_table("내진등급 구분")
    _rp_table(doc, [["구 분", "등급별 적용 구조물", "적 용"]], [
        ["내진특등급", "파괴 시 큰 재난 초래 및 기능 마비 시 사회적으로 큰 영향을 줄 수 있는 시설물",
         "⊙" if "특" in grade else ""],
        ["내진1등급", "사회·국가적으로 중요한 구조물", "⊙" if "1" in grade else ""],
        ["내진2등급", "그 외의 일반 구조물", "⊙" if "2" in grade else ""],
    ], [3.0, 11.0, 2.0], body_align=["center", "left", "center"], first_col_fill=True)
    cap_table("내진성능수준 구분")
    _rp_table(doc, [["구 분", "성능수준 정의", "적 용"]], [
        ["기능수행수준", "구조물에 심각한 구조적 손상이 발생하지 않고 지진 시나 지진 경과 후에도 "
                    "구조물의 기능이 정상적으로 유지될 수 있는 성능수준", "⊙" if "기능" in perf else ""],
        ["붕괴방지수준", "구조물에 제한적인 구조적 피해는 발생하나 긴급보수를 통해 단시간에 "
                    "구조물로서의 기능이 발휘될 수 있는 수준", "⊙" if "붕괴" in perf else ""],
    ], [3.0, 11.0, 2.0], body_align=["center", "left", "center"], first_col_fill=True)
    cap_table("재현주기별 설계지반운동 수준")
    _rp_table(doc, [["성능수준", "내진특등급", "내진1등급", "내진2등급"]], [
        ["기능수행", "재현주기 200년", "재현주기 100년", "재현주기 50년"],
        ["붕괴방지", "재현주기 2,400년", "재현주기 1,000년", "재현주기 500년"],
    ], [3.4, 4.2, 4.2, 4.2], first_col_fill=True)

    H2("2.2 지진구역 및 위험도계수", "2.2")
    cap_table("지진구역의 구분 및 지진구역계수")
    _rp_table(doc, [["지진구역", "행정구역", "구역계수 Z"]], [
        ["Ⅰ", "시 : 서울, 인천, 대전, 부산, 대구, 울산, 광주, 세종\n"
              "도 : 경기, 강원 남부, 충북, 충남, 경북, 경남, 전북, 전남 북동부", "0.11g"],
        ["Ⅱ", "도 : 강원 북부, 전남 남서부, 제주", "0.07g"],
    ], [2.4, 11.0, 2.6], body_align=["center", "left", "center"], first_col_fill=True)
    cap_table("재현주기별 위험도계수(I)")
    rps = ["50년", "100년", "200년", "500년", "1,000년", "2,400년"]
    _rp_table(doc, [["재현주기"] + rps], [["위험도계수"] + ["0.40", "0.57", "0.73", "1.00", "1.40", "2.00"]],
              [2.8] + [2.2] * 6, first_col_fill=True)

    H2("2.3 내진성능수준 결정", "2.3")
    _rp_bullet(doc, "설계지반운동수준은 유효수평지반가속도 S = Z × I 로 결정하였다.")
    if mode2:
        _rp_bullet(doc, "지반증폭은 부지 고유의 지진응답해석(1차원 등가선형해석)을 수행하여 직접 반영하였다.")
    cap_table("설계지반운동수준 결정")
    _rp_table(doc, [["내진등급", "내진성능수준", "재현주기", "지진구역계수(Z)", "위험도계수(I)", "설계운동수준(S)"]],
              [[grade, perf, rp, f"{Z:.2f}g", f"{I:.2f}", f"{S:.3f}g"]],
              [2.6, 2.8, 2.2, 2.9, 2.6, 2.9])
    kds = ctx.get("kds") or {}
    if kds:
        _rp_bullet(doc, f"지반종류 {kds.get('soil', '-')}의 단주기지반증폭계수 Fa = {kds.get('Fa', np.nan):.3f}"
                        f"(KDS 표 4.2-8, 직선보간)이며, 지표면 최대지반가속도 하한은 S × Fa = "
                        f"{kds.get('SFA', np.nan):.3f}g이다(KDS 4.7(6)① 단서).")
    if wd:
        pgas = [float(np.max(np.abs(np.asarray(a, float)))) for a in wd.values()]
        avg = float(np.mean(pgas))
        _rp_bullet(doc, f"입력 지진파의 최대가속도는 평균 {avg:.3f}g (범위 {min(pgas):.3f} ~ {max(pgas):.3f}g)로, "
                        f"설계운동수준 {S:.3f}g 대비 {avg / S * 100:.0f}% 수준이다.", level=1)

    H2("2.4 액상화 판정기준", "2.4")
    cap_table("액상화 허용 안전율")
    _rp_table(doc, [["구 분", "예비평가", "본평가"]], [
        ["판정기준", "액상화 평가 생략 가능성 검토", "FS = CRR / CSR > 1.0"],
    ], [3.0, 6.5, 6.5], first_col_fill=True)
    page_break()

    # ===================== 3. 지진응답해석 =====================
    H1("3. 지진응답해석")
    H2("3.1 개요", "3.1")
    _rp_bullet(doc, "설계기준에서 제시된 설계지반운동수준은 지표면이 자유장으로 보통암인 암반 노두에서의 "
                    "운동수준이므로, 지표면이 토사지반인 경우 기반암 운동을 입력으로 하여 지반의 동적 "
                    "변형특성에 따른 지반응답해석을 수행하여야 한다.")
    if mode2:
        _rp_bullet(doc, "본 검토에서는 기반암 노두운동 가속도 시간이력을 입력하여 시추공별 1차원 등가선형 "
                        "지반응답해석(SHAKE 방식)을 수행하고, 심도별 최대가속도(amax)를 산정하였다.")
    else:
        _rp_bullet(doc, "본 검토에서는 별도 수행된 지진응답해석(SHAKE 등)의 심도별 최대가속도(amax/g) 결과를 "
                        "입력자료로 적용하였다.")

    H2("3.2 설계 입력지진파", "3.2")
    _rp_bullet(doc, f"국내외 실지진기록 및 인공지진파 등 총 {len(wave_cols)}종의 지진파를 적용하였다"
                    f"({', '.join(map(str, wave_cols))}).")
    if mode2:
        if ctx.get("matching_basis"):
            _rp_bullet(doc, ctx["matching_basis"])
        _rp_bullet(doc, "입력 지진파는 설계자가 목표 설계응답스펙트럼에 정합(match)한 시간이력으로, 본 검토에서는 "
                        "별도의 진폭 재보정 없이 원본을 그대로 적용하였다(외부 정합 결과에 다시 보정을 가할 경우 "
                        "이중 보정이 되어 설계지진 수준의 근거가 불명확해지므로).")
    else:
        _rp_bullet(doc, "각 지진파에 대한 지진응답해석은 별도로 수행되었으며, 본 검토에는 그 결과인 시추공별·심도별 "
                        "최대가속도(amax/g)를 입력자료로 적용하였다.")
    if wd:
        cap_table("설계 입력지진파 현황")
        rows = []
        for w in wave_cols:
            if w not in wd:
                continue
            a = np.asarray(wd[w], float)
            dur = len(a) * dtp[w]
            kind = "인공지진파" if "인공" in str(w) else "실지진기록 기반"
            rows.append([w, kind, f"{np.max(np.abs(a)):.3f}", f"{dtp[w]:.3f}", f"{dur:.1f}", f"{len(a):,}"])
        _rp_table(doc, [["지진파", "구 분", "최대가속도\n(g)", "시간간격\nΔt (s)", "지속시간\n(s)", "데이터 수"]],
                  rows, [3.6, 3.2, 2.3, 2.3, 2.3, 2.3])
        figure(_rp_fig_input_waves(wd, dtp), 15.5, "설계 입력지진파 가속도 시간이력")

    H2("3.3 해석방법 및 동적물성치", "3.3")
    if mode2:
        _rp_bullet(doc, "해석방법 : 1차원 등가선형 지반응답해석(Schnabel, Lysmer & Seed, 1972; SHAKE 방식)")
        _ct = ctx.get("sr_case_table")
        _ap = (_ct[_ct["케이스"] == ctx.get("sr_applied")].iloc[0]
               if _ct is not None and len(_ct) and (_ct["케이스"] == ctx.get("sr_applied")).any() else None)
        _ratio = f"{_ap['변형률비']:.2f}" if (_ap is not None and _ap["id"] != "-") else "0.65"
        _shift_note = ("" if (_ap is None or _ap["id"] == "-" or abs(_ap["curve_shift"] - 1.0) < 1e-9)
                       else f", 적용 해석조건: {_ap['동적곡선']}")
        _rp_bullet(doc, f"층별 복소전단탄성계수 G* = G(1 + 2iβ)로 전달함수를 구성하고, 유효변형률(최대변형률 × {_ratio})에 "
                        "대응하는 G/Gmax 및 감쇠비가 허용오차 이내로 수렴할 때까지 반복 계산하였다.", level=1)
        _rp_bullet(doc, "기반암 : KDS 17 10 00 : 2024 정의(전단파속도 Vs ≥ 760 m/s인 지층)에 따라 시추공별로 판정하고, "
                        "그 상부 지층을 해석 대상 토층으로 하였다.", level=1)
        _rp_bullet(doc, "동적물성치 : 지층두께·단위중량은 시추 및 실내시험 결과, 전단파속도는 현장 검층 결과를 적용하였다.",
                   level=1)
        cap_table("지진응답해석 흐름")
        _rp_table(doc, [["입력 데이터 작성", "등가선형해석", "해석 결과"]], [[
            "∙ 층별 두께·단위중량\n∙ 전단파속도(Vs)\n∙ G/Gmax, 감쇠비 곡선\n∙ 입력지진파 시간이력",
            "∙ 층별 전달함수 산정\n∙ 유효변형률 → G, D 갱신\n∙ 허용오차 내 반복연산",
            "∙ 심도별 최대가속도\n∙ 지표면 가속도 이력\n∙ 수렴 여부 확인"]],
            [5.3, 5.3, 5.4], body_align=["left"] * 3)
        cur = ctx.get("curves")
        if cur:
            figure(_rp_fig_dyn_curves(cur["sand"], cur["clay_fn"], cur["rock"]), 15.5,
                   f"적용 동적변형특성 곡선 (점성토는 소성지수 PI에 따라 보간 적용{_shift_note})")
        if sr_log is not None and len(sr_log):
            cap_table("시추공별 지반응답해석 수행 요약")
            g = sr_log.groupby("Borehole", sort=False)
            rows = []
            for bh, s_ in g:
                conv = int(s_["수렴여부"].fillna(False).astype(bool).sum())
                rows.append([bh, str(s_["기반암조건"].iloc[0]).replace("자동판정", ""),
                             f"{conv}/{len(s_)}",
                             f"{s_['지표amax/g'].min():.3f} ~ {s_['지표amax/g'].max():.3f}",
                             f"{s_['기반암상단amax/g'].min():.3f} ~ {s_['기반암상단amax/g'].max():.3f}"])
            _rp_table(doc, [["시추공", "기반암 조건", "수렴\n(건)", "지표 amax (g)", "기반암상단 amax (g)"]],
                      rows, [2.2, 5.6, 1.6, 3.3, 3.3], body_align=["center", "left", "center", "center", "center"],
                      font_size=8.5)
    else:
        _rp_bullet(doc, "외부 지진응답해석 결과의 심도별 최대가속도를 그대로 적용하였다.")

    H2("3.4 지반가속도 산정결과", "3.4")
    _rp_bullet(doc, "지진응답해석 결과를 토대로 상부 사질토층에 대한 액상화 검토를 수행하였으며, 시추공별 "
                    "심도별 최대가속도 분포는 다음과 같다.")
    for k, chunk in enumerate(_chunks(boreholes, 6)):
        figure(_rp_fig_amax_profiles(df_raw, wave_cols, chunk), _grid_w(chunk),
               f"심도별 최대가속도 분포({k + 1})" if len(boreholes) > 6 else "심도별 최대가속도 분포")

    H2("3.5 적용 해석조건", "3.5")
    case_tbl = ctx.get("sr_case_table")
    applied = ctx.get("sr_applied") or "단일 해석결과"
    if case_tbl is not None and len(case_tbl) and (case_tbl["케이스"] == applied).any():
        r = case_tbl[case_tbl["케이스"] == applied].iloc[0]
        _rp_bullet(doc, "본 검토의 지진응답해석은 다음 해석조건을 적용하였으며, 이하 액상화 평가는 이 해석결과를 "
                        "기준으로 수행하였다.")
        cap_table("지진응답해석 적용 해석조건")
        if r["id"] != "-":
            vs_txt = "현장 실측 Vs" if abs(r["vs_factor"] - 1.0) < 1e-9 else f"현장 실측 Vs × {r['vs_factor']:.1f}"
            rows = [
                ["해석 방법", "1차원 등가선형 지반응답해석 (SHAKE 방식)"],
                ["동적물성곡선", f"{r['동적곡선']} (사질토 : Seed & Idriss, 1970 / 점성토 : Vucetic & Dobry, 1991, PI별)"],
                ["유효변형률비", f"{r['변형률비']:.2f}"],
                ["토층 전단파속도", vs_txt],
                ["기반암", f"Vs ≥ 760 m/s 지층 (시추 미도달 시 가정 Vs = {r['bedrock_vs']:,.0f} m/s)"],
                ["입력 지진파", "기반암 노두운동 (설계자가 목표 응답스펙트럼에 정합한 시간이력)"],
                ["조건 근거", r["근거"]],
            ]
        else:
            comp = ", ".join(case_tbl.loc[case_tbl["id"] != "-", "케이스"].astype(str))
            rows = [["해석 방법", "1차원 등가선형 지반응답해석 (SHAKE 방식)"], ["적용 결과", applied],
                    ["산정 방법", r["근거"]], ["구성 해석조건", comp]]
        _rp_table(doc, [["항 목", "적용 내용"]], rows, [3.6, 12.4], body_align=["center", "left"],
                  first_col_fill=True, font_size=8.5)
        _rp_bullet(doc, f"적용 해석결과의 지표면 최대가속도는 평균 {r['지표 amax 평균(g)']:.3f}g, "
                        f"최대 {r['지표 amax 최대(g)']:.3f}g이다.")
        if ctx.get("sr_reason"):
            _rp_bullet(doc, f"적용 사유 : {ctx['sr_reason']}")
    else:
        _rp_bullet(doc, f"적용 결과 : {applied}")

    kds = ctx.get("kds") or {}
    pg = kds.get("pga_df")
    if pg is not None and len(pg):
        H2("3.6 지표면 최대가속도 검토", "3.6")
        _rp_bullet(doc, "KDS 17 10 00 : 2024 4.7(6)①에 따라 지반응답해석의 지표면 최대지반가속도는 기반암의 "
                        f"유효수평지반가속도와 단주기지반증폭계수의 곱(S × Fa = {kds['S']:.3f} × {kds['Fa']:.3f} = "
                        f"{kds['SFA']:.3f}g) 이상이어야 한다. 지진파별 지표면 최대가속도의 평균으로 검토하였다.")
        cap_table("시추공별 지표면 최대가속도 검토")
        rows = [[r["Borehole"], f"{r['지표 PGA 평균(g)']:.3f}", f"{r['지표 PGA 최소(g)']:.3f} ~ {r['지표 PGA 최대(g)']:.3f}",
                 f"{r['하한 S×Fa(g)']:.3f}", f"{r['평균/하한']:.2f}", r["판정"]] for _, r in pg.iterrows()]
        _rp_table(doc, [["시추공", "지표 PGA\n평균(g)", "지진파별 범위\n(g)", "하한 S×Fa\n(g)", "평균/하한", "판 정"]], rows,
                  [2.4, 2.4, 3.6, 2.4, 2.2, 2.4], row_fills=[_RP_NG_FILL if r[-1] == "미달" else None for r in rows],
                  font_size=8.5)
        n_low = int((pg["판정"] == "미달").sum())
        cm = ctx.get("csr_method", "tau")
        if n_low == 0:
            _rp_bullet(doc, "모든 시추공에서 지표면 최대가속도가 하한(S × Fa) 이상이다.")
        elif cm == "pga_rd":
            _rp_bullet(doc, f"{n_low}개 시추공에서 지표면 최대가속도가 하한에 미달하여, 반복전단응력비 산정 시 "
                            f"지진파별로 max(PGA, {kds['SFA']:.3f}g)를 적용하였다.")
        elif cm == "amax":
            _rp_bullet(doc, f"{n_low}개 시추공에서 지표면 최대가속도가 하한에 미달한다. 가속도 기반 반복전단응력비를 "
                            "적용하는 경우 KDS 단서 조건을 만족하지 못하므로 최대전단응력(τmax) 방법 또는 하한을 적용한 "
                            "지표 PGA·응력감소계수 방법으로 재검토가 필요하다.")
        else:
            _rp_bullet(doc, f"{n_low}개 시추공에서 지표면 최대가속도가 하한에 미달한다. 본 검토는 최대전단응력(τmax) 방법"
                            "(KDS 식 4.7-2 원칙)을 적용하였으므로 하한이 직접 적용되지는 않으나, 입력 지진파와 해석조건의 "
                            "적정성 확인 지표로 기록하였다.")
    page_break()

    # ===================== 4. 액상화 평가 =====================
    H1("4. 액상화 평가")
    H2("4.1 검토 개요", "4.1")
    _rp_bullet(doc, "사업대상지의 지반조사 결과(주상도, 입도분포 등)를 분석하여 액상화 발생 대상영역을 검토하고, "
                    "경제적인 평가를 위해 단계적으로 수행하였다.")
    _rp_bullet(doc, "액상화 현상은 지반에 가해진 진동하중으로 간극수압이 상승하고 유효응력이 감소하여 포화 "
                    "사질토가 전단저항을 잃게 되는 현상이다.")
    figure(fig_kds_flow_png(), 16.0, "액상화 평가 흐름도 (KDS 17 10 00 : 2024)")

    H2("4.2 예비평가", "4.2")
    _rp_bullet(doc, "다음에 해당하는 지반은 액상화 평가를 생략할 수 있다(KDS 17 10 00 : 2024).")
    for t_ in ("연중 최고지하수위 상부에 위치한 지층",
               "주상도 상의 SPT-N값이 25를 초과하는 지층(에너지효율 보정 전 측점별 낙하횟수 기준)",
               "고소성의 점토거동 유형 지반 : 소성지수(PI), 액성한계(LL), 현장함수비(wc)가 아래 그림의 영역 C에 "
               "해당하는 지반(액상화 가능성은 낮으나 반복연화로 인한 강도저하·대변형 가능성 별도 검토)"):
        _rp_bullet(doc, t_, level=1)
    _rp_bullet(doc, "본 검토에서는 위 생략조건을 층별로 자동 판정하고, 평가심도는 지표면 하 20m 이내로 하였다. "
                    "시추공별 예비평가 결과는 다음과 같다.")
    cap_table("시추공별 예비평가 결과 (층수)")
    pre_cols = ["생략: 지하수위 상부", "생략: N > 25", "생략: 고소성 점토거동(영역 C)", "생략: 심도 20m 초과",
                "비대상: 비사질토(점성토·암반 등)"]
    rows = []
    for bh in boreholes:
        g = df_spt[df_spt["Borehole"] == bh]
        vc = g["예비평가"].value_counts() if "예비평가" in g.columns else pd.Series(dtype=int)
        n_t = int((g["대상층여부"] == "대상층").sum())
        rows.append([bh, str(len(g))] + [str(int(vc.get(c, 0))) for c in pre_cols] + [str(n_t)])
    _rp_table(doc, [["시추공", "전체\n층수", "지하수위\n상부", "N > 25", "점토거동\n영역 C", "심도\n20m 초과",
                     "비사질토\n(점성토·암반)", "본평가\n대상"]], rows, [2.2, 1.5, 1.9, 1.7, 1.9, 1.9, 2.8, 2.0],
              row_fills=[_RP_SUB_FILL if r[-1] == "0" else None for r in rows], font_size=8.5)
    if clay is not None and len(clay):
        cap_table("고소성 점토거동 판별 결과")
        cols = [c for c in ["Borehole", "토층", "시작심도(m)", "종료심도(m)", "LL(%)", "PI", "wc(%)", "wc/LL",
                            "판정영역", "액상화평가"] if c in clay.columns]
        rows = [[_fmt(r[c], 2) if isinstance(r[c], (float, np.floating)) else str(r[c]) for c in cols]
                for _, r in clay.iterrows()]
        widths = {"Borehole": 1.8, "토층": 2.6, "시작심도(m)": 1.4, "종료심도(m)": 1.4, "LL(%)": 1.2, "PI": 1.1,
                  "wc(%)": 1.2, "wc/LL": 1.2, "판정영역": 1.3, "액상화평가": 3.8}
        _rp_table(doc, [[c.replace("(m)", "\n(m)") for c in cols]], rows, [widths[c] for c in cols], font_size=8)
        figure(_rp_fig_clay(clay), 12.5, "소성지수를 바탕으로 한 액상화 가능 범위 (KDS 17 10 00 : 2024 그림 4.7-1)")
    else:
        _rp_bullet(doc, "고소성 점토거동 판별 시료(PI·LL·wc)가 입력되지 않아, 세립토층은 토층명에 따라 예비평가하였다. "
                        "세립토층의 실내시험 결과가 확보되면 아래 판별 기준으로 영역을 판정하여 반영한다.")
        cap_table("점토거동 판별 영역 경계")
        _rp_table(doc, [["영 역", "판별 기준", "평 가"]], [
            ["영역 A", "PI < 12%,  wc/LL > 0.85", "액상화 가능성 높음"],
            ["영역 B", "PI < 18%,  wc/LL > 0.80 (영역 A 제외)", "액상화 가능성 있음 (실내시험 권장)"],
            ["영역 C", "영역 A, B를 제외한 영역", "액상화 가능성 낮음 (반복연화 별도 검토)"],
        ], [3.0, 7.0, 6.0], first_col_fill=True)
        figure(_rp_fig_clay(None), 12.5, "소성지수를 바탕으로 한 액상화 가능 범위 (KDS 17 10 00 : 2024 그림 4.7-1)")

    H2("4.3 본평가 방법", "4.3")
    _rp_bullet(doc, "본평가는 지진 시 흙이 액상화되지 않고 버틸 수 있는 반복저항응력비(CRR)를, 지진하중에 의해 "
                    "특정 깊이에서 발생하는 반복전단응력비(CSR)로 나눈 안전율로 평가하며, 안전율 1.0 이하를 "
                    "액상화 발생으로 판정하였다.")
    H3("가) 반복전단응력비(CSR)")
    if ctx.get("csr_method") == "pga_rd":
        figure(_rp_eq_png([("CSR", r"$CSR = 0.65\,\dfrac{a_{max,s}}{g}\,\dfrac{\sigma_v}{\sigma_v^{\prime}}\,r_d,"
                                   r"\quad a_{max,s} \geq S \times F_a$"),
                           ("rd", r"$\ln r_d=\alpha(z)+\beta(z)M,\ \alpha=-1.012-1.126\sin(z/11.73+5.133),"
                                  r"\ \beta=0.106+0.118\sin(z/11.28+5.142)$")]),
               14.0, "반복전단응력비 산정식 (지표면 최대가속도와 응력감소계수)")
        _kd = ctx.get("kds") or {}
        _rp_bullet(doc, "KDS 17 10 00 : 2024 식(4.7-2) 단서에 따라 지반응답해석으로 산정한 지표면 최대지반가속도(amax,s)와 "
                        "응력감소계수(rd, Idriss 1999)를 적용하였다. 지표면 최대가속도는 지진파별로 하한 S × Fa"
                        f"(= {_kd.get('SFA', np.nan):.3f}g) 이상을 적용하였으며, 설계지진 규모는 KDS SPT 방법의 규모보정계수"
                        "(0.376 = 8.64·e^(−M/4) − 1.325)와 같은 M = 6.5를 사용하였다. 안전율은 각 층의 상단(시험) 심도에서 산정하였다.",
                   level=1)
    elif ctx.get("csr_method", "tau") == "tau" and mode2:
        figure(_rp_eq_png([("CSR", r"$CSR = 0.65\,\dfrac{(\tau_{max})_{d,GRA}}{\sigma_v^{\prime}}$")]),
               9.0, "반복전단응력비 산정식 (지반응답해석 최대전단응력)")
        _rp_bullet(doc, "(τmax)d,GRA : 지반응답해석(GRA)으로 산정한 평가 깊이의 최대전단응력, σv′ : 유효상재응력"
                        "(KDS 64 17 00 : 2024). τmax는 등가선형해석의 수렴 물성(G*)과 변형률로부터 각 층 상단에서 "
                        "직접 산정하였으며, 안전율은 각 층의 상단(시험) 심도에서 산정하였다.", level=1)
    else:
        figure(_rp_eq_png([("CSR", r"$CSR = 0.65\,\dfrac{a_{max}}{g}\,\dfrac{\sigma_v}{\sigma_v^{\prime}}$")]),
               9.0, "반복전단응력비 산정식 (심도별 최대가속도 근사)")
        _rp_bullet(doc, "CSR = 0.65·(τmax)d,GRA/σv′ (KDS 64 17 00 : 2024)에서 τmax ≈ σv·amax(z)/g로 근사하였다"
                        "(수정 Seed & Idriss, rd = 1). amax : 지진응답해석에 의한 심도별 최대가속도, σv : 전상재응력, "
                        "σv′ : 유효상재응력(물의 단위중량 10 kN/m³). 안전율은 각 층의 상단(시험) 심도에서 산정하였다.",
                   level=1)
    msf = ctx.get("msf", {})
    H3("나) 표준관입시험(SPT-N) 방법 - Idriss & Boulanger (2008)")
    figure(_rp_eq_png([
        ("보정N치", r"$(N_1)_{60}=C_N N_{60},\quad C_N=\left(\dfrac{P_a}{\sigma_v^{\prime}}\right)^{m}\leq 1.7,"
                  r"\quad m=0.784-0.0768\sqrt{(N_1)_{60cs}}$"),
        ("세립분보정", r"$\Delta(N_1)_{60}=\exp\left(1.63+\dfrac{9.7}{FC+0.01}-\left(\dfrac{15.7}{FC+0.01}\right)^2\right)$"),
        ("CRR7.5", r"$CRR_{7.5}=\exp\left(\dfrac{N}{14.1}+\left(\dfrac{N}{126}\right)^2-\left(\dfrac{N}{23.6}\right)^3"
                   r"+\left(\dfrac{N}{25.4}\right)^4-2.8\right),\ N=(N_1)_{60cs}$"),
        ("상재압보정", r"$K_\sigma=1-C_\sigma\ln\left(\dfrac{\sigma_v^{\prime}}{P_a}\right)\leq 1.1,\quad "
                    r"C_\sigma=\dfrac{1}{18.9-2.55\sqrt{(N_1)_{60cs}}}\leq 0.3$"),
        ("최종 CRR", r"$CRR = CRR_{7.5}\times K_\sigma \times MSF$"),
    ], fontsize=11.5), 16.0, "SPT-N을 이용한 반복저항응력비 산정식")
    _rp_bullet(doc, "Pa : 대기압(100 kPa, σv′와 같은 단위), Kσ : 유효상재응력 보정계수, "
                    "(N1)60cs : 세립분 보정 환산 N치", level=1)
    _rp_bullet(doc, f"규모보정계수 MSF = 1 + (MSFmax − 1)(8.64e^(−M/4) − 1.325), MSFmax = 1.09 + ((N1)60cs/31.5)² ≤ 2.2 "
                    f"(설계지진 규모 M = {msf.get('spt_M', 6.5)})", level=1)
    _rp_bullet(doc, "N값은 현장측정 N치를 에너지보정이 반영된 N60으로 간주하였다.", level=1)
    H3("다) 전단파속도(Vs) 방법 - Kayen et al. (2013) / Andrus & Stokoe (2000)")
    figure(_rp_eq_png([
        ("응력보정", r"$V_{s1}=V_s\,C_V,\quad C_V=\left(\dfrac{P_a}{\sigma_v^{\prime}}\right)^{0.25}$"
                   r"  ≤ 1.5 (방법1), ≤ 1.4 (방법2)"),
        ("방법1", r"$CRR_{7.5}=\exp\left(\dfrac{(0.0073V_{s1})^{2.8011}-4.8981-0.0099\ln\sigma_v^{\prime}"
                 r"+0.0028FC-0.4984}{1.946}\right)$"),
        ("방법2", r"$CRR_{7.5}=0.022\left(\dfrac{V_{s1}}{100}\right)^2+2.8\left(\dfrac{1}{V_{s1}^{*}-V_{s1}}"
                 r"-\dfrac{1}{V_{s1}^{*}}\right)$"),
        ("최종 CRR", r"$CRR=\min(CRR_{7.5,1}\,MSF_1,\ CRR_{7.5,2}\,MSF_2)$"),
    ], fontsize=11.5), 16.0, "전단파속도를 이용한 반복저항응력비 산정식")
    _rp_bullet(doc, "Vs1* : 세립분함유량(FC)에 따른 한계 전단파속도 (FC ≤ 5% : 215 m/s, 5% < FC < 35% : "
                    "215 − 0.5(FC − 5) m/s, FC ≥ 35% : 200 m/s)", level=1)
    if ctx.get("vs1_mode", "original") == "cap":
        _rp_bullet(doc, "Vs1이 Vs1*를 초과하는 경우 Vs1 = Vs1*로 제한하여 안전측으로 평가하였다.", level=1)
    else:
        _rp_bullet(doc, "Vs1 ≥ Vs1*인 경우 Andrus & Stokoe(2000)에 따라 방법2는 액상화 불가(CRR → ∞)로 보고 "
                        "방법1의 CRR을 적용하였다.", level=1)
    _rp_bullet(doc, "방법1·방법2는 KDS 17 10 00 : 2024 식(4.7-13)~(4.7-19)를 적용하였으며, 최종 CRR은 "
                    "CRR7.5에 규모보정계수(방법1 MSF = 1.2165, 방법2 MSF = 1.2987)를 곱하여 산정하고, 두 방법 중 "
                    "작은 값을 안전율 계산에 적용하였다.", level=1)
    _rp_bullet(doc, f"규모보정계수 : 방법1 MSF = {msf.get('vs_case1', 1.2165):.4f}, 방법2 MSF = "
                    f"{msf.get('vs_case2', 1.2987):.4f} (설계지진 규모 M = {msf.get('spt_M', 6.5)})", level=1)

    H2("4.4 표준관입시험(SPT) 이용 결과", "4.4")
    _rp_bullet(doc, "시추공별·지진파별 심도에 따른 안전율 분포는 다음과 같다(심벌은 각 층 상단 심도에서 "
                    "산정한 안전율, 붉은 음영은 FS ≤ 1.0 영역).")
    for k, chunk in enumerate(_chunks(boreholes, 6)):
        figure(_rp_fig_fs_profiles(df_spt, wave_cols, chunk, "SPT"), _grid_w(chunk),
               f"액상화 본평가 표준관입시험(SPT) 이용 결과({k + 1})" if len(boreholes) > 6
               else "액상화 본평가 표준관입시험(SPT) 이용 결과")

    H2("4.5 전단파속도(Vs) 이용 결과", "4.5")
    for k, chunk in enumerate(_chunks(boreholes, 6)):
        figure(_rp_fig_fs_profiles(df_vs, wave_cols, chunk, "Vs"), _grid_w(chunk),
               f"액상화 본평가 전단파속도(Vs) 이용 결과({k + 1})" if len(boreholes) > 6
               else "액상화 본평가 전단파속도(Vs) 이용 결과")

    H2("4.6 액상화 본평가 결과 정리", "4.6")
    cap_table("액상화 본평가 결과 정리")
    sm = spt_min.set_index("Borehole")
    vm = vs_min.set_index("Borehole")
    rows, fills = [], []
    n_ng_bh = 0
    for bh in boreholes:
        fs1 = sm["FS"].get(bh, np.nan) if bh in sm.index else np.nan
        fs2 = vm["FS"].get(bh, np.nan) if bh in vm.index else np.nan
        w1 = sm["최소안전율 지진파"].get(bh, "-") if bh in sm.index else "-"
        w2 = vm["최소안전율 지진파"].get(bh, "-") if bh in vm.index else "-"
        j1 = "-" if not np.isfinite(fs1) else ("O.K" if fs1 > 1.0 else "N.G")
        j2 = "-" if not np.isfinite(fs2) else ("O.K" if fs2 > 1.0 else "N.G")
        ng = "N.G" in (j1, j2)
        n_ng_bh += int(ng)
        rows.append([bh, _fmt(fs1, 2), "1.00", j1, w1 or "-", _fmt(fs2, 2), "1.00", j2, w2 or "-"])
        fills.append(_RP_NG_FILL if ng else None)
    _rp_table(doc,
              [["구 분", "표준관입시험(SPT-N)을 이용한 평가", None, None, None,
                "전단파속도(Vs)를 이용한 평가", None, None, None],
               [None, "검토\n안전율", "기준\n안전율", "판정", "지배\n지진파", "검토\n안전율", "기준\n안전율", "판정", "지배\n지진파"]],
              rows, [2.0, 1.6, 1.6, 1.3, 2.0, 1.6, 1.6, 1.3, 2.0],
              merges=[(0, 0, 1, 0), (0, 1, 0, 4), (0, 5, 0, 8)], row_fills=fills,
              ng_col_idx={3, 7}, font_size=8.5)
    figure(_rp_fig_bar(spt_min, "표준관입시험(SPT-N) 검토결과 - 시추공별 최소 안전율"), 15.5,
           "표준관입시험(SPT-N) 검토결과 그래프")
    figure(_rp_fig_bar(vs_min, "전단파속도(Vs) 검토결과 - 시추공별 최소 안전율"), 15.5,
           "전단파속도(Vs) 검토결과 그래프")

    H2("4.7 시추공별 종합판정", "4.7")
    _rp_bullet(doc, "KDS 17 10 00 : 2024에 따라 본평가 안전율 1.0을 기준으로 판정하였으며, 시추공별 판정은 "
                    "SPT-N 방법과 Vs 방법 중 작은 안전율을 적용하였다(보수적 판정).")
    judge = borehole_judgement(df_spt, df_vs)
    cap_table("시추공별 종합판정")
    rows, fills = [], []
    for _, r in judge.iterrows():
        rows.append([r["Borehole"], _fmt(r["최소FS(SPT)"], 2), _fmt(r["최소FS(Vs)"], 2), _fmt(r["최소FS"], 2),
                     f"{r['NG층수(SPT)']} / {r['NG층수(Vs)']}", r["액상화 발생구간(GL-m)"], r["종합판정"]])
        fills.append(_RP_NG_FILL if r["종합판정"].startswith("액상화 발생") else None)
    _rp_table(doc, [["시추공", "최소 FS\n(SPT)", "최소 FS\n(Vs)", "판정 FS\n(최소)", "NG 층수\n(SPT / Vs)",
                     "액상화 발생구간\n(GL-m)", "종합판정"]], rows, [1.9, 1.6, 1.6, 1.6, 2.0, 2.8, 4.5],
              row_fills=fills, font_size=8.5)
    n_judge_ng = int(judge["종합판정"].str.startswith("액상화 발생").sum())
    _rp_bullet(doc, f"검토 시추공 {len(judge)}개소 중 {n_judge_ng}개소에서 액상화 발생(FS ≤ 1.0)으로 판정되었다.")

    H2("4.8 액상화 위험구간", "4.8")
    ng_rows = []
    for name, d in (("SPT", spt_t), ("Vs", vs_t)):
        for _, r in d[d["FS"] <= 1.0].iterrows():
            ng_rows.append([r["Borehole"], name, f"{r['시작심도(m)']:.1f} ~ {r['종료심도(m)']:.1f}", r["토층"],
                            _fmt(r["CSR"], 3), _fmt(r["CRR"], 3), _fmt(r["FS"], 3), r.get("지배지진파", "-") or "-"])
    if ng_rows:
        _rp_bullet(doc, f"안전율 1.0 이하로 액상화 발생 가능성이 있는 구간은 총 {len(ng_rows)}개 층(평가방법별 합계)이며, "
                        "세부 현황은 다음과 같다.")
        if len(ng_rows) <= 60:
            cap_table("액상화 위험구간 현황 (FS ≤ 1.0)")
            show = sorted(ng_rows, key=lambda x: (str(x[0]), x[1], float(x[2].split("~")[0])))
            _rp_table(doc, [["시추공", "평가\n방법", "심도\n(GL-m)", "토층", "CSR\n(최대)", "CRR", "FS", "지배 지진파"]],
                      show, [2.0, 1.4, 2.2, 3.0, 1.7, 1.7, 1.6, 2.4], font_size=8)
        else:
            # 위험구간이 많으면 층 단위 목록 대신 시추공·평가방법별 요약 (상세는 엑셀 상세계산 시트)
            cap_table("시추공별 액상화 위험구간 요약 (FS ≤ 1.0)")
            summ_rows = []
            for name, d in (("SPT", spt_t), ("Vs", vs_t)):
                ngd = d[d["FS"] <= 1.0]
                for bh_, gg in ngd.groupby("Borehole", sort=False):
                    w = gg.loc[gg["FS"].idxmin()]
                    summ_rows.append([bh_, name, f"{gg['시작심도(m)'].min():.1f} ~ {gg['종료심도(m)'].max():.1f}",
                                      str(len(gg)), _fmt(gg["FS"].min(), 3), f"{w['시작심도(m)']:.1f}",
                                      w.get("지배지진파", "-") or "-"])
            summ_rows.sort(key=lambda x: (str(x[0]), x[1]))
            _rp_table(doc, [["시추공", "평가\n방법", "위험구간 범위\n(GL-m)", "NG\n층수", "최소 FS", "최소 FS\n심도(m)",
                             "지배 지진파"]], summ_rows, [2.3, 1.5, 3.0, 1.5, 2.0, 2.2, 3.5], font_size=8)
            _rp_para(doc, "※ 층별 상세 목록은 엑셀 결과파일의 SPT_상세계산·Vs_상세계산 시트 참조.", size=8.5,
                     color=(0x7F, 0x8C, 0x8D))
    else:
        _rp_bullet(doc, "전 시추공·전 평가대상층에서 안전율이 1.0을 초과하여, 액상화 위험구간은 없는 것으로 평가되었다.")

    H2("4.9 액상화 분포도", "4.9")
    coords = ctx.get("coords_df")
    if coords is not None and len(coords.merge(judge, on="Borehole")) > 0:
        n_match = len(coords.merge(judge, on="Borehole"))
        _rp_bullet(doc, f"시추공 좌표 {n_match}개소를 이용하여 시추공별 종합판정과 최소 안전율 분포(역거리가중 보간)를 "
                        "평면상에 도시하였다. 굵은 실선은 FS = 1.0 경계로, 경계 안쪽(붉은 영역)이 액상화 발생 "
                        "가능 구역이다.")
        figure(_rp_fig_map(coords, judge), 15.5, "액상화 분포도 (시추공별 판정 및 최소 안전율 분포)")
        if n_match < 3:
            _rp_para(doc, "※ 좌표가 있는 시추공이 3개 미만이어서 안전율 분포 보간은 생략하였다.", size=8.5,
                     color=(0x7F, 0x8C, 0x8D))
    else:
        _rp_bullet(doc, "시추공 좌표가 입력되지 않아 액상화 분포도는 작성하지 않았다(좌표 엑셀 업로드 시 자동 작성).")
    page_break()

    # ===================== 5. 대책공법 =====================
    H1("5. 액상화 대책공법 검토")
    rz, ra = ctx.get("rm_zones"), ctx.get("rm_alts")
    import re as _re
    if rz is not None and len(rz):
        H2("5.1 개량 대상 범위", "5.1")
        _rp_bullet(doc, f"액상화 발생(FS ≤ 1.0) 시추공의 영향면적과 최하단 액상화층 하단까지의 심도를 개량범위로 "
                        f"산정하였다({ctx.get('rm_area_how', '')}, 검토 부지면적 {ctx.get('rm_total_area') or 0:,.0f} m²).")
        cap_table("시추공별 개량 대상 범위")
        rows = [[r["Borehole"], f"{r['NG 상단(m)']:.1f}", f"{r['개량심도(m)']:.1f}", f"{r['개량두께(m)']:.1f}",
                 f"{r['평균FC(%)']:.1f}", f"{r['영향면적(m²)']:,.0f}", f"{r['개량체적(m³)']:,.0f}"] for _, r in rz.iterrows()]
        rows.append(["합 계", "-", f"최대 {rz['개량심도(m)'].max():.1f}", "-", "-",
                     f"{rz['영향면적(m²)'].sum():,.0f}", f"{rz['개량체적(m³)'].sum():,.0f}"])
        _rp_table(doc, [["시추공", "액상화층\n상단(GL-m)", "개량심도\n(GL-m)", "액상화층\n두께(m)", "평균 FC\n(%)",
                         "영향면적\n(m²)", "개량체적\n(m³)"]], rows, [2.2, 2.2, 2.2, 2.1, 1.9, 2.7, 2.7],
                  row_fills=[None] * (len(rows) - 1) + [_RP_HDR_FILL], font_size=8.5)
        coords = ctx.get("coords_df")
        if coords is not None and len(coords) >= 3:
            figure(_rp_fig_improve_map(coords, judge, rz), 15.0, "액상화 개량 대상 범위도 (영향면적 기준)")

        H2("5.2 공법별 적용성 검토", "5.2")
        _cdb = ctx.get("rm_cost_db") or REMED_COST_DB
        _rp_bullet(doc, "공법별 단가는 공법 고유 단위(원/m²·원/m³·원/m)의 직접공사비를 적용하고, 기둥형 공법은 배치간격, "
                        "심층혼합처리는 개량률로 평면 1m²당 공사비를 환산하였다. 단가 산정 근거는 다음과 같다.")
        cap_table("대책공법 적용 단가 및 산정 근거")
        _rp_table(doc, [["공 법", "단 가", "산정 조건", "환산단가\n(원/m³)", "산정 근거"]],
                  [[m["공법"], f"{_num(m.get('단가'), 0):,.0f} {m.get('단위', '')}", remed_quantity_text(m),
                    f"{remed_equiv_m3(m):,.0f}", str(m.get("단가 근거", ""))] for m in _cdb],
                  [3.0, 2.6, 2.6, 1.7, 6.1], body_align=["left", "right", "left", "right", "left"], font_size=7.5)
        _rp_bullet(doc, "환산단가는 개량심도 10m 기준의 개량지반 체적당 공사비로, 공법 간 비교용이다. 치환 단가는 표준시장단가"
                        "(2025 상반기) 토공 항목과 골재 단가로 산출하였다.", level=1)
        _rp_bullet(doc, f"현장 입지 : {ctx.get('site_location', '-')} / {ctx.get('space_constraint', '-')}")
        _rp_bullet(doc, "공법별 최대 개량심도, 적용 세립분함유량, 소음·진동 및 장비 제약을 개량 대상 구간에 대조하여 "
                        "적용 가능 여부를 판정하였다.")
        cap_table("공법별 적용성 및 단일공법 개략공사비")
        sg = ra["single"]
        rows, fills = [], []
        for _, r in sg.iterrows():
            rows.append([r["공법"], r["분류"], "적용 가능" if r["적용가능"] else "부적합", r["부적합 사유"] or "-",
                         f"{r['단가']:,.0f}\n{r['단위']}", _fmt(r["개략공사비(백만원)"], 0)])
            fills.append(None if r["적용가능"] else _RP_SUB_FILL)
        _rp_table(doc, [["공 법", "분 류", "적용성", "부적합 사유", "단가\n(단위)", "개략공사비\n(백만원)"]], rows,
                  [3.4, 2.3, 1.7, 4.6, 1.9, 2.1], body_align=["left", "center", "center", "left", "right", "right"],
                  row_fills=fills, font_size=8)

        H2("5.3 대안별 개략공사비 및 최적안", "5.3")
        alts = ra["alts"]
        if alts:
            cap_table("대책공법 대안별 개략공사비 비교")
            _cand_i = [i for i, x in enumerate(alts) if not str(x["대안"]).startswith("대안 4")] or list(range(len(alts)))
            best_i = min(_cand_i, key=lambda i: alts[i]["개략공사비(백만원)"])
            rows = [[x["대안"], x["공법 구성"], f"{x['개량면적(m²)']:,.0f}", f"{x['개량체적(m³)']:,.0f}",
                     f"{x['개략공사비(백만원)']:,.0f}", x["특징"]] for x in alts]
            _rp_table(doc, [["대 안", "공법 구성", "개량면적\n(m²)", "개량체적\n(m³)", "개략공사비\n(백만원)", "특 징"]],
                      rows, [3.3, 3.6, 1.9, 1.9, 2.0, 3.3], body_align=["left", "left", "right", "right", "right", "left"],
                      row_fills=[_RP_HDR_FILL if i == best_i else None for i in range(len(rows))], font_size=8)
            figure(_rp_fig_alt_cost(alts, best_i), 14.0, "대안별 개략공사비")
            _full = [x for x in alts if not str(x["대안"]).startswith("대안 4")]
            _a4 = next((x for x in alts if str(x["대안"]).startswith("대안 4")), None)
            if _full:
                bx = min(_full, key=lambda x: x["개략공사비(백만원)"])
                _rp_bullet(doc, f"액상화층 전체를 개량하는 경우, 개략공사비가 가장 작은 「{bx['대안'].split('  ')[-1]}」"
                                f"(공법 구성 : {bx['공법 구성']}, 약 {bx['개략공사비(백만원)']:,.0f}백만원)을 최적안으로 제안한다.")
            if _a4 is not None:
                _rp_bullet(doc, f"현장 실무형 대안(대안 4 : {_a4['공법 구성']}, 약 {_a4['개략공사비(백만원)']:,.0f}백만원)은 "
                                "지표 피해 억제를 위한 상부 처리 비용으로, KDS 4.7(7)에 따른 시설물 안정성 평가(기초 형식별 "
                                "액상화 영향 반영)에서 안정성이 확보되는 경우에 적용할 수 있다.")
            _ovh = float(ctx.get("overhead") or 1.0)
            _rp_bullet(doc, f"단가 기준 : {ctx.get('price_basis') or REMED_PRICE_BASIS_DEFAULT}"
                            + (f", 제경비 포함 배율 {_ovh:.2f} 적용" if abs(_ovh - 1.0) > 1e-9 else " (직접공사비)") + ".",
                       level=1)
            _rp_bullet(doc, "개략공사비는 개량체적 기준 단가에 의한 금액으로, 실시설계 시 최신 표준품셈·표준시장단가·물가자료 및 "
                            "개량 설계(치환율·배치간격·개량강도)에 따라 재산정하여야 한다.", level=1)
            _rp_bullet(doc, "시공 후 표준관입시험·콘관입시험·전단파속도 시험 등으로 보강효과를 확인하고 액상화 재평가를 "
                            "수행하여야 한다.", level=1)
        else:
            _rp_bullet(doc, "현장 조건(심도·세립분·입지)을 모두 만족하는 공법 조합이 없어 공법 조건의 재검토가 필요하다.")
        cr = ctx.get("crust_df")
        _sec_q = "5.4"
        if cr is not None and len(cr):
            H2("5.4 현장 실무형 검토 (상부 비액상화층 확보)", "5.4")
            _sec_q = "5.5"
            _rp_bullet(doc, "KDS 17 10 00 : 2024 4.7(7)은 안전율이 1 이하인 경우 액상화 영향을 반영한 시설물 안정성 평가를 "
                            "수행하고, 안정성을 확보하지 못할 때 대책방안을 수립하도록 규정한다. 이에 따라 전면 지반개량 외에 "
                            "상부 비액상화층(H1) 확보에 의한 지표 피해 억제 방안을 함께 검토하였다.")
            _rp_bullet(doc, "Ishihara(1985)는 수평지반에서 비액상화 표층 두께(H1)와 지표 최대가속도에 따라 하부 액상화층(H2)의 "
                            "지표 피해(분사·균열) 발생 경계를 제시하였으며, 본 검토에서는 경계곡선 상한(지표 PGA 0.2g : 3m, "
                            "0.3g : 6m, 0.45g 이상 : 9m)을 보수적으로 적용하였다. 평가 PGA는 max(지표 PGA 평균, S × Fa)이다.",
                       level=1)
            _rp_bullet(doc, "상부 쇄석치환은 지하수위 하부 굴착에 따른 배수·흙막이를 고려하여 실무 한계심도 이내로, 동다짐은 "
                            "유효개량심도 D = n√(W·H)(Menard & Broise 1975, Lukas 1995) 이내로 적용성을 판단하였다.", level=1)
            cap_table("시추공별 상부 비액상화층 확보 검토")
            rows = [[r["Borehole"], f"{r['액상화층 상단 H1(m)']:.1f}", f"{r['H2(m)']:.1f}", f"{r['평가 PGA(g)']:.3f}",
                     f"{r['필요 H1(m)']:.1f}", f"{r['필요 처리심도(m)']:.1f}", f"{r['동다짐 유효심도(m)']:.1f}", r["판정"],
                     r["현장 실무형 대책"]] for _, r in cr.iterrows()]
            _rp_table(doc, [["시추공", "현재 H1\n(m)", "H2\n(m)", "평가 PGA\n(g)", "필요 H1\n(m)", "처리심도\n(GL-m)",
                             "동다짐\n유효심도(m)", "판 정", "대 책"]], rows,
                      [1.6, 1.3, 1.1, 1.5, 1.3, 1.5, 1.6, 2.6, 4.0],
                      body_align=["center"] * 7 + ["left", "left"], font_size=7.5)
            _rp_bullet(doc, "상부 처리안은 수평지반의 지표 피해 억제를 목적으로 하므로, 말뚝기초의 부마찰력·수평지지력 감소, "
                            "얕은기초 지지력, 경사지·호안 인접부의 측방유동, 매설관 부상은 시설물 안정성 평가에서 별도로 "
                            "검토하여야 한다.", level=1)
        if remed_df is not None and len(remed_df):
            H2(f"{_sec_q} 공법 정성 평가 (참고)", _sec_q)
            cap_table("시공성·경제성·시공실적 종합 평가 (상위 5개 공법, " + str(ctx.get("remed_priority") or "균형") + ")")
            cols = [c for c in ["추천 순위", "보강 공법", "평가 점수", "경제성", "검토 의견"] if c in remed_df.columns]
            rows = [[_re.sub(r"[^\w\s().%가-힣A-Za-z0-9★☆,·/~≥≤>=<:-]", "", str(r[c])).strip() for c in cols]
                    for _, r in remed_df.head(5).iterrows()]
            _rp_table(doc, [cols], rows, [{"추천 순위": 2.4, "보강 공법": 3.6, "평가 점수": 1.6, "경제성": 2.6,
                                           "검토 의견": 5.8}[c] for c in cols],
                      body_align=["center"] * (len(cols) - 1) + ["left"], font_size=8)
    else:
        _rp_bullet(doc, "전 구간에서 액상화에 대해 안전한 것으로 평가되어 별도의 액상화 대책공법은 불필요한 것으로 판단된다.")

    # ===================== 6. 결론 =====================
    H1("6. 결론")
    n_bh = len(boreholes)
    spt_ng = int((spt_min["FS"] <= 1.0).sum())
    vs_ng = int((vs_min["FS"] <= 1.0).sum())
    concl = [
        f"「내진설계 일반(KDS 17 10 00 : 2024)」에 따라 {grade}, {perf}(재현주기 {rp}), 설계운동수준 "
        f"S = {S:.3f}g 조건으로 시추공 {n_bh}개소에 대해 {len(wave_cols)}종 지진파를 적용하여 액상화를 평가하였다.",
        ("기반암(Vs ≥ 760m/s) 노두운동을 입력으로 시추공별 1차원 등가선형 지반응답해석을 수행하여 심도별 최대가속도를 "
         "산정하였다." if mode2 else "외부 지진응답해석의 심도별 최대가속도를 적용하였다."),
        f"표준관입시험(SPT-N) 방법 평가 결과 {n_bh}개소 중 {spt_ng}개소, 전단파속도(Vs) 방법 평가 결과 {vs_ng}개소에서 "
        "최소 안전율이 1.0 이하로 산정되었다.",
    ]
    _kd = ctx.get("kds") or {}
    _pg = _kd.get("pga_df")
    if _pg is not None and len(_pg):
        _nl = int((_pg["판정"] == "미달").sum())
        concl.append(f"지표면 최대가속도 검토(KDS 4.7(6)①, 하한 S × Fa = {_kd['SFA']:.3f}g) 결과 "
                     + (f"{_nl}개소가 하한에 미달하였다" if _nl else "전 시추공이 하한 이상이었다")
                     + ("(반복전단응력비 산정 시 하한 적용)." if (_nl and ctx.get("csr_method") == "pga_rd") else "."))
    if n_ng_bh:
        ngj = judge[judge["종합판정"].str.startswith("액상화 발생")]
        concl.append(f"시추공별 종합판정 결과 {len(ngj)}개소({', '.join(map(str, ngj['Borehole'].head(8)))}"
                     f"{' 등' if len(ngj) > 8 else ''})에서 액상화 발생이 예상되어, 해당 구간에 대한 "
                     "액상화 대책이 필요한 것으로 판단된다.")
        if remed_df is not None and len(remed_df):
            import re as _re
            _alts = (ctx.get("rm_alts") or {}).get("alts") or []
            _full = [x for x in _alts if not str(x["대안"]).startswith("대안 4")]
            if _full:
                _b = min(_full, key=lambda x: x["개략공사비(백만원)"])
                concl.append(f"액상화층 전체를 개량하는 경우 대책공법은 개량범위(면적·심도), 지반조건 및 현장 입지를 고려한 "
                             f"대안 비교 결과 「{_b['대안'].split('  ')[-1]} : {_b['공법 구성']}」(개략공사비 약 "
                             f"{_b['개략공사비(백만원)']:,.0f}백만원)을 최적안으로 제안하며, 시공 후 보강효과 확인시험을 "
                             "시행하여야 한다.")
        _cr = ctx.get("crust_df")
        if _cr is not None and len(_cr):
            _ok = _cr[_cr["필요 처리심도(m)"] <= 0]["Borehole"].tolist()
            _up = _cr[_cr["판정"].str.contains("확보 가능")]
            _no = _cr[_cr["판정"].str.contains("불충분")]["Borehole"].tolist()
            _nm = lambda L: f"{', '.join(map(str, L[:6]))}{' 등' if len(L) > 6 else ''}({len(L)}개소)"
            _t = []
            if _ok:
                _t.append(f"상부 비액상화층 충분 {_nm(_ok)} - 지표 피해 가능성이 낮아 지반개량보다 시설물 안정성 평가·"
                          "기초 설계에서 액상화 영향을 반영하는 방안이 경제적임")
            if len(_up):
                _t.append(f"상부 처리로 확보 가능 {_nm(_up['Borehole'].tolist())} - GL-{_up['필요 처리심도(m)'].max():.1f}m까지 "
                          "상부 치환 또는 동다짐으로 비액상화층 확보")
            if _no:
                _t.append(f"상부 처리로 불충분 {_nm(_no)} - 심층개량 또는 구조물 대책 필요")
            if _t:
                concl.append("현장 실무형 검토(Ishihara 1985, KDS 4.7(7)) 결과는 다음과 같으며, 어느 경우에도 하부 "
                             "액상화층이 기초에 미치는 영향은 별도로 검토하여야 한다 : " + " / ".join(_t) + ".")
    else:
        concl.append("전 시추공에서 안전율 1.0 초과로 액상화 발생 가능성이 낮은 것으로 평가되었다.")
    for c_ in concl:
        _rp_bullet(doc, c_)
    _rp_para(doc, "")
    _rp_para(doc, "※ 본 검토서는 액상화 평가 프로그램의 자동 계산 결과이며, 설계 적용 전 책임기술자의 검토·확인이 필요하다.",
             size=8.5, color=(0x7F, 0x8C, 0x8D))
    page_break()

    # ===================== 부록 =====================
    H1("부록 A. 시추공별 상세 계산표")
    limit = int(ctx.get("detail_row_limit", 600))
    total = len(spt_t) + len(vs_t)
    if total > limit:
        _rp_para(doc, f"※ 평가대상 층이 많아(SPT·Vs 합계 {total}개 행) 상세계산표는 엑셀 결과파일"
                      "(SPT_상세계산, Vs_상세계산 시트)로 대체한다.", size=9.5)
    else:
        num.set("A")
        for bh in boreholes:
            g1 = spt_t[spt_t["Borehole"] == bh].sort_values("시작심도(m)")
            g2 = vs_t[vs_t["Borehole"] == bh].sort_values("시작심도(m)")
            if g1.empty and g2.empty:
                continue
            H3(f"■ {bh}")
            if not g1.empty:
                cap_table(f"{bh} 표준관입시험(SPT-N) 방법 상세 계산")
                rows, fills = [], []
                for _, r in g1.iterrows():
                    rows.append([f"{r['시작심도(m)']:.1f}", r["토층"], _fmt(r["σv(kPa)"], 1), _fmt(r["σv'(kPa)"], 1),
                                 _fmt(r.get("N치(N60)"), 0), _fmt(r.get("(N1)60cs"), 2), _fmt(r.get("Ksigma"), 3),
                                 _fmt(r.get("MSF"), 3), _fmt(r["CRR"], 3), _fmt(r["CSR"], 3), _fmt(r["FS"], 3),
                                 "N.G" if (np.isfinite(r["FS"]) and r["FS"] < 1) else ("O.K" if np.isfinite(r["FS"]) else "-")])
                    fills.append(_RP_NG_FILL if rows[-1][-1] == "N.G" else None)
                _rp_table(doc, [["심도\n(m)", "토층", "σv\n(kPa)", "σv′\n(kPa)", "N60", "(N1)60cs", "Kσ", "MSF",
                                 "CRR", "CSR\n(최대)", "FS", "판정"]], rows,
                          [1.1, 2.4, 1.3, 1.3, 1.0, 1.5, 1.2, 1.2, 1.3, 1.4, 1.3, 1.0],
                          row_fills=fills, ng_col_idx={11}, font_size=7.5)
            if not g2.empty:
                cap_table(f"{bh} 전단파속도(Vs) 방법 상세 계산")
                rows, fills = [], []
                for _, r in g2.iterrows():
                    rows.append([f"{r['시작심도(m)']:.1f}", r["토층"], _fmt(r["σv'(kPa)"], 1), _fmt(r.get("Vs(m/s)"), 0),
                                 _fmt(r.get("Vs1_CASE1"), 1), _fmt(r.get("Vs1_star"), 1), _fmt(r.get("CRR_CASE1"), 3),
                                 _fmt(r.get("CRR_CASE2"), 3), _fmt(r["CRR"], 3), _fmt(r["CSR"], 3), _fmt(r["FS"], 3),
                                 "N.G" if (np.isfinite(r["FS"]) and r["FS"] < 1) else ("O.K" if np.isfinite(r["FS"]) else "-")])
                    fills.append(_RP_NG_FILL if rows[-1][-1] == "N.G" else None)
                _rp_table(doc, [["심도\n(m)", "토층", "σv′\n(kPa)", "Vs\n(m/s)", "Vs1\n(m/s)", "Vs1*\n(m/s)",
                                 "CRR\n(방법1)", "CRR\n(방법2)", "CRR", "CSR\n(최대)", "FS", "판정"]], rows,
                          [1.1, 2.4, 1.3, 1.2, 1.3, 1.3, 1.4, 1.4, 1.2, 1.4, 1.2, 1.0],
                          row_fills=fills, ng_col_idx={11}, font_size=7.5)

    page_break()
    H1("부록 B. 입력 지반자료")
    num.set("B")
    in_cols = [c for c in ["Borehole", "토층", "심도(GL,m)", "층두께(m)", "γsat(kN/m3)", "γt(kN/m3)", "Vs(m/s)",
                           "SPT-N", "FC(%)"] if c in df_raw.columns]
    if len(df_raw) <= 800:
        cap_table("시추공별 입력 지반자료")
        rows = []
        for _, r in df_raw.iterrows():
            rows.append([str(r[c]) if c in ("Borehole", "토층") else _fmt(r[c], 1 if c != "SPT-N" else 0)
                         for c in in_cols])
        hdr = [c.replace("(", "\n(") for c in in_cols]
        wmap = {"Borehole": 1.9, "토층": 2.8, "심도(GL,m)": 1.6, "층두께(m)": 1.5, "γsat(kN/m3)": 1.7,
                "γt(kN/m3)": 1.7, "Vs(m/s)": 1.5, "SPT-N": 1.3, "FC(%)": 1.3}
        _rp_table(doc, [hdr], rows, [wmap[c] for c in in_cols], font_size=7.5)
    else:
        _rp_para(doc, "※ 입력자료가 많아 엑셀 결과파일의 '입력데이터' 시트로 대체한다.", size=9.5)

    out = io.BytesIO()
    doc.save(out)
    return out.getvalue()


# ---------------------------------------------------------
# 5. 최적 보강공법 추천 알고리즘
# ---------------------------------------------------------
def evaluate_remediation_options(ng_df, site_loc, space_const, priority="균형", crust_df=None):
    max_depth = ng_df["종료심도(m)"].max()
    min_depth = ng_df["시작심도(m)"].min()
    ng_span = max_depth - min_depth  # 액상화 위험구간의 수직 분포 폭(국소적/광범위 여부 판단용)
    avg_fc = ng_df["FC(%)"].mean()
    is_urban = "도심지" in site_loc
    is_narrow = "협소" in space_const

    eval_results = []
    # 경제성 우선(현장 실무형) + 모든 액상화 시추공에서 상부 처리로 비액상화층(H1) 확보가 가능하면,
    # 치환·동다짐은 '액상화층 전체'가 아니라 '필요 처리심도'까지만 개량하는 것으로 평가한다.
    crust_mode = bool(priority.startswith("경제성") and crust_df is not None and len(crust_df)
                      and (crust_df["판정"] != "상부 처리로 불충분").all()
                      and (crust_df["필요 처리심도(m)"] > 0).any())
    crust_depth = float(crust_df["필요 처리심도(m)"].max()) if crust_mode else None
    _crust_txt = " / ".join(crust_df["현장 실무형 대책"].astype(str)) if crust_mode else ""

    for m in METHODS_DB:
        _is_crust = crust_mode and (("치환" in m["공법명"] and "상부 치환" in _crust_txt)
                                    or ("동다짐" in m["공법명"] and "동다짐" in _crust_txt))
        eval_depth = crust_depth if _is_crust else max_depth
        # base_score(시공성/적용성) : econ_score(경제성) : usage_score(실무 활용빈도)
        # = 30 : 30 : 40 으로 가중 평균한다. 활용빈도에 가장 큰 비중을 둔 이유는,
        # 실제로 현장에서 많이 채택되는 공법일수록 시공사례·설계기준·품질관리 노하우가
        # 축적되어 있어 리스크가 낮다는 점을 반영하기 위함이다. 이렇게 하면 SCP/GCP처럼
        # "성능·경제성·실적"이 모두 우수한 공법이 최상위권을 유지하면서도, 치환·동다짐처럼
        # 경제성과 실적은 우수하지만 base_score(범용 적용성)가 상대적으로 낮았던 공법의
        # 순위가 실제 현장에서 체감하는 채택 빈도에 맞게 끌어올려진다.
        if priority.startswith("경제성"):
            # 현장 실무형(경제성 우선): 공사비 증액 부담이 큰 현장 여건을 반영해 경제성 비중을 높인다
            score = 0.2 * m["base_score"] + 0.5 * m["econ_score"] + 0.3 * m["usage_score"]
        else:
            score = (
                0.3 * m["base_score"]
                + 0.3 * m["econ_score"]
                + 0.4 * m["usage_score"]
            )
        reasons = []
        bonuses = []

        if eval_depth > m["max_depth"]:
            depth_diff = eval_depth - m["max_depth"]
            score -= 25 + depth_diff * 5
            reasons.append(
                f"개량 심도 초과(필요 {eval_depth:.1f}m > 한계 {m['max_depth']}m)"
            )

        if "FC < 15%" in m["FC_제약"] and avg_fc >= 15:
            score -= 30
            reasons.append(f"세골함유량 초과(평균 {avg_fc:.1f}% ≥ 15%)")
        elif "FC < 20%" in m["FC_제약"] and avg_fc >= 20:
            score -= 25
            reasons.append(f"세골함유량 초과(평균 {avg_fc:.1f}% ≥ 20%)")
        elif "FC < 25%" in m["FC_제약"] and avg_fc >= 25:
            score -= 20
            reasons.append(f"세골함유량 초과(평균 {avg_fc:.1f}% ≥ 25%)")
        elif "FC < 35%" in m["FC_제약"] and avg_fc >= 35:
            score -= 20
            reasons.append(f"세골함유량 초과(평균 {avg_fc:.1f}% ≥ 35%)")

        if is_urban and m["소음진동"] in ["높음", "매우 높음"]:
            score -= 35
            reasons.append("도심지 소음·진동 민원 우려 극심")
        elif is_urban and m["소음진동"] == "중/높음":
            score -= 15
            reasons.append("도심지 소음·진동 민원 주의 필요")

        if is_narrow and m["협소부지"] == "불가":
            score -= 45
            reasons.append("협소 부지 대형 장비 진입 불가")

        # --- 조건부 가산점 (실제 시공사례 기반 보정) ---
        # ① 동다짐: 소음·진동이 허용되는 해안/매립지/제약없는 부지에서는 채택 사례가
        #    매우 많다. 도심지가 아닐 때만 실무 활용빈도를 추가로 반영한다.
        if "동다짐" in m["공법명"] and not is_urban:
            score += 10
            bonuses.append(
                "소음·진동 허용 지역(해안/매립지 등)에서 적용사례가 많아 가산 반영"
            )

        # ② 치환공법: 자갈·쇄석·모래·순환골재 등 다양한 치환재를 활용할 수 있어,
        #    액상화 우려층이 국소적으로 얕게(대략 3~5m 이내) 분포하는 경우 경제성이
        #    특히 우수하고 채택 사례도 많다. 개량심도가 짧고 위험구간 폭도 좁을 때만 가산.
        if "치환" in m["공법명"] and max_depth <= 5.0 and ng_span <= 3.0:
            score += 12
            bonuses.append(
                f"국소 얕은 심도(구간폭 {ng_span:.1f}m, 최대심도 {max_depth:.1f}m)에 "
                "다양한 치환재(자갈·쇄석·모래·순환골재)를 활용한 경제적 시공사례가 많아 가산 반영"
            )

        # ③ 현장 실무형: 상부 비액상화층(H1) 확보로 지표 피해를 억제할 수 있는 경우 (Ishihara 1985)
        if _is_crust:
            score += 15
            bonuses.append(f"현장 실무형: 상부 GL-{crust_depth:.1f}m 처리로 비액상화층 확보(지표피해 억제), "
                           "하부 액상화층은 시설물 안정성 평가·기초 설계에 반영")

        score = round(min(100.0, max(10.0, score)), 1)

        eval_results.append(
            {
                "보강 공법": m["공법명"],
                "공법 분류": m["분류"],
                "적용 심도": m["적용심도"],
                "FC 제약": m["FC_제약"],
                "평가 점수": score,
                "경제성": m["경제성"],
                "실무 활용빈도": m["활용빈도"],
                "시공성/환경성": m["시공성"],
                "검토 의견": (
                    ", ".join(bonuses + reasons)
                    if (bonuses or reasons)
                    else "현장 지층 및 입지 조건 완벽 부합"
                ),
            }
        )

    res_df = (
        pd.DataFrame(eval_results)
        .sort_values(by="평가 점수", ascending=False)
        .reset_index(drop=True)
    )

    ranks = []
    for idx in range(len(res_df)):
        if idx == 0:
            ranks.append("🥇 1순위 (최적)")
        elif idx == 1:
            ranks.append("🥈 2순위 (추천)")
        elif idx == 2:
            ranks.append("🥉 3순위 (검토)")
        else:
            ranks.append(f"{idx+1}순위")

    res_df.insert(0, "추천 순위", ranks)
    return (
        res_df,
        f"최대 개량필요 심도: {max_depth:.1f}m, 평균 세골함유량(FC): {avg_fc:.1f}%",
    )


# ---------------------------------------------------------
# 6. 메인 프로그램 제어
# ---------------------------------------------------------
if uploaded_file is not None:
    try:
        df_raw = normalize_input_columns(pd.read_excel(uploaded_file))

        # --- 시추공 최대 100개 ---
        MAX_BOREHOLES = 100
        _bh_all = list(dict.fromkeys(df_raw["Borehole"]))
        if len(_bh_all) > MAX_BOREHOLES:
            st.warning(
                f"⚠️ 업로드한 시추공이 {len(_bh_all)}개입니다. 최대 {MAX_BOREHOLES}개까지 검토하므로 "
                f"앞의 {MAX_BOREHOLES}개({_bh_all[0]} ~ {_bh_all[MAX_BOREHOLES - 1]})만 평가합니다."
            )
            df_raw = df_raw[df_raw["Borehole"].isin(_bh_all[:MAX_BOREHOLES])].copy()

        # --- 고소성 점토거동 판별 (예비평가 생략조건 ③과 연동) ---
        df_clay_result = None
        if uploaded_clay_file is not None:
            df_clay_raw = pd.read_excel(uploaded_clay_file)
            df_clay_result = run_clay_behavior_eval(df_clay_raw)

        # --- 시추공 좌표 (액상화 분포도) ---
        coords_df = None
        if uploaded_coord_file is not None:
            _craw = (pd.read_csv(uploaded_coord_file, header=None) if uploaded_coord_file.name.lower().endswith(".csv")
                     else pd.read_excel(uploaded_coord_file, header=None))
            coords_df, _cmsg = parse_coords_table(_craw, swap_xy=coord_swap_xy)
            if coords_df is None:
                st.warning(_cmsg + " 사이드바의 '좌표 입력양식 받기'로 양식을 확인해주세요.")
            else:
                if _cmsg:
                    st.info("🗺️ " + _cmsg)
                _miss = [b_ for b_ in dict.fromkeys(df_raw["Borehole"].astype(str)) if b_ not in set(coords_df["Borehole"])]
                if _miss:
                    st.warning(f"좌표가 없는 시추공 {len(_miss)}개: {', '.join(_miss[:10])}{' 등' if len(_miss) > 10 else ''} "
                               "(분포도·영향면적에서 제외)")

        site_response_log_df = None
        wave_dict_all, dt_per_wave = {}, {}
        sr_case_table, sr_case_dfs, sr_case_logs = None, {}, {}
        sr_case_results = {}
        sr_applied = "외부 지진응답해석 결과 (amax/g 직접입력)"
        wave_names_for_check = []
        if site_response_mode == "② 시간이력 → 자체 등가선형해석":
            if uploaded_th_file is None:
                st.warning(
                    "⚠️ '② 시간이력 → 자체 등가선형해석'을 선택하셨습니다. "
                    "사이드바에서 지진파 시간이력 엑셀을 업로드해주세요."
                )
                st.stop()
            th_raw = pd.read_excel(uploaded_th_file)
            wave_names_for_check = [c for c in th_raw.columns if not str(c).startswith("시간")]
            # 지진파별 시간 컬럼('시간_지진파(s)') 형식과 공통 '시간(s)' 형식을 모두 지원
            per_wave_time = all(f"시간_{w}(s)" in th_raw.columns for w in wave_names_for_check)
            if per_wave_time:
                for w in wave_names_for_check:
                    sub = th_raw[[f"시간_{w}(s)", w]].dropna()
                    wave_dict_all[w] = sub[w].to_numpy(dtype=float)
                    dt_per_wave[w] = float(np.median(np.diff(sub[f"시간_{w}(s)"].to_numpy())))
            else:
                time_col = "시간(s)" if "시간(s)" in th_raw.columns else th_raw.columns[0]
                dt_shared = float(np.median(np.diff(th_raw[time_col].to_numpy())))
                for w in wave_names_for_check:
                    wave_dict_all[w] = th_raw[w].dropna().to_numpy(dtype=float)
                    dt_per_wave[w] = dt_shared
            # 기존에 입력엑셀에 같은 이름의 amax 열이 있으면 해석결과로 대체되도록 제거
            df_raw = df_raw.drop(columns=[w for w in wave_names_for_check if w in df_raw.columns])
            df_input_raw = df_raw.copy()

            # --- 해석 케이스 일괄 해석 (입력 지진파는 설계자가 정합한 원본 그대로 사용) ---
            sel_cases = [c for c in SR_CASES if c["케이스"] in sr_case_choice] or [SR_CASES[0]]
            if custom_case is not None:
                # 이름이 기본 케이스와 겹치지 않도록 정리
                if any(custom_case["케이스"] == c["케이스"] for c in sel_cases):
                    custom_case["케이스"] = custom_case["케이스"] + " (사용자)"
                sel_cases = sel_cases + [custom_case]
            _prog = st.progress(0.0, text="지진응답해석 준비 중...")
            for ci, c in enumerate(sel_cases):
                def _cb(i, n, ci=ci, c=c):
                    _prog.progress(min((ci + i / max(n, 1)) / len(sel_cases), 1.0),
                                   text=f"지진응답해석 {c['케이스']} ({ci + 1}/{len(sel_cases)}) - 시추공 {i + 1}/{n}")
                _k = float(c.get("input_scale", 1.0))
                _wd_case = wave_dict_all if abs(_k - 1.0) < 1e-12 else scale_wave_dict(wave_dict_all, _k)
                d_out, lg = run_sr_case_cached(
                    df_raw, _wd_case, dt_per_wave,
                    c["curve_shift"], c["변형률비"], c["vs_factor"], c["bedrock_vs"], _progress_cb=_cb,
                )
                sr_case_dfs[c["케이스"]], sr_case_logs[c["케이스"]] = d_out, lg
            _prog.empty()
            # 검토서 적용 가능 케이스(기본 케이스 + 범위 내 사용자 정의)만 평균·포락에 포함
            _reportable = [c["케이스"] for c in sel_cases if not c.get("위반")]
            _ref_only = [c["케이스"] for c in sel_cases if c.get("위반")]
            if len(_reportable) >= 2:
                _base_dfs = {k: sr_case_dfs[k] for k in _reportable}
                sr_case_dfs[SR_DERIVED[0]] = derive_case_df(_base_dfs, wave_names_for_check, "mean")
                sr_case_dfs[SR_DERIVED[1]] = derive_case_df(_base_dfs, wave_names_for_check, "max")
                _reportable += SR_DERIVED

            # --- 케이스별 액상화 평가 결과 비교 ---
            _rows = []
            sr_case_results = {}
            for lab, d in sr_case_dfs.items():
                _sp, _vs, _ = run_liquefaction_eval(d, n_skip=N_SKIP, clay_zones=df_clay_result, csr_method=CSR_METHOD,
                                                    vs1_mode=VS1_MODE, pga_floor=PGA_FLOOR)
                sr_case_results[lab] = (_sp, _vs)
                meta = next((c for c in sel_cases if c["케이스"] == lab), None)
                row = {"id": meta["id"] if meta else "-", "케이스": lab,
                       "동적곡선": meta["동적곡선"] if meta else "-",
                       "curve_shift": meta["curve_shift"] if meta else np.nan,
                       "변형률비": meta["변형률비"] if meta else np.nan,
                       "vs_factor": meta["vs_factor"] if meta else np.nan,
                       "bedrock_vs": meta["bedrock_vs"] if meta else np.nan,
                       "입력배율": float(meta.get("input_scale", 1.0)) if meta else 1.0,
                       "검토서 적용": "참고용(범위 밖)" if lab in _ref_only else "가능",
                       "근거": (meta["근거"] if meta else ("적용 가능 케이스들의 심도별 amax 평균" if lab == SR_DERIVED[0]
                                                         else "적용 가능 케이스들의 심도별 amax 최대값(포락)"))
                               + (" / 범위 밖: " + "; ".join(meta["위반"]) if meta and meta.get("위반") else "")}
                row.update(case_metrics(_sp, _vs, d, wave_names_for_check))
                _pg = surface_pga_table(d, wave_names_for_check, KDS_S, kds_soil_class)
                row["지표PGA/S×Fa 최소"] = float(_pg["평균/하한"].min()) if len(_pg) else np.nan
                _rows.append(row)
            sr_case_table = pd.DataFrame(_rows)

            _labels = [l for l in sr_case_dfs.keys() if l in _reportable]
            _safe = set(sr_case_table.loc[sr_case_table["액상화 미발생"], "케이스"])
            st.markdown("#### ✅ 적용할 지진응답해석 결과")
            sr_applied = st.selectbox(
                "해석 케이스 선택 (비교 결과는 '🌍 지반응답해석' 탭 참조)", _labels, index=0,
                format_func=lambda x: x + ("   · 액상화 미발생" if x in _safe else ""),
                key="sr_applied_sel",
            )
            if _ref_only:
                st.caption("※ 참고용(범위 밖) 케이스 " + ", ".join(_ref_only)
                           + "는 '🌍 지반응답해석' 탭에서 결과를 볼 수 있으며, 적용 결과 선택 목록에는 포함되지 않습니다.")
            df_raw = sr_case_dfs[sr_applied]
            site_response_log_df = sr_case_logs.get(sr_applied, sr_case_logs[sel_cases[0]["케이스"]])

        df_spt, df_vs, wave_cols = run_liquefaction_eval(df_raw, n_skip=N_SKIP, clay_zones=df_clay_result,
                                                         csr_method=CSR_METHOD, vs1_mode=VS1_MODE, pga_floor=PGA_FLOOR)
        judge_df = borehole_judgement(df_spt, df_vs)
        # KDS 17 10 00:2024 4.7(6)① 단서 - 지표면 최대가속도 ≥ S×Fa 검토 (적용 결과 기준)
        kds_pga_df = surface_pga_table(df_raw, wave_cols, KDS_S, kds_soil_class)
        # 계산서·검토서에는 검토서 적용 가능 케이스만 수록 (범위 밖 참고용 케이스는 화면 비교 전용)
        sr_case_table_out = (sr_case_table[sr_case_table["검토서 적용"] == "가능"].reset_index(drop=True)
                             if sr_case_table is not None and "검토서 적용" in sr_case_table.columns else sr_case_table)

        tab_clay, tab_site, tab_spt, tab_vs, tab_summary = st.tabs(
            [
                "🧱 고소성 점토거동 평가",
                "🌍 지반응답해석",
                "📌 1. SPT 액상화 평가",
                "📌 2. Vs 액상화 평가",
                "📋 3. 종합 평가",
            ]
        )

        # TAB. 지반응답해석 (SHAKE 대체) - 해석 케이스 비교 및 적용 결과
        sr_reason = ""
        sr_rep_bh = None
        with tab_site:
            def _show_kds_pga_check(_pg, title_suffix=""):
                st.markdown(f"### 📏 KDS 지표면 최대가속도 하한 검토{title_suffix}")
                st.caption(
                    f"KDS 17 10 00:2024 4.7(6)① 단서: 지표면 최대지반가속도는 기반암 유효수평지반가속도와 "
                    f"단주기지반증폭계수의 곱(S×Fa = {KDS_S:.3f}×{KDS_FA:.3f} = {KDS_SFA:.3f}g) 이상이어야 합니다. "
                    "지표면 값은 지진파별 지표 최대가속도의 평균입니다."
                )
                if _pg is None or _pg.empty:
                    st.info("지표면 가속도 자료가 없습니다.")
                    return
                _pg_sty = _pg.style.format({"지표 PGA 평균(g)": "{:.3f}", "지표 PGA 최소(g)": "{:.3f}",
                                            "지표 PGA 최대(g)": "{:.3f}", "S(g)": "{:.3f}", "Fa": "{:.3f}",
                                            "하한 S×Fa(g)": "{:.3f}", "평균/하한": "{:.2f}"})
                _pg_sty = _style_map(_pg_sty, lambda v: "color:#C0392B;font-weight:bold" if v == "미달" else
                                     ("color:#1E8449;font-weight:bold" if v == "만족" else ""), subset=["판정"])
                st.dataframe(_pg_sty, use_container_width=True, hide_index=True)
                n_low = int((_pg["판정"] == "미달").sum())
                if n_low == 0:
                    st.success("모든 시추공에서 지표면 최대가속도가 KDS 하한(S×Fa) 이상입니다.")
                elif CSR_METHOD == "pga_rd":
                    st.info(f"{n_low}개 시추공이 하한 미달이나, 'KDS 단서(PGA×rd)' 방법은 지진파별로 "
                            f"max(PGA, {KDS_SFA:.3f}g)를 적용하므로 하한이 이미 반영되어 있습니다.")
                elif CSR_METHOD == "amax":
                    st.warning(f"{n_low}개 시추공의 지표면 최대가속도가 S×Fa 미만입니다. 가속도 기반 CSR(심도별 amax 근사)은 "
                               "KDS 단서 조건을 만족하지 못하므로, CSR 산정을 '최대전단응력 τmax (원칙)' 또는 "
                               "'지표 PGA × rd (KDS 단서)'로 바꿔 검토하세요.")
                else:
                    st.info(f"{n_low}개 시추공의 지표면 최대가속도가 S×Fa 미만입니다. τmax(원칙) 방법에는 이 하한이 직접 "
                            "적용되지 않지만, 입력 지진파 크기·해석조건이 설계지반운동수준에 맞는지 확인하는 지표로 활용하세요.")

            if site_response_mode != "② 시간이력 → 자체 등가선형해석":
                st.info(
                    "현재 '① 결과(amax/g) 직접입력' 모드입니다. 사이드바에서 "
                    "'② 시간이력 → 자체 등가선형해석'을 선택하면 이 탭에서 해석 케이스 비교와 "
                    "지반응답해석 수행 결과를 확인할 수 있습니다."
                )
                _show_kds_pga_check(kds_pga_df, " (업로드한 외부 해석결과)")
            else:
                st.markdown("### 🧪 지진응답해석 케이스 비교")
                st.caption(
                    "1차원 등가선형해석(SHAKE) 결과는 입력가정(동적물성곡선, 유효변형률비, 전단파속도 "
                    "불확실성, 기반암 가정)에 따라 달라집니다. C1~C7은 기술적으로 타당한 범위의 가정이며, "
                    "입력 지진파는 설계자가 정합한 원본을 그대로 사용했습니다. 사용자 정의 케이스는 사이드바에서 설정합니다."
                )
                _show = sr_case_table.copy()
                _show.insert(0, "적용", np.where(_show["케이스"] == sr_applied, "◀ 적용", ""))
                _show["액상화"] = np.where(_show["액상화 미발생"], "미발생", "발생")
                st.dataframe(
                    _show[["적용", "케이스", "검토서 적용", "동적곡선", "변형률비", "vs_factor", "bedrock_vs", "입력배율",
                           "지표 amax 평균(g)", "지표 amax 최대(g)", "지표PGA/S×Fa 최소", "최소FS(SPT)", "최소FS(Vs)",
                           "NG 시추공", "NG 층수(SPT)", "NG 층수(Vs)", "액상화", "근거"]].rename(
                        columns={"vs_factor": "토층 Vs 배율", "bedrock_vs": "기반암 Vs(미도달 시)"}
                    ).style.format({"변형률비": "{:.2f}", "토층 Vs 배율": "{:.2f}", "기반암 Vs(미도달 시)": "{:,.0f}",
                                    "입력배율": "{:.2f}", "지표 amax 평균(g)": "{:.3f}", "지표 amax 최대(g)": "{:.3f}",
                                    "지표PGA/S×Fa 최소": "{:.2f}",
                                    "최소FS(SPT)": "{:.3f}", "최소FS(Vs)": "{:.3f}"}, na_rep="-"),
                    use_container_width=True, hide_index=True,
                )
                _safe_rep = [k for k in _safe if k in _labels]
                _safe_ref = [k for k in _safe if k not in _labels]
                if _safe_rep:
                    st.success(f"액상화가 발생하지 않는 케이스(검토서 적용 가능): {', '.join(sorted(_safe_rep))}")
                else:
                    st.info(
                        "검토서 적용이 가능한 모든 케이스에서 액상화 발생 구간이 확인됩니다. "
                        "입력가정 선택과 무관하게 액상화 대책(또는 KDS 4.7(7) 시설물 안정성 평가)이 필요한 것으로 판단됩니다."
                    )
                if _safe_ref:
                    st.warning(f"액상화 미발생 결과가 나온 참고용 케이스: {', '.join(sorted(_safe_ref))} — 검토서 적용 범위 밖 "
                               "설정이므로 적용 결과로 선택할 수 없습니다.")
                _fig_c = go.Figure()
                _fig_c.add_bar(x=sr_case_table["케이스"], y=sr_case_table["최소FS(SPT)"], name="최소 FS (SPT)",
                               marker_color="#2E86C1")
                _fig_c.add_bar(x=sr_case_table["케이스"], y=sr_case_table["최소FS(Vs)"], name="최소 FS (Vs)",
                               marker_color="#8E44AD")
                _fig_c.add_hline(y=1.0, line_dash="dash", line_color="red", annotation_text="FS = 1.0")
                _fig_c.update_layout(barmode="group", template="plotly_white", height=380,
                                     title="<b>케이스별 최소 안전율</b>", yaxis_title="최소 안전율 (FS)",
                                     legend=dict(orientation="h", y=-0.25, x=0.5, xanchor="center"))
                st.plotly_chart(_fig_c, use_container_width=True)

                # --- 사용자 정의 케이스 결과 (보기) ---
                if custom_case is not None and custom_case["케이스"] in sr_case_results:
                    _cl = custom_case["케이스"]
                    _csp, _cvs = sr_case_results[_cl]
                    with st.expander(f"🔎 사용자 정의 케이스 결과 보기 — {_cl}"
                                     + (" (참고용·검토서 미반영)" if custom_case.get("위반") else ""), expanded=True):
                        if custom_case.get("위반"):
                            st.warning("검토서 적용 범위 밖 설정: " + "; ".join(custom_case["위반"]))
                        _cj = borehole_judgement(_csp, _cvs)
                        _bj = borehole_judgement(*sr_case_results[sr_applied]) if sr_applied in sr_case_results else None
                        if _bj is not None:
                            _cj = _cj.merge(_bj[["Borehole", "최소FS"]].rename(columns={"최소FS": f"최소FS ({sr_applied})"}),
                                            on="Borehole", how="left")
                        st.dataframe(_cj, use_container_width=True, hide_index=True)
                        _need = []
                        for _lab, _d in (("SPT", _csp), ("Vs", _cvs)):
                            _t = _d[_d["대상층여부"] == "대상층"]
                            if _t["FS"].notna().any():
                                _need.append(f"{_lab} 최소 FS {float(_t['FS'].min()):.3f}")
                        st.caption("설정값: 변형률비 {:.2f}, 동적곡선 ×{:g}, Vs ×{:.2f}, 기반암 Vs {:,.0f} m/s, 입력배율 {:.2f} → {}".format(
                            custom_case["변형률비"], custom_case["curve_shift"], custom_case["vs_factor"],
                            custom_case["bedrock_vs"], custom_case["input_scale"], ", ".join(_need)))

                _min_bh = judge_df.dropna(subset=["최소FS"]).sort_values("최소FS")
                _bh_opts = list(dict.fromkeys(df_raw["Borehole"]))
                sr_rep_bh = st.selectbox(
                    "케이스별 심도별 최대가속도를 비교할 시추공", _bh_opts,
                    index=_bh_opts.index(_min_bh["Borehole"].iloc[0]) if len(_min_bh) else 0,
                    key="sr_rep_bh",
                )
                _fig_p = go.Figure()
                for lab, d in sr_case_dfs.items():
                    g = d[d["Borehole"] == sr_rep_bh].sort_values("심도(GL,m)")
                    _fig_p.add_trace(go.Scatter(
                        x=g[wave_names_for_check].mean(axis=1), y=g["심도(GL,m)"], mode="lines+markers",
                        name=lab, line=dict(width=4 if lab == sr_applied else 1.5,
                                            dash="dot" if lab not in _labels else "solid"),
                    ))
                _fig_p.add_vline(x=KDS_SFA, line_dash="dash", line_color="red",
                                 annotation_text=f"S×Fa = {KDS_SFA:.3f}g", annotation_position="bottom right")
                _fig_p.update_layout(template="plotly_white", height=520,
                                     title=f"<b>{sr_rep_bh} 케이스별 심도별 최대가속도 (지진파 평균)</b>",
                                     xaxis=dict(title="최대가속도 (g)", side="top"),
                                     yaxis=dict(title="심도 (m)", autorange="reversed"))
                st.plotly_chart(_fig_p, use_container_width=True)

                sr_reason = st.text_area(
                    "적용 케이스 선정 사유 (검토서 3.5절에 기재)",
                    value=f"{sr_applied}를 적용함. 지반조사 결과와 해석 가정의 대표성을 고려하여 선정함.",
                    key="sr_reason",
                )

                st.markdown("---")
                _show_kds_pga_check(kds_pga_df, f" (적용 케이스: {sr_applied})")

                # --- 외부 지진응답해석 결과와 비교 ---
                if uploaded_ext_file is not None:
                    st.markdown("---")
                    st.markdown("### 🔍 외부 지진응답해석 결과와 비교")
                    try:
                        _ext = normalize_input_columns(pd.read_excel(uploaded_ext_file))
                        _cmp = compare_external_sr(_ext, df_raw, wave_names_for_check, KDS_S, kds_soil_class,
                                                   n_skip=N_SKIP, clay_zones=df_clay_result, vs1_mode=VS1_MODE)
                    except Exception as _e:
                        _cmp = None
                        st.error(f"외부 결과 파일을 읽지 못했습니다: {_e}")
                    if _cmp is None:
                        st.warning("외부 결과에서 지진파 이름·시추공·심도가 일치하는 자료를 찾지 못했습니다. "
                                   "지진파 열 이름과 Borehole·심도(GL,m)를 입력엑셀과 같게 맞춰주세요.")
                    else:
                        st.caption(f"비교 기준: 프로그램 = 적용 케이스 '{sr_applied}', 지진파 {len(_cmp['waves'])}개 공통")
                        for _t in external_diagnosis_text(_cmp):
                            st.markdown(f"- {_t}")
                        st.markdown("**① 시추공별 가속도 비교**")
                        st.dataframe(_cmp["accel"].style.format({c: "{:.3f}" for c in _cmp["accel"].columns if c != "Borehole"}),
                                     use_container_width=True, hide_index=True)
                        st.markdown("**② CSR 산정방법·해석결과별 최소 안전율** (①→② 가속도 차이, ②→③ 산정방법 차이, ④·⑤ KDS 하한 적용)")
                        _c1, _c2 = st.columns(2)
                        _c1.markdown("SPT 방법")
                        _c1.dataframe(_cmp["fs_spt"].style.format({c: "{:.3f}" for c in _cmp["fs_spt"].columns if c != "Borehole"}),
                                      use_container_width=True, hide_index=True)
                        _c2.markdown("Vs 방법")
                        _c2.dataframe(_cmp["fs_vs"].style.format({c: "{:.3f}" for c in _cmp["fs_vs"].columns if c != "Borehole"}),
                                      use_container_width=True, hide_index=True)
                        _bh_e = st.selectbox("심도별 가속도 비교 시추공", list(dict.fromkeys(_cmp["prof"]["Borehole"])),
                                             key="ext_bh")
                        _g = _cmp["prof"][_cmp["prof"]["Borehole"] == _bh_e].sort_values("심도(GL,m)")
                        _fe = go.Figure()
                        _fe.add_trace(go.Scatter(x=_g[[f"{x}_외부" for x in _cmp["waves"]]].mean(axis=1), y=_g["심도(GL,m)"],
                                                 mode="lines+markers", name="외부 해석 (지진파 평균)", line=dict(color="#E67E22", width=3)))
                        _fe.add_trace(go.Scatter(x=_g[[f"{x}_프로그램" for x in _cmp["waves"]]].mean(axis=1), y=_g["심도(GL,m)"],
                                                 mode="lines+markers", name=f"프로그램 ({sr_applied})", line=dict(color="#1F4E78", width=3)))
                        _fe.add_vline(x=_cmp["sfa"], line_dash="dash", line_color="red", annotation_text="S×Fa")
                        _fe.add_vline(x=KDS_S, line_dash="dot", line_color="gray", annotation_text="입력 S")
                        _fe.update_layout(template="plotly_white", height=480, title=f"<b>{_bh_e} 심도별 최대가속도 비교</b>",
                                          xaxis=dict(title="최대가속도 (g)", side="top"),
                                          yaxis=dict(title="심도 (m)", autorange="reversed"))
                        st.plotly_chart(_fe, use_container_width=True)
                        _bc_sig = (getattr(uploaded_ext_file, "name", ""), _wave_hash(wave_dict_all, dt_per_wave),
                                   len(df_input_raw))
                        if st.button("🧮 외부 결과를 재현하는 입력지진 배율 역산 (진단용)", key="ext_backcalc"):
                            _pb = st.progress(0.0, text="역산 중...")
                            _bc = backcalc_input_scale(
                                _ext, df_input_raw, wave_dict_all, dt_per_wave,
                                _progress_cb=lambda i, n: _pb.progress(i / max(n, 1), text=f"입력배율 {i + 1}/{n} 해석 중"))
                            _pb.empty()
                            st.session_state["ext_backcalc_result"] = (_bc_sig, _bc)
                        _bc_state = st.session_state.get("ext_backcalc_result")
                        if _bc_state is not None and _bc_state[0] == _bc_sig:
                            _bc = _bc_state[1]
                            _best = _bc.iloc[(_bc["amax 비 평균(외부/프로그램)"] - 1.0).abs().argmin()]
                            st.dataframe(_bc.style.format({"입력배율": "{:.2f}", "amax 비 평균(외부/프로그램)": "{:.3f}",
                                                           "평균 상대오차": "{:.1%}"}), use_container_width=True, hide_index=True)
                            st.info(f"외부 결과는 프로그램 기본조건(C1)에서 입력지진을 약 {_best['입력배율']:.1f}배로 줄였을 때와 "
                                    "평균 가속도가 비슷합니다. 외부 해석의 입력 가속도 크기·입력 방식을 우선 확인하세요"
                                    "(동적물성곡선 등 다른 조건 차이로도 일부 설명될 수 있음).")

                st.markdown("---")
                st.markdown(f"### 🌊 적용 케이스 지반응답해석 수행 결과 — {sr_applied}")
                st.caption("시추공×지진파별 반복계산 수렴 여부와 지표/기반암상단 amax/g 요약입니다.")
                st.dataframe(site_response_log_df, use_container_width=True)
                if not site_response_log_df["수렴여부"].all():
                    st.warning(
                        "일부 시추공×지진파 조합에서 등가선형 반복계산이 허용오차 이내로 수렴하지 "
                        "못했습니다. 지반이 매우 연약하거나 입력 가속도가 큰 경우일 수 있으니 결과를 검토해주세요."
                    )

        # TAB 1. 고소성 점토거동 평가 (KDS 17 10 00:2024)
        with tab_clay:
            st.markdown(CLAY_INTRO_MD)
            st.markdown("---")
            if df_clay_result is None:
                st.info(
                    "👈 사이드바의 '고소성 점토거동 평가 엑셀 파일'을 업로드하면 판별 결과가 표시됩니다. "
                    "액상화 평가 엑셀 없이 점토 엑셀만 올려도 개별 평가 화면을 볼 수 있습니다. "
                    "입력 열: Borehole, 토층, 시작심도(m), 종료심도(m), LL(%), PI, wc(%)"
                )
            else:
                render_clay_section(df_clay_result, key="main")
                st.download_button("📥 점토거동 판별 결과 엑셀", data=build_clay_excel(df_clay_result),
                                   file_name="점토거동_판별결과.xlsx", key="clay_xlsx_main",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # TAB 2. SPT 기반 평가
        with tab_spt:
            st.markdown("""
            ### 📝 SPT(표준관입시험) 기반 액상화 평가 방법
            본 검토는 **내진설계 일반(KDS 17 10 00 : 2024)** 에 의거하여 수행되었습니다.
            """)
            st.markdown("---")

            st.subheader("1️⃣ 액상화 대상층 평가 결과")
            df_spt_target = df_spt[df_spt["대상층여부"] == "대상층"].copy()
            show_spt_target = df_spt_target[
                [
                    "Borehole",
                    "토층",
                    "시작심도(m)",
                    "종료심도(m)",
                    "N치(N60)",
                    "FC(%)",
                    "CSR",
                    "CRR",
                    "FS",
                    "판정",
                ]
            ]
            st.dataframe(show_spt_target, use_container_width=True)

            if not df_spt_target.empty:
                st.markdown(
                    "#### 📊 시추공별 지진파 안전율"
                )
                selected_bh_spt = st.selectbox(
                    "조회할 시추공 선택 (SPT)",
                    df_spt_target["Borehole"].unique(),
                    key="spt_bh_select",
                )
                bh_spt_df = df_spt_target[
                    df_spt_target["Borehole"] == selected_bh_spt
                ]
                st.plotly_chart(
                    make_depth_fs_profile_chart(
                        bh_spt_df, wave_cols, f"SPT - {selected_bh_spt}"
                    ),
                    use_container_width=True,
                )

            st.markdown("---")

            st.subheader("2️⃣ 시추공별 안전율 평가")
            if not df_spt_target.empty:
                spt_min_df = build_min_fs_summary(df_spt_target)

                col_t, col_g = st.columns([0.4, 0.6])
                with col_t:
                    st.markdown("**시추공별 최소 안전율 요약표**")
                    st.dataframe(format_min_fs_for_display(spt_min_df), use_container_width=True)
                with col_g:
                    st.plotly_chart(
                        make_fs_chart(spt_min_df, "SPT 방법"),
                        use_container_width=True,
                    )
            else:
                st.info("평가 대상층이 존재하지 않습니다.")

        # TAB 3. Vs 기반 평가
        with tab_vs:
            st.markdown("""
            ### 📝 Vs(전단파속도) 기반 액상화 평가 방법
            본 검토는 **내진설계 일반(KDS 17 10 00 : 2024)** 에 의거하여 수행되었습니다.
            """)
            st.markdown("---")

            st.subheader("1️⃣ 액상화 대상층 평가 결과")
            df_vs_target = df_vs[df_vs["대상층여부"] == "대상층"].copy()
            show_vs_target = df_vs_target[
                [
                    "Borehole",
                    "토층",
                    "시작심도(m)",
                    "종료심도(m)",
                    "Vs(m/s)",
                    "FC(%)",
                    "CSR",
                    "CRR",
                    "FS",
                    "판정",
                ]
            ]
            st.dataframe(show_vs_target, use_container_width=True)

            if not df_vs_target.empty:
                st.markdown(
                    "#### 📊 시추공별 지진파 안전율"
                )
                selected_bh_vs = st.selectbox(
                    "조회할 시추공 선택 (Vs)",
                    df_vs_target["Borehole"].unique(),
                    key="vs_bh_select",
                )
                bh_vs_df = df_vs_target[
                    df_vs_target["Borehole"] == selected_bh_vs
                ]
                st.plotly_chart(
                    make_depth_fs_profile_chart(
                        bh_vs_df, wave_cols, f"Vs - {selected_bh_vs}"
                    ),
                    use_container_width=True,
                )

            st.markdown("---")

            st.subheader("2️⃣ 시추공별 안전율 평가")
            if not df_vs_target.empty:
                vs_min_df = build_min_fs_summary(df_vs_target)

                col_t, col_g = st.columns([0.4, 0.6])
                with col_t:
                    st.markdown("**시추공별 최소 안전율 요약표**")
                    st.dataframe(format_min_fs_for_display(vs_min_df), use_container_width=True)
                with col_g:
                    st.plotly_chart(
                        make_fs_chart(vs_min_df, "Vs 방법"),
                        use_container_width=True,
                    )
            else:
                st.info("평가 대상층이 존재하지 않습니다.")

        # TAB 4. 종합 평가
        with tab_summary:
            with st.expander("🧭 액상화 평가 흐름도 (KDS 17 10 00:2024)", expanded=False):
                st.image(fig_kds_flow_png(), use_container_width=True)
                st.caption(
                    "본 프로그램의 판정 절차: 예비평가(생략조건 ① 지하수위 상부 ② N>25 ③ 고소성 점토거동 영역 C) "
                    "→ 본평가(지진응답해석 CSR, SPT·Vs 방법 CRR, FS = CRR/CSR > 1.0) → 액상화 발생 시 대책공법 수립·재평가"
                )

            st.subheader("🔎 시추공별 종합판정 (KDS 17 10 00:2024)")
            st.caption("SPT-N 방법과 Vs 방법 중 작은 안전율로 판정합니다(보수적). 기준안전율 FS = 1.0")
            _jc = judge_df["종합판정"].value_counts()
            jc1, jc2, jc3 = st.columns(3)
            jc1.metric("안전 (FS > 1.0)", f"{_jc.get('안전 (FS > 1.0)', 0)}공")
            jc2.metric("액상화 발생 (FS ≤ 1.0)", f"{_jc.get('액상화 발생 (FS ≤ 1.0)', 0)}공")
            jc3.metric("평가생략 (예비평가)", f"{_jc.get('평가생략(예비평가)', 0)}공")
            st.dataframe(
                judge_df.style.format({"최소FS(SPT)": "{:.3f}", "최소FS(Vs)": "{:.3f}", "최소FS": "{:.3f}"},
                                      na_rep="-"),
                use_container_width=True, hide_index=True,
            )
            if coords_df is not None and len(coords_df):
                _nm = len(coords_df.merge(judge_df, on="Borehole"))
                if _nm:
                    st.plotly_chart(make_liq_map_plotly(coords_df, judge_df), use_container_width=True)
                    if _nm < len(judge_df):
                        st.caption(f"※ 좌표가 일치하는 시추공 {_nm}/{len(judge_df)}개소만 도시했습니다.")
                else:
                    st.warning("좌표 파일의 시추공명이 액상화 평가 입력의 시추공명과 일치하지 않습니다.")
            else:
                st.caption("🗺️ 사이드바에서 시추공 좌표 엑셀을 업로드하면 액상화 분포도가 표시됩니다.")
            st.markdown("---")

            st.subheader("📋 1. 액상화 평가 종합 요약")

            if not df_spt_target.empty and not df_vs_target.empty:
                summary_df = pd.merge(
                    spt_min_df[
                        ["Borehole", "FS", "최소안전율 지진파", "최종판정"]
                    ],
                    vs_min_df[
                        ["Borehole", "FS", "최소안전율 지진파", "최종판정"]
                    ],
                    on="Borehole",
                    suffixes=(" (SPT)", " (Vs)"),
                )

                m1, m2, m3 = st.columns(3)
                m1.metric("검토 시추공 수", f"{len(summary_df)}개소")
                m2.metric(
                    "SPT 평가 결과",
                    (
                        "전 시추공 안전"
                        if (summary_df["최종판정 (SPT)"] == "안전(만족)").all()
                        else "NG 구간 존재"
                    ),
                )
                m3.metric(
                    "Vs 평가 결과",
                    (
                        "전 시추공 안전"
                        if (summary_df["최종판정 (Vs)"] == "안전(만족)").all()
                        else "NG 구간 존재"
                    ),
                )

                st.dataframe(summary_df, use_container_width=True)

                st.markdown("#### 📊 전체 시추공 SPT / Vs 안전율")
                st.plotly_chart(
                    make_summary_comparison_chart(summary_df),
                    use_container_width=True,
                )

                ng_spt_layers = df_spt[df_spt["판정"] == "NG(불만족)"]
                ng_vs_layers = df_vs[df_vs["판정"] == "NG(불만족)"]
                all_ng_layers = pd.concat(
                    [ng_spt_layers, ng_vs_layers]
                ).drop_duplicates(
                    subset=["Borehole", "시작심도(m)", "종료심도(m)"]
                )

                st.markdown("---")
                remed_df, site_info_msg = None, None
                crust_df = None
                rm_price_basis, rm_overhead = REMED_PRICE_BASIS_DEFAULT, 1.0
                rm_cost_db = [dict(m) for m in REMED_COST_DB]
                if len(all_ng_layers) > 0:
                    st.subheader(
                        "💡 2. 최적 액상화 보강대책 선정 (1, 2, 3 순위 산정)"
                    )
                    st.warning(
                        f"⚠️ 총 {len(all_ng_layers)}개 지층 구간에서 액상화 위험(FS ≤ 1.0)이 감지되었습니다."
                    )

                    # --- 현장 실무형 검토 (상부 비액상화층 확보) ---
                    with st.expander("⚙️ 현장 실무형 검토 조건 (상부 쇄석치환·동다짐)", expanded=False):
                        _e1, _e2, _e3 = st.columns(3)
                        crust_rep_limit = _e1.number_input("상부 치환 최대심도 (m)", 1.0, 6.0, 3.0, 0.5, key="crust_rep",
                                                           help="지하수위 하부 굴착은 배수·흙막이가 필요해 실무상 2~3m, 최대 5m 정도")
                        crust_dc_w = _e2.number_input("동다짐 추 무게 W (t)", 5.0, 40.0, 15.0, 1.0, key="crust_dcw",
                                                        help="적용 단가(16,000원/m²)는 15t급 기준")
                        crust_dc_h = _e3.number_input("동다짐 낙하고 H (m)", 5.0, 40.0, 20.0, 1.0, key="crust_dch")
                        st.caption("동다짐 유효심도 D = n·√(W·H), n = 0.5 (FC ≤ 15%), 0.4 (15% < FC ≤ 35%) "
                                   "(Menard & Broise 1975, Lukas 1995), 기본 15t×20m. 필요 비액상화층 두께는 Ishihara(1985) 경계곡선 "
                                   "상한(지표 PGA 0.2g: 3m, 0.3g: 6m, 0.45g 이상: 9m)을 보수적으로 적용합니다.")
                    _surf = (dict(zip(kds_pga_df["Borehole"], kds_pga_df["지표 PGA 평균(g)"]))
                             if kds_pga_df is not None and len(kds_pga_df) else {})
                    crust_df = crust_assessment(df_spt, df_vs, judge_df, _surf, KDS_SFA, rep_limit=crust_rep_limit,
                                                dc_weight=crust_dc_w, dc_drop=crust_dc_h,
                                                urban="도심지" in site_location, narrow="협소" in space_constraint)
                    remed_df, site_info_msg = evaluate_remediation_options(
                        all_ng_layers, site_location, space_constraint, priority=remed_priority, crust_df=crust_df
                    )

                    st.markdown("#### 🔧 현장 실무형 검토 — 상부 비액상화층 확보 (쇄석치환·동다짐)")
                    st.caption(
                        "KDS 17 10 00:2024 4.7(7)은 안전율 1 이하일 때 곧바로 지반개량을 요구하지 않고, 액상화 영향을 반영한 "
                        "시설물 안정성 평가 후 안정성을 확보하지 못하면 대책을 세우도록 합니다. 현장에서 상부 쇄석치환·동다짐을 "
                        "많이 쓰는 근거는 비액상화 표층(H1)이 충분하면 하부에 액상화층(H2)이 있어도 지표 피해(분사·균열)가 "
                        "억제된다는 Ishihara(1985) 기준입니다."
                    )
                    if crust_df is not None and len(crust_df):
                        st.dataframe(crust_df.style.format({c: "{:.2f}" if "PGA" in c else "{:.1f}"
                                                            for c in crust_df.columns
                                                            if c not in ("Borehole", "판정", "현장 실무형 대책")}),
                                     use_container_width=True, hide_index=True)
                        _n_ok = int((crust_df["필요 처리심도(m)"] <= 0).sum())
                        _n_up = int(crust_df["판정"].str.contains("확보 가능").sum())
                        _n_no = len(crust_df) - _n_ok - _n_up
                        _lines = []
                        for _, _cr in crust_df.iterrows():
                            _lines.append(f"* **{_cr['Borehole']}** : {_cr['판정']} → {_cr['현장 실무형 대책']}")
                        _head = ("**🔧 현장 실무형 추천** — " +
                                 ("모든 액상화 발생 시추공에서 상부 처리(또는 무처리)로 지표 피해 억제가 가능합니다."
                                  if _n_no == 0 else f"{_n_no}개소는 상부 처리만으로 불충분하여 심층개량 또는 구조물 대책이 필요합니다."))
                        (st.success if _n_no == 0 else st.warning)(_head + "\n" + "\n".join(_lines[:30])
                                                                   + ("\n* …" if len(_lines) > 30 else ""))
                        st.caption("상부 처리안은 수평지반 지표 피해 억제가 목적입니다. 말뚝 부마찰력·수평저항 감소, 얕은기초 "
                                   "지지력, 경사지·호안부 측방유동, 매설관 부상은 시설물 안정성 평가(KDS 4.7(7))에서 별도 검토하세요. "
                                   "공사비는 '3. 대안 비교'의 대안 4에 반영됩니다.")

                    st.markdown("#### 🏆 지층 및 입지 조건별 보강공법 추천 순위")
                    st.info(
                        f"📌 **지층 검토 조건**: {site_info_msg} | **입지 조건**: {site_location}, {space_constraint} | "
                        f"**추천 기준**: {remed_priority}"
                    )
                    st.dataframe(remed_df, use_container_width=True)

                    top1 = remed_df.iloc[0]
                    top2 = remed_df.iloc[1]
                    top3 = remed_df.iloc[2]

                    st.success(f"""
                    **🎯 최적 추천 공법 조합 분석** ({remed_priority})
                    * **🥇 1순위**: **{top1['보강 공법']}** ({top1['평가 점수']}점) - {top1['검토 의견']}
                    * **🥈 2순위**: **{top2['보강 공법']}** ({top2['평가 점수']}점) - {top2['검토 의견']}
                    * **🥉 3순위**: **{top3['보강 공법']}** ({top3['평가 점수']}점) - {top3['검토 의견']}
                    """)
                else:
                    st.success("""
                    **✅ 종합 평가: 전 지층 액상화에 대해 안전함 (만족)**
                    * 모든 검토 시추공 및 지층에서 안전율($FS$)이 1.0을 초과하는 것으로 산정되었습니다.
                    """)

                # --- 대책공법 적용규모·개략공사비 (대안 비교) ---
                rm_zones, rm_alts, rm_total_area, rm_area_how = None, None, None, None
                if len(all_ng_layers) > 0:
                    st.markdown("---")
                    st.subheader("💰 3. 대책공법 적용규모 및 개략공사비 (대안 비교)")
                    _hull = None
                    if coords_df is not None and len(coords_df) >= 3:
                        _, _hull, _ = influence_areas(coords_df, list(judge_df["Borehole"]))
                    rc1, rc2 = st.columns([1, 2])
                    rm_site_area = rc1.number_input(
                        "검토 대상 부지면적 (m²)", min_value=100.0, value=float(round(_hull or 10000.0)), step=100.0,
                        help="좌표가 있으면 시추공 볼록다각형 면적이 기본값이며, 영향면적(티센)을 이 면적으로 비례 환산합니다. "
                        "좌표가 없으면 부지면적을 평가 시추공 수로 균등 배분합니다.",
                    )
                    rc2.caption(
                        "개략공사비 = Σ 영향면적 × 공법별 환산단가(개량심도 반영). 단가는 공법 고유 단위(원/m²·원/m³·원/m)로 "
                        "적용하며, 기둥형 공법은 배치간격, DCM은 개량률로 평면 1m²당 공사비를 환산합니다. 개량심도는 최하단 "
                        "액상화층 하단 심도입니다."
                    )
                    with st.expander("📝 공법별 적용 단가·산정조건 (2026.10 검토 확정단가, 직접 수정 가능)", expanded=False):
                        _base_db = pd.DataFrame(REMED_COST_DB)
                        _base_db.insert(6, "환산단가(원/m³, 심도10m)", [round(remed_equiv_m3(m)) for m in REMED_COST_DB])
                        _cost_edit = st.data_editor(
                            _base_db, use_container_width=True, hide_index=True, key="remed_cost_editor",
                            disabled=["환산단가(원/m³, 심도10m)"],
                            column_config={"단가": st.column_config.NumberColumn("단가", format="%d"),
                                           "단가 근거": st.column_config.TextColumn("단가 근거", width="large")},
                        )
                        st.caption("단가: 설계자 제시 단가를 공개 자료(표준시장단가 토공 항목, 관급 골재 단가, SGR 공사비 내역, "
                                   "FHWA·일본 액상화 대책 단가)와 비교 검토해 확정한 직접공사비입니다. 상세 비교는 "
                                   "'공법별 단가 검토표'를 참고하세요. 간격·개량률을 바꾸면 환산단가가 다시 계산됩니다.")
                        _p1, _p2 = st.columns([2, 1])
                        rm_price_basis = _p1.text_input("단가 기준 (검토서 기재)", value=REMED_PRICE_BASIS_DEFAULT,
                                                        key="price_basis")
                        rm_overhead = _p2.number_input("제경비 포함 배율", 1.00, 2.00, 1.00, 0.05, key="price_overhead",
                                                       help="1.00 = 직접공사비. 간접노무비·산재보험·일반관리비·이윤·부가세 등을 "
                                                            "포함하려면 해당 배율을 입력하세요.")
                    rm_cost_db = [{k: v for k, v in r.items() if k != "환산단가(원/m³, 심도10m)"}
                                  for r in _cost_edit.to_dict("records")]
                    rm_zones, rm_total_area, rm_area_how = improvement_zones(
                        df_spt, df_vs, judge_df, coords_df, rm_site_area)
                    rm_alts = remediation_alternatives(
                        rm_zones, rm_cost_db, urban="도심지" in site_location, narrow="협소" in space_constraint,
                        crust_df=crust_df, overhead=rm_overhead)
                    st.markdown(f"**개량 대상 범위** ({rm_area_how}) — 액상화 발생 {len(rm_zones)}개소, "
                                f"개량면적 {rm_zones['영향면적(m²)'].sum():,.0f} m², 개량체적 {rm_zones['개량체적(m³)'].sum():,.0f} m³")
                    st.dataframe(rm_zones.style.format({"NG 상단(m)": "{:.1f}", "개량심도(m)": "{:.1f}", "개량두께(m)": "{:.1f}",
                                                        "평균FC(%)": "{:.1f}", "영향면적(m²)": "{:,.0f}",
                                                        "개량체적(m³)": "{:,.0f}"}),
                                 use_container_width=True, hide_index=True)
                    st.markdown("**공법별 적용성 및 단일공법 개략공사비**")
                    st.dataframe(rm_alts["single"].style.format({"개량체적(m³)": "{:,.0f}", "단가": "{:,.0f}",
                                                                 "환산단가(원/m³)": "{:,.0f}",
                                                                 "개략공사비(백만원)": "{:,.1f}"}, na_rep="-"),
                                 use_container_width=True, hide_index=True)
                    if rm_alts["alts"]:
                        _alt_df = pd.DataFrame([{k: v for k, v in a.items() if not k.startswith("_")} for a in rm_alts["alts"]])
                        _full_df = _alt_df[~_alt_df["대안"].astype(str).str.startswith("대안 4")]
                        _best = (_full_df if len(_full_df) else _alt_df).loc[
                            (_full_df if len(_full_df) else _alt_df)["개략공사비(백만원)"].idxmin()]
                        _a4row = _alt_df[_alt_df["대안"].astype(str).str.startswith("대안 4")]
                        st.markdown("**대안별 개략공사비 비교**")
                        st.dataframe(_alt_df.style.format({"개량면적(m²)": "{:,.0f}", "개량체적(m³)": "{:,.0f}",
                                                           "개략공사비(백만원)": "{:,.1f}"}),
                                     use_container_width=True, hide_index=True)
                        _fig_a = go.Figure(go.Bar(x=_alt_df["대안"], y=_alt_df["개략공사비(백만원)"],
                                                  text=[f"{v:,.0f}" for v in _alt_df["개략공사비(백만원)"]],
                                                  textposition="outside", marker_color="#1F4E78"))
                        _fig_a.update_layout(template="plotly_white", height=360, yaxis_title="개략공사비 (백만원)",
                                             title="<b>대안별 개략공사비</b>")
                        st.plotly_chart(_fig_a, use_container_width=True)
                        st.success(f"💡 전면 개량 시 공사비 최소 대안: **{_best['대안']}** — {_best['공법 구성']} "
                                   f"(약 {_best['개략공사비(백만원)']:,.0f} 백만원)")
                        if len(_a4row):
                            _a4 = _a4row.iloc[0]
                            st.info(f"🔧 현장 실무형(대안 4): {_a4['공법 구성']} — 약 {_a4['개략공사비(백만원)']:,.0f} 백만원. "
                                    "KDS 4.7(7) 시설물 안정성 평가(기초 형식별 액상화 영향 반영)에서 안정성이 확보될 때 적용할 수 있습니다.")
                        st.caption(f"단가 기준: {rm_price_basis} / 제경비 포함 배율 {rm_overhead:.2f}"
                                   + (" (직접공사비)" if abs(rm_overhead - 1.0) < 1e-9 else ""))
                        if len(_a4row):
                            st.caption("※ 대안 4는 지표 피해 억제를 위한 상부 처리 비용만 포함합니다. 하부 액상화층의 영향은 "
                                       "KDS 4.7(7)에 따른 시설물 안정성 평가와 기초 설계(말뚝 부마찰력·수평저항 저감 등)에 반영해야 합니다.")
                    else:
                        st.warning("현장 조건(심도·세립분·입지)을 만족하는 공법 조합이 없습니다. 단가표의 적용 조건을 확인하세요.")

                # 엑셀 다운로드 / 검토서(Word) 다운로드
                st.markdown("---")
                st.subheader("📑 4. 결과물 다운로드")
                excel_bytes = build_excel_report(
                    df_spt, df_vs, wave_cols, summary_df, df_raw, df_clay_result,
                    site_response_log_df=site_response_log_df,
                    sr_case_table=sr_case_table_out, judge_df=judge_df, csr_method=CSR_METHOD,
                    extra_tables=[
                        ("KDS_지표가속도검토", f"KDS 17 10 00:2024 4.7(6)① 지표면 최대가속도 ≥ S×Fa 검토 "
                                              f"(S={KDS_S:.3f}g, {kds_soil_class}, Fa={KDS_FA:.3f}, S×Fa={KDS_SFA:.3f}g)",
                         kds_pga_df),
                        ("현장실무형검토", "현장 실무형 검토 - 상부 비액상화층(H1) 확보 (Ishihara 1985, KDS 4.7(7))", crust_df),
                        ("공법단가", f"대책공법 적용 단가 및 산정 근거 (직접공사비) - {rm_price_basis}",
                         pd.DataFrame([{**{k: m.get(k) for k in ("공법", "분류", "단가", "단위", "직경(m)", "간격(m)", "개량률",
                                                                  "최대심도(m)", "FC상한(%)")},
                                        "환산단가(원/m³, 심도10m)": round(remed_equiv_m3(m)),
                                        "단가 근거": m.get("단가 근거", "")} for m in rm_cost_db])),
                    ],
                )
                st.download_button(
                    label="📥 전체 평가 결과 엑셀 다운로드",
                    data=excel_bytes,
                    file_name="Liquefaction_Evaluation_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                with st.expander("📝 검토서 작성 정보 (표지·내진설계기준 표에 기재)", expanded=False):
                    ic1, ic2, ic3 = st.columns(3)
                    rp_project = ic1.text_input("과업명", value="○○ 조성공사", key="rp_project")
                    rp_org = ic2.text_input("검토기관", value="", key="rp_org")
                    rp_author = ic3.text_input("작성자", value="", key="rp_author")
                    ic4, ic5, ic6 = st.columns(3)
                    rp_grade = ic4.selectbox("내진등급", ["내진특등급", "내진1등급", "내진2등급"], index=1, key="rp_grade")
                    rp_perf = ic5.selectbox("내진성능수준", ["붕괴방지수준", "기능수행수준"], key="rp_perf")
                    rp_rp, rp_Z, rp_I = kds_rp, KDS_Z, KDS_I
                    ic6.markdown(f"{kds_zone_label}, 재현주기 {kds_rp}<br>S = Z×I = **{KDS_S:.3f}g**, "
                                 f"{kds_soil_class} Fa = {KDS_FA:.3f}", unsafe_allow_html=True)
                    st.caption(
                        "※ 지진구역·재현주기·지반종류는 사이드바 '설계지진 조건'에서 바꿉니다. 이 값은 검토서 기재와 "
                        "KDS 지표 최대가속도 하한(S×Fa) 검토에 쓰이며, 지진응답해석 입력은 업로드한 지진파 그대로입니다."
                    )

                if st.button("📄 액상화 검토서 생성 (Word, 한글에서 열림)", type="primary"):
                    _ng_all = pd.concat(
                        [df_spt_target[df_spt_target["FS"] <= 1.0], df_vs_target[df_vs_target["FS"] <= 1.0]]
                    )
                    _ctx = dict(
                        df_raw=df_raw, df_spt=df_spt, df_vs=df_vs, wave_cols=wave_cols,
                        spt_min_df=spt_min_df, vs_min_df=vs_min_df,
                        site_response_mode=site_response_mode,
                        site_response_log_df=site_response_log_df,
                        wd=wave_dict_all, dtp=dt_per_wave, matching_basis=matching_basis,
                        df_clay_result=df_clay_result,
                        coords_df=coords_df,
                        sr_case_table=sr_case_table_out, sr_applied=sr_applied, sr_reason=sr_reason,
                        sr_case_dfs={k: v for k, v in sr_case_dfs.items()
                                     if sr_case_table_out is None or k in set(sr_case_table_out["케이스"])},
                        sr_rep_bh=sr_rep_bh,
                        rm_zones=rm_zones, rm_alts=rm_alts, rm_total_area=rm_total_area, rm_area_how=rm_area_how,
                        csr_method=CSR_METHOD, vs1_mode=VS1_MODE,
                        kds={"S": KDS_S, "Fa": KDS_FA, "SFA": KDS_SFA, "soil": kds_soil_class, "pga_df": kds_pga_df},
                        crust_df=crust_df, price_basis=rm_price_basis, overhead=rm_overhead, rm_cost_db=rm_cost_db,
                        remed_priority=remed_priority,
                        remed_df=remed_df if len(_ng_all) else None,
                        remed_msg=site_info_msg,
                        site_location=site_location, space_constraint=space_constraint,
                        info={
                            "과업명": rp_project, "검토기관": rp_org, "작성자": rp_author,
                            "내진등급": rp_grade, "내진성능수준": rp_perf, "재현주기": rp_rp,
                            "Z": rp_Z, "I": rp_I,
                        },
                        curves={
                            "sand": DEFAULT_CURVES["sand"], "rock": DEFAULT_CURVES["rock"],
                            "clay_fn": get_clay_curve_for_pi,
                        },
                        msf={"spt_M": 6.5, "vs_case1": MSF_VS_CASE1, "vs_case2": MSF_VS_CASE2},
                    )
                    with st.spinner("검토서 작성 중 (그래프·표 생성)..."):
                        st.session_state["report_bytes"] = build_review_report_docx_v2(_ctx)
                if st.session_state.get("report_bytes"):
                    st.download_button(
                        label="⬇️ 생성된 검토서 다운로드 (.docx)",
                        data=st.session_state["report_bytes"],
                        file_name="액상화_평가_검토서.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                    st.caption("입력 파일이나 조건을 바꾼 경우 '검토서 생성'을 다시 눌러주세요.")

            else:
                st.info("비교할 대상 지층 데이터가 없습니다.")

    except Exception as e:
        st.error(f"데이터 처리 과정에서 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())

elif uploaded_clay_file is not None:
    # 액상화 평가 엑셀 없이 고소성 점토거동 엑셀만 올린 경우 - 점토거동 판별 단독(개별) 평가
    st.subheader("🧱 고소성 점토거동 평가 (단독 평가)")
    st.markdown(CLAY_INTRO_MD)
    try:
        _dfc = run_clay_behavior_eval(pd.read_excel(uploaded_clay_file))
        render_clay_section(_dfc, key="solo")
        st.download_button("📥 점토거동 판별 결과 엑셀", data=build_clay_excel(_dfc),
                           file_name="점토거동_판별결과.xlsx", key="clay_xlsx_solo",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.caption("액상화 평가 엑셀을 함께 올리면 영역 C 지층은 예비평가에서 자동으로 생략되고, "
                   "영역 A·B 지층은 토층명이 세립토여도 본평가 대상에 포함됩니다.")
    except Exception as e:
        st.error(f"점토거동 판별 중 오류가 발생했습니다: {e}")
        st.code(traceback.format_exc())
else:
    st.info(
        "👈 왼쪽 사이드바에서 엑셀 파일(.xlsx)을 업로드하면 액상화 평가가 수행됩니다. "
        "고소성 점토거동 평가 엑셀만 올리면 점토거동 판별을 단독으로 볼 수 있습니다."
    )