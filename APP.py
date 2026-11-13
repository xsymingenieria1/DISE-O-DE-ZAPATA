# -*- coding: utf-8 -*-
"""
DISEÑO DE ZAPATAS ACI 318-19 - XSYM (Versión Web Streamlit)
Tipos: aislada | excentrica | esquinada | combinada
Despliegue: GitHub + Streamlit Community Cloud
"""
import math
import io
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib.figure import Figure
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ═══════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN GLOBAL Y ESTILOS
# ═══════════════════════════════════════════════════════════════════════
BAR_DATA = {
    "Ø 6 mm": (0.60, 0.283), "Ø 8 mm": (0.80, 0.503), "Ø 12 mm": (1.20, 1.131),
    "Ø 1/4\"": (0.635, 0.317), "Ø 3/8\"": (0.953, 0.713), "Ø 1/2\"": (1.270, 1.267),
    "Ø 5/8\"": (1.588, 1.979), "Ø 3/4\"": (1.905, 2.850), "Ø 1\"": (2.540, 5.067),
    "Ø 1 1/4\"": (3.175, 7.917), "Ø 1 1/2\"": (3.810, 11.401),
}

COLORS = {
    'zapata_fill': '#e3f2fd', 'zapata_edge': '#1565c0',
    'columna_fill': '#cfd8dc', 'columna_edge': '#455a64',
    'nucleo_fill': '#fff9c4', 'nucleo_edge': '#fbc02d',
    'presion_fill': '#ffe0b2', 'presion_line': '#e65100',
    'varilla_top': '#d32f2f', 'varilla_bot': '#1976d2',
    'critico_fill': '#ffcdd2', 'punz_fill': '#fff3e0', 'punz_edge': '#ef6c00',
    'banda': '#f8bbd0', 'viga': '#b3e5fc', 'axis': '#424242', 'text': '#212121'
}
REC = 0.075  # Recubrimiento en metros

TIPOS = ['aislada', 'excentrica', 'esquinada', 'combinada']
TIT = {'aislada': 'ZAPATA AISLADA CÉNTRICA', 'excentrica': 'ZAPATA EXCÉNTRICA',
       'esquinada': 'ZAPATA ESQUINADA', 'combinada': 'ZAPATA COMBINADA'}

COMMON = ['fc', 'fy', 'lam', 'qa', 'gamma_s', 'gamma_z', 'gamma_sol', 'Df', 'surcharge', 'h_sol', 'h', 'd', 'rho_min', 'B', 'L']
SIM_COL = ['Pcm', 'Pcv', 'b_col', 't_col']
ENABLED = {
    'aislada': COMMON + SIM_COL,
    'excentrica': COMMON + SIM_COL + ['Mcm', 'Mcv'],
    'esquinada': COMMON + SIM_COL + ['Mcm', 'Mcv', 'b_viga', 'h_viga'],
    'combinada': COMMON + ['Pcm1', 'Pcv1', 'Pcm2', 'Pcv2', 'b1', 't1', 'b2', 't2', 'sep', 'x1'],
}
DEF_AIS = dict(fc=210, fy=4200, lam=1.0, qa=2.125, gamma_s=1.75, gamma_z=2.40, gamma_sol=1.75, Df=2.2, surcharge=0.5, h_sol=0.0, h=80, d=70, B=3.05, L=3.35, b_col=30, t_col=60, Pcm=92.5, Pcv=73.0, rho_min=0.0018)
DEF_EXC = dict(fc=210, fy=4200, lam=1.0, qa=2.0, gamma_s=1.8, gamma_z=2.4, gamma_sol=1.75, Df=1.8, surcharge=0.2, h_sol=0.10, h=55, d=45, B=3.0, L=1.5, b_col=30, t_col=25, Pcm=32, Pcv=12, Mcm=6, Mcv=4, rho_min=0.0024152)
DEF_ESQ = dict(fc=210, fy=4200, lam=1.0, qa=2.0, gamma_s=1.8, gamma_z=2.4, gamma_sol=1.75, Df=1.8, surcharge=0.2, h_sol=0.10, h=65, d=55, B=1.8, L=1.8, b_col=30, t_col=30, Pcm=26, Pcv=8, Mcm=5, Mcv=4, rho_min=0.0024152, b_viga=30, h_viga=60)
DEF_COMB = dict(fc=280, fy=4200, lam=1.0, qa=2.4, gamma_s=1.8, gamma_z=2.4, gamma_sol=2.2, Df=1.8, surcharge=0.0, h_sol=0.0, h=70, d=60, B=2.62, L=4.90, Pcm1=55, Pcv1=45, Pcm2=90, Pcv2=70, b1=45, t1=45, b2=50, t2=50, sep=3.6, x1=0.225, rho_min=0.0027889)
DEF_TIPO = {'aislada': DEF_AIS, 'excentrica': DEF_EXC, 'esquinada': DEF_ESQ, 'combinada': DEF_COMB}
BAR_DEF = {'aislada': "Ø 3/4\"", 'excentrica': "Ø 5/8\"", 'esquinada': "Ø 5/8\"", 'combinada': "Ø 3/4\""}

def _trapz(y, x):
    y = np.asarray(y, dtype=float); x = np.asarray(x, dtype=float)
    if y.size < 2 or x.size < 2: return 0.0
    return float(np.sum(0.5 * (y[1:] + y[:-1]) * np.diff(x)))

# ═══════════════════════════════════════════════════════════════════════
# MOTOR DE CÁLCULO (Lógica original intacta)
# ═══════════════════════════════════════════════════════════════════════
def calc_simple(dt, modo):
    fc, fy, lam = dt['fc'], dt['fy'], dt['lam']
    B, L, h, d = dt['B'], dt['L'], dt['h'], dt['d']
    b_col, t_col = dt['b_col'], dt['t_col']
    Pcm, Pcv = dt['Pcm'], dt['Pcv']
    Mcm, Mcv = dt.get('Mcm', 0.0), dt.get('Mcv', 0.0)
    qa, gs, gz, gsol = dt['qa'], dt['gamma_s'], dt['gamma_z'], dt['gamma_sol']
    Df, SC, h_sol = dt['Df'], dt['surcharge'], dt['h_sol']
    rho_min = dt['rho_min']; db, Ab = BAR_DATA[dt['bar_name']]
    phi_v, phi_f, phi_p = 0.75, 0.90, 0.75
    h_m, d_m = h / 100.0, d / 100.0; b_m, t_m = b_col / 100.0, t_col / 100.0
    esq = (modo == 'esquinada')

    qa_t = qa * 10.0
    q_desc = gs * (Df - h_m - h_sol) + gz * h_m + gsol * h_sol + SC
    qn = qa_t - q_desc
    if qn <= 0: raise ValueError("Capacidad portante neta negativa.")

    Ps, Ms = Pcm + Pcv, Mcm + Mcv
    A_req = Ps / qn; A_prov = B * L; e_serv = abs(Ms) / Ps if Ps else 0.0
    I = B * L ** 3 / 12.0; c = L / 2.0
    qmax_s = Ps / A_prov + abs(Ms) * c / I; qmin_s = Ps / A_prov - abs(Ms) * c / I

    combos = [("1.4D", 1.4 * Pcm, 1.4 * Mcm), ("1.2D+1.6L", 1.2 * Pcm + 1.6 * Pcv, 1.2 * Mcm + 1.6 * Mcv)]
    if dt.get('combo_extra', True): combos.append(("1.4D+1.7L", 1.4 * Pcm + 1.7 * Pcv, 1.4 * Mcm + 1.7 * Mcv))
    nombre_combo, Pu, Mu = max(combos, key=lambda t: abs(t[1]))

    e = abs(Mu) / Pu if Pu else 0.0; L6 = L / 6.0
    if e < 1e-6: caso, caso_desc = "I", "e = 0 → UNIFORME"
    elif e < L6 - 1e-4: caso, caso_desc = "II", "TRAPEZOIDAL"
    elif abs(e - L6) <= 1e-4: caso, caso_desc = "III", "TRIANGULAR"
    else: caso, caso_desc = "IV", "CONTACTO PARCIAL"
    
    a_cont = None
    if caso in ("I", "II", "III"):
        qu_max = Pu / A_prov + abs(Mu) * c / I; qu_min = max(Pu / A_prov - abs(Mu) * c / I, 0.0)
    else:
        a_cont = 3.0 * (L / 2.0 - e); qu_max = 2.0 * Pu / (B * a_cont); qu_min = 0.0
    qu_prom = (qu_max + qu_min) / 2.0

    def q_x(x):
        if caso in ("I", "II", "III"): return qu_max - (qu_max - qu_min) * x / L
        return qu_max * max(0.0, 1.0 - x / a_cont) if a_cont else 0.0

    xs = np.linspace(0, L, 400); qs = np.array([q_x(v) for v in xs])
    col_c = (L - t_m / 2.0) if esq else L / 2.0; col_cB = (B - b_m / 2.0) if esq else B / 2.0
    f_near, f_far = col_c - t_m / 2.0, col_c + t_m / 2.0
    vol_max, vol_min = f_near, L - f_far

    xcr_max = vol_max - d_m
    Vu_max = B * _trapz(qs[xs <= xcr_max], xs[xs <= xcr_max]) if xcr_max > 0 else 0.0
    xcr_min = f_far + d_m
    Vu_min = B * _trapz(qs[xs >= xcr_min], xs[xs >= xcr_min]) if (vol_min - d_m) > 0 else 0.0
    Vu_L, x_gob = (Vu_max, xcr_max) if Vu_max >= Vu_min else (Vu_min, xcr_min)
    qu_crit = q_x(x_gob)
    Vc_L = 0.53 * lam * math.sqrt(fc) * B * 100.0 * d / 1000.0; phiVc_L = phi_v * Vc_L

    fB = col_cB - b_m / 2.0; volB = max(fB, B - (col_cB + b_m / 2.0))
    q_avg_B = qu_prom if not esq else (qu_max + q_x(vol_max)) / 2.0
    Vu_B = q_avg_B * L * max(volB - d_m, 0.0)
    Vc_B = 0.53 * lam * math.sqrt(fc) * L * 100.0 * d / 1000.0; phiVc_B = phi_v * Vc_B
    MB = q_avg_B * volB ** 2 / 2.0

    if esq:
        px, py = (t_col + d / 2.0) / 100.0, (b_col + d / 2.0) / 100.0
        bo = (b_col + d / 2.0) + (t_col + d / 2.0); Ap = px * py
        uu = np.linspace(L - px, L, 60); quu = np.array([q_x(v) for v in uu])
        Vu_p = Pu - py * _trapz(quu, uu); alfa = 30.0
    else:
        px, py = (t_col + d) / 100.0, (b_col + d) / 100.0
        bo = 2.0 * ((b_col + d) + (t_col + d)); Ap = (b_col + d) * (t_col + d) / 1e4
        Vu_p = Pu - qu_prom * Ap * 10.0; alfa = 40.0
    
    lam_s = min(math.sqrt(2.0 / (1.0 + 0.004 * d * 10.0)), 1.0)
    beta = max(b_col, t_col) / min(b_col, t_col)
    Vc1 = 0.53 * (1 + 2 / beta) * lam * math.sqrt(fc) * bo * d / 1000.0
    Vc2 = 0.27 * (2 + alfa * d / bo) * lam * math.sqrt(fc) * bo * d / 1000.0
    Vc3 = 1.06 * lam_s * lam * math.sqrt(fc) * bo * d / 1000.0
    Vc_p = min(Vc1, Vc2, Vc3); phiVc_p = phi_p * Vc_p; punz_ok = Vu_p <= phiVc_p

    q_cara = q_x(vol_max); F1 = q_cara * vol_max; F2 = 0.5 * (qu_max - q_cara) * vol_max
    x1, x2 = vol_max / 2.0, 2.0 * vol_max / 3.0
    ML = F1 * x1 + F2 * x2
    if vol_min > 0.05:
        q_cm = q_x(f_far); F1n = q_cm * vol_min; F2n = 0.5 * (q_cm - qu_min) * vol_min
        ML = max(ML, F1n * vol_min / 2.0 + F2n * 2.0 * vol_min / 3.0)

    rho_max_fac = 0.5 if dt.get('zona_sismica') == "Sísmica" else 0.75
    rho_b = 0.85 * 0.85 * (fc / fy) * (6000.0 / (6000.0 + fy)); rho_max = rho_max_fac * rho_b
    def disenar(Mt):
        Rn = Mt * 1e5 / (phi_f * 100.0 * d ** 2)
        term = max(1.0 - 2.0 * Rn / (0.85 * fc), 0.0)
        rho = min(max((0.85 * fc / fy) * (1.0 - math.sqrt(term)), rho_min), rho_max)
        return Rn, rho, rho * 100.0 * d

    Rn_L, rho_L, As_m_L = disenar(ML); Rn_B, rho_B, As_m_B = disenar(MB)
    As_min_m = rho_min * 100.0 * d
    As_m_L = max(As_m_L, As_min_m); As_m_B = max(As_m_B, As_min_m)
    As_tot_L, As_tot_B = As_m_L * B, As_m_B * L
    n_L = max(2, math.ceil(As_tot_L / Ab)); n_B = max(2, math.ceil(As_tot_B / Ab))
    s_L = (B * 100.0 - 15.0 - n_L * db) / (n_L - 1) if n_L > 1 else 0.0
    s_B = (L * 100.0 - 15.0 - n_B * db) / (n_B - 1) if n_B > 1 else 0.0
    s_max = min(3.0 * h, 45.0); As_prov_L, As_prov_B = n_L * Ab, n_B * Ab

    kg_m = Ab * 0.785; len_L, len_B = L - 2 * REC, B - 2 * REC
    kg_L, kg_B = n_L * len_L * kg_m, n_B * len_B * kg_m; kg_tot = kg_L + kg_B
    ld_cm = max(0.06 * db * fy / math.sqrt(fc), 0.006 * db * fy, 30.0)
    ld_disp = max(vol_max, volB) * 100.0 - 7.5

    if esq:
        Mu_tb = Pu * max(L / 2.0 - t_m / 2.0, B / 2.0 - b_m / 2.0)
        b_v, h_v = dt.get('b_viga', 30), dt.get('h_viga', 60); d_v = h_v - 6.0
        Rn_v = Mu_tb * 1e5 / (phi_f * b_v * d_v ** 2)
        term_v = max(1.0 - 2.0 * Rn_v / (0.85 * fc), 0.0)
        rho_v = (0.85 * fc / fy) * (1.0 - math.sqrt(term_v)); viga_ok = rho_v <= rho_max
        As_v = max(rho_v, rho_min) * b_v * d_v
        db_v, Ab_v = BAR_DATA[dt.get('bar_viga', "Ø 3/4\"")]; n_v = max(2, math.ceil(As_v / Ab_v)); As_prov_v = n_v * Ab_v
    else:
        Mu_tb=b_v=h_v=d_v=Rn_v=rho_v=As_v=n_v=As_prov_v=0.0; viga_ok=True; db_v, Ab_v = db, Ab

    checks = [
        ("Az ≥ A req", A_prov >= A_req, f"{A_prov:.2f} ≥ {A_req:.2f}"),
        ("qmax serv ≤ qa", qmax_s <= qa_t, f"{qmax_s:.2f} ≤ {qa_t:.2f}"),
        ("qmin serv ≥ 0", qmin_s >= -1e-6, f"{qmin_s:.2f} ≥ 0"),
        ("Cortante L Vu≤φVc", Vu_L <= phiVc_L, f"{Vu_L:.2f} ≤ {phiVc_L:.2f}"),
        ("Cortante B Vu≤φVc", Vu_B <= phiVc_B, f"{Vu_B:.2f} ≤ {phiVc_B:.2f}"),
        ("Punzonamiento Vu≤φVc", punz_ok, f"{Vu_p:.2f} ≤ {phiVc_p:.2f}"),
        ("Acero L ≥ req", As_prov_L >= As_tot_L, f"{As_prov_L:.2f} ≥ {As_tot_L:.2f}"),
        ("Acero B ≥ req", As_prov_B >= As_tot_B, f"{As_prov_B:.2f} ≥ {As_tot_B:.2f}"),
        ("S ≤ s_max", max(s_L, s_B) <= s_max, f"{max(s_L, s_B):.1f} ≤ {s_max:.0f}"),
        ("Ld ≤ disp", ld_cm <= ld_disp, f"{ld_cm:.1f} ≤ {ld_disp:.1f}"),
    ]
    if esq: checks.append(("Viga conex. ρ≤ρmax", viga_ok, f"ρ={rho_v:.5f} ≤ {rho_max:.5f}"))

    return dict(tipo=modo, qa_t=qa_t, q_desc=q_desc, qn=qn, Ps=Ps, Ms=Ms, A_req=A_req, A_prov=A_prov,
                e_serv=e_serv, I=I, c=c, qmax_s=qmax_s, qmin_s=qmin_s, combos=combos, nombre_combo=nombre_combo,
                Pu=Pu, Mu=Mu, e=e, L6=L6, caso=caso, caso_desc=caso_desc, a_cont=a_cont,
                qu_max=qu_max, qu_min=qu_min, qu_prom=qu_prom, q_x=q_x, col_c=col_c, col_cB=col_cB,
                vol_max=vol_max, vol_min=vol_min, x_gob=x_gob, Vu_L=Vu_L, qu_crit=qu_crit,
                Vc_L=Vc_L, phiVc_L=phiVc_L, volB=volB, q_avg_B=q_avg_B, Vu_B=Vu_B, Vc_B=Vc_B, phiVc_B=phiVc_B,
                bo=bo, Ap=Ap, Vu_p=Vu_p, Vc_p=Vc_p, phiVc_p=phiVc_p, punz_ok=punz_ok,
                q_cara=q_cara, F1=F1, F2=F2, ML=ML, MB=MB, As_min_m=As_min_m,
                As_m_L=As_m_L, As_m_B=As_m_B, As_tot_L=As_tot_L, As_tot_B=As_tot_B, n_L=n_L, n_B=n_B,
                s_L=s_L, s_B=s_B, s_max=s_max, rho_b=rho_b, rho_max=rho_max, As_prov_L=As_prov_L,
                As_prov_B=As_prov_B, kg_m=kg_m, len_L=len_L, len_B=len_B, kg_L=kg_L, kg_B=kg_B,
                kg_tot=kg_tot, ld_cm=ld_cm, ld_disp=ld_disp, Mu_tb=Mu_tb, b_v=b_v, h_v=h_v, d_v=d_v,
                rho_v=rho_v, viga_ok=viga_ok, As_v=As_v, n_v=n_v, As_prov_v=As_prov_v,
                db_v=db_v, Ab_v=Ab_v, checks=checks, db=db, Ab=Ab)

def calc_combinada(dt):
    fc, fy, lam = dt['fc'], dt['fy'], dt['lam']
    B, L, h, d = dt['B'], dt['L'], dt['h'], dt['d']
    b1c, t1c, b2c, t2c = dt['b1'], dt['t1'], dt['b2'], dt['t2']
    sep, x1c = dt['sep'], dt['x1']
    Pcm1, Pcv1, Pcm2, Pcv2 = dt['Pcm1'], dt['Pcv1'], dt['Pcm2'], dt['Pcv2']
    qa, gs, gz, gsol = dt['qa'], dt['gamma_s'], dt['gamma_z'], dt['gamma_sol']
    Df, SC, h_sol = dt['Df'], dt['surcharge'], dt['h_sol']
    rho_min = dt['rho_min']; db, Ab = BAR_DATA[dt['bar_name']]
    phi_v, phi_f, phi_p = 0.75, 0.90, 0.75
    h_m, d_m = h / 100.0, d / 100.0
    b1, t1 = b1c / 100.0, t1c / 100.0; b2, t2 = b2c / 100.0, t2c / 100.0
    x2c = x1c + sep

    qa_t = qa * 10.0
    q_desc = gs * (Df - h_m - h_sol) + gz * h_m + gsol * h_sol + SC
    qn = qa_t - q_desc
    if qn <= 0: raise ValueError("Capacidad portante neta negativa.")
    Ps1, Ps2 = Pcm1 + Pcv1, Pcm2 + Pcv2; Ps = Ps1 + Ps2
    Xo = Ps2 * sep / Ps; A_req = Ps / qn
    L_auto = 2.0 * (x1c + Xo); B_auto = A_req / L_auto
    A_prov = B * L; e_res = (x1c + Xo) - L / 2.0
    q_serv = Ps / A_prov
    qmax_s = q_serv * (1 + 6 * abs(e_res) / L); qmin_s = q_serv * (1 - 6 * abs(e_res) / L)

    c1 = [("1.4D", 1.4 * (Pcm1 + Pcm2)), ("1.2D+1.6L", 1.2 * (Pcm1 + Pcm2) + 1.6 * (Pcv1 + Pcv2))]
    if dt.get('combo_extra', True): c1.append(("1.4D+1.7L", 1.4 * (Pcm1 + Pcm2) + 1.7 * (Pcv1 + Pcv2)))
    nombre_combo, Pu = max(c1, key=lambda t: t[1])
    if nombre_combo.startswith("1.4D+1.7"): Pu1, Pu2 = 1.4 * Pcm1 + 1.7 * Pcv1, 1.4 * Pcm2 + 1.7 * Pcv2
    elif nombre_combo.startswith("1.2"): Pu1, Pu2 = 1.2 * Pcm1 + 1.6 * Pcv1, 1.2 * Pcm2 + 1.6 * Pcv2
    else: Pu1, Pu2 = 1.4 * Pcm1, 1.4 * Pcm2
    
    qup = Pu / A_prov; w = Pu / L
    def V_f(x):
        v = w * x
        if x > x1c: v -= Pu1
        if x > x2c: v -= Pu2
        return v
    def M_f(x):
        m = w * x * x / 2.0
        if x > x1c: m -= Pu1 * (x - x1c)
        if x > x2c: m -= Pu2 * (x - x2c)
        return m

    x0 = Pu1 / w; M_neg = M_f(x0)
    M_pos_c2 = max(M_f(x2c), w * (L - x2c) ** 2 / 2.0)
    xa, xb = x1c + t1 / 2.0 + d_m, x2c - t2 / 2.0 - d_m
    Vu_a, Vu_b = abs(V_f(xa)), abs(V_f(xb))
    Vu_cort, x_gob = (Vu_a, xa) if Vu_a >= Vu_b else (Vu_b, xb)
    Vc_c = 0.53 * lam * math.sqrt(fc) * B * 100.0 * d / 1000.0; phiVc_c = phi_v * Vc_c

    bo_ext = (b1c + d) + 2.0 * (t1c / 2.0 + d / 2.0)
    Ap_ext = (t1c / 2.0 + d / 2.0 + t1c / 2.0) * (b1c + d) / 1e4
    Vu_pe = Pu1 - qup * Ap_ext
    bo_int = 2.0 * ((b2c + d) + (t2c + d))
    Ap_int = (b2c + d) * (t2c + d) / 1e4
    Vu_pi = Pu2 - qup * Ap_int
    lam_s = min(math.sqrt(2.0 / (1.0 + 0.004 * d * 10.0)), 1.0)
    def trio(bo, alfa, bet):
        v1 = 0.53 * (1 + 2 / bet) * lam * math.sqrt(fc) * bo * d / 1000.0
        v2 = 0.27 * (2 + alfa * d / bo) * lam * math.sqrt(fc) * bo * d / 1000.0
        v3 = 1.06 * lam_s * lam * math.sqrt(fc) * bo * d / 1000.0
        return v1, v2, v3, min(v1, v2, v3)
    Vc1e, Vc2e, Vc3e, Vc_e = trio(bo_ext, 30.0, max(b1c, t1c) / min(b1c, t1c))
    Vc1i, Vc2i, Vc3i, Vc_i = trio(bo_int, 40.0, max(b2c, t2c) / min(b2c, t2c))
    phiVc_pe, phiVc_pi = phi_p * Vc_e, phi_p * Vc_i

    rho_max_fac = 0.5 if dt.get('zona_sismica') == "Sísmica" else 0.75
    rho_b = 0.85 * 0.85 * (fc / fy) * (6000.0 / (6000.0 + fy)); rho_max = rho_max_fac * rho_b
    b_dis = B * 100.0
    def disenar(Mt, bcm):
        Rn = Mt * 1e5 / (phi_f * bcm * d ** 2)
        term = max(1.0 - 2.0 * Rn / (0.85 * fc), 0.0)
        rho = min(max((0.85 * fc / fy) * (1.0 - math.sqrt(term)), rho_min), rho_max)
        return Rn, rho, rho * bcm * d
    Asmin_B = rho_min * b_dis * d
    Rn_t, rho_t, As_t = disenar(abs(M_neg), b_dis); As_t = max(As_t, Asmin_B)
    Rn_b, rho_b2, As_b = disenar(M_pos_c2, b_dis); As_b = max(As_b, Asmin_B)
    n_top = max(2, math.ceil(As_t / Ab)); n_bot = max(2, math.ceil(As_b / Ab))
    s_top = (b_dis - 15.0 - n_top * db) / (n_top - 1) if n_top > 1 else 0.0
    s_bot = (b_dis - 15.0 - n_bot * db) / (n_bot - 1) if n_bot > 1 else 0.0

    trans = []
    for Pu_i, b_ci, t_ci, tag in [(Pu1, b1, t1, "Col. ext"), (Pu2, b2, t2, "Col. int")]:
        m_i = (B - b_ci) / 2.0; w_i = Pu_i / B; Mu_i = w_i * m_i ** 2 / 2.0
        Rn_i, rho_i, As_i = disenar(Mu_i, 100.0); As_i = max(As_i, rho_min * 100.0 * d)
        banda = b_ci * 100.0 + d; As_banda = As_i * banda / 100.0
        n_b = max(2, math.ceil(As_banda / Ab))
        s_b = (banda - 15.0 - n_b * db) / (n_b - 1) if n_b > 1 else 0.0
        trans.append(dict(tag=tag, m=m_i, w=w_i, Mu=Mu_i, Rn=Rn_i, rho=rho_i, As_m=As_i, banda=banda, As_banda=As_banda, n=n_b, s=s_b))

    As_mont = 0.0018 * 100.0 * h; s_mont = 100.0 * Ab / As_mont
    n_mont = max(2, math.ceil(B * 100.0 / s_mont))
    ld_cm = max(0.06 * db * fy / math.sqrt(fc), 0.006 * db * fy, 30.0)
    ld_disp = min(L - x2c - t2 / 2.0, x2c - t2 / 2.0) * 100.0 - 7.5
    kg_m = Ab * 0.785; len_L, len_B = L - 2 * REC, B - 2 * REC
    grupos = [("Sup. tramo", n_top, len_L), ("Inf. apoyo", n_bot, len_L), ("Transv. c1", trans[0]['n'], len_B), ("Transv. c2", trans[1]['n'], len_B), ("Montaje", n_mont, len_L)]
    metrado = [(nom, nn, ll, nn*ll, nn*ll*kg_m) for nom, nn, ll in grupos]
    kg_tot = sum(m[4] for m in metrado)
    s_max = min(3.0 * h, 45.0)
    checks = [
        ("Az ≥ A req", A_prov >= A_req, f"{A_prov:.2f} ≥ {A_req:.2f}"),
        ("qmax serv ≤ qa", qmax_s <= qa_t, f"{qmax_s:.2f} ≤ {qa_t:.2f}"),
        ("qmin serv ≥ 0", qmin_s >= -1e-6, f"{qmin_s:.2f} ≥ 0"),
        ("Cortante Vu≤φVc", Vu_cort <= phiVc_c, f"{Vu_cort:.2f} ≤ {phiVc_c:.2f}"),
        ("Punz. ext Vu≤φVc", Vu_pe <= phiVc_pe, f"{Vu_pe:.2f} ≤ {phiVc_pe:.2f}"),
        ("Punz. int Vu≤φVc", Vu_pi <= phiVc_pi, f"{Vu_pi:.2f} ≤ {phiVc_pi:.2f}"),
        ("Acero sup. ≥ req", n_top * Ab >= As_t, f"{n_top * Ab:.2f} ≥ {As_t:.2f}"),
        ("Acero inf. ≥ req", n_bot * Ab >= As_b, f"{n_bot * Ab:.2f} ≥ {As_b:.2f}"),
        ("Franja c1 ≥ req", trans[0]['n'] * Ab >= trans[0]['As_banda'], f"{trans[0]['n'] * Ab:.2f} ≥ {trans[0]['As_banda']:.2f}"),
        ("Franja c2 ≥ req", trans[1]['n'] * Ab >= trans[1]['As_banda'], f"{trans[1]['n'] * Ab:.2f} ≥ {trans[1]['As_banda']:.2f}"),
        ("S ≤ s_max", max(s_top, s_bot) <= s_max, f"{max(s_top, s_bot):.1f} ≤ {s_max:.0f}"),
        ("Ld ≤ disp", ld_cm <= ld_disp, f"{ld_cm:.1f} ≤ {ld_disp:.1f}"),
    ]
    return dict(tipo='combinada', qa_t=qa_t, q_desc=q_desc, qn=qn, Ps1=Ps1, Ps2=Ps2, Ps=Ps, Xo=Xo,
                A_req=A_req, L_auto=L_auto, B_auto=B_auto, A_prov=A_prov, e_res=e_res, q_serv=q_serv,
                qmax_s=qmax_s, qmin_s=qmin_s, combos=c1, nombre_combo=nombre_combo, Pu=Pu, Pu1=Pu1, Pu2=Pu2,
                qup=qup, w=w, V_f=V_f, M_f=M_f, x0=x0, x1c=x1c, x2c=x2c, M_neg=M_neg, M_pos_c2=M_pos_c2,
                xa=xa, xb=xb, Vu_a=Vu_a, Vu_b=Vu_b, Vu_cort=Vu_cort, x_gob=x_gob, Vc_c=Vc_c, phiVc_c=phiVc_c,
                bo_ext=bo_ext, Ap_ext=Ap_ext, Vu_pe=Vu_pe, Vc1e=Vc1e, Vc2e=Vc2e, Vc3e=Vc3e, Vc_e=Vc_e,
                phiVc_pe=phiVc_pe, bo_int=bo_int, Ap_int=Ap_int, Vu_pi=Vu_pi, Vc1i=Vc1i, Vc2i=Vc2i,
                Vc3i=Vc3i, Vc_i=Vc_i, phiVc_pi=phiVc_pi, rho_b=rho_b, rho_max=rho_max, Asmin_B=Asmin_B,
                Rn_t=Rn_t, rho_t=rho_t, As_t=As_t, n_top=n_top, s_top=s_top, Rn_b=Rn_b, rho_b2=rho_b2,
                As_b=As_b, n_bot=n_bot, s_bot=s_bot, trans=trans, As_mont=As_mont, s_mont=s_mont,
                n_mont=n_mont, ld_cm=ld_cm, ld_disp=ld_disp, kg_m=kg_m, metrado=metrado, kg_tot=kg_tot,
                s_max=s_max, checks=checks, db=db, Ab=Ab)

# ═══════════════════════════════════════════════════════════════════════
# GENERADORES DE GRÁFICOS (Retornan objetos Figure para Streamlit)
# ═══════════════════════════════════════════════════════════════════════
def plot_plano(r, d):
    fig = Figure(figsize=(8, 6), dpi=100)
    ax = fig.add_subplot(111)
    B, L = d['B'], d['L']; dm = d['d'] / 100
    ax.add_patch(patches.Rectangle((0, 0), L, B, fc=COLORS['zapata_fill'], ec=COLORS['zapata_edge'], lw=2.5, hatch='xxx'))
    
    if r['tipo'] != 'combinada':
        b_m, t_m = d['b_col'] / 100, d['t_col'] / 100
        xc, yc = r['col_c'], r['col_cB']
        ax.add_patch(patches.Rectangle((xc - t_m / 2, yc - b_m / 2), t_m, b_m, fc=COLORS['columna_fill'], ec=COLORS['columna_edge'], lw=2, hatch='///', zorder=3))
        nw = L / 3.0
        ax.add_patch(patches.Rectangle(((L - nw) / 2, (B - nw) / 2), nw, nw, fc=COLORS['nucleo_fill'], ec=COLORS['nucleo_edge'], ls='--', lw=1.5, alpha=0.6, zorder=2))
        if r['tipo'] == 'esquinada':
            px, py = t_m + dm / 2, b_m + dm / 2
            poly = [(L, B), (L, B - py), (L - px, B - py), (L - px, B), (L, B)]
            ax.add_patch(patches.Polygon(poly, fc=COLORS['punz_fill'], ec=COLORS['punz_edge'], ls='--', lw=2, alpha=0.7, zorder=2))
        else:
            ax.add_patch(patches.Rectangle((xc - (t_m + dm) / 2, yc - (b_m + dm) / 2), t_m + dm, b_m + dm, fc='none', ec=COLORS['punz_edge'], ls='--', lw=2, alpha=0.8, zorder=2))
    else:
        x1, x2 = r['x1c'], r['x2c']
        b1, t1 = d['b1'] / 100, d['t1'] / 100; b2, t2 = d['b2'] / 100, d['t2'] / 100
        ax.add_patch(patches.Rectangle((x1 - t1 / 2, (B - b1) / 2), t1, b1, fc=COLORS['columna_fill'], ec=COLORS['columna_edge'], lw=2, hatch='///', zorder=3))
        ax.add_patch(patches.Rectangle((x2 - t2 / 2, (B - b2) / 2), t2, b2, fc=COLORS['columna_fill'], ec=COLORS['columna_edge'], lw=2, hatch='///', zorder=3))
        px1 = t1 + dm / 2; py1 = b1 + dm
        poly1 = [(0, (B-py1)/2), (px1, (B-py1)/2), (px1, (B+py1)/2), (0, (B+py1)/2)]
        ax.add_patch(patches.Polygon(poly1, fc=COLORS['punz_fill'], ec=COLORS['punz_edge'], ls='--', lw=2, alpha=0.7, zorder=2))
        ax.add_patch(patches.Rectangle((x2 - (t2/2 + dm/2), (B - (b2 + dm))/2), t2 + dm, b2 + dm, fc='none', ec=COLORS['punz_edge'], ls='--', lw=2, alpha=0.8, zorder=2))

    ax.set_title(f"PLANTA ESTRUCTURAL - {TIT[r['tipo']]}", fontsize=12, fontweight='bold', pad=10)
    ax.set_aspect('equal'); ax.set_xlim(-0.5, L + 0.5); ax.set_ylim(-0.5, B + 0.5); ax.axis('off')
    fig.tight_layout()
    return fig

def plot_pres(r, d):
    fig = Figure(figsize=(8, 5), dpi=100)
    L = d['L']
    if r['tipo'] != 'combinada':
        ax = fig.add_subplot(111)
        xs = np.linspace(0, L, 100)
        qs = np.array([r['q_x'](x) for x in xs])
        ax.fill_between(xs, 0, qs, color=COLORS['presion_fill'], alpha=0.6, step='mid')
        ax.plot(xs, qs, color=COLORS['presion_line'], lw=3, label='Presión de Diseño')
        ax.axhline(0, color='black', lw=1)
        ax.text(0.05*L, qs[0]*0.9, f"qu_max={r['qu_max']:.2f}", color='red', fontweight='bold')
        ax.text(0.95*L, qs[-1]*0.9, f"qu_min={r['qu_min']:.2f}", color='blue', fontweight='bold')
        ax.set_xlabel("Distancia (m)"); ax.set_ylabel("Presión (t/m²)")
        ax.grid(True, linestyle=':', alpha=0.6)
        ax.set_title(f"DIAGRAMA DE PRESIONES - CASO {r['caso']}", fontsize=11, fontweight='bold')
    else:
        xs = np.linspace(0, L, 300)
        Vs = np.array([r['V_f'](x) for x in xs]); Ms = np.array([r['M_f'](x) for x in xs])
        ax1 = fig.add_subplot(211)
        ax1.plot(xs, Vs, 'b-', lw=2.5)
        ax1.fill_between(xs, 0, Vs, where=Vs>0, color='#bbdefb', alpha=0.5)
        ax1.fill_between(xs, 0, Vs, where=Vs<0, color='#ffcdd2', alpha=0.5)
        ax1.axhline(0, color='black', lw=1); ax1.axvline(r['x_gob'], color='red', ls='--', lw=2)
        ax1.set_ylabel("Cortante V (ton)"); ax1.grid(True, alpha=0.3)
        
        ax2 = fig.add_subplot(212)
        ax2.plot(xs, Ms, 'g-', lw=2.5)
        ax2.fill_between(xs, 0, Ms, where=Ms>0, color='#c8e6c9', alpha=0.5)
        ax2.fill_between(xs, 0, Ms, where=Ms<0, color='#e1bee7', alpha=0.5)
        ax2.axhline(0, color='black', lw=1)
        ax2.set_xlabel("Longitud (m)"); ax2.set_ylabel("Momento M (t-m)"); ax2.grid(True, alpha=0.3)
    fig.tight_layout()
    return fig

def generate_excel(r, d):
    wb = Workbook()
    ws = wb.active
    ws.title = "Memoria de Cálculo"
    ws['A1'] = f"MEMORIA DE CÁLCULO - {TIT[r['tipo']]} (ACI 318-19)"
    ws['A1'].font = Font(bold=True, size=14)
    
    row = 3
    ws[f'A{row}'] = "PARÁMETROS DE DISEÑO"; ws[f'A{row}'].font = Font(bold=True)
    row += 1
    for k, v in d.items():
        ws[f'A{row}'] = str(k); ws[f'B{row}'] = str(v); row += 1
        
    row += 2
    ws[f'A{row}'] = "VERIFICACIONES"; ws[f'A{row}'].font = Font(bold=True)
    row += 1
    for check_name, ok, detail in r['checks']:
        status = "CUMPLE" if ok else "NO CUMPLE"
        ws[f'A{row}'] = check_name; ws[f'B{row}'] = detail; ws[f'C{row}'] = status
        ws[f'C{row}'].font = Font(color="008000" if ok else "FF0000", bold=True)
        row += 1
        
    row += 2
    ws[f'A{row}'] = "METRADO DE ACERO"; ws[f'A{row}'].font = Font(bold=True)
    row += 1
    if r['tipo'] != 'combinada':
        ws[f'A{row}'] = f"Dir. L: {r['n_L']} var x {r['len_L']:.2f} m = {r['n_L']*r['len_L']:.1f} m ({r['kg_L']:.1f} kg)"
        row += 1
        ws[f'A{row}'] = f"Dir. B: {r['n_B']} var x {r['len_B']:.2f} m = {r['n_B']*r['len_B']:.1f} m ({r['kg_B']:.1f} kg)"
        row += 1
    else:
        for m in r['metrado']:
            ws[f'A{row}'] = f"{m[0]}: {m[1]} var x {m[2]:.2f} m = {m[3]:.1f} m ({m[4]:.1f} kg)"
            row += 1
    row += 1
    ws[f'A{row}'] = f"PESO TOTAL DE ACERO: {r['kg_tot']:.1f} kg"
    ws[f'A{row}'].font = Font(bold=True)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

# ═══════════════════════════════════════════════════════════════════════
# INTERFAZ STREAMLIT
# ═══════════════════════════════════════════════════════════════════════
st.set_page_config(page_title="Diseño de Zapatas ACI 318-19", layout="wide", page_icon="🏗️")
st.title("🏗️ DISEÑO DE ZAPATAS ACI 318-19 - XSYM")
st.markdown("Herramienta profesional para el diseño de zapatas aisladas, excéntricas, esquinadas y combinadas.")

with st.sidebar:
    st.header("⚙️ Configuración")
    tipo = st.selectbox("Tipo de Zapata", TIPOS, index=0)
    
    st.subheader("Materiales / Suelo")
    fc = st.number_input("f'c (kg/cm²)", value=float(DEF_TIPO[tipo]['fc']))
    fy = st.number_input("fy (kg/cm²)", value=float(DEF_TIPO[tipo]['fy']))
    qa = st.number_input("qa (kg/cm²)", value=float(DEF_TIPO[tipo]['qa']))
    gamma_s = st.number_input("γ suelo (t/m³)", value=float(DEF_TIPO[tipo]['gamma_s']))
    gamma_z = st.number_input("γ zapata (t/m³)", value=float(DEF_TIPO[tipo]['gamma_z']))
    Df = st.number_input("Df (m)", value=float(DEF_TIPO[tipo]['Df']))
    
    st.subheader("Cargas")
    Pcm = st.number_input("Pcm (t)", value=float(DEF_TIPO[tipo]['Pcm']))
    Pcv = st.number_input("Pcv (t)", value=float(DEF_TIPO[tipo]['Pcv']))
    Mcm = st.number_input("Mcm (t-m)", value=float(DEF_TIPO[tipo].get('Mcm', 0.0))) if tipo in ['excentrica', 'esquinada'] else 0.0
    Mcv = st.number_input("Mcv (t-m)", value=float(DEF_TIPO[tipo].get('Mcv', 0.0))) if tipo in ['excentrica', 'esquinada'] else 0.0
    
    st.subheader("Geometría")
    B = st.number_input("Ancho B (m)", value=float(DEF_TIPO[tipo]['B']))
    L = st.number_input("Largo L (m)", value=float(DEF_TIPO[tipo]['L']))
    h = st.number_input("Altura h (cm)", value=float(DEF_TIPO[tipo]['h']))
    d = st.number_input("Peralte d (cm)", value=float(DEF_TIPO[tipo]['d']))
    b_col = st.number_input("Col. b (cm)", value=float(DEF_TIPO[tipo]['b_col']))
    t_col = st.number_input("Col. t (cm)", value=float(DEF_TIPO[tipo]['t_col']))
    
    st.subheader("Refuerzo")
    bar_name = st.selectbox("Varilla zapata", list(BAR_DATA.keys()), index=list(BAR_DATA.keys()).index(BAR_DEF[tipo]))
    zona_sismica = st.selectbox("Zona sísmica", ["No sísmica", "Sísmica"])
    combo_extra = st.checkbox("Incluir combo 1.4D+1.7L", value=True)

    calcular = st.button("▶ CALCULAR DISEÑO", type="primary", use_container_width=True)

# Estado inicial
if 'resultado' not in st.session_state:
    st.session_state.resultado = None

if calcular:
    try:
        dt = {
            'fc': fc, 'fy': fy, 'lam': 1.0, 'qa': qa, 'gamma_s': gamma_s, 'gamma_z': gamma_z,
            'gamma_sol': 1.75, 'Df': Df, 'surcharge': 0.2, 'h_sol': 0.10, 'h': h, 'd': d,
            'rho_min': 0.0018, 'B': B, 'L': L, 'Pcm': Pcm, 'Pcv': Pcv, 'Mcm': Mcm, 'Mcv': Mcv,
            'b_col': b_col, 't_col': t_col, 'bar_name': bar_name, 'zona_sismica': zona_sismica,
            'combo_extra': combo_extra
        }
        # Añadir campos específicos si aplican
        if tipo == 'combinada':
            dt.update({'Pcm1': Pcm, 'Pcv1': Pcv, 'Pcm2': Pcm*1.5, 'Pcv2': Pcv*1.5, 'b1': b_col, 't1': t_col, 'b2': b_col, 't2': t_col, 'sep': L/2, 'x1': 0.5})
            res = calc_combinada(dt)
        elif tipo == 'esquinada':
            dt.update({'b_viga': 30, 'h_viga': 60, 'bar_viga': "Ø 3/4\""})
            res = calc_simple(dt, 'esquinada'); res['tipo'] = 'esquinada'
        elif tipo == 'excentrica':
            res = calc_simple(dt, 'excentrica'); res['tipo'] = 'excentrica'
        else:
            dt['Mcm'] = 0.0; dt['Mcv'] = 0.0
            res = calc_simple(dt, 'aislada'); res['tipo'] = 'aislada'; res['caso'] = "I"; res['caso_desc'] = "e = 0 → UNIFORME"
        
        st.session_state.resultado = res
        st.session_state.datos = dt
        st.success("✅ Cálculo completado exitosamente.")
    except Exception as e:
        st.error(f"❌ Error en el cálculo: {str(e)}")

if st.session_state.resultado:
    r = st.session_state.resultado
    d = st.session_state.datos
    
    # Resumen rápido
    st.subheader("📊 Resumen de Verificaciones")
    cols = st.columns(2)
    with cols[0]:
        st.info(f"**Tipo:** {TIT[r['tipo']]} | **Comb:** {r['nombre_combo']} | **Pu:** {r['Pu']:.1f} t")
    with cols[1]:
        checks_html = ""
        for name, ok, detail in r['checks']:
            icon = "✅" if ok else "❌"
            checks_html += f"{icon} {name}: {detail}<br>"
        st.markdown(checks_html, unsafe_allow_html=True)

    # Pestañas de gráficos
    tab1, tab2, tab3, tab4 = st.tabs(["📐 Planta", "📈 Presiones / Diagramas", "✂️ Cortante", "🏋️ Punzonamiento"])
    
    with tab1:
        fig_plano = plot_plano(r, d)
        st.pyplot(fig_plano)
        
    with tab2:
        fig_pres = plot_pres(r, d)
        st.pyplot(fig_pres)
        
    with tab3:
        st.markdown(f"**Sección crítica:** x = {r.get('x_gob', 0):.2f} m | **Vu:** {r.get('Vu_L', r.get('Vu_cort', 0)):.2f} t ≤ **φVc:** {r.get('phiVc_L', r.get('phiVc_c', 0)):.2f} t")
        
    with tab4:
        if r['tipo'] != 'combinada':
            st.markdown(f"**Vu:** {r['Vu_p']:.2f} t ≤ **φVc:** {r['phiVc_p']:.2f} t → {'✅ CUMPLE' if r['punz_ok'] else '❌ NO CUMPLE'}")
        else:
            st.markdown(f"**Ext:** Vu={r['Vu_pe']:.2f} ≤ φVc={r['phiVc_pe']:.2f} | **Int:** Vu={r['Vu_pi']:.2f} ≤ φVc={r['phiVc_pi']:.2f}")

    # Exportación
    st.divider()
    st.subheader("💾 Exportar Resultados")
    excel_buffer = generate_excel(r, d)
    st.download_button(
        label="📥 Descargar Memoria de Cálculo en Excel",
        data=excel_buffer,
        file_name=f"Memoria_Zapata_{tipo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )