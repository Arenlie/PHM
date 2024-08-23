from math import floor, ceil
import numpy as np
from numpy import real, sqrt
from scipy import fftpack
from scipy.signal import hilbert, argrelextrema


# fft变换
def fft_spectrum(signal: np.ndarray, fs: float):
    """
    傅里叶变换频谱
    :param signal:np.ndarray 时序振动信号
    :param fs:float 采样频率
    :return:
    f:np.ndarray, 频率
    y:np.ndarray, 幅值
    """
    n = len(signal)
    f = fs * np.arange(n // 2) / n
    fft_data = abs(np.fft.fft(signal))[:int(n / 2)] * 2 / n  # 归一化处理，双边频谱反转后需要×2
    return f, fft_data


# 频域滤波
def fft_filter(signal, lpf1, lpf2, fs):
    """
    signal: np.ndarray 原始振动信号
    fs: float 采样频率
    lpf1: float 滤波频率-低频
    lpf2: float 滤波频率-高频
    :return
    y: np.ndarray 滤波信号
    """
    yy = fftpack.fft(signal)
    m = len(yy)
    k = m / fs
    for i in range(0, floor(k * lpf1)):
        yy[i] = 0
    for i in range(ceil(k * lpf2 - 1), m):
        yy[i] = 0
    y = 2 * real(fftpack.ifft(yy))
    return y


# 希尔伯特包络
def hilbert_envelop(signal):
    """
    输入原始信号，输出包络时域信号
    """
    hx = hilbert(signal)  # 对信号进行希尔伯特变换
    analytic_signal = signal - hx * 1j  # 进行hilbert变换后的解析信号
    return np.abs(analytic_signal)


# 时域积分
def acc2dis(data: np.ndarray, fs: float):
    """
    采用时域积分的方式，将振动加速度信号转化为速度信号和位移信号
    Parameters
    ----------
    data: np.ndarray, 振动加速度信号
    fs: float, 采样频率
    Return
    ------
    s_ifft: np.ndarray, 积分速度信号
    d_ifft：np.ndarray, 积分位移信号
    """
    n = len(data)
    a_mul_dt = data / fs

    s = []
    total = a_mul_dt[0]
    for i in range(n - 1):
        total = total + a_mul_dt[i + 1]
        s.append(total)
    s_fft = np.fft.fft(s)
    s_fft[:2] = 0  # 去除直流分量
    s_fft[-3:] = 0  # 去除直流分量
    s_ifft = np.real(np.fft.ifft(s_fft))

    s_mut_dt = s_ifft / fs
    d = []
    total = s_mut_dt[0]
    for i in range(n - 2):
        total = total + s_mut_dt[i + 1]
        d.append(total)
    d_fft = np.fft.fft(d)
    d_fft[:2] = 0
    d_fft[-3:] = 0
    d_ifft = np.real(np.fft.ifft(d_fft))
    return s_ifft * 1000, d_ifft * 1000000  # 单位转换


# 通频速度有效值（10-1000Hz）
def calc_vel_pass_rms(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    :return: 通频速度有效值（10-1000Hz）
    """
    x1 = fft_filter(signal, 10, 1000, fs)
    x2, x3 = acc2dis(x1, fs)
    x4 = fft_filter(x2, 10, 1000, fs)
    vel_pass_rms = np.sqrt(np.mean(x4 ** 2))
    return vel_pass_rms


# 低频速度有效值（3-1000Hz）
def calc_vel_low_rms(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    v_rms: 速度有效值
    """
    x1 = fft_filter(signal, 3, 1000, fs)
    x2, x3 = acc2dis(x1, fs)
    vel_low_rms = np.sqrt(np.mean(x2 ** 2))
    return vel_low_rms


# 加速度有效值（3-10KHz）
def calc_acc_rms(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    a_rms: 加速度有效值
    """
    x1 = fft_filter(signal, 3, 10000, fs)
    acc_rms = np.sqrt(np.mean(x1 ** 2))
    return acc_rms


# 加速度峰值（3-10KHz）
def calc_acc_p(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    a_p: 加速度有效值
    """
    x1 = fft_filter(signal, 3, 10000, fs)
    x2 = sorted(abs(x1), reverse=True)
    x3 = x2[0:100]
    acc_p = np.mean(x3)
    return acc_p


# 振动冲击值（5K~10KHz）
def calc_vibration_impulse(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    impulse: 振动冲击值
    """
    x1 = fft_filter(signal, 5000, 10000, fs)
    x2 = hilbert_envelop(x1)
    x3 = fft_filter(x2, 3, 500, fs)
    impulse = np.sqrt(np.mean(x3 ** 2))
    return impulse


# 加速度峭度指标(3~10KHz)
def calc_acc_kurtosis(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    acc_k: 峭度指标
    """
    x1 = fft_filter(signal, 3, 10000, fs)
    K = sum(x1 ** 4) / len(x1)
    a = sqrt(sum(x1 ** 2) / len(x1))
    acc_kurtosis = K / (a ** 4)
    return acc_kurtosis


# 加速度歪度指标(3-10KHz)
def calc_acc_skew(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    acc_s: 歪度指标
    """
    x1 = fft_filter(signal, 3, 10000, fs)
    S = sum(x1 ** 3) / len(x1)
    a = sqrt(sum(x1 ** 2) / len(x1))
    acc_skew = S / (a ** 3)
    return acc_skew


# 速度峰值(3-1000Hz)
def calc_vel_p(signal, fs):
    """
    signal: 振动加速度信号
    fs: 采样频率
    Return
    vel_p: 速度峰值
    """
    x1 = fft_filter(signal, 10, 1000, fs)
    x2, x3 = acc2dis(x1, fs)
    x4 = sorted(abs(x2), reverse=True)
    x5 = x4[0:100]
    vel_p = np.mean(x5)
    return vel_p


# 加速度传感器特征值计算
def calc_fea_acc(signal, fs):
    """
    加速度传感器特征值计算
    """
    # 通频速度有效值
    vel_pass_rms = calc_vel_pass_rms(signal, fs)
    # 低频速度有效值
    vel_low_rms = calc_vel_low_rms(signal, fs)
    # 加速度有效值
    acc_rms = calc_acc_rms(signal, fs)
    # 加速度峰值
    acc_p = calc_acc_p(signal, fs)
    # 振动冲击值
    vibration_impulse = calc_vibration_impulse(signal, fs)
    # 加速度峭度指标
    acc_kurtosis = calc_acc_kurtosis(signal, fs)
    # 加速度歪度指标
    acc_skew = calc_acc_skew(signal, fs)
    # 速度峰值
    vel_p = calc_vel_p(signal, fs)
    return vel_pass_rms, vel_low_rms, acc_rms, acc_p, vibration_impulse, acc_kurtosis, acc_skew, vel_p


# 速度有效值(3-1000Hz)
def vel_rms(x, fs):
    """
    x: 速度信号
    fs: 采样频率
    Return
    a_rms: 速度有效值
    """
    x1 = fft_filter(x, 3, 1000, fs)
    v_rms = np.sqrt(np.mean(x1 ** 2))
    return v_rms


# 最大正向峰值(5-500Hz)
def calc_max_positive_p(signal, fs):
    """
    signal: 振动位移信号
    fs: 采样频率
    Return
    max_positive_p: 最大正向峰值
    """
    x1 = fft_filter(signal, 5, 500, fs)
    x2 = []
    for i in range(1, len(x1)):
        if x1[i - 1] > 0:
            k = x1[i - 1]
            x2.append(k)
    max_positive_p = abs(max(x2) - np.mean(x1))
    return max_positive_p


# 最大负向峰值(5-500Hz)
def calc_max_negative_p(signal, fs):
    """
    signal: 振动位移信号
    fs: 采样频率
    Return
    max_negative_p: 最大负向峰值
    """
    x1 = fft_filter(signal, 5, 500, fs)
    x2 = []
    for i in range(1, len(x1)):
        if x1[i - 1] < 0:
            k = x1[i - 1]
            x2.append(k)
    max_negative_p = abs(min(x2) - np.mean(x1))
    return max_negative_p


# 监测峰峰值(5-500Hz)
def calc_monitor_pp(signal, fs, L):
    """
    signal: 振动位移信号
    fs: 采样频率
    L: 传感器量程
    Return
    mon_pp: 监测峰峰值
    """
    x1 = fft_filter(signal, 5, 500, fs)
    for i in range(1, len(x1)):
        if x1[i - 1] < x1[i]:
            x1[i] = x1[i - 1] + (1000 / fs) * (L * 0.05)
        elif x1[i - 1] > x1[i]:
            x1[i] = x1[i - 1] - x1[i - 1] * 0.63 * (1 / fs)
    mon_pp = max(x1) - min(x1)
    return mon_pp


# 诊断峰峰值(5-500Hz)
def calc_diagnosis_pp(signal, fs):
    """
    signal: 振动位移信号
    fs: 采样频率
    Return
    dia_pp: 诊断峰峰值
    """
    x1 = fft_filter(signal, 5, 500, fs)
    dia_pp = max(x1) - min(x1)
    return dia_pp


# 峰值(5-500Hz)
def calc_peak(signal, fs):
    """
    signal: 振动位移信号
    fs: 采样频率
    Return
    Peak: 峰值
    """
    x1 = fft_filter(signal, 5, 500, fs)
    x2 = sorted(abs(x1), reverse=True)
    x3 = x2[0:100]
    peak = np.mean(x3)
    return peak


# 峰值因子(5-500Hz)
def calc_peaking_factor(signal, fs):
    """
    signal: 振动位移信号
    fs: 采样频率
    Return
    Cf: 峰值因子
    """
    x1 = fft_filter(signal, 5, 500, fs)
    PP = (max(x1) - min(x1)) / 2
    rms = np.sqrt(np.mean(x1 ** 2))
    peaking_factor = PP / rms
    return peaking_factor


# 推导峰值(5-500Hz)
def calc_derive_p(signal, fs):
    """
    signal: 振动位移信号
    fs: 采样频率
    Return
    de_p: 推导峰值
    """
    x1 = fft_filter(signal, 5, 500, fs)
    rms = np.sqrt(np.mean(x1 ** 2))
    de_p = 1.414 * rms
    return de_p


# 位移传感器特征值计算
def calc_fea_displacement(signal, fs, L):
    """
    位移传感器特征值计算
    """
    # 最大正向峰值
    max_positive_p = calc_max_positive_p(signal, fs)
    # 最大负向峰值
    max_negative_p = calc_max_negative_p(signal, fs)
    # 监测峰峰值
    mon_pp = calc_monitor_pp(signal, fs, L)
    # 诊断峰峰值
    dia_pp = calc_diagnosis_pp(signal, fs)
    # 峰值
    peak = calc_peak(signal, fs)
    # 峰值因子
    peaking_factor = calc_peaking_factor(signal, fs)
    # 推导峰值
    de_p = calc_derive_p(signal, fs)
    return max_positive_p, max_negative_p, mon_pp, dia_pp, peak, peaking_factor, de_p


def envelope_detection(x: np.ndarray):
    """
    提取振动信号的包络曲线
    Parameters
    ----------
    x:  np.ndarray, 原始振动信号
    y： np.ndarray, 希尔伯特变换变换后的解析信号
    Return
    ------
    z: np.ndarray, 包络曲线
    """
    x1 = hilbert(x)
    y = x - x1 * 1j
    z = abs(y)
    return z


# 应力波传感器特征值计算
def calc_fea_ylb(signal):
    """
    应力波传感器特征值计算
    :return ylb_SWE, ylb_SWPE, ylb_SWPA: 应力波能量 脉冲能量 最大峰高
    """
    fs_ylb = 131072
    f_min = 36000  # 低频截至频率
    f_max = 45000  # 高频截至频率
    L_times = 0.1  # 最小排序法时有效:人工设置极限阈值
    signal_filter_data = fft_filter(signal, f_min, f_max, fs_ylb)
    signal_envelope = envelope_detection(signal_filter_data)
    signal_envelope_less = signal_envelope[argrelextrema(signal_envelope, np.less)]
    L_sort = np.sort(signal_envelope_less)
    L_Limit = L_sort[round(len(L_sort) * L_times)]  # 设置最小阀值
    signal_ylb = signal_envelope.copy()
    signal_ylb[signal_ylb <= L_Limit] = 0  # 低于阀值设置为0
    signal_ylb[0] = 0  # 首元素设置为0
    signal_ylb[len(signal_envelope) - 1] = 0  # 尾元素设置为0
    signal_ylb[signal_ylb > L_Limit] = 1
    signal_dt = np.diff(signal_ylb)
    dt_s = [i for i in range(len(signal_dt)) if signal_dt[i] == 1]
    dt_x = [j for j in range(len(signal_dt)) if signal_dt[j] == -1]
    signal_ylb_SWPA = np.linspace(0, 0, len(dt_s))  # 应力波脉冲能量峰值(SWPA)赋初值0
    signal_ylb_SWPE = np.linspace(0, 0, len(dt_s))  # 应力波脉冲能量值(SWPE)赋初值0
    for i in range(len(dt_s)):
        signal_ylb_SWPA[i] = max(signal_envelope[dt_s[i] + 1:dt_x[i] + 1])
        signal_ylb_SWPE[i] = sum(signal_envelope[dt_s[i] + 1:dt_x[i] + 1] - L_Limit)
    ylb_SWE = sum(signal_envelope[signal_envelope > 0])
    ylb_SWPE = sum(signal_ylb_SWPE)
    ylb_SWPA = max(signal_ylb_SWPA)  # 最大峰高 SWPA
    return ylb_SWE, ylb_SWPE, ylb_SWPA


# 获取给定频率的幅值
def calc_multiple_frequency(fft_data, ift, sp):
    """
    获取给定频率的幅值
    Parameters
    ----------
    fft_data: 频谱数据 [频率, 幅值]
    ift: 给定频率
    sp: 采样频率除以采样点数
    Returns
    -------
    f_iX:i倍频幅值
    """
    try:
        f_iX = max(fft_data[floor((ift - 0.5) / sp):ceil((ift + 0.5) / sp)])
    except ValueError:
        f_iX = 0
    return f_iX


# 获取给定阶次倍频的幅值
def calc_fea_HS(fft_data, fea, sp, f_st, f_ord):
    """
    :param fft_data: 频谱数据
    :param fea: 特征值
    :param sp: 采样频率除以采样点数
    :param f_st: 起始阶次
    :param f_ord: 计算阶次
    :return: fea_HS: f_st~(f_st+f_ord-1)倍频幅值
    """
    fea_HS = []
    for i in range(f_st, f_st + f_ord):
        fea_HS.append(calc_multiple_frequency(fft_data, i * fea, sp))
    return fea_HS


# 获取选定频段内的能量和
def calc_HRS(fft_data, fea_min, fea_max, sp):
    """
    :param fft_data: 频谱数据
    :param fea_min: 频率下限
    :param fea_max: 频率上限
    :param sp: 采样频率除以采样点数
    :return: HRS: 能量和
    """
    mul = 0.707
    fea_X = fft_data[floor((fea_min - 0.5) / sp):ceil((fea_max + 0.5) / sp)]
    HRS = sqrt(sum(fea_X ** 2)) * mul
    return HRS


# 获取给定中心频率和边带的上边带能量和
def calc_fea_HCS_up(fft_data, fc, fb, sp, f_st, f_ord):
    """
    获取给定中心频率和边带的上边带能量和
    Parameters
    ----------
    fft_data: 频谱数据 [频率, 幅值]
    fc: 中心频率
    fb: 边带频率
    sp: 采样频率除以采样点数
    f_st: 起始阶次
    f_ord: 计算阶次
    Returns
    -------
    HRS_up:f_st~(f_st+f_ord-1)倍上边带能量和
    """
    mul = 0.707
    HCS_up = []
    for i in range(f_st, f_st + f_ord):
        xb = []
        for j in range(1, 6):
            xb.append(calc_multiple_frequency(fft_data, i * fc + j * fb, sp))
        xb.append(calc_multiple_frequency(fft_data, i * fc, sp))
        HCS_up.append(sqrt(sum([x ** 2 for x in xb])) * mul)
    return HCS_up


# 获取给定中心频率和边带的下边带能量和
def calc_fea_HCS_low(fft_data, fc, fb, sp, f_st, f_ord):
    """
    获取给定中心频率和边带的下边带能量和
    Parameters
    ----------
    fft_data: 频谱数据 [频率, 幅值]
    fc: 中心频率
    fb: 边带频率
    sp: 采样频率除以采样点数
    f_st: 起始阶次
    f_ord: 计算阶次
    Returns
    -------
    HRS_low:f_st~(f_st+f_ord-1)倍下边带能量和
    """
    mul = 0.707
    HCS_low = []
    for i in range(f_st, f_st + f_ord):
        xb = []
        for j in range(1, 6):
            xb.append(calc_multiple_frequency(fft_data, i * fc - j * fb, sp))
        xb.append(calc_multiple_frequency(fft_data, i * fc, sp))
        HCS_low.append(sqrt(sum([x ** 2 for x in xb])) * mul)
    return HCS_low


# 获取给定分频倍频的幅值
def calc_fea_HDS(fft_data, fea, sp, f_st, f_ord):
    """
    :param fft_data: 频谱数据
    :param fea: 特征值
    :param sp: 采样频率除以采样点数
    :param f_st: 起始阶次
    :param f_ord: 计算阶次
    :return: fea_HS: 1/f_st~1/(f_st+f_ord-1)倍频幅值
    """
    fea_HDS = []
    for i in range(f_st, f_st + f_ord):
        fea_HDS.append(calc_multiple_frequency(fft_data, 1 / i * fea, sp))
    return fea_HDS


# 获取给定中心频率和边带频率的边带能量比
def calc_fea_HCR(fft_data, cf, sf, sp, f_st, f_ord):
    """
    计算给定中心频率和边带频率的边带能量比
    Parameters
    ----------
    fft_data: 频谱数据 [频率, 幅值]
    cf: 中心频率
    sf: 边带频率
    sp: 采样频率除以采样点数
    f_st: 起始阶次
    f_ord: 计算阶次
    Returns
    -------
    fea_HCR:f_st~(f_st+f_ord-1)倍能量比
    """
    fea_HCR = []
    for i in range(f_st, f_st + f_ord):
        xb1 = []
        xb2 = []
        for j in range(1, 6):
            xb1.append(calc_multiple_frequency(fft_data, i * cf - j * sf, sp))
            xb2.append(calc_multiple_frequency(fft_data, i * cf + j * sf, sp))
        try:
            Hb1 = sum([x ** 2 for x in xb1])  # 下边带能量和
        except ValueError:
            Hb1 = 0
        try:
            Hb2 = sum([x ** 2 for x in xb2])  # 上边带能量和
        except ValueError:
            Hb2 = 0
        xc = calc_multiple_frequency(fft_data, i * cf, sp)  # 给定中心频率幅值
        if xc == 0:
            H = 0
        else:
            H = sqrt(Hb1 + Hb2) / xc
        fea_HCR.append(H)
    return fea_HCR
