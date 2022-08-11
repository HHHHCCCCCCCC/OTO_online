from scipy.signal import savgol_filter
import numpy as np
from scipy.interpolate import make_interp_spline  # 导入插值模块

class pretreatment():
###############预处理方法################
################平滑MA、SG################
    def smoothMA(self, a, WSZ):
        # a:原始数据，NumPy 1-D array containing the data to be smoothed
        # 必须是1-D的，如果不是，请使用 np.ravel()或者np.squeeze()转化
        # WSZ: smoothing window size needs, which must be odd number,
        a = np.ravel(a)
        a = a.tolist()
        out0 = np.convolve(a, np.ones(WSZ, dtype=int), 'valid') / WSZ
        r = np.arange(1, WSZ - 1, 2)
        start = np.cumsum(a[:WSZ - 1])[::2] / r
        stop = (np.cumsum(a[:-WSZ:-1])[::2] / r)[::-1]
        return np.concatenate((start, out0, stop))


    def Diff(self, data_x, windowsize, deriv):
        data_x = np.mat(data_x)
        data_x = data_x.T
        # print(data_x)
        ans = savgol_filter(data_x, windowsize, 2, deriv, axis=0, mode='nearest')
        ans[0, 0] = 0
        ans[-1, -1] = 0
        # ans = ans.tolist()
        return ans


    # def SG(self,X):
    #     # X = np.array(X).flatten()
    #     # X = X.tolist()
    #     X_SG = savgol_filter(X , 3, 2, deriv=1, axis=1, mode='nearest')  # 拟合数据，拟合阶次，求几阶导数
    #     X_SG[0, 0] = 0
    #     X_SG[-1, -1] = 0
    #     return X_SG

    ################归一化################
    def MaxNorm(self, x):  # 最大值归一化
        x = np.mat(x)
        Abs = np.abs(x)
        _Max_ = np.max(Abs, axis=1)
        _x_ = x / _Max_
        return _x_


    def MaxMinNormalize(self, x):  # 最大最小值归一化 ############## 输入的数据必须是以列排列的，每一列是一组光谱数据   axis=0表示取每一列的最值
        x = np.mat(x)
        _Max_ = np.max(x, axis=1)
        _Min_ = np.min(x, axis=1)
        _x_ = (x - _Min_) / (_Max_ - _Min_)
        return _x_


    # X_max = np.max(X, axis=0)  # axis =1 是每一列的最小值  而  axis =0 是每一行的最小值
    # X_min = np.min(X, axis=0)  # axis =1 是每一列的最小值  而  axis =0 是每一行的最小值

    def vector_Normalize(self, x):  # 矢量归一化# 输入的数据必须是以列排列的，每一列是一组光谱数据   axis=0表示取每一列的最值
        x = np.mat(x)
        _mean_ = np.mean(x, axis=1)
        _squre_ = np.square(x)
        _squresum_ = _squre_.sum(axis=1)
        _Sqrtmean_ = np.sqrt(_squresum_)
        _x_ = (x - _mean_) / _Sqrtmean_
        return _x_


    # X_max = np.max(x, axis=0)  # axis =1 是每一列的最小值  而  axis =0 是每一行的最小值
    # X_min = np.min(X, axis=0)  # axis =1 是每一列的最小值  而  axis =0 是每一行的最小值

    ################求导################
    def D1(self, x, y):  # 一阶求导
        dydx = np.diff(y) / np.diff(x)
        dydx = dydx.tolist()  # 求导后y个数减少一个，最后添加一个0，数据保持原有个数
        dydx.append(0)
        return dydx


    def D2(self, x, y):  # 二阶求导
        dydx = np.diff(y) / np.diff(x)
        dydx = dydx.tolist()
        dydx.append(0)  # 求导后y个数减少一个，最后添加一个0，数据保持原有个数
        dydx2 = np.diff(dydx) / np.diff(x)
        dydx2 = dydx2.tolist()
        dydx2.append(0)
        return dydx2


    def cen_difference_l(self, x, y):  # 一阶中心差分求导，x表示波长列表，y表示密度列表
        list = []
        n = len(y)
        i = 1
        while (i < (n - 1)):
            X_cen = (y[i + 1] - y[i - 1]) / (x[i + 1] - x[i - 1])
            i += 1
            list.append(X_cen)
        list.insert(0, 0)
        list.append(0)  # 在列表末尾和第一个位置添加0，保持元素个数一致
        return list


    def cen_difference_2(self, x, y):  # 二阶中心差分求导
        list = []
        n = len(y)
        i = 1
        while (i < (n - 1)):
            X_cen = ((y[i + 1] + y[i - 1] - 2 * y[i]) / ((x[i + 1] - x[i - 1]) ** 2))
            i += 1
            list.append(X_cen)
        list.insert(0, 0)
        list.append(0)
        return list


    ################标准正态分布################
    def SNV(self, x):  # SNV中标准差分为总体标准差和样本标准差，ddof=1则为样本标准差，即分母除以n-1个数
        Intensities_SNV = x  # 定义一个列表，用于存放SNV计算结果
        Intensities_Mean = np.mean(Intensities_SNV)
        Intensities_Std = np.std(Intensities_SNV, ddof=1)
        for i in range(0, len(Intensities_SNV)):
            Intensities_SNV[i] = (Intensities_SNV[i] - Intensities_Mean) / Intensities_Std
        return Intensities_SNV


    ################多元散射校正################
    def MSC(self, data_x):  # 多元散射校正
        data_x = np.array(data_x)
        mean = np.mean(data_x, axis=0)
        [n] = data_x.shape
        msc_x = np.ones((n))
        poly = np.polyfit(mean, data_x, 1)  # 每个样本做一次一元线性回归
        for i in range(n):
            msc_x[i] = (data_x[i] - poly[1]) / poly[0]
        return msc_x


    ################均值中心化################
    def MeanCenter(self, x):  # 平均值
        _x_bar_ = np.mean(x, axis=1)
        _x_ = x - _x_bar_
        _x_ = _x_.tolist()
        return _x_


    def AutosCaling(self, x):
        _std_ = np.std(x, ddof=1,axis=1)  ############ # ### std c = np.sqrt(((a - np.mean(a)) ** 2).sum() / (a.size - 1))  axis =1 是针对行求方差 ####
        _x_ = x / _std_
        return _x_


    def Spline(self, x, y , x_min , x_max ,n):  # 三次样条插值
        _x_ = np.array(x)
        _y_ = np.array(y)  # 取出y中的一列
        x_smooth = np.linspace(x_min, x_max, n)  # np.linspace 等差数列,从x.min()到x.max()生成750个数，便于后续插值
        y_smooth = make_interp_spline(x, y)(x_smooth)
        return x_smooth, y_smooth  # 返回插值后的光谱数据
    ###############以上为预处理方法#################

