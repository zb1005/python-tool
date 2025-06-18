import numpy as np
import matplotlib.pyplot as plt
from mpl_toolkits.basemap import Basemap
import random

# 设置图形
plt.figure(figsize=(16, 10))

# 创建地图实例
m = Basemap(projection='mill', llcrnrlat=-60, urcrnrlat=80,
            llcrnrlon=-180, urcrnrlon=180, resolution='c')

# 绘制地图特征
m.drawcoastlines()
m.drawcountries()
m.drawmapboundary(fill_color='#99ccff')
m.fillcontinents(color='#cc9966', lake_color='#99ccff')

# 中美主要城市坐标 (纬度, 经度)
us_cities = [
    (40.7128, -74.0060),   # 纽约
    (34.0522, -118.2437),  # 洛杉矶
    (37.7749, -122.4194),  # 旧金山
    (38.9072, -77.0369),   # 华盛顿
    (41.8781, -87.6298)    # 芝加哥
]

cn_cities = [
    (39.9042, 116.4074),   # 北京
    (31.2304, 121.4737),   # 上海
    (23.1291, 113.2644),   # 广州
    (30.5728, 104.0668),   # 成都
    (22.5431, 114.0579)    # 深圳
]

# 全球其他主要节点
other_cities = [
    (51.5074, -0.1278),    # 伦敦
    (48.8566, 2.3522),     # 巴黎
    (35.6762, 139.6503),   # 东京
    (55.7558, 37.6173),    # 莫斯科
    (37.5665, 126.9780)    # 首尔
]

all_cities = us_cities + cn_cities + other_cities

# 生成中美之间的网络攻防连接
connections = []
for us in us_cities:
    for cn in cn_cities:
        # 增加中美之间的连接密度
        for _ in range(5):
            connections.append((us, cn))

# 改进的连接线绘制方法
def plot_connection(m, lon1, lat1, lon2, lat2, color, alpha=0.3, linewidth=1):
    # 将连接线分成多段绘制，避免断开
    steps = 50
    lons = np.linspace(lon1, lon2, steps)
    lats = np.linspace(lat1, lat2, steps)
    x, y = m(lons, lats)
    m.plot(x, y, color=color, alpha=alpha, linewidth=linewidth)

# 绘制连接线
for (lat1, lon1), (lat2, lon2) in connections:
    if (lat1, lon1) in us_cities and (lat2, lon2) in cn_cities:
        plot_connection(m, lon1, lat1, lon2, lat2, 'red')
    elif (lat1, lon1) in cn_cities and (lat2, lon2) in us_cities:
        plot_connection(m, lon1, lat1, lon2, lat2, 'blue')

# 绘制城市节点
x, y = m(np.array([lon for lat, lon in us_cities + cn_cities]), 
         np.array([lat for lat, lon in us_cities + cn_cities]))
m.scatter(x, y, color='green', s=30, alpha=0.8, zorder=5)

# 添加标题和图例
plt.title('DeepSeek上线时中美网络攻防数据流向 (2023)', fontsize=16)
plt.legend(handles=[
    plt.Line2D([0], [0], color='red', lw=2, label='美国到中国 (攻击)'),
    plt.Line2D([0], [0], color='blue', lw=2, label='中国到美国 (防御)'),
    plt.Line2D([0], [0], marker='o', color='w', markerfacecolor='green', markersize=8, label='网络节点')
], loc='lower left')

plt.show()