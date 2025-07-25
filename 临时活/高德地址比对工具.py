import requests
import pandas as pd
import logging
import time

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)  

# 高德地图API Key（需自行申请）
AMAP_KEY = "0ab7ac8e13b9e6d793a7ea83763b5f8a"
QPS_LIMIT_DELAY = 0.3  # 每次请求间隔0.3秒（3次/秒）
MAX_RETRIES = 1  # 最大重试次数

def get_amap_coordinates(province, city, district, address):
    """通过高德地图API获取地址的经纬度"""
    full_address = ''.join([province or '', city or '', district or '', address or ''])
    url = "https://restapi.amap.com/v3/geocode/geo"
    params = {
        "key": AMAP_KEY,
        "address": full_address,
        "output": "json"
    }
    for retry in range(MAX_RETRIES):
        try:
            time.sleep(QPS_LIMIT_DELAY)
            response = requests.get(url, params=params)
            data = response.json()
            if data.get("status") == "1" and int(data.get("count", 0)) > 0:
                return data["geocodes"][0]["location"]  # 直接返回"lng,lat"格式字符串
            logger.warning(f"地址解析失败（第{retry+1}次）: {data.get('info', '未知错误')}")
        except Exception as e:
            logger.error(f"地址解析异常（第{retry+1}次）: {e}")
    return None

def get_amap_distance(origin, destination):
    "使用高德API计算两点间距离（米）"
    url = "https://restapi.amap.com/v3/distance"
    params = {
        "key": AMAP_KEY,
        "origins": origin,
        "destination": destination,
        "type": 1  # 1:直线距离，2:驾车距离（需考虑路况）
    }
    for retry in range(MAX_RETRIES):
        try:
            time.sleep(QPS_LIMIT_DELAY)
            response = requests.get(url, params=params)
            data = response.json()
            if data.get("status") == "1" and data.get("results"):
                return int(data["results"][0]["distance"])  # 直接返回米数
            logger.warning(f"距离计算失败（第{retry+1}次）: {data.get('info', '未知错误')}")
        except Exception as e:
            logger.error(f"距离计算异常（第{retry+1}次）: {e}")
    return None

class ExcelAddressComparator:
    def __init__(self):
        pass

    def compare_excel_files(self, file_path, address_cols):
        df = pd.read_excel(file_path)
        required_cols = address_cols['deliver'] + address_cols['install']
        if not all(col in df.columns for col in required_cols):
            raise ValueError(f"缺失必要列: {[col for col in required_cols if col not in df.columns]}")

        # 新增结果列
        df['省份匹配'] = df[address_cols['deliver'][0]] == df[address_cols['install'][0]]
        df['城市匹配'] = df[address_cols['deliver'][1]] == df[address_cols['install'][1]]
        df['区县匹配'] = df[address_cols['deliver'][2]] == df[address_cols['install'][2]]
        df['地址名称一致'] = df[address_cols['deliver'][3]] == df[address_cols['install'][3]]  # 新增：详细地址名称一致性检查
        df['高德测距(米)'] = None
        df['最终匹配'] = False

        for idx, row in df.iterrows():
            logger.info(f"处理第 {idx+1}/{len(df)} 条记录...")

            # 新增：名字完全一致时直接标记匹配
            if row['地址名称一致']:
                logger.info(f"第 {idx+1} 条记录地址名称完全一致，跳过API调用")
                df.at[idx, '高德测距(米)'] = 0  # 距离设为0表示完全一致
                df.at[idx, '最终匹配'] = True
                continue

            # 原行政区域匹配检查
            if not (row['省份匹配'] and row['城市匹配'] and row['区县匹配']):
                logger.info(f"第 {idx+1} 条记录行政区域不匹配，跳过API调用")
                df.at[idx, '最终匹配'] = False
                continue

            # 获取经纬度（格式："lng,lat"）
            origin = get_amap_coordinates(*row[address_cols['deliver']])
            dest = get_amap_coordinates(*row[address_cols['install']])
            if not (origin and dest):
                logger.warning(f"第 {idx+1} 条记录经纬度获取失败")
                continue

            # 调用高德测距API
            distance = get_amap_distance(origin, dest)
            if distance is not None:
                df.at[idx, '高德测距(米)'] = distance
                df.at[idx, '最终匹配'] = (distance <= 300)  # 行政区域已匹配，仅需判断距离

        output_path = file_path.replace('.xlsx', '_高德测距结果.xlsx')
        df.to_excel(output_path, index=False)
        logger.info(f"结果已保存至: {output_path}")
        return df

if __name__ == "__main__":
    comparator = ExcelAddressComparator()
    comparator.compare_excel_files(
        file_path=r"C:\Users\zhangbon\Desktop\送安一致率数据测试.xlsx",
        address_cols={
            'deliver': ['送货省份', '送货城市', '送货区域', '详细地址'],
            'install': ['安装省', '安装市', '安装区县', '安装详细地址']
        }
    )