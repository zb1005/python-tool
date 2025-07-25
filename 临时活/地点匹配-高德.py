import requests
import math

# 高德地图API Key（需自行申请）
AMAP_KEY = "0ab7ac8e13b9e6d793a7ea83763b5f8a"

def get_amap_coordinates(address, city=None):
    """通过高德地图API获取地址的经纬度"""
    url = "https://restapi.amap.com/v3/geocode/geo"
    params = {
        "key": AMAP_KEY,
        "address": address,
        "city": city or "",  # 可选参数，指定城市
        "output": "json"
    }
    try:
        response = requests.get(url, params=params)
        data = response.json()
        if data.get("status") == "1" and data.get("count") > "0":
            location = data["geocodes"][0]["location"]
            lng, lat = map(float, location.split(','))
            return lat, lng
        else:
            print(f"地址解析失败: {data.get('info', '未知错误')}")
    except Exception as e:
        print(f"请求异常: {e}")
    return None, None

def haversine_distance(lat1, lon1, lat2, lon2):
    """计算两个经纬度之间的球面距离（千米）"""
    R = 6371.0
    lat1_rad = math.radians(lat1)
    lon1_rad = math.radians(lon1)
    lat2_rad = math.radians(lat2)
    lon2_rad = math.radians(lon2)
    
    dlat = lat2_rad - lat1_rad
    dlon = lon2_rad - lon1_rad
    a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    return R * c

# 示例：计算两个地址的距离
def compare_addresses(address1, address2, threshold=0.5):
    lat1, lon1 = get_amap_coordinates(address1)
    lat2, lon2 = get_amap_coordinates(address2)
    
    if None in (lat1, lon1, lat2, lon2):
        return False, float('inf')
    
    distance = haversine_distance(lat1, lon1, lat2, lon2)
    return distance <= threshold, distance

# 使用示例
result, dist = compare_addresses("北京市海淀区中关村", "北京市海淀区中关村大街1号")
print(f"匹配结果: {result}, 距离: {dist:.3f}千米")