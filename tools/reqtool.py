import requests


def login(username, password, url):
    # 创建 session 对象
    session = requests.Session()

    # 准备 POST 请求参数
    data = {
        'name': username,
        'password': password
    }

    # 发送登录请求
    response = session.post(f'{url}', data=data)

    # 检查响应状态码
    if response.status_code != 200:
        raise Exception(f"Failed to login. Status code: {response.status_code}")

    # 返回 session 对象
    return session


def get_outreach_info(session, start_time):
    # 准备请求参数
    params = {
        'outStartTime2': start_time,
    }

    # 发送GET请求
    response = session.get('https://zjzww.ywwlgj.net:8443/abps/boundaryInoutnetworkmerge/list', params=params)
    outreach = response.json()
    return outreach


def get_outnet_details(session, in_ip, gateway_series):
    # 准备请求参数
    params = {
        'gatewaySeries': gateway_series,
        'inIp': in_ip,
        "page": 1,
    }

    # 发送GET请求
    response = session.get('https://zjzww.ywwlgj.net:8443/abps/boundaryInoutnetworkmerge/listDetail', params=params)
    detail_dict = response.json()

    return detail_dict
