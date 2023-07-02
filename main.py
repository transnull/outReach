import os

from tools.tool import *


def test():
    # 获取固定会话
    session = login(url="https://xxxx.com/login/dologin",
                    username="xxxxx",
                    password="xxxxx"
                    )

    # 获取当前日期
    #today = "2023-04-03"
    today = get_current_date()
    time_str = today + ' 00:00:00'
    # 得到外联信息词典
    outreach_info_dict = get_outreach_info(session, time_str)

    # 得到细节信息
    illicit_outnet_sum = process_data(session, outreach_info_dict, today)
    for table_info in illicit_outnet_sum:
        if table_info:
            local_network = table_info[0]['local_network']

            output_file_name = generate_filename(local_network,today=today)

            if not os.path.exists(output_file_name):
                save_to_excel(ip_list=table_info, template_file_path="template.xlsx", output_file_path=output_file_name)
            else:
                print('该单位已经处理')


if __name__ == '__main__':
    # 初始化处理
    path = r"./"
    os.chdir(path)  # 修改工作路径
    test()
