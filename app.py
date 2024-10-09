import pandas as pd
import os


# 定义函数读取Excel中的数据
def read_data_from_file(file_path):
    try:
        df = pd.read_excel(file_path)
        if "订单号" in df.columns:
            return df
        else:
            print(f"文件 {file_path} 不包含 '订单号' 列")
            return pd.DataFrame()
    except Exception as e:
        print(f"读取文件 {file_path} 时发生错误: {e}")
        return pd.DataFrame()


# 从指定文件夹中读取所有Excel文件的数据
def read_data_from_folder(folder_path):
    all_data = pd.DataFrame()  # 初始化一个空的DataFrame
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".xlsx") or file_name.endswith(
            ".xls"
        ):  # 检查是否是Excel文件
            file_path = os.path.join(folder_path, file_name)
            print(f"正在读取文件: {file_path}")
            file_data = read_data_from_file(file_path)

            # 对每个文件进行重复处理
            file_data = handle_duplicates_in_file(file_data, file_path)

            # 合并当前文件数据
            all_data = pd.concat([all_data, file_data], ignore_index=True)
    return all_data


# 检查文件中的重复项，并根据用户选择删除并保存
def handle_duplicates_in_file(data, file_path):
    duplicates = data[data.duplicated(subset=["订单号"], keep=False)]
    if not duplicates.empty:
        print(f"文件 {file_path} 中发现重复的订单号：")
        print(duplicates[["订单号"]])  # 打印重复的订单号
        action = input(f"是否要删除文件 {file_path} 中的重复订单号？ (y/n): ")
        if action.lower() == "y":
            data = data.drop_duplicates(
                subset=["订单号"], keep="first"
            )  # 删除重复，保留第一个出现的
            data.to_excel(file_path, index=False)  # 将去重后的数据保存到原文件
            print(f"已删除 {file_path} 中的重复项并保存。")
        else:
            print(f"未删除 {file_path} 中的重复项。")
    return data


# 检查合并前的重复项
def check_merge_duplicates(existing_data, new_data):
    duplicates = new_data[new_data["订单号"].isin(existing_data["订单号"])]
    if not duplicates.empty:
        print("即将合并的数据中存在重复订单号：")
        print(duplicates[["订单号"]])  # 打印重复的订单号
        action = input("是否继续合并这些重复订单号？ (y/n): ")
        return action.lower() == "y"
    return True


# 保存数据到Excel，并进行去重
def save_data(file_path, data):
    data.drop_duplicates(subset=["订单号"], keep="first", inplace=True)  # 最终去重
    data.to_excel(file_path, index=False)


# 如果 merged_orders.xlsx 不存在则创建空文件
def create_empty_merged_file(file_path):
    if not os.path.exists(file_path):
        print(f"{file_path} 不存在，正在创建一个空的文件...")
        df = pd.DataFrame(columns=["订单号"])  # 创建空的DataFrame，只有'订单号'列
        df.to_excel(file_path, index=False)


# 主函数，执行整个流程
def main():
    # 桌面路径和目标文件夹
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    folder_path = os.path.join(desktop_path, "order")  # 要检查的文件夹路径
    merged_file = os.path.join(
        desktop_path, "merged_orders.xlsx"
    )  # 合并后订单号的保存文件

    # 检查文件夹是否存在，如果不存在则创建
    if not os.path.exists(folder_path):
        print(f"文件夹 {folder_path} 不存在，正在创建该文件夹...")
        os.makedirs(folder_path)

    # 从文件夹中读取所有Excel文件的数据
    all_data = read_data_from_folder(folder_path)

    if all_data.empty:
        print("没有找到任何有效的订单数据，或文件夹中没有有效的Excel文件。")
        return

    # 检查并创建空的 merged_orders.xlsx 文件
    create_empty_merged_file(merged_file)

    # 读取合并文件（merged_orders.xlsx）
    existing_data = read_data_from_file(merged_file)

    # 检查合并前的重复
    if not check_merge_duplicates(existing_data, all_data):
        print("用户选择不合并。")
        return

    # 合并并保存
    combined_data = pd.concat([existing_data, all_data], ignore_index=True)
    save_data(merged_file, combined_data)
    print(f"数据已合并并保存到 {merged_file}")


if __name__ == "__main__":
    main()
