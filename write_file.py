import os

def write_file(file_path: str, content: str, encoding: str = "utf-8",
               mode: str = "w", create_dirs: bool = True) -> None:
    """
    将内容写入本地文本文件

    :param file_path:   目标文件路径
    :param content:     要写入的字符串内容
    :param encoding:    编码格式,默认 utf-8
    :param mode:        'w' 覆盖写 / 'a' 追加写 / 'x' 文件存在则报错
    :param create_dirs: 是否自动创建父目录
    """
    if create_dirs:
        parent = os.path.dirname(os.path.abspath(file_path))
        os.makedirs(parent, exist_ok=True)

    with open(file_path, mode, encoding=encoding) as f:
        f.write(content)


# 使用示例
write_file("output.txt", "你好,世界\n第二行")
write_file("logs/app.log", "追加一行\n", mode="a")  # 追加模式
