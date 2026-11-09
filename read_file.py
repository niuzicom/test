import os

def read_file(file_path: str, encoding: str = "utf-8", errors: str = "strict") -> str:
    """
    读取本地文本文件内容

    :param file_path: 文件路径
    :param encoding:  编码格式,默认 utf-8(Windows 下的 txt 可能是 gbk)
    :param errors:    编码错误处理方式: 'strict' / 'ignore' / 'replace'
    :return:          文件内容字符串
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")
    if not os.path.isfile(file_path):
        raise IsADirectoryError(f"路径不是一个文件: {file_path}")

    with open(file_path, "r", encoding=encoding, errors=errors) as f:
        return f.read()


# 使用示例
content = read_file("test.txt")
print(content)
