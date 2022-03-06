from os import path


def get_file_id(filepath):
    head, _ = path.split(filepath)
    return path.split(head)[1]