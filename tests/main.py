#!/usr/bin/python3

from libs.Manager import Manager

manager = Manager("file.xls")

file_content = manager.parse_file()

print(file_content)
