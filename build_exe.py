# -*- coding: utf-8 -*-
"""
本地打包脚本（本脚本不进仓库也行，但绝不包含真实Key）
用法：
1. 在项目目录创建 key.txt，里面只放一行你的高德Key（key.txt 已在 .gitignore 中）
2. 如需内置数据：把 data_all.json、embed/采集数据缓存.json、embed/已导出电话记录.json 放好
   （这些是本地数据文件，已在 .gitignore 中，永远不会被提交）
3. 运行：python build_exe.py
产物：dist/商户数据查看器.exe（Key已焊入、数据已内嵌）
"""
import os
import sys
import shutil
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
KEY_FILE = os.path.join(BASE, 'key.txt')
SRC = os.path.join(BASE, 'viewer_app.py')
TMP = os.path.join(BASE, '_viewer_app_build.py')  # 临时注入Key的副本，打完即删


def main():
    if not os.path.exists(KEY_FILE):
        sys.exit('缺少 key.txt：请创建该文件并写入你的高德Key（一行，无空格）')
    key = open(KEY_FILE, encoding='utf-8').read().strip()
    if not key:
        sys.exit('key.txt 是空的')

    with open(SRC, encoding='utf-8') as f:
        code = f.read()
    placeholder = "AMAP_KEY = ''"
    if placeholder not in code:
        sys.exit('viewer_app.py 里没找到 Key 占位符，文件可能被改过')
    code = code.replace(placeholder, f"AMAP_KEY = '{key}'", 1)
    with open(TMP, 'w', encoding='utf-8') as f:
        f.write(code)

    try:
        cmd = [sys.executable, '-m', 'PyInstaller', '--onefile', '--windowed',
               '--name', '商户数据查看器', '--hidden-import', 'openpyxl',
               '--clean', '--noconfirm']
        # 内置数据存在才打包进去（开源用户没有这些文件也能正常打包）
        for p in ('data_all.json',
                  os.path.join('embed', '采集数据缓存.json'),
                  os.path.join('embed', '已导出电话记录.json')):
            if os.path.exists(os.path.join(BASE, p)):
                cmd += ['--add-data', f'{p};.']
            else:
                print(f'[提示] {p} 不存在，跳过内嵌')
        cmd.append(TMP)
        subprocess.run(cmd, cwd=BASE, check=True)
        print('\n打包完成：dist/商户数据查看器.exe')
    finally:
        if os.path.exists(TMP):
            os.remove(TMP)  # 含Key的临时文件必须删掉
        # 清掉 PyInstaller 生成的 spec（里面可能带路径信息，无所谓但保持目录干净）
        spec = os.path.join(BASE, '商户数据查看器.spec')
        if os.path.exists(spec):
            os.remove(spec)


if __name__ == '__main__':
    main()
