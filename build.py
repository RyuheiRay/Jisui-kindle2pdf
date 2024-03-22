import os
import shutil
import subprocess
from pathlib import Path

EXE_NAME = 'Jisui-kindle2pdf'
PROJECT_PATH = Path(__file__).parent
DIST_PATH = PROJECT_PATH / 'dist'


def copy_assets_to_dist():
    # リソースファイルの元のパス
    assets_path = Path(__file__).parent / 'src/assets'

    # distディレクトリのパス
    dist_path = DIST_PATH / EXE_NAME

    # distディレクトリの作成
    if not os.path.exists(dist_path):
        os.makedirs(dist_path)

    try:
        # distディレクトリにリソースファイルをコピー
        shutil.copytree(assets_path, os.path.join(dist_path, 'assets'))
        print("リソースファイルをコピーしました。")
    except Exception as e:
        print(f"リソースファイルのコピーに失敗しました: {e}")


def main():
    # distディレクトリの作成
    if not os.path.exists(DIST_PATH):
        os.makedirs(DIST_PATH)
    # 前回の生成物を削除する
    shutil.rmtree(DIST_PATH)
    # PyInstallerのコマンドを作成します
    pyinstaller_command = [
        'pyinstaller',
        '--noconsole',  # コンソールを非表示にします
        '--paths',  # パスを追加します
        f'{PROJECT_PATH}',  # パスを指定します
        '--name',  # 実行ファイル名を指定します
        f'{EXE_NAME}',  # 実行ファイル名
        f"{PROJECT_PATH/'src/main.py'}"  # ビルドするスクリプトファイルを指定します
    ]

    # PyInstallerを実行します
    try:
        subprocess.run(pyinstaller_command, check=True)
        print("Build successful.")
    except subprocess.CalledProcessError as e:
        print(f"Build failed: {e}")

    copy_assets_to_dist()


if __name__ == "__main__":
    main()
