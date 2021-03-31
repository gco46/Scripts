import shutil
from pathlib import Path
import sys
import time
import codecs


class CopytreeIgnore(object):
    """
    Utility for shutil.copytree(), 'ignore' argument.
    can be used instead of shutil.ignore_patterns().
    """
    log_path = Path().home() / "out.log"

    def __init__(self, in_pattern=None, ex_pattern=None):
        if in_pattern:
            self.set_include_pattern(in_pattern)
        if ex_pattern:
            self.set_exclude_pattern(ex_pattern)

    def set_include_pattern(self, ptn):
        try:
            # check whether ptn is iterable
            some_iterator = iter(ptn)
            self.in_patterns = tuple(ptn)
        except TypeError:
            print(ptn, "is not iterable")

    def set_exclude_pattern(self, ptn):
        try:
            # check whether ptn is iterable
            some_iterator = iter(ptn)
            self.ex_patterns = tuple(ptn)
        except TypeError:
            print(ptn, "is not iterable")

    def exclude(self, directory, files):
        cwd = Path(directory)
        # 標準出力
        print('now copying', str(cwd))
        # ファイル出力
        print('now copying', str(cwd), file=codecs.open(
            str(self.log_path), mode="a"))
        while True:
            disk_usage_gb = shutil.disk_usage('C:/').free / 1024 / 1024 / 1024
            if disk_usage_gb > 20:
                break
            print('wait for GDrive Sync...: free is {:>6.2f} GB'.format(
                disk_usage_gb))
            # wait for an half hour
            time.sleep(3600)

        # ignore file/directories list
        ignores = []
        for ptn in self.ex_patterns:
            for file_path in cwd.glob(ptn):
                ignores.append(file_path.name)

        return ignores

    def include(self, directory, files):
        # TODO: メソッド修正(excludeと同じになっている)
        ignores = []
        cwd = Path(directory)
        for ptn in self.in_patterns:
            for file_path in cwd.glob(ptn):
                ignores.append(file_path.name)
        return ignores

    def include_and_exclude(self, directory, files):
        """
        include patterns are prefered to exclude patterns.
        """
        result = []
        cwd = Path(directory)
        for ptn in self.in_patterns:
            for file_path in cwd.glob(ptn):
                # TODO: excludeパターンと一致していれば除外
                pass

    def _is_pattern_match(self, dir_path):
        # TODO: パターンと一致確認
        pass


def main():
    ignore_list = [
        '.svn',
        'vssver.scc',
    ]
    # パスを指定--------------------------------------
    srs = Path("C:/Workspace/Scripts/scripts/test")
    dst = Path("C:/Workspace/Scripts/scripts/test2")
    # ------------------------------------------------
    IgnPtn = CopytreeIgnore(ex_pattern=ignore_list)
    shutil.copytree(srs, dst, ignore=IgnPtn.exclude, dirs_exist_ok=True)


if __name__ == "__main__":
    main()
