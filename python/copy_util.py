import shutil
from pathlib import Path
import time
import codecs

from typing import List


class CopytreeIgnore(object):
    """
    Utility for shutil.copytree(), 'ignore' argument.
    can be used instead of shutil.ignore_patterns().
    """
    log_path: Path = Path().home() / "out.log"

    def __init__(self, in_pattern=None, ex_pattern=None):
        if in_pattern:
            self.set_include_pattern(in_pattern)
        if ex_pattern:
            self.set_exclude_pattern(ex_pattern)

    def set_include_pattern(self, ptn) -> None:
        try:
            # check whether ptn is iterable
            some_iterator = iter(ptn)
            self.in_patterns: List[str] = list(ptn)
        except TypeError:
            print(ptn, "is not iterable")

    def set_exclude_pattern(self, ptn) -> None:
        try:
            # check whether ptn is iterable
            some_iterator = iter(ptn)
            self.ex_patterns: List[str] = list(ptn)
        except TypeError:
            print(ptn, "is not iterable")

    def exclude(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that doesn't match glob patterns in
        'ex_patterns'.
        """
        cwd: Path = Path(directory)
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
        ignores: set = set()
        for ptn in self.ex_patterns:
            for file_path in cwd.glob(ptn):
                ignores.add(file_path.name)

        return list(ignores)

    def include(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that matches glob patterns in 'in_patterns'.
        """
        includes: set = set()
        cwd = Path(directory)
        for ptn in self.in_patterns:
            for file_path in cwd.glob(ptn):
                includes.add(file_path.name)
        ignores: set = set(files)

        return list(ignores - includes)

    def include_and_exclude(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        include patterns are prefered to exclude patterns.
        """
        excludes: set = set()
        includes: set = set()
        cwd = Path(directory)
        for ptn in self.ex_patterns:
            for file_path in cwd.glob(ptn):
                excludes.add(file_path.name)
        for ptn in self.in_patterns:
            for file_path in cwd.glob(ptn):
                includes.add(file_path.name)
        includes = set(files) - includes

        return list(includes | excludes)

    def _is_pattern_match(self, dir_path):
        # TODO: パターンと一致確認
        pass


def main():
    in_list = [
        '.svn',
        'vssver.scc',
    ]
    # パスを指定--------------------------------------
    srs = Path("C:/Workspace/Scripts/scripts/test")
    dst = Path("C:/Workspace/Scripts/scripts/test2")
    # ------------------------------------------------
    IgnPtn = CopytreeIgnore(in_pattern=in_list)
    shutil.copytree(srs, dst, ignore=IgnPtn.include_and_exclude,
                    dirs_exist_ok=True)


if __name__ == "__main__":
    main()
