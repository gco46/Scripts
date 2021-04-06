import shutil
from pathlib import Path
import time
import codecs

from typing import List, Iterable, Callable, Optional


class CopytreeIgnore(object):
    """
    Utility for shutil.copytree(), 'ignore' argument.
    following methods can be used instead of shutil.ignore_patterns().
      - exclude
      - include
      - include_and_exclude
    """
    DirCall = Callable[[str, List[str]], None]

    def __init__(self,
                 in_patterns: Iterable[str] = [],
                 ex_patterns: Iterable[str] = [],
                 callback: Optional[DirCall] = None):
        self._set_include_pattern(in_patterns)
        self._set_exclude_pattern(ex_patterns)
        self._set_callback(callback)

    def _set_include_pattern(self, ptn: Iterable[str]) -> None:
        try:
            # check whether ptn is iterable
            iter(ptn)
            self.in_patterns: List[str] = list(ptn)
        except TypeError:
            print(ptn, "is not iterable")

    def _set_exclude_pattern(self, ptn: Iterable[str]) -> None:
        try:
            # check whether ptn is iterable
            iter(ptn)
            self.ex_patterns: List[str] = list(ptn)
        except TypeError:
            print(ptn, "is not iterable")

    def _set_callback(self, callback: Optional[DirCall]) -> None:
        self.callback = callback

    def exclude(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that doesn't match glob patterns in
        'ex_patterns'.
        """
        if self.callback:
            self.callback(directory, files)

        cwd: Path = Path(directory)
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
        cwd: Path = Path(directory)
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
        cwd: Path = Path(directory)
        if self.in_patterns:
            for ptn in self.in_patterns:
                for file_path in cwd.glob(ptn):
                    includes.add(file_path.name)
            includes = set(files) - includes
        else:
            # if in_patterns is empty, all files and directories become target
            # so ignore file is none
            includes = set([])
        for ptn in self.ex_patterns:
            for file_path in cwd.glob(ptn):
                excludes.add(file_path.name)

        return list(includes | excludes)

    def _is_pattern_match(self, dir_path):
        # TODO: パターンと一致確認
        pass


def wait_for_gdrive_sync(directory: str, files: List[str]) -> None:
    """
    it is used for FileSream upload, more than C drive capacity.
    """
    # save out.log in home directory
    log_path: Path = Path().home() / 'out.log'
    cwd: Path = Path(directory)
    # print next copy directory
    print('now copying', str(cwd))
    # print next copy directory (file output)
    print('now copying', str(cwd), file=codecs.open(str(log_path), mode="a"))
    while True:
        # if C drive(default FileStream cache) capacity is less than 20GB,
        # wait for google drive sync and delete cache file.
        disk_usage_gb = shutil.disk_usage('C:/').free / 1024 / 1024 / 1024
        if disk_usage_gb > 20:
            break
        print('wait for GDrive Sync...: free is {:>6.2f} GB'.format(disk_usage_gb))
        # wait for an hour
        time.sleep(3600)


def main():
    ex_list = [
        '.svn',
        'vssver.scc',
        'Shortcut*'
    ]
    # decide path ------------------------------------
    srs = Path("C:/Workspace/Scripts/scripts/vba_macro")
    dst = Path("C:/Workspace/Scripts/scripts/test")
    # ------------------------------------------------
    IgnPtn = CopytreeIgnore(ex_patterns=iter(ex_list), callback=wait_for_gdrive_sync)
    shutil.copytree(srs, dst, ignore=IgnPtn.include_and_exclude,
                    dirs_exist_ok=True)


if __name__ == "__main__":
    main()
