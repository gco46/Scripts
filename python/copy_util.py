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

    def ignore_exclude(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that doesn't match glob patterns in
        'ex_patterns'.
        """
        if self.callback:
            self.callback(directory, files)

        ignores: set = self._create_exclude_set(directory)
        return list(ignores)

    def ignore_include(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that matches glob patterns in 'in_patterns'.
        """
        ignores: set = self._create_include_set(directory, files)
        return list(ignores)

    def ignore(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        in_patterns and ex_patterns are used to create ignore list.
        """
        in_ignores: set = set()
        ex_ignores: set = set()
        if self.in_patterns:
            # if in_patterns is empty, all files and directories become target
            # so ignore file is none
            in_ignores = self._create_include_set(directory, files)
        ex_ignores = self._create_exclude_set(directory)

        return list(in_ignores | ex_ignores)

    def _create_exclude_set(self, dir_path: str) -> set:
        cwd: Path = Path(dir_path)
        ignores: set = set()
        for ptn in self.ex_patterns:
            for file_path in cwd.glob(ptn):
                ignores.add(file_path.name)
        return ignores

    def _create_include_set(self, dir_path: str, files: List[str]) -> set:
        includes: set = set()
        cwd: Path = Path(dir_path)
        for ptn in self.in_patterns:
            for file_path in cwd.glob(ptn):
                includes.add(file_path.name)
        ignores: set = set(files) - includes
        return ignores


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
    in_list = [
        '*.bas'
    ]
    # decide path ------------------------------------
    srs = Path("C:/Workspace/Scripts/scripts/vba_macro")
    dst = Path("C:/Workspace/Scripts/scripts/test")
    # ------------------------------------------------
    IgnPtn = CopytreeIgnore(ex_patterns=ex_list,
                            in_patterns=in_list,
                            callback=wait_for_gdrive_sync)
    shutil.copytree(srs, dst, ignore=IgnPtn.ignore,
                    dirs_exist_ok=True)


if __name__ == "__main__":
    main()
