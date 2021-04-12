import shutil
from pathlib import Path
import time
import codecs

from typing import List, Iterable, Callable, Optional, Union


class CopytreeIgnore(object):
    """
    Utility for shutil.copytree(), 'ignore' argument.
    following methods can be used instead of shutil.ignore_patterns().
    """
    DirCall = Union[
        Callable[[str, List[str]], None],
        Callable[[str, List[str]], List[str]]
    ]

    def __init__(self,
                 in_patterns: Iterable[str] = [],
                 ex_patterns: Iterable[str] = [],
                 callbacks: Optional[List[DirCall]] = None):
        """
        in_patterns: include file/directory glob patterns
        ex_patterns: exclude file/directory glob patterns
        callbacks: optional callback function, which is called every recursive search
                per directory.
                callback can return ignore file list called 'optional ignore list'
                if optional ignore lists are set, ignore() method gives them copytree()
        """
        self._set_include_pattern(in_patterns)
        self._set_exclude_pattern(ex_patterns)
        self._set_callback(callbacks)

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

    def _set_callback(self, callbacks: Optional[List[DirCall]]) -> None:
        self.beforeDirCopy = callbacks

    def ignore_exclude(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that doesn't match glob patterns in
        'ex_patterns'.
        """
        opt_ignores: set = self._create_optional_set(directory, files)
        ignores: set = self._create_exclude_set(directory)
        return list(ignores | opt_ignores)

    def ignore_include(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that matches glob patterns in 'in_patterns'.
        """
        opt_ignores: set = self._create_optional_set(directory, files)
        ignores: set = self._create_include_set(directory, files)
        return list(ignores | opt_ignores)

    def ignore(self, directory: str, files: List[str]) -> List[str]:
        """
        callback function for shutil.copytree() 'ignore' argument.
        copytree() only copies that matches glob patterns in 'in_patterns',
        excluding 'ex_patterns'.
        """
        opt_ignores: set = self._create_optional_set(directory, files)
        in_ignores: set = set()
        ex_ignores: set = set()
        if self.in_patterns:
            # if in_patterns is empty, all files and directories become targets
            # so ignore file is none
            in_ignores = self._create_include_set(directory, files)
        ex_ignores = self._create_exclude_set(directory)

        return list(in_ignores | ex_ignores | opt_ignores)

    def _create_optional_set(self, dir_path: str, files: List[str]) -> set:
        """
        create optional ignore set by callback function
        """
        ignores: set = set()
        if self.beforeDirCopy:
            ignore_list: List[str] = []
            for func in self.beforeDirCopy:
                tmp_ignore_list: Optional[List[str]] = func(dir_path, files)
                if tmp_ignore_list:
                    ignore_list += tmp_ignore_list
            ignores = set(ignore_list)
        return ignores

    def _create_exclude_set(self, dir_path: str) -> set:
        # directory path which contains target files
        parent_dir: Path = Path(dir_path)
        ignores: set = set()
        for ptn in self.ex_patterns:
            for file_path in parent_dir.glob(ptn):
                ignores.add(file_path.name)
        return ignores

    def _create_include_set(self, dir_path: str, files: List[str]) -> set:
        includes: set = set()
        parent_dir: Path = Path(dir_path)
        for ptn in self.in_patterns:
            for file_path in parent_dir.glob(ptn):
                includes.add(file_path.name)
        ignores: set = set(files) - includes
        return ignores


def wait_for_gdrive_sync(directory: str, files: List[str]) -> None:
    """
    it is used for FileSream upload, more than C drive capacity.
    """
    # save out.log in home directory
    log_path: Path = Path().home() / 'out.log'
    parent_dir: Path = Path(directory)
    # print next copy directory
    print('now copying', str(parent_dir))
    # print next copy directory (file output)
    print('now copying', str(parent_dir), file=codecs.open(str(log_path), mode="a"))
    while True:
        # if C drive(default FileStream cache) capacity is less than 20GB,
        # wait for google drive sync and delete cache file.
        disk_usage_gb = shutil.disk_usage(parent_dir.home()).free / 1024 / 1024 / 1024
        if disk_usage_gb > 20:
            break
        print('wait for GDrive Sync...: free is {:>6.2f} GB'.format(disk_usage_gb))
        # wait for an hour
        time.sleep(3600)


def ignore_duplicate_dir(directory: str, files: List[str]) -> List[str]:
    """
    ignore duplicate directory, if compressed file exists in same directory.
    """
    ignores: List[str] = []
    # large directory rarely include compressed file,
    # so skip for performance
    if len(files) > 30:
        return ignores
    compressed_file_list: List[str] = [
        '*.7z',
        '*.zip'
    ]
    # to determine whether directory or not, get absolute path
    target_dir: Path = Path(directory).resolve()
    for ptn in compressed_file_list:
        # if compressed file exists, directory whose name is same is added
        # to ignore list
        for cmp_file in target_dir.glob(ptn):
            if cmp_file.stem in files:
                idx: int = files.index(cmp_file.stem)
                tmp_path: Path = target_dir / files[idx]
                if tmp_path.is_dir():
                    ignores.append(files[idx])
            else:
                continue

    return ignores


def main():
    ex_list: List[str] = [
        '.svn',
        'vssver.scc',
        'Shortcut*'
    ]
    in_list: List[str] = [
    ]
    cb_list: List[CopytreeIgnore.DirCall] = [
        wait_for_gdrive_sync,
        ignore_duplicate_dir
    ]
    # decide path ------------------------------------
    srs = Path("vba_macro/")
    dst = Path("test/")
    # ------------------------------------------------
    IgnPtn = CopytreeIgnore(ex_patterns=ex_list,
                            in_patterns=in_list,
                            callbacks=cb_list)
    shutil.copytree(srs, dst, ignore=IgnPtn.ignore,
                    dirs_exist_ok=True)


if __name__ == "__main__":
    main()
