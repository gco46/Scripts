import shutil
from pathlib import Path


class CopytreeIgnore(object):
    """
    Utility for shutil.copytree(), 'ignore' argument.
    can be used instead of shutil.ignore_patterns().
    """

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
        result = []
        cwd = Path(directory)
        for ptn in self.ex_patterns:
            for file_path in cwd.glob(ptn):
                result.append(file_path.name)
        return result

    def include(self, directory, files):
        result = []
        cwd = Path(directory)
        for ptn in self.in_patterns:
            for file_path in cwd.glob(ptn):
                result.append(file_path.name)
        return result

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
        'vssver.scc'
    ]
    srs = Path("C:/Workspace/CommonTool/")
    tgt = Path("C:/Workspace/test/")
    IgnPtn = CopytreeIgnore(ignore_list)
    shutil.copytree(srs, tgt, ignore=IgnPtn.exclude)


if __name__ == "__main__":
    main()
