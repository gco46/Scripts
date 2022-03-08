from pathlib import Path
from typing import List
import argparse


class ToggleSrc():
    COMMENT_SET: dict = {
        ".c": "// ",
        ".h": "// "
    }

    def __init__(self,
                 tgt_path: str,
                 start_word: str = "#pragma asm",
                 end_word: str = "#pragma endasm",
                 code_flag: str = "DISABLED:") -> None:
        # 対象ファイル or ディレクトリのpath
        self.tgt_path = tgt_path
        # 対象行の判定用キーワード
        self.start_word = start_word
        self.end_word = end_word
        # 本クラスの処理による編集箇所を明示するためのコードフラグ
        self.code_flag = code_flag

    def exe_toggle(self) -> bool:
        """
        ソースファイルのトグル実行
        Return:
            bool: 成功 or 失敗
        """
        target = Path(self.tgt_path)

        if target.is_file():
            self._scan_src_file(target)
            return True
        elif target.is_dir():
            for file in target.glob("**/*.*"):
                self._scan_src_file(file)
            return True
        return False

    def _scan_src_file(self, file: Path) -> None:
        """
        ソースファイルを走査し、対象行を変更して上書き保存する
        Args:
            file (Path): 対象のテキストファイル
        """
        # 対象ソースに含まれない場合はスキップ
        extension: str = file.suffix
        if extension not in list(ToggleSrc.COMMENT_SET.keys()):
            return

        result_buf: List[str] = []
        is_tgt_line: bool = False
        update_file: bool = False

        # デコード不能なマルチバイト文字を含むファイルは無視
        # error="ignore"
        with open(file, mode="r", encoding="shift-jis", errors="ignore") as f:
            for line in f.readlines():
                if self.start_word in line:
                    is_tgt_line = True

                if is_tgt_line:
                    # 対象Lineが見つかった時点でファイル更新実施を決定
                    update_file = True
                    line_buf = self._toggle_tgt_line(line, extension)
                else:
                    line_buf = line

                if self.end_word in line:
                    is_tgt_line = False

                result_buf.append(line_buf)

        if update_file:
            with open(file, mode="w", encoding="shift-jis", errors="ignore") as f:
                f.write("".join(result_buf))

    def _toggle_tgt_line(self, line: str, ext: str) -> str:
        """
        対象行をコメントアウト or コメント解除する
        Args:
            line (str): ファイルから抽出した行
            ext (str): 対象ファイルの拡張子
        Return:
            str: コメントアウト or コメント解除した行
        """
        if self.code_flag in line:
            result_line = line.replace(ToggleSrc.COMMENT_SET[ext] + self.code_flag, "")
        else:
            result_line = ToggleSrc.COMMENT_SET[ext] + self.code_flag + line
        return result_line


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("tgt_path", help="target path of file or directory", type=str)
    args = parser.parse_args()

    TglObj = ToggleSrc(args.tgt_path)
    TglObj.exe_toggle()

    # テスト用
    # tglobj = ToggleSrc(str(Path("C:/Workspace/A4_MEB/RV019PP_SRC/trunk/Apli/PJ/")))
    # tglobj.exe_toggle()
