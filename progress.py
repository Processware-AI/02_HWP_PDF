"""프로그레스 바 유틸리티 (3줄 표시)"""

import sys
import time


class ProgressBar:
    """3줄 프로그레스 바.

    1줄: 현재 파일명
    2줄: 진행 정보 (퍼센트, 완료/전체, 경과시간)
    3줄: 프로그레스 바 (100칸)
    """

    def __init__(self, total: int):
        self.total = total
        self.current = 0
        self.start_time = time.time()
        self._filename = ""
        self._printed = False

    def update(self, filename: str):
        self.current += 1
        self._filename = filename
        self._draw()

    def _draw(self):
        # 이전 출력 지우기 (3줄 위로 이동)
        if self._printed:
            sys.stdout.write("\033[3A\033[J")

        pct = int(self.current / self.total * 100) if self.total else 0
        elapsed = time.time() - self.start_time
        filled = int(self.current / self.total * 100) if self.total else 0

        bar = "#" * filled + " " * (100 - filled)

        sys.stdout.write(f"  {self._filename}\n")
        sys.stdout.write(f"  {pct:3d}% | {self.current}/{self.total} 완료 | {elapsed:.1f}초 경과\n")
        sys.stdout.write(f"  [{bar}]\n")
        sys.stdout.flush()
        self._printed = True

    def close(self):
        pass
