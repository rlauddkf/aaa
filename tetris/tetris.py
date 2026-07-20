"""
테트리스 (Tetris)
- 방향키로 이동/회전, Space로 하드드롭
- ESC를 누르면 프로그램이 즉시 종료됩니다.
"""

import random
import sys

import pygame

# ---------------------------------------------------------------------------
# 설정 (Config)
# ---------------------------------------------------------------------------
CELL = 30                 # 셀 한 칸 픽셀 크기
COLS = 10                 # 보드 가로 칸 수
ROWS = 20                 # 보드 세로 칸 수
SIDEBAR = 6 * CELL        # 우측 정보 패널 너비

BOARD_W = COLS * CELL
BOARD_H = ROWS * CELL
WIDTH = BOARD_W + SIDEBAR
HEIGHT = BOARD_H

FPS = 60

# 색상
BLACK = (15, 15, 20)
GRID = (40, 40, 50)
WHITE = (235, 235, 235)
GRAY = (120, 120, 130)

# 테트로미노 모양 정의 (4x4 회전 기준 좌표)
SHAPES = {
    "I": [(0, 1), (1, 1), (2, 1), (3, 1)],
    "O": [(1, 0), (2, 0), (1, 1), (2, 1)],
    "T": [(1, 0), (0, 1), (1, 1), (2, 1)],
    "S": [(1, 0), (2, 0), (0, 1), (1, 1)],
    "Z": [(0, 0), (1, 0), (1, 1), (2, 1)],
    "J": [(0, 0), (0, 1), (1, 1), (2, 1)],
    "L": [(2, 0), (0, 1), (1, 1), (2, 1)],
}

COLORS = {
    "I": (0, 240, 240),
    "O": (240, 240, 0),
    "T": (160, 0, 240),
    "S": (0, 240, 0),
    "Z": (240, 0, 0),
    "J": (0, 0, 240),
    "L": (240, 160, 0),
}


class Piece:
    """현재 떨어지는 조각."""

    def __init__(self, kind):
        self.kind = kind
        self.color = COLORS[kind]
        self.cells = list(SHAPES[kind])
        # 시작 위치: 가운데 상단
        self.x = COLS // 2 - 2
        self.y = 0

    def blocks(self, cells=None, x=None, y=None):
        """절대 좌표 리스트를 반환."""
        cells = self.cells if cells is None else cells
        x = self.x if x is None else x
        y = self.y if y is None else y
        return [(x + cx, y + cy) for cx, cy in cells]

    def rotated(self):
        """시계방향 90도 회전한 셀 좌표(로컬)를 반환. O 조각은 회전 안 함."""
        if self.kind == "O":
            return list(self.cells)
        # 4x4 기준 회전: (x, y) -> (3 - y, x) 후 최소값으로 정규화
        rot = [(3 - cy, cx) for cx, cy in self.cells]
        min_x = min(c[0] for c in rot)
        min_y = min(c[1] for c in rot)
        return [(cx - min_x, cy - min_y) for cx, cy in rot]


class Tetris:
    def __init__(self):
        self.board = [[None] * COLS for _ in range(ROWS)]
        self.bag = []
        self.current = self._next_piece()
        self.next_piece = self._next_piece()
        self.score = 0
        self.lines = 0
        self.level = 1
        self.game_over = False
        self.fall_timer = 0.0

    # -- 조각 생성 (7-bag 방식) ------------------------------------------
    def _next_piece(self):
        if not self.bag:
            self.bag = list(SHAPES.keys())
            random.shuffle(self.bag)
        return Piece(self.bag.pop())

    @property
    def fall_speed(self):
        """레벨이 오를수록 빨라짐 (초 단위)."""
        return max(0.05, 0.5 - (self.level - 1) * 0.04)

    # -- 충돌 판정 --------------------------------------------------------
    def _valid(self, cells, x, y):
        for bx, by in self.current.blocks(cells, x, y):
            if bx < 0 or bx >= COLS or by >= ROWS:
                return False
            if by >= 0 and self.board[by][bx] is not None:
                return False
        return True

    # -- 조작 ------------------------------------------------------------
    def move(self, dx):
        if self._valid(self.current.cells, self.current.x + dx, self.current.y):
            self.current.x += dx

    def rotate(self):
        rot = self.current.rotated()
        # 벽 킥: 좌우로 조금 밀어보며 회전 시도
        for kick in (0, -1, 1, -2, 2):
            if self._valid(rot, self.current.x + kick, self.current.y):
                self.current.cells = rot
                self.current.x += kick
                return

    def soft_drop(self):
        if self._valid(self.current.cells, self.current.x, self.current.y + 1):
            self.current.y += 1
            self.score += 1
            return True
        self._lock()
        return False

    def hard_drop(self):
        while self._valid(self.current.cells, self.current.x, self.current.y + 1):
            self.current.y += 1
            self.score += 2
        self._lock()

    # -- 고정 & 라인 제거 -------------------------------------------------
    def _lock(self):
        for bx, by in self.current.blocks():
            if by < 0:
                self.game_over = True
                return
            self.board[by][bx] = self.current.color
        self._clear_lines()
        self.current = self.next_piece
        self.next_piece = self._next_piece()
        if not self._valid(self.current.cells, self.current.x, self.current.y):
            self.game_over = True

    def _clear_lines(self):
        remaining = [row for row in self.board if any(c is None for c in row)]
        cleared = ROWS - len(remaining)
        if cleared:
            for _ in range(cleared):
                remaining.insert(0, [None] * COLS)
            self.board = remaining
            self.lines += cleared
            # 점수: 한 번에 여러 줄 지울수록 보너스
            self.score += (0, 100, 300, 500, 800)[cleared] * self.level
            self.level = 1 + self.lines // 10

    # -- 시간 진행 --------------------------------------------------------
    def update(self, dt):
        if self.game_over:
            return
        self.fall_timer += dt
        if self.fall_timer >= self.fall_speed:
            self.fall_timer = 0.0
            self.soft_drop()

    def reset(self):
        self.__init__()


# ---------------------------------------------------------------------------
# 렌더링 (Rendering)
# ---------------------------------------------------------------------------
def draw_cell(surface, x, y, color):
    rect = pygame.Rect(x * CELL, y * CELL, CELL, CELL)
    pygame.draw.rect(surface, color, rect)
    pygame.draw.rect(surface, BLACK, rect, 2)


def draw(surface, game, font, big_font):
    surface.fill(BLACK)

    # 격자
    for x in range(COLS):
        for y in range(ROWS):
            rect = pygame.Rect(x * CELL, y * CELL, CELL, CELL)
            pygame.draw.rect(surface, GRID, rect, 1)

    # 고정된 블록
    for y in range(ROWS):
        for x in range(COLS):
            color = game.board[y][x]
            if color:
                draw_cell(surface, x, y, color)

    # 현재 조각
    if not game.game_over:
        for bx, by in game.current.blocks():
            if by >= 0:
                draw_cell(surface, bx, by, game.current.color)

    # 사이드바 구분선
    pygame.draw.line(surface, GRAY, (BOARD_W, 0), (BOARD_W, HEIGHT), 2)

    sx = BOARD_W + 20
    surface.blit(font.render("SCORE", True, GRAY), (sx, 20))
    surface.blit(big_font.render(str(game.score), True, WHITE), (sx, 45))
    surface.blit(font.render("LINES", True, GRAY), (sx, 100))
    surface.blit(big_font.render(str(game.lines), True, WHITE), (sx, 125))
    surface.blit(font.render("LEVEL", True, GRAY), (sx, 180))
    surface.blit(big_font.render(str(game.level), True, WHITE), (sx, 205))

    # 다음 조각 미리보기
    surface.blit(font.render("NEXT", True, GRAY), (sx, 270))
    for cx, cy in game.next_piece.cells:
        rect = pygame.Rect(sx + cx * CELL, 300 + cy * CELL, CELL, CELL)
        pygame.draw.rect(surface, game.next_piece.color, rect)
        pygame.draw.rect(surface, BLACK, rect, 2)

    # 조작 안내
    tips = ["← → : 이동", "↑ : 회전", "↓ : 소프트드롭",
            "Space : 하드드롭", "ESC : 종료"]
    for i, tip in enumerate(tips):
        surface.blit(font.render(tip, True, GRAY), (sx, 440 + i * 24))

    # 게임 오버
    if game.game_over:
        overlay = pygame.Surface((BOARD_W, HEIGHT))
        overlay.set_alpha(200)
        overlay.fill(BLACK)
        surface.blit(overlay, (0, 0))
        msg = big_font.render("GAME OVER", True, (240, 80, 80))
        surface.blit(msg, (BOARD_W // 2 - msg.get_width() // 2, HEIGHT // 2 - 40))
        sub = font.render("R: 다시 시작   ESC: 종료", True, WHITE)
        surface.blit(sub, (BOARD_W // 2 - sub.get_width() // 2, HEIGHT // 2 + 10))


# ---------------------------------------------------------------------------
# 메인 루프
# ---------------------------------------------------------------------------
def main():
    pygame.init()
    screen = pygame.display.set_mode((WIDTH, HEIGHT))
    pygame.display.set_caption("Tetris")
    clock = pygame.time.Clock()
    font = pygame.font.SysFont("malgungothic,arial", 20)
    big_font = pygame.font.SysFont("malgungothic,arial", 32, bold=True)

    game = Tetris()

    while True:
        dt = clock.tick(FPS) / 1000.0

        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                pygame.quit()
                sys.exit()
            if event.type == pygame.KEYDOWN:
                # ESC -> 프로그램 즉시 종료
                if event.key == pygame.K_ESCAPE:
                    pygame.quit()
                    sys.exit()
                if game.game_over:
                    if event.key == pygame.K_r:
                        game.reset()
                    continue
                if event.key == pygame.K_LEFT:
                    game.move(-1)
                elif event.key == pygame.K_RIGHT:
                    game.move(1)
                elif event.key == pygame.K_UP:
                    game.rotate()
                elif event.key == pygame.K_DOWN:
                    game.soft_drop()
                elif event.key == pygame.K_SPACE:
                    game.hard_drop()

        game.update(dt)
        draw(screen, game, font, big_font)
        pygame.display.flip()


if __name__ == "__main__":
    main()
