"""
테트리스 (Tetris) - 엑셀 스프레드시트 위장 테마
- 방향키로 이동/회전(↑), Space로 하드드롭
- ←/→/↓ 키는 누르고 있으면 연속 이동
- 낙하 지점은 '선택된 셀' 테두리처럼 표시
- ESC를 누르면 프로그램이 즉시 종료됩니다.
"""

import random
import sys

import pygame

# ---------------------------------------------------------------------------
# 게임판 설정 (Game config)
# ---------------------------------------------------------------------------
COLS = 10                 # 보드 가로 칸 수
ROWS = 20                 # 보드 세로 칸 수

FPS = 60

# 키를 누르고 있을 때 연속 이동 설정 (초 단위)
DAS_DELAY = 0.16   # 처음 누른 뒤 자동 반복이 시작되기까지의 지연
ARR_RATE = 0.04    # 자동 반복 시 한 칸 이동 간격
SOFT_DROP_RATE = 0.03  # ↓ 를 누르고 있을 때 낙하 간격

# ---------------------------------------------------------------------------
# 엑셀 스타일 화면 레이아웃 (Excel-look layout)
# ---------------------------------------------------------------------------
CELL = 28                 # 셀 한 칸 픽셀 크기
GUTTER = 34               # 행 번호(왼쪽) 칸 너비
COLHDR_H = 20             # 열 머리글(A,B,C...) 높이
COLS_VIEW = 16            # 화면에 보이는 열 수 (A~P)
ROWS_VIEW = 22            # 화면에 보이는 행 수

TABS_H = 26               # 리본 탭 줄 높이
RIBBON_H = 78             # 리본 본문 높이
FORMULA_H = 26            # 이름상자 + 수식 입력줄 높이
CHROME_H = TABS_H + RIBBON_H + FORMULA_H  # 격자 위쪽 엑셀 UI 전체 높이

GRID_X0 = GUTTER                    # 셀 격자가 시작되는 x
GRID_Y0 = CHROME_H + COLHDR_H       # 셀 격자가 시작되는 y

WIDTH = GUTTER + COLS_VIEW * CELL
HEIGHT = GRID_Y0 + ROWS_VIEW * CELL

INFO_COL = 11             # 점수/라인 등 정보를 표시할 시작 열 (L열)

# 색상 (엑셀 2007 느낌)
TABBAR_BG = (183, 208, 234)
TAB_ACTIVE_BG = (245, 248, 252)
TAB_TEXT = (40, 45, 60)
RIBBON_BG = (239, 243, 248)
BTN_BG = (250, 251, 253)
GROUP_LABEL = (110, 120, 135)
SEP = (200, 208, 220)
HDR_BG = (232, 236, 242)
HDR_TEXT = (90, 100, 115)
HDR_SEL_BG = (255, 221, 148)
HDR_SEL_TEXT = (120, 70, 0)
GRIDLINE = (214, 220, 230)
CELL_WHITE = (255, 255, 255)
BLOCK = (38, 38, 38)          # 블록 = 검게 칠해진 셀
GHOST_BORDER = (55, 110, 70)  # 낙하 지점 = 선택된 셀 테두리처럼
PAGEBREAK = (150, 160, 175)
INK = (25, 25, 30)

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

    def ghost_y(self):
        """현재 조각이 그대로 떨어졌을 때 놓일 y 좌표를 반환."""
        y = self.current.y
        while self._valid(self.current.cells, self.current.x, y + 1):
            y += 1
        return y

    def reset(self):
        self.__init__()


# ---------------------------------------------------------------------------
# 렌더링 (Rendering) - 엑셀 스프레드시트처럼 보이게
# ---------------------------------------------------------------------------
def col_letter(c):
    """0 -> A, 1 -> B, ... 엑셀 열 문자."""
    s = ""
    c += 1
    while c > 0:
        c, rem = divmod(c - 1, 26)
        s = chr(65 + rem) + s
    return s


def cell_rect(c, r):
    """격자 상의 (열 c, 행 r) 셀 픽셀 사각형."""
    return pygame.Rect(GRID_X0 + c * CELL, GRID_Y0 + r * CELL, CELL, CELL)


def fill_cell(surface, c, r):
    """셀을 검게 칠함 (블록 표현)."""
    if 0 <= c < COLS_VIEW and 0 <= r < ROWS_VIEW:
        pygame.draw.rect(surface, BLOCK, cell_rect(c, r))


def active_cell(game):
    """현재 조각의 좌상단 셀 = 엑셀의 '선택된 셀'로 취급."""
    if game.game_over:
        return 0, 0
    blocks = game.current.blocks()
    c = min(b[0] for b in blocks)
    r = max(min(b[1] for b in blocks), 0)
    return c, r


def draw_ribbon(surface, fonts):
    # 리본 탭 줄
    pygame.draw.rect(surface, TABBAR_BG, (0, 0, WIDTH, TABS_H))
    tabs = ["Home", "Insert", "Page Layout", "Formulas",
            "Data", "Review", "View"]
    active = "View"
    x = 6
    for name in tabs:
        surf = fonts["tab"].render(name, True, TAB_TEXT)
        w = surf.get_width() + 14
        if name == active:
            pygame.draw.rect(surface, TAB_ACTIVE_BG, (x, 3, w, TABS_H - 3))
            pygame.draw.line(surface, SEP, (x, 3), (x, TABS_H - 1))
            pygame.draw.line(surface, SEP, (x + w, 3), (x + w, TABS_H - 1))
        surface.blit(surf, (x + 7, (TABS_H - surf.get_height()) // 2))
        x += w + 3

    # 리본 본문
    pygame.draw.rect(surface, RIBBON_BG, (0, TABS_H, WIDTH, RIBBON_H))
    pygame.draw.line(surface, SEP, (0, TABS_H + RIBBON_H - 1),
                     (WIDTH, TABS_H + RIBBON_H - 1))

    # View 탭의 그룹들 (버튼은 형태만 흉내)
    groups = [("Workbook Views", 150), ("Show/Hide", 118), ("Zoom", 112)]
    gx = 4
    ry = TABS_H + 6
    for name, gw in groups:
        for i in range(3):
            bx = gx + 8 + i * (min(gw, 130) // 3)
            pygame.draw.rect(surface, BTN_BG, (bx, ry, 30, 42))
            pygame.draw.rect(surface, SEP, (bx, ry, 30, 42), 1)
        lbl = fonts["group"].render(name, True, GROUP_LABEL)
        surface.blit(lbl, (gx + (gw - lbl.get_width()) // 2,
                           TABS_H + RIBBON_H - 15))
        pygame.draw.line(surface, SEP, (gx + gw, ry),
                         (gx + gw, TABS_H + RIBBON_H - 16))
        gx += gw + 2


def draw_formula_bar(surface, game, fonts):
    y = TABS_H + RIBBON_H
    pygame.draw.rect(surface, CELL_WHITE, (0, y, WIDTH, FORMULA_H))
    # 이름 상자 (선택된 셀 주소 표시)
    nb_w = GUTTER + CELL
    pygame.draw.rect(surface, CELL_WHITE, (2, y + 3, nb_w, FORMULA_H - 6))
    pygame.draw.rect(surface, SEP, (2, y + 3, nb_w, FORMULA_H - 6), 1)
    c, r = active_cell(game)
    ref = fonts["cell"].render(f"{col_letter(c)}{r + 1}", True, INK)
    surface.blit(ref, (8, y + (FORMULA_H - ref.get_height()) // 2))
    # fx 기호
    fx = fonts["fx"].render("fx", True, (90, 120, 180))
    surface.blit(fx, (nb_w + 12, y + (FORMULA_H - fx.get_height()) // 2))
    # 수식 입력줄
    fx0 = nb_w + 34
    pygame.draw.rect(surface, SEP, (fx0, y + 3, WIDTH - fx0 - 4, FORMULA_H - 6), 1)
    pygame.draw.line(surface, SEP, (0, y + FORMULA_H - 1), (WIDTH, y + FORMULA_H - 1))


def draw_headers(surface, game, fonts):
    ac, ar = active_cell(game)
    # 좌상단 코너
    pygame.draw.rect(surface, HDR_BG, (0, CHROME_H, GUTTER, COLHDR_H))
    pygame.draw.rect(surface, GRIDLINE, (0, CHROME_H, GUTTER, COLHDR_H), 1)
    # 열 머리글 (A, B, C ...)
    for c in range(COLS_VIEW):
        x = GRID_X0 + c * CELL
        sel = (c == ac)
        pygame.draw.rect(surface, HDR_SEL_BG if sel else HDR_BG,
                         (x, CHROME_H, CELL, COLHDR_H))
        pygame.draw.rect(surface, GRIDLINE, (x, CHROME_H, CELL, COLHDR_H), 1)
        t = fonts["hdr"].render(col_letter(c), True,
                                HDR_SEL_TEXT if sel else HDR_TEXT)
        surface.blit(t, (x + (CELL - t.get_width()) // 2,
                         CHROME_H + (COLHDR_H - t.get_height()) // 2))
    # 행 번호 (1, 2, 3 ...)
    for r in range(ROWS_VIEW):
        y = GRID_Y0 + r * CELL
        sel = (r == ar)
        pygame.draw.rect(surface, HDR_SEL_BG if sel else HDR_BG,
                         (0, y, GUTTER, CELL))
        pygame.draw.rect(surface, GRIDLINE, (0, y, GUTTER, CELL), 1)
        t = fonts["hdr"].render(str(r + 1), True,
                                HDR_SEL_TEXT if sel else HDR_TEXT)
        surface.blit(t, (GUTTER - 5 - t.get_width(),
                         y + (CELL - t.get_height()) // 2))


def draw_grid(surface):
    grid_w = COLS_VIEW * CELL
    grid_h = ROWS_VIEW * CELL
    pygame.draw.rect(surface, CELL_WHITE, (GRID_X0, GRID_Y0, grid_w, grid_h))
    for c in range(COLS_VIEW + 1):
        x = GRID_X0 + c * CELL
        pygame.draw.line(surface, GRIDLINE, (x, GRID_Y0), (x, GRID_Y0 + grid_h))
    for r in range(ROWS_VIEW + 1):
        y = GRID_Y0 + r * CELL
        pygame.draw.line(surface, GRIDLINE, (GRID_X0, y), (GRID_X0 + grid_w, y))
    # 게임판 오른쪽 경계 = 페이지 나누기 선처럼
    xb = GRID_X0 + COLS * CELL
    pygame.draw.line(surface, PAGEBREAK, (xb, GRID_Y0), (xb, GRID_Y0 + grid_h), 2)


def draw_blocks(surface, game):
    # 고정된 블록
    for r in range(ROWS):
        for c in range(COLS):
            if game.board[r][c]:
                fill_cell(surface, c, r)
    if game.game_over:
        return
    # 고스트(낙하 지점) = 선택된 셀 테두리처럼
    gy = game.ghost_y()
    for bx, by in game.current.blocks(y=gy):
        if 0 <= by < ROWS_VIEW:
            pygame.draw.rect(surface, GHOST_BORDER, cell_rect(bx, by), 2)
    # 현재 조각
    for bx, by in game.current.blocks():
        if by >= 0:
            fill_cell(surface, bx, by)


def draw_info(surface, game, fonts):
    def put(c, r, text, color=INK, font=None):
        font = font or fonts["cell"]
        rect = cell_rect(c, r)
        t = font.render(text, True, color)
        surface.blit(t, (rect.x + 3, rect.y + (CELL - t.get_height()) // 2))

    col = INFO_COL
    put(col, 0, "Score:")
    put(col + 2, 0, str(game.score))
    put(col, 2, "Lines:")
    put(col + 2, 2, str(game.lines))
    put(col, 4, "Level:")
    put(col + 2, 4, str(game.level))
    put(col, 6, "Next:")
    for cx, cy in game.next_piece.cells:
        fill_cell(surface, col + cx, 7 + cy)

    tips = ["[ Controls ]", "← →  move (hold)", "↑  rotate",
            "↓  soft drop (hold)", "Space  hard drop",
            "R  restart", "Esc  quit"]
    for i, tip in enumerate(tips):
        put(col, 12 + i, tip, GROUP_LABEL, fonts["hdr"])


def draw_gameover(surface, fonts):
    bw, bh = 260, 120
    bx = GRID_X0 + (COLS * CELL - bw) // 2
    bx = max(bx, GRID_X0 + 4)
    by = GRID_Y0 + (ROWS_VIEW * CELL - bh) // 2
    pygame.draw.rect(surface, (250, 250, 252), (bx, by, bw, bh))
    pygame.draw.rect(surface, (120, 130, 145), (bx, by, bw, bh), 1)
    pygame.draw.rect(surface, TABBAR_BG, (bx, by, bw, 22))
    title = fonts["hdr"].render("Microsoft Excel", True, (25, 25, 45))
    surface.blit(title, (bx + 8, by + 4))
    msg = fonts["cell"].render("GAME OVER", True, (190, 40, 40))
    surface.blit(msg, (bx + (bw - msg.get_width()) // 2, by + 42))
    sub = fonts["kr"].render("R: 다시 시작    Esc: 종료", True, (50, 50, 60))
    surface.blit(sub, (bx + (bw - sub.get_width()) // 2, by + 78))


def draw(surface, game, fonts):
    surface.fill(CELL_WHITE)
    draw_grid(surface)
    draw_blocks(surface, game)
    draw_info(surface, game, fonts)
    draw_headers(surface, game, fonts)
    draw_ribbon(surface, fonts)
    draw_formula_bar(surface, game, fonts)
    if game.game_over:
        draw_gameover(surface, fonts)


# ---------------------------------------------------------------------------
# 메인 루프
# ---------------------------------------------------------------------------
def main():
    pygame.init()
    screen = pygame.display.set_mode((WIDTH, HEIGHT))
    pygame.display.set_caption("Book1 - Microsoft Excel")
    clock = pygame.time.Clock()
    fonts = {
        "tab": pygame.font.SysFont("segoeui,arial", 14),
        "group": pygame.font.SysFont("segoeui,arial", 10),
        "hdr": pygame.font.SysFont("arial", 12),
        "cell": pygame.font.SysFont("arial", 13),
        "fx": pygame.font.SysFont("arial", 13, italic=True),
        "kr": pygame.font.SysFont("malgungothic,gulim,nanumgothic,arial", 13),
    }

    game = Tetris()

    # 키를 누르고 있을 때의 연속 이동(DAS/ARR) 상태.
    # 값이 None 이면 "떼어진 상태", 실수면 "누른 뒤 경과 시간".
    hold = {pygame.K_LEFT: None, pygame.K_RIGHT: None, pygame.K_DOWN: None}
    charged = {k: False for k in hold}

    def do_action(key):
        if key == pygame.K_LEFT:
            game.move(-1)
        elif key == pygame.K_RIGHT:
            game.move(1)
        elif key == pygame.K_DOWN:
            game.soft_drop()

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
                # 한 번만 반응하는 키
                if event.key == pygame.K_UP:
                    game.rotate()
                elif event.key == pygame.K_SPACE:
                    game.hard_drop()

        # 누르고 있는 동안 연속 이동 처리 (←, →, ↓)
        if game.game_over:
            for k in hold:
                hold[k] = None
                charged[k] = False
        else:
            pressed = pygame.key.get_pressed()
            for key in hold:
                repeat = ARR_RATE if key != pygame.K_DOWN else SOFT_DROP_RATE
                if pressed[key]:
                    if hold[key] is None:
                        # 방금 눌림 -> 즉시 한 칸 이동 후 DAS 대기 시작
                        do_action(key)
                        hold[key] = 0.0
                        charged[key] = False
                    else:
                        hold[key] += dt
                        if not charged[key]:
                            if hold[key] >= DAS_DELAY:
                                charged[key] = True
                                hold[key] = 0.0
                                do_action(key)
                        elif hold[key] >= repeat:
                            hold[key] = 0.0
                            do_action(key)
                else:
                    hold[key] = None
                    charged[key] = False

        game.update(dt)
        draw(screen, game, fonts)
        pygame.display.flip()


if __name__ == "__main__":
    main()
