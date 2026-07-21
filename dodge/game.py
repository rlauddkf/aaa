# game.py - Pygame 버전 무한 피하기 게임

import pygame
import random
import json
import os
import sys

# Pygame 초기화
pygame.init()

# 상수
SCREEN_WIDTH = 600
SCREEN_HEIGHT = 800
FPS = 60

# 색상
BLACK = (15, 23, 42)
BLUE = (59, 130, 246)
RED = (239, 68, 68)
WHITE = (229, 231, 235)
GRAY = (165, 180, 252)

# 한글이 보이도록 시스템 폰트 사용 (Windows: 맑은 고딕, mac: AppleGothic)
KOREAN_FONT = "malgungothic,applegothic,applesdgothicneo,nanumgothic,arial"


def make_font(size, bold=False):
    return pygame.font.SysFont(KOREAN_FONT, size, bold=bold)


# 게임 클래스
class Game:
    def __init__(self):
        self.screen = pygame.display.set_mode((SCREEN_WIDTH, SCREEN_HEIGHT))
        pygame.display.set_caption("무한 피하기 게임")
        self.clock = pygame.time.Clock()
        self.font_large = make_font(40, bold=True)
        self.font_medium = make_font(26)
        self.font_small = make_font(18)

        self.running = True
        self.game_started = False
        self.game_active = False
        self.high_score = self.load_high_score()

        self.reset_game()

    def reset_game(self):
        """게임 초기화"""
        self.score = 0
        self.level = 1
        self.time = 0
        self.spawn_rate = 60
        self.obstacle_speed = 3

        self.player = {
            'x': SCREEN_WIDTH // 2 - 20,
            'y': SCREEN_HEIGHT - 60,
            'width': 40,
            'height': 40,
            'speed': 7
        }

        self.obstacles = []
        self.keys_pressed = {}

    def load_high_score(self):
        """하이스코어 파일에서 로드"""
        try:
            if os.path.exists('highscore.json'):
                with open('highscore.json', 'r') as f:
                    data = json.load(f)
                    return data.get('high_score', 0)
        except Exception:
            pass
        return 0

    def save_high_score(self):
        """하이스코어 파일에 저장"""
        try:
            with open('highscore.json', 'w') as f:
                json.dump({'high_score': self.high_score}, f)
        except Exception:
            pass

    def handle_input(self):
        """입력 처리"""
        for event in pygame.event.get():
            if event.type == pygame.QUIT:
                self.running = False

            if event.type == pygame.KEYDOWN:
                if event.key == pygame.K_ESCAPE:  # ESC 키로 종료
                    self.running = False

                if event.key == pygame.K_SPACE or event.key == pygame.K_RETURN:
                    if not self.game_active:
                        self.game_started = True
                        self.game_active = True
                        self.reset_game()

                if event.key == pygame.K_LEFT or event.key == pygame.K_a:
                    self.keys_pressed['left'] = True
                if event.key == pygame.K_RIGHT or event.key == pygame.K_d:
                    self.keys_pressed['right'] = True

            if event.type == pygame.KEYUP:
                if event.key == pygame.K_LEFT or event.key == pygame.K_a:
                    self.keys_pressed['left'] = False
                if event.key == pygame.K_RIGHT or event.key == pygame.K_d:
                    self.keys_pressed['right'] = False

    def update(self):
        """게임 업데이트"""
        if not self.game_active:
            return

        self.time += 1

        # 플레이어 이동
        if self.keys_pressed.get('left'):
            self.player['x'] -= self.player['speed']
        if self.keys_pressed.get('right'):
            self.player['x'] += self.player['speed']

        # 화면 경계 체크
        if self.player['x'] < 0:
            self.player['x'] = 0
        if self.player['x'] + self.player['width'] > SCREEN_WIDTH:
            self.player['x'] = SCREEN_WIDTH - self.player['width']

        # 장애물 생성
        if self.time % self.spawn_rate == 0:
            obstacle_x = random.randint(0, SCREEN_WIDTH - 40)
            self.obstacles.append({
                'x': obstacle_x,
                'y': -40,
                'width': 40,
                'height': 40
            })

        # 장애물 업데이트
        for obstacle in self.obstacles[:]:
            obstacle['y'] += self.obstacle_speed

            # 화면 밖으로 나간 장애물 제거
            if obstacle['y'] > SCREEN_HEIGHT:
                self.obstacles.remove(obstacle)
                self.score += 10

                # 난이도 조정
                if self.score % 100 == 0:
                    self.level = self.score // 100 + 1
                    self.spawn_rate = max(20, 60 - self.level * 5)
                    self.obstacle_speed = 3 + self.level * 0.5
            else:
                # 충돌 체크
                if self.check_collision(self.player, obstacle):
                    self.game_active = False
                    if self.score > self.high_score:
                        self.high_score = self.score
                        self.save_high_score()

    def check_collision(self, rect1, rect2):
        """충돌 감지"""
        return (rect1['x'] < rect2['x'] + rect2['width'] and
                rect1['x'] + rect1['width'] > rect2['x'] and
                rect1['y'] < rect2['y'] + rect2['height'] and
                rect1['y'] + rect1['height'] > rect2['y'])

    def draw(self):
        """화면 그리기"""
        self.screen.fill(BLACK)

        # 플레이어 그리기
        pygame.draw.rect(self.screen, BLUE,
                         (self.player['x'], self.player['y'],
                          self.player['width'], self.player['height']))

        # 장애물 그리기
        for obstacle in self.obstacles:
            pygame.draw.rect(self.screen, RED,
                             (obstacle['x'], obstacle['y'],
                              obstacle['width'], obstacle['height']))

        # UI 그리기 (항상 표시)
        score_text = self.font_medium.render(f"점수: {self.score}", True, WHITE)
        level_text = self.font_medium.render(f"레벨: {self.level}", True, WHITE)
        high_score_text = self.font_small.render(
            f"하이스코어: {self.high_score}", True, GRAY)

        self.screen.blit(score_text, (20, 20))
        self.screen.blit(level_text, (20, 50))
        self.screen.blit(high_score_text, (20, SCREEN_HEIGHT - 30))

        # 게임 오버 화면
        if self.game_started and not self.game_active:
            overlay = pygame.Surface((SCREEN_WIDTH, SCREEN_HEIGHT))
            overlay.set_alpha(200)
            overlay.fill(BLACK)
            self.screen.blit(overlay, (0, 0))

            game_over_text = self.font_large.render("게임 오버!", True, WHITE)
            final_score_text = self.font_medium.render(
                f"최종 점수: {self.score}", True, WHITE)
            restart_text = self.font_small.render(
                "스페이스바로 다시 시작 / ESC로 종료", True, GRAY)

            self.screen.blit(game_over_text,
                             (SCREEN_WIDTH // 2 - game_over_text.get_width() // 2,
                              SCREEN_HEIGHT // 2 - 60))
            self.screen.blit(final_score_text,
                             (SCREEN_WIDTH // 2 - final_score_text.get_width() // 2,
                              SCREEN_HEIGHT // 2))
            self.screen.blit(restart_text,
                             (SCREEN_WIDTH // 2 - restart_text.get_width() // 2,
                              SCREEN_HEIGHT // 2 + 50))

        # 시작 화면
        if not self.game_started:
            overlay = pygame.Surface((SCREEN_WIDTH, SCREEN_HEIGHT))
            overlay.set_alpha(180)
            overlay.fill(BLACK)
            self.screen.blit(overlay, (0, 0))

            title_text = self.font_large.render("무한 피하기", True, WHITE)
            controls_text = self.font_small.render(
                "← → 또는 A / D 키로 이동", True, GRAY)
            instruction_text = self.font_small.render(
                "빨간 상자들을 피해보세요!", True, GRAY)
            start_text = self.font_medium.render("스페이스바로 시작", True, WHITE)
            high_score_display = self.font_small.render(
                f"하이스코어: {self.high_score}", True, GRAY)
            esc_text = self.font_small.render("ESC로 종료", True, GRAY)

            self.screen.blit(title_text,
                             (SCREEN_WIDTH // 2 - title_text.get_width() // 2,
                              SCREEN_HEIGHT // 2 - 100))
            self.screen.blit(controls_text,
                             (SCREEN_WIDTH // 2 - controls_text.get_width() // 2,
                              SCREEN_HEIGHT // 2 - 20))
            self.screen.blit(instruction_text,
                             (SCREEN_WIDTH // 2 - instruction_text.get_width() // 2,
                              SCREEN_HEIGHT // 2 + 10))
            self.screen.blit(start_text,
                             (SCREEN_WIDTH // 2 - start_text.get_width() // 2,
                              SCREEN_HEIGHT // 2 + 80))
            self.screen.blit(high_score_display,
                             (SCREEN_WIDTH // 2 - high_score_display.get_width() // 2,
                              SCREEN_HEIGHT // 2 + 140))
            self.screen.blit(esc_text,
                             (SCREEN_WIDTH // 2 - esc_text.get_width() // 2,
                              SCREEN_HEIGHT - 40))

        pygame.display.flip()

    def run(self):
        """게임 메인 루프"""
        while self.running:
            self.handle_input()
            self.update()
            self.draw()
            self.clock.tick(FPS)

        pygame.quit()
        sys.exit()


if __name__ == "__main__":
    game = Game()
    game.run()
