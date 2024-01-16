import pygame
import sys
import random

pygame.init()

# Các biến cơ bản
screen_width, screen_height = 600, 400
gravity = 0.25
bird_velocity = 0
bird_flap = -5
bird_y = screen_height // 2
bird_size = 20
pipe_width = 50
pipe_height = 300
pipe_gap = 100
pipe_distance = 200
pipes = []

# Màu sắc
white = (255, 255, 255)
green = (0, 255, 0)

# Khởi tạo màn hình
screen = pygame.display.set_mode((screen_width, screen_height))
pygame.display.set_caption("Flappy Bird")

# Hàm vẽ con chim
def draw_bird():
    pygame.draw.circle(screen, white, (50, bird_y), bird_size)

# Hàm vẽ ống
def draw_pipe(pipe_x, top_pipe_height):
    pygame.draw.rect(screen, green, (pipe_x, 0, pipe_width, top_pipe_height))
    pygame.draw.rect(screen, green, (pipe_x, top_pipe_height + pipe_gap, pipe_width, screen_height))

# Hàm kiểm tra va chạm
def check_collision(pipe_x, top_pipe_height):
    if bird_y < top_pipe_height or bird_y > top_pipe_height + pipe_gap:
        if 50 < pipe_x < 50 + pipe_width:
            return True
    return False

# Vòng lặp chính
clock = pygame.time.Clock()
running = True
while running:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False
        elif event.type == pygame.KEYDOWN:
            if event.key == pygame.K_SPACE:
                bird_velocity = bird_flap

    # Cập nhật vị trí của con chim
    bird_velocity += gravity
    bird_y += bird_velocity

    # Tạo ống mới và kiểm tra va chạm
    if len(pipes) == 0 or pipes[-1][0] < screen_width - pipe_distance:
        top_pipe_height = random.randint(50, screen_height - pipe_gap - 50)
        pipes.append((screen_width, top_pipe_height))

    # Di chuyển và vẽ ống
    pipes = [(pipe[0] - 2, pipe[1]) for pipe in pipes]
    pipes = [pipe for pipe in pipes if pipe[0] > -pipe_width]
    for pipe in pipes:
        draw_pipe(pipe[0], pipe[1])
        if check_collision(pipe[0], pipe[1]):
            print("Game Over!")
            running = False

    # Vẽ con chim
    draw_bird()

    # Cập nhật màn hình
    pygame.display.flip()

    # Đặt tốc độ khung hình
    clock.tick(60)

    # Xóa màn hình
    screen.fill((0, 0, 0))

pygame.quit()
sys.exit()
 