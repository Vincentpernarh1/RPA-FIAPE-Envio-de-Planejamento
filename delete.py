import pyautogui as pag
import time



time.sleep(10)  # Aguarda 10 segundos para o usuário posicionar o mouse
# x,y = pos = pag.position()
# print(f"Posição do mouse: {pos}")

x,y = 557, 426

for _ in range(200):
    pag.click(x, y)
    time.sleep(0.6)