import pyautogui

img = pyautogui.locateOnScreen("imagens/gerandoRelatorio.png", confidence=0.9, grayscale=True)
print("Imagem encontrada" if img else "Imagem NÃO encontrada")
