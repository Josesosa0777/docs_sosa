import cv2
import numpy as np
import pytesseract
import pandas as pd
from pytesseract import Output

# Carga el video
cap = cv2.VideoCapture('velocidad.mp4')

# Inicializa variables
fps = cap.get(cv2.CAP_PROP_FPS)  # fotogramas por segundo
frame_count = 0
data = []  # Lista para almacenar los resultados

while cap.isOpened():
    ret, frame = cap.read()
    if not ret:
        break

    # Preprocesamiento del fotograma para mejorar la detección de texto
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    edged = cv2.Canny(blurred, 50, 200)

    # Detección de contornos
    contours, _ = cv2.findContours(edged.copy(), cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if w > 20 and h > 20:  # Filtrar contornos pequeños
            roi = frame[y:y+h, x:x+w]  # Región de interés (posible señal de velocidad)
            
            # Aplicar OCR para detectar el texto
            text = pytesseract.image_to_string(roi, config='--psm 7', output_type=Output.STRING)
            
            # Filtrar los resultados de OCR para encontrar velocidades
            text = ''.join(filter(str.isdigit, text))
            if text.isdigit():
                speed = int(text)
                time = frame_count / fps
                data.append({"Velocidad (km/h)": speed, "Tiempo (s)": time})
    
    frame_count += 1

# Liberar el video
cap.release()

# Crear un DataFrame de Pandas
df = pd.DataFrame(data)

# Mostrar la tabla
print(df)

# Guardar la tabla en un archivo CSV (opcional)
df.to_csv('detected_speeds.csv', index=False)

# Imprimir la velocidad máxima y el tiempo en que ocurre
if not df.empty:
    max_row = df.loc[df['Velocidad (km/h)'].idxmax()]
    print(f"\nVelocidad máxima detectada: {max_row['Velocidad (km/h)']} km/h")
    print(f"Tiempo en el que ocurre: {max_row['Tiempo (s)']} segundos")
else:
    print("No se detectaron señales de velocidad.")
