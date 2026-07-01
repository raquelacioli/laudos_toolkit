#!/usr/bin/env python3
"""
scan_documento.py
------------------
Recorta a foto de um documento (removendo o fundo/mesa), corrige a
perspectiva (como um scanner) e melhora nitidez/contraste, deixando a
imagem pronta para leitura em um laudo pericial.

Uso:
    python3 scan_documento.py entrada.jpg saida.jpg
"""

import sys
import cv2
import numpy as np


def order_points(pts):
    rect = np.zeros((4, 2), dtype="float32")
    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]      # top-left
    rect[2] = pts[np.argmax(s)]      # bottom-right
    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]   # top-right
    rect[3] = pts[np.argmax(diff)]   # bottom-left
    return rect


def four_point_transform(image, pts):
    rect = order_points(pts)
    (tl, tr, br, bl) = rect

    widthA = np.linalg.norm(br - bl)
    widthB = np.linalg.norm(tr - tl)
    maxWidth = max(int(widthA), int(widthB))

    heightA = np.linalg.norm(tr - br)
    heightB = np.linalg.norm(tl - bl)
    maxHeight = max(int(heightA), int(heightB))

    dst = np.array([
        [0, 0],
        [maxWidth - 1, 0],
        [maxWidth - 1, maxHeight - 1],
        [0, maxHeight - 1]], dtype="float32")

    M = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(image, M, (maxWidth, maxHeight))
    return warped


def find_document_contour(image):
    """Usa GrabCut para segmentar o documento do fundo (mesa/madeira) e
    retorna os 4 cantos do retangulo minimo que envolve o documento
    encontrado, ja alinhados para correcao de perspectiva/rotacao."""
    h, w = image.shape[:2]
    scale = 800.0 / max(h, w)
    small = cv2.resize(image, (int(w * scale), int(h * scale)))
    sh, sw = small.shape[:2]

    mask = np.zeros((sh, sw), np.uint8)
    # retangulo inicial: assume que o documento ocupa a regiao central,
    # com uma margem de fundo (mesa) nas bordas
    margin_x, margin_y = int(sw * 0.06), int(sh * 0.06)
    rect = (margin_x, margin_y, sw - 2 * margin_x, sh - 2 * margin_y)

    bgdModel = np.zeros((1, 65), np.float64)
    fgdModel = np.zeros((1, 65), np.float64)
    try:
        cv2.grabCut(small, mask, rect, bgdModel, fgdModel, 5, cv2.GC_INIT_WITH_RECT)
    except cv2.error:
        return None

    mask2 = np.where((mask == cv2.GC_FGD) | (mask == cv2.GC_PR_FGD), 255, 0).astype("uint8")
    mask2 = cv2.morphologyEx(mask2, cv2.MORPH_CLOSE, np.ones((15, 15), np.uint8))
    mask2 = cv2.morphologyEx(mask2, cv2.MORPH_OPEN, np.ones((9, 9), np.uint8))

    contours, _ = cv2.findContours(mask2, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return None
    largest = max(contours, key=cv2.contourArea)
    if cv2.contourArea(largest) < (sh * sw) * 0.15:
        return None

    rect_min = cv2.minAreaRect(largest)
    box = cv2.boxPoints(rect_min)
    corners = box.astype("float32") / scale
    return corners


def enhance(image):
    lab = cv2.cvtColor(image, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    l = clahe.apply(l)
    lab = cv2.merge((l, a, b))
    image = cv2.cvtColor(lab, cv2.COLOR_LAB2BGR)

    blur = cv2.GaussianBlur(image, (0, 0), sigmaX=3)
    sharpened = cv2.addWeighted(image, 1.5, blur, -0.5, 0)

    hsv = cv2.cvtColor(sharpened, cv2.COLOR_BGR2HSV).astype("float32")
    h, s, v = cv2.split(hsv)
    s = np.clip(s * 1.15, 0, 255)
    hsv = cv2.merge((h, s, v)).astype("uint8")
    result = cv2.cvtColor(hsv, cv2.COLOR_HSV2BGR)
    return result


def process(in_path, out_path, margin_pct=0.01):
    image = cv2.imread(str(in_path))
    if image is None:
        raise RuntimeError(f"nao foi possivel ler {in_path}")

    contour = find_document_contour(image)
    if contour is not None:
        h, w = image.shape[:2]
        contour[:, 0] = np.clip(contour[:, 0], 0, w - 1)
        contour[:, 1] = np.clip(contour[:, 1], 0, h - 1)
        warped = four_point_transform(image, contour)
        # se o recorte ficou minusculo ou quase igual a imagem toda, mantem original
        area_ratio = (warped.shape[0] * warped.shape[1]) / (image.shape[0] * image.shape[1])
        if area_ratio < 0.15:
            warped = image
    else:
        warped = image

    result = enhance(warped)
    cv2.imwrite(str(out_path), result, [cv2.IMWRITE_JPEG_QUALITY, 95])


if __name__ == "__main__":
    process(sys.argv[1], sys.argv[2])
