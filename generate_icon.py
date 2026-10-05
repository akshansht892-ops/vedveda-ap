"""
generate_icon.py — Creates the app icon (static/icon.png + icon.ico)
Run once before building: python generate_icon.py
"""
import os
from PIL import Image, ImageDraw, ImageFont

SIZE = 256

def make_icon():
    img  = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Rounded-rect background (sky-blue gradient approximated as solid)
    r = 56
    draw.rounded_rectangle([0, 0, SIZE - 1, SIZE - 1], radius=r, fill="#0ea5e9")

    # Inner lighter highlight
    draw.rounded_rectangle([10, 10, SIZE - 11, SIZE // 2], radius=36, fill="#38bdf8")

    # Bar-chart icon (three bars)
    bar_color = "white"
    bars = [(52, 160, 92, 220), (108, 120, 148, 220), (164, 80, 204, 220)]
    for bx0, by0, bx1, by1 in bars:
        draw.rounded_rectangle([bx0, by0, bx1, by1], radius=8, fill=bar_color)

    # "V" letter
    try:
        font = ImageFont.truetype("arial.ttf", 72)
    except Exception:
        font = ImageFont.load_default()
    draw.text((SIZE // 2, 34), "V", font=font, fill="white", anchor="mt")

    os.makedirs("static", exist_ok=True)
    img.save("static/icon.png", "PNG")

    # Also save .ico for the exe
    ico_sizes = [(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)]
    img.save("static/icon.ico", "ICO", sizes=ico_sizes)
    print("Icons generated: static/icon.png  static/icon.ico")

if __name__ == "__main__":
    make_icon()
