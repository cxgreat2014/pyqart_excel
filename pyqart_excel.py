import sys
import time
import argparse

import keyboard
from pyqart import QrPainter, QrHalftonePrinter
from pyqart.qr.data import QrData

QR_VERSION = 3
POINT_PIXEL = 1


def tab():
    keyboard.press_and_release('tab')


def down():
    keyboard.press_and_release('down')


def next_box(x, y, img_x):
    if (y % 2 == 0 and x != img_x - 1) or (y % 2 == 1 and x != 0):
        if y % 2 == 0:
            keyboard.press_and_release('right')
        else:
            keyboard.press_and_release('left')
    time.sleep(0.005)


def set_color(r, g, b):
    keyboard.press_and_release('alt,h,h,m')
    time.sleep(0.03)
    keyboard.press_and_release('right,tab,tab,tab,tab')
    keyboard.write(str(r))
    tab()
    keyboard.write(str(g))
    tab()
    keyboard.write(str(b))
    tab()
    keyboard.press_and_release('enter')
    time.sleep(0.2)


def end_to_first(img_x, img_y):
    if img_y % 2 == 1:
        for _ in range(img_x - 1):
            keyboard.press_and_release('left')
    for _ in range(img_y):
        keyboard.press_and_release('up')


def paint_to_excle(img_data):
    img_x, img_y = img_data.size
    set_color(0, 0, 0)
    for y in range(img_y):
        if y % 2 == 0:
            x_range = range(img_x)
        else:
            x_range = range(img_x - 1, -1, -1)
        for x in x_range:
            pix_color = img_data.getpixel((x, y))
            if pix_color == (0, 0, 0):
                keyboard.press_and_release('f4')
            next_box(x, y, img_x)
        down()
    end_to_first(img_x, img_y)
    set_color(255, 255, 255)
    for y in range(img_y):
        if y % 2 == 0:
            x_range = range(img_x)
        else:
            x_range = range(img_x - 1, -1, -1)
        for x in x_range:
            pix_color = img_data.getpixel((x, y))
            if pix_color == (255, 255, 255):
                keyboard.press_and_release('f4')
            next_box(x, y, img_x)
        down()
    end_to_first(img_x, img_y)
    old_color = (255, 255, 255)
    for y in range(img_y):
        if y % 2 == 0:
            x_range = range(img_x)
        else:
            x_range = range(img_x - 1, -1, -1)
        for x in x_range:
            pix_color = img_data.getpixel((x, y))
            if pix_color in ((0, 0, 0), (255, 255, 255)):
                pass
            elif pix_color == old_color:
                keyboard.press_and_release('f4')
            else:
                set_color(*pix_color)
                old_color = pix_color
            next_box(x, y, img_x)
        down()


def main():
    global QR_VERSION
    parser = argparse.ArgumentParser(
        prog="pyqart_excel",
        description="A program of generate QArt Codes and write to excel.",
        epilog="Writen by CxGreat2014. (https://github.com/cxgreat2014/pyqart_excle)",
    )
    parser.add_argument(
        'url', type=str,
        help="url will be encode, like http://example.com/",
    )
    parser.add_argument(
        'img', type=str,
        help="target image the QrCode will look like",
    )
    parser.add_argument(
        '-d', '--delay', type=float, default=5,
        help="delay for you switch to excel window",
    )

    argv = sys.argv[1:]

    args = parser.parse_args(argv)
    print(f'Please switch to excel window in {args.delay} seconds')
    time.sleep(args.delay)

    start = time.time()
    QR_VERSION = QrData(args.url).version_used_available[0]
    qr_target = QrPainter(args.url, QR_VERSION)
    qr_img = QrHalftonePrinter.print(qr_target, img=args.img, point_width=POINT_PIXEL)
    paint_to_excle(qr_img)
    end = time.time()

    print("Used time:", end - start, 'second.')


if __name__ == '__main__':
    main()
