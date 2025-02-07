import argparse
import os
import win32print
import win32ui
from PIL import Image, UnidentifiedImageError, ImageWin


WIDTH = 150
HEIGHT = 100
INCH = 25.4
DPI = 300

SPACING = 1
STEPS = 10

LEFT = 7 * SPACING
RIGHT = 7 * SPACING
TOP = 4 * SPACING
BOTTOM = 4 * SPACING

WIDTH_PX = int(WIDTH / INCH * DPI)
HEIGHT_PX = int(HEIGHT / INCH * DPI)

MAX_BORDER = 30

X0 = int(LEFT / INCH * DPI)
X1 = int(RIGHT / INCH * DPI)
Y0 = int(TOP / INCH * DPI)
Y1 = int(BOTTOM / INCH * DPI)
W = WIDTH_PX - X1 - X0
H = HEIGHT_PX - Y1 - Y0

PRINTER_NAME = "SELPHY"
PHYSICAL_WIDTH = 110
PHYSICAL_HEIGHT = 111

def process_image(input_filename, border_pixels):
    background = Image.new(mode="RGB", size=(WIDTH_PX, HEIGHT_PX), color=(255, 255, 255))
    try:
        with Image.open(input_filename, "r") as im:
            width, height = im.width, im.height
            if im.height > im.width:
                im = im.rotate(angle=270, expand=True)
                width, height = height, width
            scale_x = (W - 2 * border_pixels) / width
            scale_y = (H - 2 * border_pixels) / height
            scale = min(scale_x, scale_y)
            w = int(scale * width)
            h = int(scale * height)
            offset_x = int((WIDTH_PX - w) / 2)
            offset_y = int((HEIGHT_PX - h) / 2)
            im = im.resize(size=(w, h), resample=Image.Resampling.NEAREST)
            background.paste(im, box=(offset_x, offset_y))
            return background
    except UnidentifiedImageError:
        print(f"Unsupported image file \"{input_filename}\"")


def get_printer_name():
    for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
        if PRINTER_NAME in p[2].upper():
            return p[2]
    return None


def print_image(im, print_job_name):
    printer_name = get_printer_name()
    if printer_name is None:
        print("Unable to find Canon SELPHY amongst local printers")
        exit(4)

    # bmp = im.convert("RGB").tobytes("raw", "BGR")
    # printer = win32print.OpenPrinter(printer_name)


    context = win32ui.CreateDC()
    context.CreatePrinterDC(printer_name)
    printer_width = context.GetDeviceCaps(PHYSICAL_WIDTH)
    printer_height = context.GetDeviceCaps(PHYSICAL_HEIGHT)

    context.StartDoc(print_job_name)
    context.StartPage()

    if im.width < im.height:
        im = im.rotate(angle=90, expand=True)

    dib = ImageWin.Dib(im)
    output_handle = context.GetHandleOutput()
    dib.draw(output_handle, (0, 0, printer_width, printer_height))

    context.EndPage()
    context.EndDoc()
    context.DeleteDC()
    #
    #
    # win32print.StartDocPrinter(printer, 1, (print_job_name, None, "RAW"))
    # win32print.StartPagePrinter(printer)
    # win32print.WritePrinter(printer, bmp)
    # win32print.EndPagePrinter(printer)
    # win32print.EndDocPrinter(printer)
    # win32print.ClosePrinter(printer)


def output_image(im, output_filename, print_job_name="image"):
    if len(output_filename) > 0:
        im.save(output_filename, dpi=(DPI, DPI))
    else:
        print_image(im, print_job_name)


def main():
    parser = argparse.ArgumentParser(
        description="Adjust image to print on Canon Selphy without cropping, print it or save to file")
    parser.add_argument("--input", "-i",
                        help="Input image or directory with multiple images",
                        type=str,
                        required=True)
    parser.add_argument("--border", "-b",
                        help="(optional) add border around the image (in mm)",
                        type=float,
                        default=0.0,
                        required=False)
    parser.add_argument("--output", "-o",
                        help="(optional) Output filename or directory for processed images."
                             "If not provided the image will be sent to the printer",
                        type=str,
                        default="",
                        required=False)
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Input path \"{args.input}\" does not exist.")
        exit(1)

    if len(args.output) > 0:
        output_parent = os.path.dirname(args.output)
        if not os.path.isdir(output_parent):
            print(f"Parent directory \"{output_parent}\" must exist.")
            exit(2)

    if args.border > MAX_BORDER:
        print(f"Border cannot exceed {MAX_BORDER} mm.")
        exit(3)

    border_px = int(args.border / INCH * DPI)

    if os.path.isfile(args.input):
        result = process_image(args.input, border_px)
        output_image(result, args.output, args.input)
    else:
        root, subdirectory, files = next(os.walk(args.input))
        for filename in files:
            fn = os.path.join(root, filename)
            basename, ext = os.path.splitext(filename)
            fn_output = os.path.join(args.output, basename + "-print" + ext)
            result = process_image(fn, border_px)
            output_image(result, fn_output, filename)


if __name__ == "__main__":
    main()
