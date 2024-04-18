import cv2
from pyzbar.pyzbar import decode


def barcode_reader(image_path):
    img = cv2.imread(image_path)
    detectedBarcodes = decode(img)

    if not detectedBarcodes:
        print("Barcode Not Detected or your barcode is blank/corrupted!")
    else:
        for barcode in detectedBarcodes:
            (x, y, w, h) = barcode.rect
            cv2.rectangle(img, (x-10, y-10),
                          (x + w+10, y + h+10),
                          (255, 0, 0), 2)
            if barcode.data != "":
                decoded_string = barcode.data.decode("utf-8")
                return decoded_string


if __name__ == "__main__":
    image_path = "images/barcode.png"
    info = barcode_reader(image_path)
    print(info)
