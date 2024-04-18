import numpy as np
import keras as K
from keras.models import load_model
import cv2
import matplotlib.pyplot as plt


def f1(y_true, y_pred):
    y_true = K.cast(y_true, 'float32')
    y_pred = K.cast(y_pred, 'float32')

    def recall(y_true, y_pred):
        true_positives = K.sum(K.round(K.clip(y_true * y_pred, 0, 1)))
        possible_positives = K.sum(K.round(K.clip(y_true, 0, 1)))
        recall = true_positives / (possible_positives + K.epsilon())
        return recall

    def precision(y_true, y_pred):
        true_positives = K.sum(K.round(K.clip(y_true * y_pred, 0, 1)))
        predicted_positives = K.sum(K.round(K.clip(y_pred, 0, 1)))
        precision = true_positives / (predicted_positives + K.epsilon())
        return precision

    precision = precision(y_true, y_pred)
    recall = recall(y_true, y_pred)
    return 2*((precision*recall)/(precision+recall+K.epsilon()))


model = load_model(r'../models/best_cnn.hdf5', custom_objects={"f1": f1})


def segment_digits(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, binary = cv2.threshold(
        gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    contours, _ = cv2.findContours(
        binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    segmented_regions = []

    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        digit_roi = binary[max(0, y - 10):min(y + h + 10, image.shape[0]),
                           max(0, x - 10):min(x + w + 10, image.shape[1])]
        if w > 10 and h > 10:
            segmented_regions.append((x, y, w, h))
            cv2.rectangle(image, (x, y), (x+w, y+h), (0, 255, 0), 2)

    return segmented_regions, image


def crop_and_pad_images(image, segmented_regions):
    cropped_images = []

    for x, y, w, h in segmented_regions:
        digit_roi = image[max(0, y - 10):min(y + h + 10, image.shape[0]),
                          max(0, x - 10):min(x + w + 10, image.shape[1])]
        digit_roi_gray = cv2.cvtColor(digit_roi, cv2.COLOR_BGR2GRAY)
        resized_digit_roi = cv2.resize(digit_roi_gray, (28, 28))
        padded_digit_roi = np.zeros((28, 28), dtype=np.uint8)
        padded_digit_roi[:resized_digit_roi.shape[0],
                         :resized_digit_roi.shape[1]] = resized_digit_roi

        cropped_images.append((padded_digit_roi, x))

    return cropped_images


def predict_and_display(model, image):
    prediction = model.predict(np.array([image]))
    return np.argmax(prediction)


def get_prediction(image_path):
    input_image = cv2.imread(image_path)
    segmented_regions, image_with_boxes = segment_digits(input_image.copy())
    cropped_images = crop_and_pad_images(image_with_boxes, segmented_regions)

    number = []

    for cropped_image, x in cropped_images:
        contrast = 2.25
        brightness = -100
        cropped_image = cv2.convertScaleAbs(
            cropped_image, alpha=contrast, beta=brightness)
        cropped_image = cv2.bitwise_not(cropped_image)
        cropped_image = cropped_image.astype(np.float64) / 255.0
        cropped_image = np.expand_dims(cropped_image, axis=-1)
        # plt.imshow(cropped_image, cmap='gray')
        digit = predict_and_display(model, cropped_image)
        number.append((digit, x))
        # plt.show()

    final_number = int("".join([str(t[0])
                       for t in sorted(number, key=lambda x: x[1])]))
    return final_number


if __name__ == "__main__":
    print(get_prediction(image_path=r'../images/image3.jpeg'))
